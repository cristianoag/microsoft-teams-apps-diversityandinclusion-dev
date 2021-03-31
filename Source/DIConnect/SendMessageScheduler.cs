﻿// <copyright file="SendMessageScheduler.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues.DataQueue;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues.PrepareToSendQueue;

    /// <summary>
    /// Register background timed service to send scheduled messages.
    /// </summary>
    public class SendMessageScheduler : IHostedService, IDisposable
    {
        private readonly ILogger<SendMessageScheduler> smslogger;
        private readonly NotificationDataRepository notificationDataRepository;
        private readonly SentNotificationDataRepository sentNotificationDataRepository;
        private readonly PrepareToSendQueue prepareToSendQueue;
        private readonly DataQueue dataQueue;
        private readonly double forceCompleteMessageDelayInSeconds;
        private Timer smstimer;

        /// <summary>
        /// Initializes a new instance of the <see cref="SendMessageScheduler"/> class.
        /// </summary>
        /// <param name="logger">system logger</param>
        /// <param name="factory">factory</param>
        public SendMessageScheduler(ILogger<SendMessageScheduler> logger, IServiceScopeFactory factory)
        {
            this.smslogger = logger;
            this.notificationDataRepository = factory.CreateScope().ServiceProvider.GetRequiredService<NotificationDataRepository>();
            this.sentNotificationDataRepository = factory.CreateScope().ServiceProvider.GetRequiredService<SentNotificationDataRepository>();
            this.prepareToSendQueue = factory.CreateScope().ServiceProvider.GetRequiredService<PrepareToSendQueue>();
            this.dataQueue = factory.CreateScope().ServiceProvider.GetRequiredService<DataQueue>();
            this.forceCompleteMessageDelayInSeconds = 86400;
        }

        /// <summary>
        /// Start the service <see cref="StartAsync"/>.
        /// </summary>
        /// <param name="stoppingToken">system logger</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public Task StartAsync(CancellationToken stoppingToken)
        {
            this.smslogger.LogInformation("[DIConnect Scheduler] Hosted Service is running.");

            this.smstimer = new Timer(this.DoWork, null, TimeSpan.Zero, TimeSpan.FromMinutes(1));

            return Task.CompletedTask;
        }

        /// <summary>
        /// Stops the service <see cref="StopAsync"/>
        /// </summary>
        /// <param name="stoppingToken">This is the cancellation token</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public Task StopAsync(CancellationToken stoppingToken)
        {
            this.smslogger.LogInformation("[DIConnect Scheduler] Hosted Service is stopping.");

            this.smstimer?.Change(Timeout.Infinite, 0);

            return Task.CompletedTask;
        }

        /// <summary>
        /// Disposes the service
        /// </summary>
        public void Dispose()
        {
            this.smstimer?.Dispose();
        }

        private async void DoWork(object state)
        {
            DateTime now = DateTime.Now;

            this.smslogger.LogInformation(
                "[DIConnect Scheduler] is processing unsent scheduled messages before {Now}.", now);

            try
            {
                var notificationEntities = await this.notificationDataRepository.GetAllPendingScheduledNotificationsAsync();
                foreach (var notificationEntity in notificationEntities)
                {
                    this.smslogger.LogInformation("[DIConnect Scheduler] sending notification: {0}", notificationEntity.Title);
                    this.SendNotification(notificationEntity.Id);
                }
            }
            catch (Exception ex)
            {
                this.smslogger.LogError(ex.ToString());
            }
        }

        private async void SendNotification(string id)
        {
            var draftNotificationDataEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.DraftNotificationsPartition,
                id);

            if (draftNotificationDataEntity == null)
            {
                throw new Exception($"Draft notification, Id: {id}, could not be found.");
            }
            else
            {
                var newSentNotificationId =
                    await this.notificationDataRepository.MoveDraftToSentPartitionAsync(draftNotificationDataEntity);

                // Ensure the data table needed by the Azure Functions to send the notifications exist in Azure storage.
                await this.sentNotificationDataRepository.EnsureSentNotificationDataTableExistsAsync();

                var prepareToSendQueueMessageContent = new PrepareToSendQueueMessageContent
                {
                    NotificationId = newSentNotificationId,
                };

                await this.prepareToSendQueue.SendAsync(prepareToSendQueueMessageContent);

                // Send a "force complete" message to the data queue with a delay to ensure that
                // the notification will be marked as complete no matter the counts
                var forceCompleteDataQueueMessageContent = new DataQueueMessageContent
                {
                    NotificationId = newSentNotificationId,
                    ForceMessageComplete = true,
                };
                await this.dataQueue.SendDelayedAsync(
                    forceCompleteDataQueueMessageContent,
                    this.forceCompleteMessageDelayInSeconds);
            }
        }
    }
}
