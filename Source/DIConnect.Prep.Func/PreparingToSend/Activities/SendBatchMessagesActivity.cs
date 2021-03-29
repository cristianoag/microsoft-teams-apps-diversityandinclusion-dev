﻿// <copyright file="SendBatchMessagesActivity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.DurableTask;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues.SendQueue;

    /// <summary>
    /// Sends batch messages to Send Queue.
    /// </summary>
    public class SendBatchMessagesActivity
    {
        private readonly SendQueue sendQueue;

        /// <summary>
        /// Initializes a new instance of the <see cref="SendBatchMessagesActivity"/> class.
        /// </summary>
        /// <param name="sendQueue">Send queue service.</param>
        public SendBatchMessagesActivity(
            SendQueue sendQueue)
        {
            this.sendQueue = sendQueue ?? throw new ArgumentNullException(nameof(sendQueue));
        }

        /// <summary>
        /// Sends batch messages to Send Queue.
        /// </summary>
        /// <param name="input">Input.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName(FunctionNames.SendBatchMessagesActivity)]
        public async Task RunAsync(
            [ActivityTrigger](NotificationDataEntity notification, List<SentNotificationDataEntity> batch) input)
        {
            var messageBatch = input.batch.Select(
                recipient =>
                {
                    return new SendQueueMessageContent()
                    {
                        NotificationId = input.notification.Id,
                        RecipientData = this.ConvertToRecipientData(recipient),
                    };
                });

            await this.sendQueue.SendAsync(messageBatch);
        }

        /// <summary>
        /// Converts sent notification data entity to Recipient data.
        /// </summary>
        /// <returns>Recipient data.</returns>
        private RecipientData ConvertToRecipientData(SentNotificationDataEntity recipient)
        {
            if (recipient.RecipientType == SentNotificationDataEntity.UserRecipientType)
            {
                return new RecipientData
                {
                    RecipientType = RecipientDataType.User,
                    RecipientId = recipient.RecipientId,
                    UserData = new UserDataEntity
                    {
                        AadId = recipient.RecipientId,
                        UserId = recipient.UserId,
                        ConversationId = recipient.ConversationId,
                        ServiceUrl = recipient.ServiceUrl,
                        TenantId = recipient.TenantId,
                    },
                };
            }
            else if (recipient.RecipientType == SentNotificationDataEntity.TeamRecipientType)
            {
                return new RecipientData
                {
                    RecipientType = RecipientDataType.Team,
                    RecipientId = recipient.RecipientId,
                    TeamData = new TeamDataEntity
                    {
                        TeamId = recipient.RecipientId,
                        ServiceUrl = recipient.ServiceUrl,
                        TenantId = recipient.TenantId,
                    },
                };
            }

            throw new ArgumentException($"Invalid recipient type: {recipient.RecipientType}.");
        }
    }
}