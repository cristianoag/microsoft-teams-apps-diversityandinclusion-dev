﻿// <copyright file="HandleExportFailureActivity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Prep.Func.Export.Activities
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using global::Azure.Storage.Blobs;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.DurableTask;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.ExportData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.DIConnect.Common.Resources;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.CommonBot;
    using Microsoft.Teams.Apps.DIConnect.Prep.Func.PreparingToSend;
    using Polly;

    /// <summary>
    /// This class contains the "clean up" durable activity.
    /// If exceptions happen in the "export" operation, this method is called to clean up and send the error message.
    /// </summary>
    public class HandleExportFailureActivity
    {
        private readonly ExportDataRepository exportDataRepository;
        private readonly string storageConnectionString;
        private readonly BlobContainerClient blobContainerClient;
        private readonly UserDataRepository userDataRepository;
        private readonly string microsoftAppId;
        private readonly BotFrameworkHttpAdapter botAdapter;
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="HandleExportFailureActivity"/> class.
        /// </summary>
        /// <param name="exportDataRepository">the export data repository.</param>
        /// <param name="repositoryOptions">the repository options.</param>
        /// <param name="botOptions">the bot options.</param>
        /// <param name="botAdapter">the users service.</param>
        /// <param name="userDataRepository">the user data repository.</param>
        /// <param name="localizer">Localization service.</param>
        public HandleExportFailureActivity(
            ExportDataRepository exportDataRepository,
            IOptions<RepositoryOptions> repositoryOptions,
            IOptions<BotOptions> botOptions,
            BotFrameworkHttpAdapter botAdapter,
            UserDataRepository userDataRepository,
            IStringLocalizer<Strings> localizer)
        {
            this.exportDataRepository = exportDataRepository;
            this.storageConnectionString = repositoryOptions.Value.StorageAccountConnectionString;
            this.blobContainerClient = new BlobContainerClient(this.storageConnectionString, Common.Constants.BlobContainerName);
            this.botAdapter = botAdapter;
            this.microsoftAppId = botOptions.Value.MicrosoftAppId;
            this.userDataRepository = userDataRepository;
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
        }

        /// <summary>
        /// Run the activity.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <param name="exportDataEntity">export data entity.</param>
        /// <param name="log">Logging service.</param>
        /// <returns>instance of metadata.</returns>
        public async Task RunAsync(
            IDurableOrchestrationContext context,
            ExportDataEntity exportDataEntity,
            ILogger log)
        {
            await context.CallActivityWithRetryAsync<Task>(
                      nameof(HandleExportFailureActivity.HandleFailureActivityAsync),
                      FunctionSettings.DefaultRetryOptions,
                      exportDataEntity);
        }

        /// <summary>
        /// This method represents the "clean up" durable activity.
        /// If exceptions happen in the "export" operation,
        /// this method is called to do the clean up work, e.g. delete the files,records and etc.
        /// </summary>
        /// <param name="exportDataEntity">export data entity.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName(nameof(HandleFailureActivityAsync))]
        public async Task HandleFailureActivityAsync(
            [ActivityTrigger] ExportDataEntity exportDataEntity)
        {
            await this.DeleteFileAsync(exportDataEntity.FileName);
            await this.SendFailureMessageAsync(exportDataEntity.PartitionKey);
            await this.exportDataRepository.DeleteAsync(exportDataEntity);
        }

        private async Task DeleteFileAsync(string fileName)
        {
            if (fileName == null)
            {
                return;
            }

            await this.blobContainerClient.CreateIfNotExistsAsync();
            await this.blobContainerClient
                    .GetBlobClient(fileName)
                    .DeleteIfExistsAsync();
        }

        private async Task SendFailureMessageAsync(string userId)
        {
            var user = await this.userDataRepository.GetAsync(UserDataTableNames.UserDataPartition, userId);

            // Set the service URL in the trusted list to ensure the SDK includes the token in the request.
            MicrosoftAppCredentials.TrustServiceUrl(user.ServiceUrl);

            var conversationReference = new ConversationReference
            {
                ServiceUrl = user.ServiceUrl,
                Conversation = new ConversationAccount
                {
                    Id = user.ConversationId,
                },
            };
            string failureText = this.localizer.GetString("ExportFailureText");

            int maxNumberOfAttempts = 10;
            await this.botAdapter.ContinueConversationAsync(
               botAppId: this.microsoftAppId,
               reference: conversationReference,
               callback: async (turnContext, cancellationToken) =>
               {
                   // Retry it in addition to the original call.
                   var retryPolicy = Policy.Handle<Exception>().WaitAndRetryAsync(maxNumberOfAttempts, p => TimeSpan.FromSeconds(p));
                   await retryPolicy.ExecuteAsync(async () =>
                   {
                       var failureMessage = MessageFactory.Text(failureText);
                       failureMessage.TextFormat = "xml";
                       await turnContext.SendActivityAsync(failureMessage, cancellationToken);
                   });
               },
               cancellationToken: CancellationToken.None);
        }
    }
}