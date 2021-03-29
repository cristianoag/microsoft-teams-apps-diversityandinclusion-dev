// <copyright file="DIConnectCleanUpFunction.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Data.Func
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using global::Azure.Storage.Blobs;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.ExportData;
    using Microsoft.Teams.Apps.DIConnect.Data.Func.Services.FileCardServices;

    /// <summary>
    /// Azure Function App triggered as per scheduled.
    /// Used for house keeping activities.
    /// </summary>
    public class DIConnectCleanUpFunction
    {
        private readonly int cleanUpFileOlderThanDays;
        private readonly ExportDataRepository exportDataRepository;
        private readonly IFileCardService fileCardService;
        private readonly BlobContainerClient blobContainerClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="DIConnectCleanUpFunction"/> class.
        /// </summary>
        /// <param name="exportDataRepository">The export data repository.</param>
        /// <param name="blobContainerClient">The Azure Blob storage container client.</param>
        /// <param name="fileCardService">The service to manage the file card.</param>
        /// <param name="cleanUpFileOptions">The options to clean up file.</param>
        public DIConnectCleanUpFunction(
            ExportDataRepository exportDataRepository,
            BlobContainerClient blobContainerClient,
            IFileCardService fileCardService,
            IOptions<CleanUpFileOptions> cleanUpFileOptions)
        {
            this.exportDataRepository = exportDataRepository;
            this.fileCardService = fileCardService;
            this.blobContainerClient = blobContainerClient;
            this.cleanUpFileOlderThanDays = int.Parse(cleanUpFileOptions.Value.CleanUpFile);
        }

        /// <summary>
        /// Azure Function App triggered as per scheduled.
        /// Used for house keeping activities.
        /// </summary>
        /// <param name="myTimer">The timer schedule.</param>
        /// <param name="log">The logger.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName("DIConnectCleanUpFunction")]
        public async Task Run([TimerTrigger("%CleanUpScheduleTriggerTime%")] TimerInfo myTimer, ILogger log)
        {
            var cleanUpDateTime = DateTime.UtcNow.AddDays(-this.cleanUpFileOlderThanDays);
            var exportDataEntities = await this.exportDataRepository.GetAllLessThanDateTimeAsync(cleanUpDateTime);
            exportDataEntities = exportDataEntities.Where(exportDataEntity => exportDataEntity.Status.Equals(ExportStatus.Completed.ToString()));
            await this.DeleteFilesAndCards(exportDataEntities);
            await this.exportDataRepository.BatchDeleteAsync(exportDataEntities);

            log.LogInformation($"DI Connect Clean Up function executed at: {DateTime.Now}");
        }

        /// <summary>
        /// This deletes the files in Azure Blob storage and file cards sent to users.
        /// </summary>
        /// <param name="exportDataEntities">the list of export data entity.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        private async Task DeleteFilesAndCards(IEnumerable<ExportDataEntity> exportDataEntities)
        {
            await this.blobContainerClient.CreateIfNotExistsAsync();

            var tasks = new List<Task>();
            foreach (var exportData in exportDataEntities)
            {
                tasks.Add(this.fileCardService.DeleteAsync(exportData.PartitionKey, exportData.FileConsentId));
                tasks.Add(this.blobContainerClient
                    .GetBlobClient(exportData.FileName)
                    .DeleteIfExistsAsync());
            }

            await Task.WhenAll(tasks);
        }
    }
}