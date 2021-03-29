﻿// <copyright file="Startup.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

[assembly: Microsoft.Azure.Functions.Extensions.DependencyInjection.FunctionsStartup(
    typeof(Microsoft.Teams.Apps.DIConnect.Data.Func.Startup))]

namespace Microsoft.Teams.Apps.DIConnect.Data.Func
{
    using System;
    using System.Globalization;
    using global::Azure.Storage.Blobs;
    using Microsoft.Azure.Functions.Extensions.DependencyInjection;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.ExportData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.CommonBot;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues.DataQueue;
    using Microsoft.Teams.Apps.DIConnect.Data.Func.Services.FileCardServices;
    using Microsoft.Teams.Apps.DIConnect.Data.Func.Services.NotificationDataServices;

    /// <summary>
    /// Register services in DI container of the Azure functions system.
    /// </summary>
    public class Startup : FunctionsStartup
    {
        /// <inheritdoc/>
        public override void Configure(IFunctionsHostBuilder builder)
        {
            // Add all options set from configuration values.
            builder.Services.AddOptions<RepositoryOptions>()
                .Configure<IConfiguration>((repositoryOptions, configuration) =>
                {
                    repositoryOptions.StorageAccountConnectionString =
                        configuration.GetValue<string>("StorageAccountConnectionString");

                    // Defaulting this value to true because the main app should ensure all
                    // tables exist. It is here as a possible configuration setting in
                    // case it needs to be set differently.
                    repositoryOptions.EnsureTableExists =
                        !configuration.GetValue<bool>("IsItExpectedThatTableAlreadyExists", true);
                });
            builder.Services.AddOptions<MessageQueueOptions>()
                .Configure<IConfiguration>((messageQueueOptions, configuration) =>
                {
                    messageQueueOptions.ServiceBusConnection =
                        configuration.GetValue<string>("ServiceBusConnection");
                });
            builder.Services.AddOptions<BotOptions>()
               .Configure<IConfiguration>((botOptions, configuration) =>
               {
                   botOptions.MicrosoftAppId =
                       configuration.GetValue<string>("MicrosoftAppId");

                   botOptions.OnlyAdminsRegisterERG =
                       configuration.GetValue<string>("OnlyAdminsRegisterERG");

                   botOptions.MicrosoftAppPassword =
                       configuration.GetValue<string>("MicrosoftAppPassword");
               });
            builder.Services.AddOptions<CleanUpFileOptions>()
               .Configure<IConfiguration>((cleanUpFileOptions, configuration) =>
               {
                   cleanUpFileOptions.CleanUpFile =
                       configuration.GetValue<string>("CleanUpFile");
               });
            builder.Services.AddOptions<DataQueueMessageOptions>()
                .Configure<IConfiguration>((dataQueueMessageOptions, configuration) =>
                {
                    dataQueueMessageOptions.FirstTenMinutesRequeueMessageDelayInSeconds =
                        configuration.GetValue<double>("FirstTenMinutesRequeueMessageDelayInSeconds", 20);

                    dataQueueMessageOptions.RequeueMessageDelayInSeconds =
                        configuration.GetValue<double>("RequeueMessageDelayInSeconds", 120);
                });

            builder.Services.AddLocalization();

            // Set current culture.
            var culture = Environment.GetEnvironmentVariable("i18n:DefaultCulture");
            CultureInfo.DefaultThreadCurrentCulture = new CultureInfo(culture);
            CultureInfo.DefaultThreadCurrentUICulture = new CultureInfo(culture);

            // Add blob client.
            builder.Services.AddSingleton(sp => new BlobContainerClient(
                sp.GetService<IConfiguration>().GetValue<string>("StorageAccountConnectionString"),
                Common.Constants.BlobContainerName));

            // Add bot services.
            builder.Services.AddSingleton<CommonMicrosoftAppCredentials>();
            builder.Services.AddSingleton<ICredentialProvider, CommonBotCredentialProvider>();
            builder.Services.AddSingleton<BotFrameworkHttpAdapter>();

            // Add services.
            builder.Services.AddSingleton<IFileCardService, FileCardService>();

            // Add notification data services.
            builder.Services.AddTransient<AggregateSentNotificationDataService>();
            builder.Services.AddTransient<UpdateNotificationDataService>();

            // Add repositories.
            builder.Services.AddSingleton<NotificationDataRepository>();
            builder.Services.AddSingleton<SentNotificationDataRepository>();
            builder.Services.AddSingleton<UserDataRepository>();
            builder.Services.AddSingleton<ExportDataRepository>();

            // Add service bus message queues.
            builder.Services.AddSingleton<DataQueue>();
        }
    }
}