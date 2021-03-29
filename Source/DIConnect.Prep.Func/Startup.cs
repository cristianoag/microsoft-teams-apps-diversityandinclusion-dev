﻿// <copyright file="Startup.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

[assembly: Microsoft.Azure.Functions.Extensions.DependencyInjection.FunctionsStartup(
    typeof(Microsoft.Teams.Apps.DIConnect.Prep.Func.Startup))]

namespace Microsoft.Teams.Apps.DIConnect.Prep.Func
{
    extern alias BetaLib;

    using System;
    using System.Globalization;
    using Microsoft.Azure.Functions.Extensions.DependencyInjection;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.EmployeeResourceGroup;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.ExportData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.UserPairupMapping;
    using Microsoft.Teams.Apps.DIConnect.Common.Services;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.AdaptiveCard;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.CommonBot;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues.DataQueue;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues.ExportQueue;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues.SendQueue;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MessageQueues.UserPairupQueue;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.Teams;
    using Microsoft.Teams.Apps.DIConnect.Prep.Func.Export.Activities;
    using Microsoft.Teams.Apps.DIConnect.Prep.Func.Export.Orchestrator;
    using Microsoft.Teams.Apps.DIConnect.Prep.Func.Export.Streams;
    using Microsoft.Teams.Apps.DIConnect.Prep.Func.PreparePairUpMatchesToSend.Activities;
    using Microsoft.Teams.Apps.DIConnect.Prep.Func.PreparingToSend;

    using Beta = BetaLib::Microsoft.Graph;

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
            builder.Services.AddOptions<DataQueueMessageOptions>()
                .Configure<IConfiguration>((dataQueueMessageOptions, configuration) =>
                {
                    dataQueueMessageOptions.MessageDelayInSeconds =
                        configuration.GetValue<double>("DataQueueMessageDelayInSeconds", 5);
                });

            builder.Services.AddOptions<TeamsConversationOptions>()
                .Configure<IConfiguration>((options, configuration) =>
                {
                    options.ProactivelyInstallUserApp =
                        configuration.GetValue<bool>("ProactivelyInstallUserApp", true);

                    options.MaxAttemptsToCreateConversation =
                        configuration.GetValue<int>("MaxAttemptsToCreateConversation", 2);
                });

            builder.Services.AddOptions<ConfidentialClientApplicationOptions>().
                Configure<IConfiguration>((confidentialClientApplicationOptions, configuration) =>
                {
                 confidentialClientApplicationOptions.ClientId = configuration.GetValue<string>("MicrosoftAppId");
                 confidentialClientApplicationOptions.ClientSecret = configuration.GetValue<string>("MicrosoftAppPassword");
                 confidentialClientApplicationOptions.TenantId = configuration.GetValue<string>("TenantId");
                });

            builder.Services.AddLocalization();

            // Set current culture.
            var culture = Environment.GetEnvironmentVariable("i18n:DefaultCulture");
            CultureInfo.DefaultThreadCurrentCulture = new CultureInfo(culture);
            CultureInfo.DefaultThreadCurrentUICulture = new CultureInfo(culture);

            // Add orchestration.
            builder.Services.AddTransient<ExportOrchestration>();

            // Add activities.
            builder.Services.AddTransient<UpdateExportDataActivity>();
            builder.Services.AddTransient<GetMetadataActivity>();
            builder.Services.AddTransient<UploadActivity>();
            builder.Services.AddTransient<SendFileCardActivity>();
            builder.Services.AddTransient<HandleExportFailureActivity>();

            // Add bot services.
            builder.Services.AddSingleton<CommonMicrosoftAppCredentials>();
            builder.Services.AddSingleton<ICredentialProvider, CommonBotCredentialProvider>();
            builder.Services.AddSingleton<BotFrameworkHttpAdapter>();

            // Add repositories.
            builder.Services.AddSingleton<NotificationDataRepository>();
            builder.Services.AddSingleton<SendingNotificationDataRepository>();
            builder.Services.AddSingleton<SentNotificationDataRepository>();
            builder.Services.AddSingleton<UserDataRepository>();
            builder.Services.AddSingleton<TeamDataRepository>();
            builder.Services.AddSingleton<ExportDataRepository>();
            builder.Services.AddSingleton<AppConfigRepository>();
            builder.Services.AddSingleton<EmployeeResourceGroupRepository>();
            builder.Services.AddSingleton<TeamUserPairUpMappingRepository>();

            // Add service bus message queues.
            builder.Services.AddSingleton<SendQueue>();
            builder.Services.AddSingleton<DataQueue>();
            builder.Services.AddSingleton<ExportQueue>();
            builder.Services.AddSingleton<UserPairUpQueue>();

            // Add miscellaneous dependencies.
            builder.Services.AddTransient<TableRowKeyGenerator>();
            builder.Services.AddTransient<AdaptiveCardCreator>();
            builder.Services.AddSingleton<SendPairUpMatchesActivity>();
            builder.Services.AddSingleton<IAppSettingsService, AppSettingsService>();

            // Add Teams services.
            builder.Services.AddTransient<ITeamMembersService, TeamMembersService>();
            builder.Services.AddTransient<IConversationService, ConversationService>();

            // Add graph services.
            this.AddGraphServices(builder);

            builder.Services.AddTransient<IDataStreamFacade, DataStreamFacade>();
        }

        /// <summary>
        /// Adds Graph Services and related dependencies.
        /// </summary>
        /// <param name="builder">Builder.</param>
        private void AddGraphServices(IFunctionsHostBuilder builder)
        {
            // Options
            builder.Services.AddOptions<ConfidentialClientApplicationOptions>().
                Configure<IConfiguration>((confidentialClientApplicationOptions, configuration) =>
                {
                    confidentialClientApplicationOptions.ClientId = configuration.GetValue<string>("MicrosoftAppId");
                    confidentialClientApplicationOptions.ClientSecret = configuration.GetValue<string>("MicrosoftAppPassword");
                    confidentialClientApplicationOptions.TenantId = configuration.GetValue<string>("TenantId");
                });

            // Graph Token Services
            builder.Services.AddSingleton<IConfidentialClientApplication>(provider =>
            {
                var options = provider.GetRequiredService<IOptions<ConfidentialClientApplicationOptions>>();
                return ConfidentialClientApplicationBuilder
                    .Create(options.Value.ClientId)
                    .WithClientSecret(options.Value.ClientSecret)
                    .WithAuthority(new Uri($"https://login.microsoftonline.com/{options.Value.TenantId}"))
                    .Build();
            });

            builder.Services.AddSingleton<IAuthenticationProvider, MsalAuthenticationProvider>();

            // Add Graph Clients.
            builder.Services.AddSingleton<IGraphServiceClient>(
                serviceProvider =>
                new GraphServiceClient(serviceProvider.GetRequiredService<IAuthenticationProvider>()));
            builder.Services.AddSingleton<Beta.IGraphServiceClient>(
                sp => new Beta.GraphServiceClient(sp.GetRequiredService<IAuthenticationProvider>()));

            // Add Service Factory
            builder.Services.AddSingleton<IGraphServiceFactory, GraphServiceFactory>();

            // Add Graph Services
            builder.Services.AddScoped<IUsersService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetUsersService());
            builder.Services.AddScoped<IGroupMembersService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetGroupMembersService());
            builder.Services.AddScoped<IAppManagerService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetAppManagerService());
            builder.Services.AddScoped<IChatsService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetChatsService());
        }
    }
}