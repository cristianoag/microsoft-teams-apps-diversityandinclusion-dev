﻿// <copyright file="IConversationService.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Common.Services.Teams
{
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Conversation service interface.
    /// </summary>
    public interface IConversationService
    {
        /// <summary>
        /// Creates a new conversation with the user.
        /// </summary>
        /// <param name="teamsUserId">User's Id in Teams.</param>
        /// <param name="tenantId">Tenant Id.</param>
        /// <param name="serviceUrl">Service url.</param>
        /// <param name="maxAttempts">Maximum attempts.</param>
        /// <param name="log">Logger.</param>
        /// <returns><see cref="CreateConversationResponse"/>.</returns>
        public Task<CreateConversationResponse> CreateConversationAsync(
            string teamsUserId,
            string tenantId,
            string serviceUrl,
            int maxAttempts,
            ILogger log);
    }
}