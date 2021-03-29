﻿// <copyright file="TeamDataEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Common.Repositories.TeamData
{
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Teams data entity class.
    /// This entity holds the information about a team.
    /// </summary>
    public class TeamDataEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets the team id.
        /// </summary>
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets the name of the team.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the service url for the team.
        /// </summary>
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets tenant id for the team.
        /// </summary>
        public string TenantId { get; set; }
    }
}