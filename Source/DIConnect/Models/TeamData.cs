﻿// <copyright file="TeamData.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Models
{
    /// <summary>
    /// Teams data model class.
    /// </summary>
    public class TeamData
    {
        /// <summary>
        /// Gets or sets team Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets description.
        /// </summary>
        public string Description { get; set; }
    }
}