﻿// <copyright file="CreateConversationResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Common.Services.Teams
{
    /// <summary>
    /// The class for the create conversation response.
    /// </summary>
    public class CreateConversationResponse
    {
        /// <summary>
        /// Gets or sets the status code.
        /// </summary>
        public int StatusCode { get; set; }

        /// <summary>
        /// Gets or sets the result type.
        /// </summary>
        public Result Result { get; set; }

        /// <summary>
        /// Gets or sets the conversation ID.
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets the error message when trying to create the conversation.
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}