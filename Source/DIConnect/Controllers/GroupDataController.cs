﻿// <copyright file="GroupDataController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Controllers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.DIConnect.Authentication;
    using Microsoft.Teams.Apps.DIConnect.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.DIConnect.Models;

    /// <summary>
    /// Controller for getting groups.
    /// </summary>
    [Route("api/groupData")]
    [Authorize(PolicyNames.MustBeAdminTeamMemberPolicy)]
    public class GroupDataController : Controller
    {
        private readonly NotificationDataRepository notificationDataRepository;
        private readonly IGroupsService groupsService;

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupDataController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository instance.</param>
        /// <param name="groupsService">Microsoft Graph service instance.</param>
        public GroupDataController(
            NotificationDataRepository notificationDataRepository,
            IGroupsService groupsService)
        {
            this.notificationDataRepository = notificationDataRepository;
            this.groupsService = groupsService;
        }

        /// <summary>
        /// check if user has access.
        /// </summary>
        /// <returns>indicating user access to group.</returns>
        [HttpGet("verifyaccess")]
        [Authorize(PolicyNames.MSGraphGroupDataPolicy)]
        public bool VerifyAccess()
        {
            return true;
        }

        /// <summary>
        /// Action method to get groups.
        /// </summary>
        /// <param name="query">user input.</param>
        /// <returns>list of group data.</returns>
        [HttpGet("search/{query}")]
        [Authorize(PolicyNames.MSGraphGroupDataPolicy)]
        public async Task<IEnumerable<GroupData>> SearchAsync(string query)
        {
            int minQueryLength = 3;
            if (string.IsNullOrEmpty(query) || query.Length < minQueryLength)
            {
                return default;
            }

            var groups = await this.groupsService.SearchAsync(query);
            return groups.Select(group => new GroupData()
            {
                Id = group.Id,
                Name = group.DisplayName,
                Mail = group.Mail,
            });
        }

        /// <summary>
        /// Get Group Data by Id.
        /// </summary>
        /// <param name="id">Draft notification Id.</param>
        /// <returns>List of Group Names.</returns>
        [HttpGet("{id}")]
        public async Task<ActionResult<IEnumerable<GroupData>>> GetGroupsAsync(string id)
        {
            var notificationEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.DraftNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var groups = await this.groupsService.GetByIdsAsync(notificationEntity.Groups)
                .Select(group => new GroupData()
                {
                    Id = group.Id,
                    Name = group.DisplayName,
                    Mail = group.Mail,
                }).ToListAsync();

            return this.Ok(groups);
        }
    }
}