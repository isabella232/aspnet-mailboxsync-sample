/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.Graph;
using System.Net.Http.Headers;
using System;

namespace MailboxSync.Helpers
{
    /// <summary>
    /// Designed to help with authentication
    /// Checks if the token exists
    /// Fetches the token silently or forces a log in if the silent way fails
    /// </summary>
    public class GraphSdkHelper
    {

        /// <summary>
        /// Get an authenticated Microsoft Graph Service client.
        /// </summary>
        /// <param name="userId">optional parameter to fetch token for a user by id </param>
        /// <returns></returns>
        public static GraphServiceClient GetAuthenticatedClient(string userId = "")
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async requestMessage =>
                    {
                        string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync(userId);

                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                    }));
            return graphClient;
        }

    }
}