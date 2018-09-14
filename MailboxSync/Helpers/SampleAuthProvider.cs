/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using MailboxSync.TokenStorage;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System;

namespace MailboxSync.Helpers
{
    public sealed class SampleAuthProvider : IAuthProvider
    {

        // Properties used to get and manage an access token.
        private string _redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
        private readonly string _appId = ConfigurationManager.AppSettings["ida:AppId"];
        private readonly string _appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
        private readonly string _nonAdminScopes = ConfigurationManager.AppSettings["ida:NonAdminScopes"];
        private readonly string _adminScopes = ConfigurationManager.AppSettings["ida:AdminScopes"];
        private TokenCache TokenCache { get; set; }

        private SampleAuthProvider() { }

        public static SampleAuthProvider Instance { get; } = new SampleAuthProvider();

        /// <summary>
        /// Gets an access token and its expiration date. First tries to get the token from the token cache.
        /// </summary>
        /// <param name="userId">optional parameter to get the access token using a user Id</param>
        /// <returns></returns>
        public async Task<string> GetUserAccessTokenAsync(string userId)
        {

            // user Id will be passed when trying to authenticate a notification
            var currentUserId = !string.IsNullOrEmpty(userId) ? userId : ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

            HttpContextBase context = HttpContext.Current.GetOwinContext().Environment["System.Web.HttpContextBase"] as HttpContextBase;
            TokenCache = new SessionTokenCache(
                currentUserId).GetMsalCacheInstance();

            if (!_redirectUri.EndsWith("/")) _redirectUri = _redirectUri + "/";
            string[] segments = context?.Request.Path.Split('/');
            ConfidentialClientApplication cca = new ConfidentialClientApplication(_appId, _redirectUri + segments?[1], new ClientCredential(_appSecret), TokenCache, null);
            bool? isAdmin = HttpContext.Current.Session["IsAdmin"] as bool?;

            string allScopes = _nonAdminScopes;
            if (isAdmin.GetValueOrDefault())
            {
                allScopes += " " + _adminScopes;
            }

            string[] scopes = allScopes.Split(' ');
            try
            {
                AuthenticationResult result = await cca.AcquireTokenSilentAsync(scopes, cca.Users.First());
                return result.AccessToken;
            }

            // Unable to retrieve the access token silently.
            catch (Exception)
            {
                HttpContext.Current.Request.GetOwinContext().Authentication.Challenge(
                    new AuthenticationProperties { RedirectUri = _redirectUri + segments?[1] },
                    OpenIdConnectAuthenticationDefaults.AuthenticationType);

                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.AuthenticationFailure.ToString(),
                        Message = "Caller needs to authenticate.",
                    });
            }
        }


    }
}
