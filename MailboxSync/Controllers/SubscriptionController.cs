using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Mvc;
using MailboxSync.Models.Subscription;
using MailboxSync.Helpers;
using Microsoft.Graph;

namespace MailboxSync.Controllers
{
    public class SubscriptionController : Controller
    {
        public static string clientId = ConfigurationManager.AppSettings["ida:AppId"];
        private static string appKey = ConfigurationManager.AppSettings["ida:AppSecret"];
        private static string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];

        // GET: Subscription
        public ActionResult Index()
        {
            return View();
        }

        [Authorize]
        public async Task<ActionResult> CreateSubscription()
        {
            string subscriptionsEndpoint = "https://graph.microsoft.com/v1.0/subscriptions/";

            // This sample subscribes to get notifications when the user receives an email.
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, subscriptionsEndpoint);

            Microsoft.Graph.Subscription subscription = new Microsoft.Graph.Subscription
            {
                Resource = "me/messages",
                ChangeType = "created",
                NotificationUrl = ConfigurationManager.AppSettings["ida:NotificationUrl"],
                ClientState = Guid.NewGuid().ToString(),
                ExpirationDateTime = DateTime.UtcNow + new TimeSpan(0, 0, 15, 0) // shorter duration useful for testing
            };

            string contentString = JsonConvert.SerializeObject(subscription,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            request.Content = new StringContent(contentString, System.Text.Encoding.UTF8, "application/json");

            string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            HttpResponseMessage response = await client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                ViewBag.AuthorizationRequest = null;
                string stringResult = await response.Content.ReadAsStringAsync();
                SubscriptionViewModel viewModel = new SubscriptionViewModel
                {
                    Subscription = JsonConvert.DeserializeObject<Microsoft.Graph.Subscription>(stringResult)
                };

                // This sample temporarily stores the current subscription ID, client state, user object ID, and tenant ID. 
                // This info is required so the NotificationController, which is not authenticated, can retrieve an access token from the cache and validate the subscription.
                // Production apps typically use some method of persistent storage.
                SubscriptionStore.SaveSubscriptionInfo(viewModel.Subscription.Id,
                    viewModel.Subscription.ClientState,
                    ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value,
                    ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value);

                // This sample just saves the current subscription ID to the session so we can delete it later.
                Session["SubscriptionId"] = viewModel.Subscription.Id;
                return View("Subscription", viewModel);
            }
            return View("Subscription", null);
        }



        // Delete the current webhooks subscription and sign out the user.
        [Authorize]
        public async Task<ActionResult> DeleteSubscription()
        {
            string subscriptionsEndpoint = "https://graph.microsoft.com/v1.0/subscriptions/";
            string subscriptionId = (string)Session["SubscriptionId"];

            // Build the request.
            HttpClient client = new HttpClient();

            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, subscriptionsEndpoint + subscriptionId);

            // try to get token silently
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            TokenCache userTokenCache = new MSALSessionCache(signedInUserID, this.HttpContext).GetMsalCacheInstance();
            ConfidentialClientApplication cca = new ConfidentialClientApplication(clientId, redirectUri, new ClientCredential(appKey), userTokenCache, null);
            if (cca.Users.Count() > 0)
            {
                string[] scopes = { "Mail.Read" };
                try
                {
                    AuthenticationResult result = await cca.AcquireTokenSilentAsync(scopes, cca.Users.First());
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);

                    // Send the `DELETE subscriptions/id` request.
                    HttpResponseMessage response = await client.SendAsync(request);

                    if (!response.IsSuccessStatusCode)
                    {
                        return RedirectToAction("Index", "Error", new { message = response.StatusCode, debug = response.Content.ReadAsStringAsync() });
                    }
                }
                catch (MsalUiRequiredException)
                {
                    try
                    {// when failing, manufacture the URL and assign it
                        string authReqUrl = await MailboxSync.Utils.OAuth2RequestManager.GenerateAuthorizationRequestUrl(scopes, cca, this.HttpContext, Url);
                        ViewBag.AuthorizationRequest = authReqUrl;
                    }
                    catch (Exception ee)
                    {

                    }
                }
            }
            else { }
            return RedirectToAction("SignOut", "Account");
        }


    }
}