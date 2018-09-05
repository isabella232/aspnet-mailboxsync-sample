using System;
using System.Configuration;
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
            GraphServiceClient graphClient = GraphSdkHelper.GetAuthenticatedClient();

            var subscription = new Microsoft.Graph.Subscription
            {
                Resource = "me/messages",
                ChangeType = "created,updated",
                NotificationUrl = ConfigurationManager.AppSettings["ida:NotificationUrl"],
                ClientState = Guid.NewGuid().ToString(),
                ExpirationDateTime = DateTime.UtcNow + new TimeSpan(0, 0, 15, 0) // shorter duration useful for testing
            };

            var response = await graphClient.Subscriptions.Request().AddAsync(subscription);

            if (response != null)
            {
                SubscriptionViewModel viewModel = new SubscriptionViewModel
                {
                    Subscription = response
                };

                // This sample temporarily stores the current subscription ID, client state, user object ID, and tenant ID. 
                // This info is required so the NotificationController, which is not authenticated, can retrieve an access token from the cache and validate the subscription.
                // Production apps typically use some method of persistent storage.
                SubscriptionStore.SaveSubscriptionInfo(viewModel.Subscription.Id,
                    viewModel.Subscription.ClientState,
                    ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value,
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
            GraphServiceClient graphClient = GraphSdkHelper.GetAuthenticatedClient();

            string subscriptionId = (string)Session["SubscriptionId"];

            try
            {
                // TODO 
                // Delete subscription request
                //var response = await graphClient.Subscriptions[subscriptionId].Request().DeleteAsync();

            }
            catch (Exception e)
            {
                // ignored
            }

            return View("Subscription", null);
        }


    }
}