using System;
using System.Configuration;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Mvc;
using MailboxSync.Helpers;
using MailboxSync.Models;
using MailboxSync.Services;
using Microsoft.Graph;

namespace MailboxSync.Controllers
{
    public class SubscriptionController : Controller
    {
        GraphServiceClient graphClient = GraphSdkHelper.GetAuthenticatedClient();

        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Create a subscription to the 'notificationUrl'
        /// </summary>
        [Authorize]
        public async Task<ActionResult> CreateSubscription()
        {
            try
            {

                var subscription = new Subscription
                {
                    Resource = "me/mailFolders('Inbox')/messages",
                    ChangeType = "created",
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

                    // This sample temporarily stores :
                    // - the current subscription ID, 
                    // - client state, 
                    // - user object ID, and 
                    // - tenant ID. 
                    // This info is required so the NotificationController, which is not authenticated,
                    // can retrieve an access token from the cache and validate the subscription.
                    // Production apps typically use some method of persistent storage
                    SubscriptionStore.SaveSubscriptionInfo(viewModel.Subscription.Id,
                        viewModel.Subscription.ClientState,
                        ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value,
                        ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value);

                    // This sample just saves the current subscription ID to the session so we can delete it later.
                    Session["SubscriptionId"] = viewModel.Subscription.Id;
                    return View("Subscription", viewModel);
                }
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "AuthenticationFailure")
                {
                    return new EmptyResult();
                }
                return RedirectToAction("Index", "Error", new
                {
                    message =
                    $"Error in {Request.RawUrl}: {se.Error.Code} {se.Error.Message}"
                });

            }

            return View("Subscription", null);
        }



        /// <summary>
        /// Delete the current webhooks subscription and sign out the user.
        /// </summary>
        /// <param name="subscriptionId">the subscription id you want to delete</param>
        [Authorize]
        public async Task<ActionResult> DeleteSubscription(string subscriptionId)
        {
            try
            {
                SubscriptionStore.DeleteSubscriptionInfo();
                GraphServiceClient graphClient = GraphSdkHelper.GetAuthenticatedClient();
                await graphClient.Subscriptions[subscriptionId].Request().DeleteAsync();
                return RedirectToAction("SignOut", "Account");
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "AuthenticationFailure")
                {
                    return new EmptyResult();
                }

                return RedirectToAction("Index", "Error", new
                {
                    message = $"Error in {Request.RawUrl}: {se.Error.Code} {se.Error.Message}"
                });

            }
        }


    }
}