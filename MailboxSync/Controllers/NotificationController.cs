using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Threading;
using System.Web.Mvc;
using MailboxSync.Models.Subscription;
using Microsoft.Graph;
using MailboxSync.Helpers;
using System.Threading.Tasks;
using MailboxSync.Services;
using MailBoxSync.Models.Subscription;
using DateTimeOffset = System.DateTimeOffset;

namespace MailboxSync.Controllers
{
    public class NotificationController : Controller
    {
        public static string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static string appKey = ConfigurationManager.AppSettings["ida:ClientSecret"];
        private static string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];

        private static ReaderWriterLockSlim SessionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);

        [Authorize]
        public ActionResult Index()
        {
            ViewBag.CurrentUserId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value;

            //Store the notifications in session state. A production
            //application would likely queue for additional processing.
            //Store the notifications in application state. A production
            //application would likely queue for additional processing.                                                                             
            var notificationArray = (ConcurrentBag<Notification>)HttpContext.Application["notifications"];
            if (notificationArray == null)
            {
                notificationArray = new ConcurrentBag<Notification>();
            }
            HttpContext.Application["notifications"] = notificationArray;
            return View(notificationArray);
        }

        // The `notificationUrl` endpoint that's registered with the webhook subscription.
        [HttpPost]
        public async Task<ActionResult> Listen()
        {

            // Validate the new subscription by sending the token back to Microsoft Graph.
            // This response is required for each subscription.
            if (Request.QueryString["validationToken"] != null)
            {
                var token = Request.QueryString["validationToken"];
                return Content(token, "plain/text");
            }

            // Parse the received notifications.
            else
            {
                try
                {
                    using (var inputStream = new System.IO.StreamReader(Request.InputStream))
                    {
                        JObject jsonObject = JObject.Parse(inputStream.ReadToEnd());
                        var notificationArray = (ConcurrentBag<Notification>)HttpContext.Application["notifications"];

                        if (jsonObject != null)
                        {

                            // Notifications are sent in a 'value' array. The array might contain multiple notifications for events that are
                            // registered for the same notification endpoint, and that occur within a short timespan.
                            JArray value = JArray.Parse(jsonObject["value"].ToString());
                            foreach (var notification in value)
                            {
                                Notification current = JsonConvert.DeserializeObject<Notification>(notification.ToString());

                                // Check client state to verify the message is from Microsoft Graph. 
                                SubscriptionStore subscription = SubscriptionStore.GetSubscriptionInfo(current.SubscriptionId);

                                // This sample only works with subscriptions that are still cached.
                                if (subscription != null)
                                {
                                    if (current.ClientState == subscription.ClientState)
                                    {
                                        //Store the notifications in application state. A production
                                        //application would likely queue for additional processing.                                                                             
                                        if (notificationArray == null)
                                        {
                                            notificationArray = new ConcurrentBag<Notification>();
                                        }
                                        notificationArray.Add(current);
                                        HttpContext.Application["notifications"] = notificationArray;
                                    }
                                }
                            }

                            if (notificationArray.Count > 0)
                            {
                                Task.Run(() => GetChangedMessagesAsync(notificationArray)).ContinueWith((prevTask) =>
                                {
                                    return new HttpStatusCodeResult(202);
                                });
                                // Query for the changed messages.
                                //await GetChangedMessagesAsync(notificationArray);
                            }
                        }
                    }
                }
                catch (Exception)
                {

                    // TODO: Handle the exception.
                    // Still return a 202 so the service doesn't resend the notification.
                }
                return new HttpStatusCodeResult(202);
            }
        }


        public async Task GetChangedMessagesAsync(IEnumerable<Notification> notifications)
        {
            List<MessageViewModel> messages = new List<MessageViewModel>();
            GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();
            MailService mailService = new MailService();
            DataService dataService = new DataService();

            foreach (var notification in notifications)
            {
                var message = await mailService.GetMessage(graphClient, notification.ResourceData.Id);
                var mI = message.FirstOrDefault();
                if (mI != null)
                {
                    var messageItem = new MessageItem
                    {
                        BodyPreview = mI.BodyPreview,
                        ChangeKey = mI.ChangeKey,
                        ConversationId = mI.ConversationId,
                        CreatedDateTime = (DateTimeOffset)mI.CreatedDateTime,
                        Id = mI.Id,
                        IsRead = (bool)mI.IsRead,
                        Subject = mI.Subject
                    };
                    var messageItems = new List<MessageItem>();
                    messageItems.Add(messageItem);
                    dataService.StoreMessage(messageItems, mI.ParentFolderId);

                }
            }
            if (messages.Count > 0)
            {
                //NotificationService notificationService = new NotificationService();
                //notificationService.SendNotificationToClient(messages);
            }
        }
    }
}