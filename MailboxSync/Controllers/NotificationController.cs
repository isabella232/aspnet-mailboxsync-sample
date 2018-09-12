using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Security.Claims;
using System.Web.Mvc;
using System.Threading.Tasks;
using MailboxSync.Helpers;
using MailboxSync.Models;
using MailboxSync.Services;
using MailboxSync.Services.SignalR;

namespace MailboxSync.Controllers
{
    public class NotificationController : Controller
    {
        public static string ClientId = ConfigurationManager.AppSettings["ida:ClientId"];

        [Authorize]
        public ActionResult Index()
        {
            ViewBag.CurrentUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

            //Store the notifications in session state. A production
            //application would likely queue for additional processing.
            //Store the notifications in application state. A production
            //application would likely queue for additional processing.                                                                             
            var notificationArray = (ConcurrentBag<NotificationItem>)HttpContext.Application["notifications"];
            if (notificationArray == null)
            {
                notificationArray = new ConcurrentBag<NotificationItem>();
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
                        var notificationArray = (ConcurrentBag<NotificationItem>)HttpContext.Application["notifications"];

                        if (jsonObject != null)
                        {

                            // Notifications are sent in a 'value' array. The array might contain multiple notifications for events that are
                            // registered for the same notification endpoint, and that occur within a short timespan.
                            JArray value = JArray.Parse(jsonObject["value"].ToString());
                            foreach (var notification in value)
                            {
                                NotificationItem current = JsonConvert.DeserializeObject<NotificationItem>(notification.ToString());

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
                                            notificationArray = new ConcurrentBag<NotificationItem>();
                                        }
                                        notificationArray.Add(current);
                                        HttpContext.Application["notifications"] = notificationArray;
                                    }
                                }
                            }

                            if (notificationArray.Count > 0)
                            {
                                await GetChangedMessagesAsync(notificationArray);
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


        public async Task GetChangedMessagesAsync(IEnumerable<NotificationItem> notifications)
        {
            DataService dataService = new DataService();
            int newMessages = 0;
            foreach (var notification in notifications)
            {
                var subscription = SubscriptionStore.GetSubscriptionInfo(notification.SubscriptionId);

                var graphClient = GraphSdkHelper.GetAuthenticatedClient(subscription.UserId);

                try
                {
                    // Get the message
                    var message = await graphClient.Me.Messages[notification.ResourceData.Id].Request()
                        .Select("id,subject,bodyPreview,createdDateTime,isRead,parentFolderId,conversationId,changeKey")
                        .GetAsync();

                    if (message != null)
                    {
                        var messageItem = new MessageItem
                        {
                            BodyPreview = message.BodyPreview,
                            ChangeKey = message.ChangeKey,
                            ConversationId = message.ConversationId,
                            CreatedDateTime = (DateTimeOffset)message.CreatedDateTime,
                            Id = message.Id,
                            IsRead = (bool)message.IsRead,
                            Subject = message.Subject
                        };
                        var messageItems = new List<MessageItem>();
                        messageItems.Add(messageItem);
                        dataService.StoreMessage(messageItems, message.ParentFolderId, null);
                        newMessages += 1;

                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
            }

            if (newMessages > 0)
            {
                NotificationService notificationService = new NotificationService();
                notificationService.SendNotificationToClient(1);
            }
        }
    }
}