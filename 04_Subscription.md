#  Create subscription and receive notifications
Subscriptions are used to track changes to the mailbox so that when something happens, whether a create message event or an update message event, your client is able to receive a notification of this change and you can update your local data accordingly.


They are nicely documented here: https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/subscription


    
## Definitions
1. SubscriptionStore

It helps store details of the subscription in a cache so the **NotificationController** can retrieve an access token from the cache and validate the notification using the stored subscription Id. In production you should use some method of persistent storage

## Create subscription
For a successful create subscription request, you must set up a controller endpoint that is able to listen out for notifications. This endpoint has to be publicly available. This would not be a problem during production. However, since you are on development mode, you have two options:
1. Set up your project for [azure](https://azure.microsoft.com/) debugging. This is quite simple if you already have an azure account and an active subscription since you would need to publish the application to azure for you to test the request.

    Instructions on how to set-up for azure debugging are detailed [here](https://azure.microsoft.com/en-in/blog/introduction-to-remote-debugging-on-azure-web-sites/)


2. Use [ngrok](https://ngrok.com/) for local debugging. ngrok provides a tunnel that exposes your local endpoint to the public hence you are able to debug your application without hosting it. The resources [here](https://ngrok.com/docs) show the process to use ngrok after installing.

    You can run the below command, changing **21942** with the **http** port on your solution

    ```
    ngrok http 21942 -host-header=localhost:21942
    ```

Your **create** method on `SubscriptionController.cs` should have this if you are subscribing to newly created messages in the inbox folder. You can change it up to a different folder if you would like to see how it reacts
```csharp
var subscription = new Subscription
{
    Resource = "me/mailFolders('Inbox')/messages",
    ChangeType = "created",
    NotificationUrl = ConfigurationManager.AppSettings["ida:NotificationUrl"],
    ClientState = Guid.NewGuid().ToString(),
    ExpirationDateTime = DateTime.UtcNow + new TimeSpan(0, 0, 15, 0) // shorter duration useful for testing
};

var response = await graphClient.Subscriptions.Request().AddAsync(subscription);
```

You should have a **listen** method on `NotificationController.cs` that must return a status 200 for the subscription to be successful. This is because it has to check whether the notification endpoint is alive before going over to the graph to perform the request.

```csharp
// Validates the new subscription by sending the token back to Microsoft Graph.
// This response is required for each new subscription.
if (Request.QueryString["validationToken"] != null)
{
    var token = Request.QueryString["validationToken"];
    return Content(token, "plain/text");
}
```

Once you receive the response about a successful subscription creation, you can go ahead and store the subscription ID in your persistent storage.
It is important to keep a bit more detail about the subscription especially the current user ID so that the notification is able to authenticate while performing requests later.

```csharp
SubscriptionStore.SaveSubscriptionInfo(
    Subscription.Id,
    Subscription.ClientState,
    ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value, // the user Id
    ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value);
```
## Process notifications
Once a change happens and a subscription exists, a notification from Microsoft Graph is sent to your listening endpoint with details of what happened.

To process it, you would need to check the values of the notifications and look at the type of change that occured together with the id of the resource affected and change your local data accordingly.
In this example, we only subscribed to new messages in the inbox folder. This is how we process the notification

```csharp
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
```
### Update the folder with new details
Once the notification array has all the notifications in an easy to loop list, you pass it through to the **GetChangedMessagesAsync** method where it gets the details of the messages and leverages `DataService.cs` to update the values of the local data store.
```csharp
foreach (var notification in notifications)
{
    var subscription = SubscriptionStore.GetSubscriptionInfo(notification.SubscriptionId);
    var graphClient = GraphServiceClientProvider.GetAuthenticatedClient(subscription.UserId);

    // Get the message
    var message = await mailService.GetMessage(graphClient, notification.ResourceData.Id);

    // update the local json file
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
        var messageItems = new List<MessageItem> { messageItem };
        dataService.StoreMessage(messageItems, message.ParentFolderId, null);
        newMessages += 1;
    }
}
```
## Notify the UI screen
[SignalR](https://docs.microsoft.com/en-us/aspnet/signalr/overview/getting-started/introduction-to-signalr) helps create a real-time experience for ASP .Net web apps. 
Setting it up requires that you 

1. add the nuget package, like so:

    ```
    install-package Microsoft.AspNet.SignalR
    ```
2. create a `NotificationService.cs` class in the services folder.

    ```csharp
    public class NotificationService
    {
        /// <summary>
        /// Fires the notification to the client
        /// </summary>
        public void SendNotificationToClient()
        {
            var hubContext = GlobalHost.ConnectionManager.GetHubContext<NotificationHub>();
            if (hubContext != null)
            {
                hubContext.Clients.All.showNotification();
            }
        }
    }
    ```

3. add this to your scripts section in the **Layout.cshtml** file

    ```html
    <script src="~/signalr/hubs"></script>
    ```


The **newMessages** in the NotificationController is used to track the successful messages that were fetched and updated. If greater than zero they inform signalR that notifies the UI that there are new messages. This gives the end user information about what is happening in the background.
```csharp
if (newMessages > 0)
{
    NotificationService notificationService = new NotificationService();
    notificationService.SendNotificationToClient();
}
```

The **cshtml** pages have to be listening out for the notification from SignalR and so, add the following code to the end of your `Home/Index.cshtml` file inside a `<script></script>` tag.

```javascript
var connection = $.hubConnection();
var hub = connection.createHubProxy("NotificationHub");
hub.on("showNotification", function () {
    var card = $("<div class=\"card\"></div>");
    var cardBody = $("<div class=\"card-body\"></div>").appendTo(card);
    $("<h5 class=\"\"><b><b></h5>").text("New Message(s) received, refresh page to view messages").appendTo(cardBody);
    $("#notice").empty();
    $("#notice").append(card);
    $("#notice");
});
connection.start();
```
The page will have show some text when your receive new messages and prompt you to reload the page. Since we subscribed to the inbox folder, you can pay special attention to it before reloading and after reloading.


# Bringing it together

The full code is available on [`SubscriptionController.cs`](MailboxSync/Controllers/SubscriptionController.cs) , [`NotificationController.cs`](MailboxSync/Controllers/NotificationController.cs) and [`SubscriptionStore.cs`](MailboxSync/Services/SubscriptionStore.cs)