# Mail Sync
Mail sync gets your messages from the graph api. Messages have to be tied to the folder they belong to, so make sure you get the folders first.

## Definitions

### 1. MessageItem
A local message model that is separate from the models provided by the 
client library. 

See the full message object here: https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/message

#### Properties:

|        Property         |                        Function                        |
| -------------------- | ------------------------------------------------------- |
| Id        | The Id of the message                      | 
| Subject        | The subject of the message                        | 
| BodyPreview        | Small preview of the message                        | 
| ConversaitionId        | optional Id added to show how much further you can extend                       | 
| CreatedDateTime        | Date the message was created                       | 
| IsRead        | Boolean value to show whether the message was read                       | 

### 2. FolderMessages
A local model that helps encapsulate folder messages. It contains a list of Message Items and a skip token  

#### Properties:

|        Property         |                        Function                        |
| -------------------- | ------------------------------------------------------- |
| MessageItems        | The list of message items                     | 
| SkipToken        | The token used to skip items when performing pagination                        | 


## Get Folder Messages
This results in a call to get the signed in user's messages based on the folder Id. 


> **Request:** GET graph.microsoft.com/v1.0/me/mailfolders/{id}/messages

You would need to pass an instance of the authenticated graph client. 
Thus a graph request looks like this

```csharp 
var request = graphClient.Me.MailFolders[folderId].Messages.GetAsync();
```

Here is an example of how to fetch the messages . 
Add these to your `MailService.cs` class.

#### Getting the mail messages for a folder
```csharp 
public async Task<FolderMessage> GetMyFolderMessages(GraphServiceClient graphClient, string folderId, int? skip)
{
    var top = Convert.ToInt32(ConfigurationManager.AppSettings["ida:PageSize"]);
    var folderMessages = new FolderMessage { SkipToken = null };

    // Initialise the request
    var request = graphClient.Me.MailFolders[folderId].Messages.Request();

    // if the pagination skip token has a value, add it to the request
    if (skip.HasValue)
    {
        request = request.Skip(skip.Value);
    }
    var messages = await request.Top(top).GetAsync();

    // if there are  other pages in the response, store the skip token
    if (messages.NextPageRequest != null)
    {
        foreach (var x in messages.NextPageRequest.QueryOptions)
        {
            if (x.Name == "$skip")
                folderMessages.SkipToken = Convert.ToInt32(x.Value);
        }
    }

    if (messages.Count > 0)
    {
        foreach (Message message in messages)
        {
            folderMessages.Messages.Add(new MessageItem
            {
                ConversationId = message.ConversationId,
                Id = message.Id,
                Subject = message.Subject,
                BodyPreview = message.BodyPreview,
                IsRead = (bool)message.IsRead,
                CreatedDateTime = (DateTimeOffset)message.CreatedDateTime
            });
        }
    }
    return folderMessages;
}
```

You can learn more about these requests on [List mail Messages](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/user_list_messages)


## Bringing it together
Go back to the **GetMyMailFolders** method in  `MailService.cs`. Update it so that as it is fetching the folders, it can go ahead and take the messages in the folder.
The new updated file should look like this: 

```csharp 
public async Task<List<FolderItem>> GetMyMailFolders(GraphServiceClient graphClient)
{
    List<FolderItem> items = new List<FolderItem>();
    var folders = await graphClient.Me.MailFolders.Request().GetAsync();
    if (folders?.Count > 0)
    {
        foreach (var folder in folders)
        {
            var folderMessages = await GetMyFolderMessages(graphClient, folder.Id, null);
            // checks if it is the mailbox folder so that when displayed,
            // it can show up first in the list
            var isStartUpFolder = (folder.DisplayName == "Inbox");
            items.Add(new FolderItem
            {
                Name = folder.DisplayName,
                Id = folder.Id,
                MessageItems = folderMessages.Messages,
                ParentId = null,
                SkipToken = folderMessages.SkipToken,
                StartupFolder = isStartUpFolder
            });
            var clientFolders = await GetChildFolders(graphClient, folder.Id, isStartUpFolder);
            items.AddRange(clientFolders);
        }
    }

    // order folder results, showing the startup folders first
    items = OrderFolderResults(items);
    return items;
}
```
The **OrderFolderResults** method helps rearrange the folders so that they can show the startup folders first. 
```csharp
private List<FolderItem> OrderFolderResults(List<FolderItem> folderResults)
{
    var startupFolderFolderItems = new List<FolderItem>();
    var nonStartupFolderFolderItems = new List<FolderItem>();
    var listOfFolders = new List<FolderItem>();
    foreach (var item in folderResults)
    {
        if (item.StartupFolder)
        {
            startupFolderFolderItems.Add(item);
        }
        else
        {
            nonStartupFolderFolderItems.Add(item);
        }
    }
    listOfFolders.AddRange(startupFolderFolderItems);
    listOfFolders.AddRange(nonStartupFolderFolderItems);
    return listOfFolders;
}
```

The **GetChildFolders** method in `MailService.cs` would also need an update to:

```csharp 
private async Task<List<FolderItem>> GetChildFolders(GraphServiceClient graphClient, string id, bool isStartUpFolder)
    {
        List<FolderItem> children = new List<FolderItem>();

        var childFolders = await graphClient.Me.MailFolders[id].ChildFolders.Request().GetAsync();

        if (childFolders?.Count > 0)
        {
            foreach (var child in childFolders)
            {
                var folderMessages = await GetMyFolderMessages(graphClient, child.Id, null);
                children.Add(new FolderItem
                {
                    Name = "-- " + child.DisplayName,
                    Id = child.Id,
                    MessageItems = folderMessages.Messages,
                    ParentId = child.ParentFolderId,
                    SkipToken = folderMessages.SkipToken,
                    StartupFolder = isStartUpFolder
                });
            }
        }
        return children;
    }
```



Also head back to the `HomeController.cs` file and update loop in the **GetMyMailfolders()** method to allow it to store the messages received into the json file.
Like so: 

```csharp 
public async Task<ActionResult> GetMyMailfolders()
{
    try
    {
        // Get the folders.
        var folders = await mailService.GetMyMailFolders(graphClient);

        foreach (var folder in folders)
        {
            if (dataService.FolderExists(folder.Id))
            {
                dataService.StoreMessage(folder.MessageItems, folder.Id, folder.SkipToken);
            }
            else
            {
                dataService.StoreFolder(folder);
            }

        }
    }
    ...
    // omitted for brevity
}
```



After its successful completion, it will redirect you to the `Index` action. 
The index action fetches the stored messages from the storage using the DataService's GetMessages method.

Go to [HomeController.cs](MailboxSync/Controllers/HomeController.cs) and [`MailService.cs`](MailboxSync/Services/MailService.cs) to see the entire code flow


