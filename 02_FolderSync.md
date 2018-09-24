# Folder Sync
Folder sync gets your folders from the graph api. Sometimes folders have child folders within them and it is important to fetch them too.
For this example, we are going to go down to one level deep and save the results in a json file for easy investigation of the data received.

## Definitions

### 1. FolderItem
A local message model that is separate representation of the models provided by the 
client library. 


See the actual folder object here: https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/mailfolder

#### Properties:

|        Property         |                        Function                        |
| -------------------- | ------------------------------------------------------- |
| Id        | The Id of the folder                      | 
| Name        | The name of the folder                        | 
| ParentId        | Id of the folder parent folder if it is a child folder                  | 
| MessageItems        | list of messages in the folder                        | 
| SkipToken        | Optional nullable property that is updated during pagination to store the skip token                       | 


## Get Mail Folders
This results in a call to get the signed in user's mailfolder. 


> **Request:** GET graph.microsoft.com/v1.0/me/mailfolders


You would need to pass an instance of the authenticated graph client

```csharp 
var folders = await graphClient.Me.MailFolders.Request().GetAsync();
```

After receiving the list of folders, you should loop through them to get the list of child folders using the folder Id. We've added a few dashes at the front of the name to show child folders.

> **Request:** GET graph.microsoft.com/v1.0/me/mailfolders/{id}/childFolders

The graph request looks like this

```csharp 
var childFolders = await graphClient.Me.MailFolders[folder.Id].ChildFolders.Request().GetAsync();
``` 

Here is an example of how to fetch the folders and the child folders. 
Add these to your `MailService.cs` class.

#### Getting the mail folders
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
            items.Add(new FolderItem
            {
                Name = folder.DisplayName,
                Id = folder.Id,
                ParentId = null,
            });
            var clientFolders = await GetChildFolders(graphClient, folder.Id);
            items.AddRange(clientFolders);
        }
    }
    return items;
}
```
#### Getting the child folders
```csharp 
private async Task<List<FolderItem>> GetChildFolders(GraphServiceClient graphClient, string id)
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
                ParentId = child.ParentFolderId,
            });
        }
    }
    return children;
}
``` 
You can learn more about these requests on [List mailFolders](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/user_list_mailfolders) 
and [List child folders](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/mailfolder_list_childfolders)


## Bringing it together
Create an action in `HomeController.cs` called GetMyMailfolders. 
This should pass the results of fetching the folders from the graph API to the data service for storage of the details.
```csharp 
public async Task<ActionResult> GetMyMailfolders()
{
    var results = new FoldersViewModel();
    try
    {
        // Initialize the GraphServiceClient.
        GraphServiceClient graphClient = GraphSdkHelper.GetAuthenticatedClient();

        // Get the folders.
        results.Items = await mailService.GetMyMailFolders(graphClient);

        foreach (var folder in results.Items)
        {
            dataService.StoreFolder(folder);
        }
    }
    catch (ServiceException se)
    {
        if (se.Error.Code == "AuthenticationFailure")
        {
            return new EmptyResult();
        }

        // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
        return RedirectToAction("Index", "Error", new { message = string.Format("Error in {0}: {1} {2}", Request.RawUrl, se.Error.Code, se.Error.Message) });
    }
    return RedirectToAction("Index");
}
```


After its successful completion, it will redirect you to the `Index` action. 
The index action fetches the stored folders from the storage using the DataService's GetFolders method.

The action goes through the items stored in the json file, creates a list of folder items
and passes it to the UI. 

```csharp 
public ActionResult Index()
{
    var folderResults = new FoldersViewModel();
    var folders = dataService.GetFolders();
    var resultItems = new List<FolderItem>();
    folderResults.Items.ToList().AddRange(folders);
    folderResults.Items = resultItems;
    return View("Index", folderResults);
}
```
