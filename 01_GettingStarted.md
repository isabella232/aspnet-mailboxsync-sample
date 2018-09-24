### Registering the app

This app uses the Azure AD v2 endpoint, so you'll register it in the [Application Registration Portal](https://apps.dev.microsoft.com).

1. Sign in to the portal with either your Microsoft account, or your work or school account.
1. Choose **Add an app**.
1. Enter a friendly name for the application and choose **Create application**.
1. Locate the **Application Secrets** section and choose **Generate New Password**. Copy the password now and save it to a safe place. Once you've copied the password, click **Ok**.
1. Locate the **Platforms** section, and choose **Add Platform**. Choose **Web**, then enter `https://localhost:44300` under **Redirect URIs**.
1. Choose **Save** at the bottom of the page.

You'll use the application ID and secret to configure the app in Visual Studio.
 
## Configure the app

1. Expose a public HTTPS notification endpoint. It can run on a service such as Microsoft Azure, or you can create a proxy web server by [using ngrok](https://ngrok.com/docs) or a similar tool.

1. Open **MailboxSync.sln** in the sample files.

    > **Note:** You may be prompted to trust certificates for localhost.

1. In Solution Explorer, open the **app.config** file in the root directory of the project.
    - For the **ida:AppId** key, replace *ENTER_YOUR_APP_ID* with the application ID of your registered application.
    - For the **ida:AppSecret** key, replace *ENTER_YOUR_SECRET* with the secret of your registered application.
    - For the **ida:NotificationUrl** key, replace *ENTER_YOUR_NOTIFICATION_URL* with the HTTPS URL. Keep the */notification/listen* portion. If you're using ngrok, use the HTTPS URL that you copied. The value will look something like this:
    ```xml
    <add key="ida:NotificationUrl" value="https://0f6fd138.ngrok.io/notification/listen" />
    ```
    - For the **ida:PageSize** key, replace *15* with the value you want to paginate with

1. Make sure that the ngrok console is still running, then press F5 to build and run the solution in debug mode.
    > **Note:** If you get errors while installing packages, make sure the local path where you placed the solution is not too long/deep. Moving the solution closer to the root drive resolves this issue.


# Key Components

### 1. MailService
A class that abstracts all the requests to the graph explorer so that the requests can be reusable everywhere in the solution.

### 2. DataService
A class that abstracts all the storage functions. The current implementation is to a json file which is not recommended for production.
You can use DocumentDb [DocumentDb](https://azure.microsoft.com/en-us/resources/videos/introduction-to-azure-documentdb) 

### 3. Local models
A local model that is separate from the models provided by the 
client library. It helps keep the code simple and you get to use only the
properties of the model that you need.

You can also add some properties to it to help out in your logic; while you can not add any properties to the client library models.

> **Note:** The Microsoft.Graph models should be looked at as convenient containers for the 
request and response bodies - that is, snapshots of the data at a point in time, 
and not to be used as client-side models.

Examples include:
- [`FolderItem.cs`](MailboxSync/Models/FolderItem.cs)
- [`MessageItem.cs`](MailboxSync/Models/MessageItem.cs)
