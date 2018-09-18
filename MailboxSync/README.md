# Mailbox Sync Sample

This is an ASP.NET web application that uses Microsoft Graph to access a user’s Microsoft account resources from within. 
This sample uses REST calls through the Graph Client to the Microsoft Graph endpoint to work with user resources--in this case, to sync emails as the user.
The sample also uses Bootstrap for styling and formatting the user experience.


## Highlights

The following are common tasks that a registered application performs:
- Get consent to fetch users' folders and messages and then get an access token.
- Use the access token to [list mailboxes](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/user_list_mailfolders), [list child folders](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/mailfolder_list_childfolders) and [list messages in the folder](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/mailfolder_list_messages).
- Store folders and messages in local data storage
- Access more message pages in folders through pagination
- Use the access token to [create a subscription](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/subscription_post_subscriptions) to a resource.
- Send back a validation token to confirm the notification URL.
- Send test email to activate notifications
- Listen for notifications from Microsoft Graph and respond with a 202 status code.
- Request more information about changed resources using data in the notification.
  
## Prerequisites

To use the Mailbox Sync sample, you need the following:

- Visual Studio 2017 installed on your development computer.
- A [work or school account](http://dev.office.com/devprogram).
- The application ID and key from the application that you [register on the Application Registration Portal](#register-the-app).
- A public HTTPS endpoint to receive and send HTTP requests. You can host this on Microsoft Azure or another service, or you can [use ngrok](#ngrok) or a similar tool while testing.


### Registering the app

This app uses the Azure AD v2 endpoint, so you'll register it in the [Application Registration Portal](https://apps.dev.microsoft.com).

1. Sign in to the portal with either your Microsoft account, or your work or school account.
1. Choose **Add an app**.
1. Enter a friendly name for the application and choose **Create application**.
1. Locate the **Application Secrets** section and choose **Generate New Password**. Copy the password now and save it to a safe place. Once you've copied the password, click **Ok**.
1. Locate the **Platforms** section, and choose **Add Platform**. Choose **Web**, then enter `https://localhost:44300` under **Redirect URIs**.
1. Choose **Save** at the bottom of the page.

You'll use the application ID and secret to configure the app in Visual Studio.
 
## Configure the applicatoin

1. Expose a public HTTPS notification endpoint. It can run on a service such as Microsoft Azure, or you can create a proxy web server by [using ngrok](https://github.com/microsoftgraph/aspnet-webhooks-rest-sample#ngrok) or a similar tool.

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


## Run the app

1. Sign in with your work or school account.

1. Consent to the **Read your mail** and **Sign you in and read your profile** permissions.

    If you don't see the **Read your mail** permission, choose **Cancel** and then add the **Read user mail** permission to the app in the Azure Portal. See the [Register the app](#register-the-app) section for instructions.

1. Choose the **Sync Mail From Server** button. The page reloads with the different mail folders, children folders and messages in each folder. Check the *mail.json* file in the root of the project to see the folder structure
       
    > **Note to self:** Add app page showing properties of the mail.json file

1. Choose the **Create subscription** button. The **Subscription** page loads with information about the subscription.

    > **Note:** This sample sets the subscription expiration to 15 minutes for testing purposes.

    > **Note to self:** Add app page showing properties of the new subscription

1. Click the **Home** link.

1. Send an email to your work or school account using the **Send Message** button. The reloads and then after a moment the page displays an alert about a new message being received. It may take several seconds for the page to update.

    > **Note to self:** Add app page showing properties of the new notification

1. Choose a folder like inbox / sent and view the list of messages associated. At the end of the list, click the **Load More** button to fetch messages using pagination.
The page will reload with a new number on the particular folder.
   
1. Click the **Notification** link to view notifications and see their structure
   

## Additional resources

* [Microsoft Graph documentation](http://graph.microsoft.io)


## Copyright

Copyright © 2018 Microsoft Corporation. All rights reserved.
