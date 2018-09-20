# Mailbox Sync Sample

This is an ASP.NET web application that uses Microsoft Graph to access a user’s Microsoft account resources from within. 
This sample uses REST calls through the Graph Client to the Microsoft Graph endpoint to work with user resources--in this case, to sync emails as the user.
The sample also uses Bootstrap for styling and formatting the user experience.

## Highlights

The following are common tasks that the application performs:
- Gets consent to fetch users' folders and messages and then get an access token.
- Uses the access token to [list mailboxes](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/user_list_mailfolders), [list child folders](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/mailfolder_list_childfolders) and [list messages in the folder](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/mailfolder_list_messages).
- Stores folders and messages in local data storage
- Accesses more message pages in folders through pagination
- Uses the access token to [create a subscription](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/subscription_post_subscriptions) to a resource.
- Sends back a validation token to confirm the notification URL.
- Sends test email to activate notifications
- Listens for notifications from Microsoft Graph and respond with a 202 status code.
- Requests more information about changed resources using data in the notification.
  
## Prerequisites

To use the Mailbox Sync sample, you need the following:

- Visual Studio 2017 installed on your development computer.
- A work or school account.
- The application ID and key from the application that you [register on the Application Registration Portal](#register-the-app).
- A public HTTPS endpoint to receive and send HTTP requests. You can host this on Microsoft Azure or another service, or you can [use ngrok](#ngrok) or a similar tool while testing.


# In this lab

There are 3 components that make up the Mailbox Sample.
1. [Exercise 1: Getting Started](01_GettingStarted.md)
1. [Exercise 2: Sync folders](02_FolderSync.md)
1. [Exercise 3: Sync email messages](MailSync.md)
1. [Exercise 4: Create subscription and receive notifications](Subscription.md)

## Additional resources

* [Microsoft Graph documentation](http://graph.microsoft.io)


## Copyright

Copyright © 2018 Microsoft Corporation. All rights reserved.
