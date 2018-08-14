/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using MailBoxSync.Models.Subscription;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MailboxSync.Models
{
    public class MailService
    {
        // Get folders in the current mail.
        public async Task<List<FolderItem>> GetMyMailFolders(GraphServiceClient graphClient)
        {
            List<FolderItem> items = new List<FolderItem>();

            // Get messages in the Inbox folder.
            var folders = await graphClient.Me.MailFolders.Request().GetAsync();

            if (folders?.Count > 0)
            {
                foreach (var folder in folders)
                {
                    items.Add(new FolderItem
                    {
                        Name = folder.DisplayName,
                        Id = folder.Id,
                        Messages = await GetMyFolderMessages(graphClient, folder.Id),
                        ParentId = null
                    });
                    var clientFolders = await GetChildFolders(graphClient, folder.Id);
                    items.AddRange(clientFolders);
                }
            }
            return items;
        }

        private async Task<List<FolderItem>> GetChildFolders(GraphServiceClient graphClient, string id)
        {
            List<FolderItem> children = new List<FolderItem>();

            // Get messages in the Child folder.
            var childFolders = await graphClient.Me.MailFolders[id].ChildFolders.Request().GetAsync();

            if (childFolders?.Count > 0)
            {
                foreach (var child in childFolders)
                {

                    children.Add(new FolderItem
                    {
                        Name = "-- " + child.DisplayName,
                        Id = child.Id,
                        Messages = await GetMyFolderMessages(graphClient, child.Id),
                        ParentId = child.ParentFolderId
                    });
                }
            }
            return children;
        }

        private List<MessageItem> CreateMessages(IMailFolderMessagesCollectionPage messages)
        {
            var items = new List<MessageItem>();
            if (messages?.Count > 0)
            {
                foreach (Message message in messages)
                {
                    items.Add(new MessageItem
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
            return items;
        }

        public async Task<List<MessageItem>> GetMyFolderMessages(GraphServiceClient graphClient, string folderId)
        {
            var items = new List<MessageItem>();
            IMailFolderMessagesCollectionPage messages = await graphClient.Me.MailFolders[folderId].Messages.Request().GetAsync();
            items = CreateMessages(messages);
            return items;
        }

        // Send an email message.
        // This snippet sends a message to the current user on behalf of the current user.
        public async Task<List<FolderItem>> SendMessage(GraphServiceClient graphClient)
        {
            List<FolderItem> items = new List<FolderItem>();

            // Create the recipient list. This snippet uses the current user as the recipient.
            User me = await graphClient.Me.Request().Select("Mail, UserPrincipalName").GetAsync();
            string address = me.Mail ?? me.UserPrincipalName;
            string guid = Guid.NewGuid().ToString();

            List<Recipient> recipients = new List<Recipient>();
            recipients.Add(new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = address
                }
            });

            // Create the message.
            Message email = new Message
            {
                Body = new ItemBody
                {
                    Content = "Body" + guid,
                    ContentType = BodyType.Text,
                },
                Subject = "Subject" + guid.Substring(0, 8),
                ToRecipients = recipients
            };

            // Send the message.
            await graphClient.Me.SendMail(email, true).Request().PostAsync();

            return items;
        }

        // Get a specified message.
        public async Task<List<ResultItem>> GetMessage(GraphServiceClient graphClient, string id)
        {
            List<ResultItem> items = new List<ResultItem>();

            // Get the message.
            Message message = await graphClient.Me.Messages[id].Request().GetAsync();

            if (message != null)
            {
                items.Add(new ResultItem
                {

                    // Get message properties.
                    Display = message.Subject,
                    Id = message.Id,
                    Properties = new Dictionary<string, object>
                    {
                        { "BodyPreview", message.BodyPreview },
                        { "IsDraft", message.IsDraft.ToString() },
                        { "Id", message.Id }
                    }
                });
            }
            return items;
        }

        // Reply to a specified message.
        public async Task<List<ResultItem>> ReplyToMessage(GraphServiceClient graphClient, string id)
        {
            List<ResultItem> items = new List<ResultItem>();

            // Reply to the message.
            await graphClient.Me.Messages[id].Reply("Some text content.").Request().PostAsync();

            items.Add(new ResultItem
            {

                // This operation doesn't return anything.
                Properties = new Dictionary<string, object>
                {
                    { "Operation completed. This call doesn't return anything.", "" }
                }
            });
            return items;
        }


        // Delete a specified message.
        public async Task<List<ResultItem>> DeleteMessage(GraphServiceClient graphClient, string id)
        {
            List<ResultItem> items = new List<ResultItem>();

            // Delete the message.
            await graphClient.Me.Messages[id].Request().DeleteAsync();

            items.Add(new ResultItem
            {

                // This operation doesn't return anything.
                Properties = new Dictionary<string, object>
                {
                    { "Operation completed. This call doesn't return anything.", "" }
                }
            });
            return items;
        }
    }
}