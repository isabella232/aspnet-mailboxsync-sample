/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading.Tasks;
using MailboxSync.Models;
using MailBoxSync.Models.Subscription;
using Microsoft.Graph;
using WebGrease.Css.Extensions;

namespace MailboxSync.Services
{
    public class MailService
    {
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
                        Messages = folderMessages.Messages,
                        ParentId = null,
                        SkipToken = folderMessages.SkipToken
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
                        Messages = folderMessages.Messages,
                        ParentId = child.ParentFolderId,
                        SkipToken = folderMessages.SkipToken
                    });
                }
            }
            return children;
        }

        private FolderMessage GenerateFolderMessages(IMailFolderMessagesCollectionPage messages)
        {
            var holder = new FolderMessage { SkipToken = null };
            if (messages.NextPageRequest != null)
            {
                foreach (var x in messages.NextPageRequest.QueryOptions)
                {
                    if (x.Name == "$skip")
                        holder.SkipToken = Convert.ToInt32(x.Value);
                }
            }

            if (messages.Count > 0)
            {
                foreach (Message message in messages)
                {
                    holder.Messages.Add(new MessageItem
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
            return holder;
        }

        public async Task<FolderMessage> GetMyFolderMessages(GraphServiceClient graphClient, string folderId, int? skip)
        {
            var top = Convert.ToInt32(ConfigurationManager.AppSettings["ida:PageSize"]);
            var request = graphClient.Me.MailFolders[folderId].Messages.Request();
            if (skip.HasValue)
            {
                request = request.Skip(skip.Value);
            }
            IMailFolderMessagesCollectionPage messages = await request.Top(top).GetAsync();
            return GenerateFolderMessages(messages);
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
        public async Task<List<Message>> GetMessage(GraphServiceClient graphClient, string id)
        {
            List<Message> items = new List<Message>();

            // Get the message.
            Message message = await graphClient.Me.Messages[id].Request().GetAsync();

            if (message != null)
            {
                items.Add(message);
            }
            return items;
        }

        // Reply to a specified message.
        public async Task<List<ResultItem>> ReplyToMessage(GraphServiceClient graphClient, string id)
        {
            var items = new List<ResultItem>();

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
            var items = new List<ResultItem>();

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