/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/


using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Hosting;
using MailboxSync.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MailboxSync.Services
{
    /// <summary>
    /// Interfaces with the data storage
    /// Current implementation uses a json file which is not recommended for production.
    /// You can use DocumentDb ; learn more about it here: https://azure.microsoft.com/en-us/resources/videos/introduction-to-azure-documentdb/
    /// </summary>
    public class DataService
    {
        private readonly string _jsonFile = HostingEnvironment.MapPath("~/mail.json");

        /// <summary>
        /// Gets the folders that are stored locally 
        /// </summary>
        /// <returns></returns>
        public List<FolderItem> GetFolders()
        {
            List<FolderItem> folderItems = new List<FolderItem>();
            if (!File.Exists(_jsonFile))
            {
                return folderItems;
            }
            try
            {
                var mailData = File.ReadAllText(_jsonFile);
                var jObject = JObject.Parse(mailData);
                JArray folders = (JArray)jObject["folders"];
                if (folders != null)
                {
                    foreach (var item in folders)
                    {
                        folderItems.Add(new FolderItem
                        {
                            Name = item["Name"].ToString(),
                            Id = item["Id"].ToString(),
                            MessageItems = GenerateMessages(item["MessageItems"].ToString()),
                            SkipToken = (int?)item["SkipToken"],
                            StartupFolder = (bool)item["StartupFolder"],
                        });
                    }
                    return folderItems;
                }
            }
            catch (Exception)
            {
                // ignored
            }

            return folderItems;
        }


        /// <summary>
        /// Generates a list of Message items from the string value of the messages in the json file
        /// </summary>
        /// <param name="messageString">the json string version of the messages in the json file</param>
        /// <returns></returns>
        private List<MessageItem> GenerateMessages(string messageString)
        {
            var messageItem = new List<MessageItem>();
            try
            {
                var messageArray = JArray.Parse(messageString);
                foreach (var item in messageArray)
                {
                    var mItem = JObject.Parse(item.ToString());
                    messageItem.Add(new MessageItem
                    {
                        Id = mItem["id"].ToString(),
                        Subject = mItem["subject"].ToString(),
                        IsRead = (bool)mItem["isRead"],
                        BodyPreview = mItem["bodyPreview"].ToString(),
                        CreatedDateTime = (DateTimeOffset)mItem["createdDateTime"]
                    });
                }
            }
            catch (Exception)
            {
                // ignored
            }
            return messageItem.OrderByDescending(k => k.CreatedDateTime).ToList();
        }


        /// <summary>
        /// Checks if the folder with the id exists
        /// </summary>
        /// <param name="folderId">The id of the folder</param>
        /// <returns>bool</returns>
        public bool FolderExists(string folderId)
        {
            bool exists = false;
            try
            {
                var mailBox = JObject.Parse(File.ReadAllText(_jsonFile));
                var folders = mailBox.GetValue("folders") as JArray;
                if (folders != null)
                {
                    if (!string.IsNullOrEmpty(folderId))
                    {
                        var folder = folders.Where(obj => obj["Id"].Value<string>() == folderId);
                        if (folder.ToList().Count > 0)
                        {
                            exists = true;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // ignored
            }

            return exists;
        }


        /// <summary>
        /// stores a folder item in the json file
        /// </summary>
        /// <param name="folder">the folder item generated</param>
        public void StoreFolder(FolderItem folder)
        {
            try
            {
                var mailBox = File.ReadAllText(_jsonFile);
                var mailBoxObject = JObject.Parse(mailBox);
                var folderArrary = mailBoxObject.GetValue("folders") as JArray;

                if (folderArrary == null)
                    folderArrary = new JArray();

                if (folderArrary.All(obj => obj["Id"].Value<string>() != folder.Id))
                {
                    folderArrary.Add(JObject.Parse(JsonConvert.SerializeObject(folder)));
                }

                mailBoxObject["folders"] = folderArrary;
                string newFolderContents = JsonConvert.SerializeObject(mailBoxObject, Formatting.Indented);
                File.WriteAllText(_jsonFile, newFolderContents);
            }
            catch (Exception)
            {
                // ignored
            }
        }


        /// <summary>
        /// Adds a list of messages to a particular folder in the json file.
        /// Can be used to save one or more messages.
        /// To save one message, add it to a list and pass the list through to the methid
        /// </summary>
        /// <param name="messages">the list of messages to be stored</param>
        /// <param name="folderId">the id of the folder where the messages will be added</param>
        /// <param name="messagesSkipToken">in case the messages are coming from a pagination request, the skip token is stored for the next request</param>
        public void StoreMessage(List<MessageItem> messages, string folderId, int? messagesSkipToken)
        {
            try
            {
                var mailBox = JObject.Parse(File.ReadAllText(_jsonFile));
                var folders = mailBox.GetValue("folders") as JArray;
                if (folders != null)
                {
                    if (!string.IsNullOrEmpty(folderId))
                    {
                        var folder = folders.Where(obj => obj["Id"].Value<string>() == folderId);
                        foreach (var item in folder)
                        {
                            var newFolderItem = new FolderItem
                            {
                                Name = item["Name"].ToString(),
                                Id = item["Id"].ToString(),
                                MessageItems = GenerateMessages(item["MessageItems"].ToString())
                            };
                            newFolderItem.MessageItems.AddRange(messages);
                            newFolderItem.SkipToken = messagesSkipToken;
                            newFolderItem.MessageItems = newFolderItem.MessageItems.GroupBy(p => new { p.Id }).Select(g => g.First()).ToList();
                            UpdateFolder(newFolderItem);
                        }
                    }
                    else
                    {
                        Console.Write(" Try Again!");
                    }
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }


        /// <summary>
        /// updates the values in the folder 
        /// </summary>
        /// <param name="folder">the folder whose properties need to change</param>
        private void UpdateFolder(FolderItem folder)
        {
            
            try
            {
                var json = File.ReadAllText(_jsonFile);
                var folderObject = JObject.Parse(json);
                var folderArrary = folderObject.GetValue("folders") as JArray;
                if (folderArrary == null) return;
                var mailData = JObject.Parse(json);
                var messageObject = JArray.Parse(JsonConvert.SerializeObject(folder.MessageItems));
                if (string.IsNullOrEmpty(folder.Id)) return;
                foreach (var mailFolder in folderArrary.Where(obj => obj["Id"].Value<string>() == folder.Id))
                {
                    mailFolder["MessageItems"] = messageObject;
                    mailFolder["SkipToken"] = folder.SkipToken;
                }
                mailData["folders"] = folderArrary;
                string output = JsonConvert.SerializeObject(mailData, Formatting.Indented);
                File.WriteAllText(_jsonFile, output);
            }
            catch (Exception)
            {
                // ignored
            }
        }
    }
}