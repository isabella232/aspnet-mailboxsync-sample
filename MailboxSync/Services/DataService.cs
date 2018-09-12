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
    /// You can use DocumentDb ; you can learn more about it here: https://azure.microsoft.com/en-us/resources/videos/introduction-to-azure-documentdb/
    /// </summary>
    public class DataService
    {
        /// <summary>
        /// Gets the folders that are stored locally 
        /// </summary>
        /// <returns></returns>
        public List<FolderItem> GetFolders()
        {
            string jsonFile = HostingEnvironment.MapPath("~/mail.json");
            List<FolderItem> folderItems = new List<FolderItem>();
            if (!File.Exists(jsonFile))
            {
                return folderItems;
            }
            try
            {
                var mailData = File.ReadAllText(jsonFile);
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
                            SkipToken = (int?)item["SkipToken"]
                        });
                    }
                    return folderItems;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Add Error : " + ex.Message);
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
            catch (Exception ex)
            {
                Console.WriteLine("Add Error : " + ex.Message);
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
            string jsonFile = HostingEnvironment.MapPath("~/mail.json");
            try
            {
                var mailBox = JObject.Parse(File.ReadAllText(jsonFile));
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
            catch (Exception ex)
            {

            }
            return exists;
        }


        /// <summary>
        /// stores a folder item in the json file
        /// </summary>
        /// <param name="folder">the folder item generated</param>
        public void StoreFolder(FolderItem folder)
        {
            string jsonFile = HostingEnvironment.MapPath("~/mail.json");
            try
            {
                var mailBox = File.ReadAllText(jsonFile);
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
                File.WriteAllText(jsonFile, newFolderContents);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Add Error : " + ex.Message);
            }
        }


        /// <summary>
        /// Adds a list of messages to a particular folder in the json file
        /// </summary>
        /// <param name="messages">the list of messages to be stored</param>
        /// <param name="folderId">the id of the folder where the messages will be added</param>
        /// <param name="messagesSkipToken">in case the messages are coming from a pagination request, the skip token is stored for the next request</param>
        public void StoreMessage(List<MessageItem> messages, string folderId, int? messagesSkipToken)
        {
            string jsonFile = HostingEnvironment.MapPath("~/mail.json");
            try
            {
                var mailBox = JObject.Parse(File.ReadAllText(jsonFile));
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
            catch (Exception ex)
            {
                Console.WriteLine("Add Error : " + ex.Message);
            }
        }


        /// <summary>
        /// updates the values in the folder 
        /// </summary>
        /// <param name="folder">the folder whose properties need to change</param>
        private void UpdateFolder(FolderItem folder)
        {
            string jsonFile = HostingEnvironment.MapPath("~/mail.json");
            try
            {
                var json = File.ReadAllText(jsonFile);
                var folderObject = JObject.Parse(json);
                var folderArrary = folderObject.GetValue("folders") as JArray;
                if (folderArrary != null)
                {
                    var mailData = JObject.Parse(json);
                    JArray messageObject = JArray.Parse(JsonConvert.SerializeObject(folder.MessageItems));

                    if (!string.IsNullOrEmpty(folder.Id))
                    {
                        foreach (var mailFolder in folderArrary.Where(obj => obj["Id"].Value<string>() == folder.Id))
                        {
                            mailFolder["MessageItems"] = messageObject;
                            mailFolder["SkipToken"] = folder.SkipToken;
                        }

                        mailData["folders"] = folderArrary;
                        string output = JsonConvert.SerializeObject(mailData, Formatting.Indented);
                        File.WriteAllText(jsonFile, output);
                    }
                    else
                    {
                        Console.Write(" Try Again!");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Add Error : " + ex.Message);
            }
        }
    }
}