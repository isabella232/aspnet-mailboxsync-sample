using MailBoxSync.Models.Subscription;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MailboxSync.Models
{
    public class DataService
    {
        public List<FolderItem> GetFolders()
        {
            string jsonFile = System.Web.Hosting.HostingEnvironment.MapPath("~/mail.json");
            List<FolderItem> folderItems = new List<FolderItem>();
            if (!System.IO.File.Exists(jsonFile))
            {
                return folderItems;
            }
            else
            {
                var mailData = System.IO.File.ReadAllText(jsonFile);
                if (mailData == null)
                {
                    return folderItems;
                }
                else
                {
                    try
                    {
                        var jObject = JObject.Parse(mailData);
                        JArray folders = (JArray)jObject["folders"];
                        if (folders != null)
                        {
                            foreach (var item in folders)
                            {
                                var name = item["Name"].ToString();
                                folderItems.Add(new FolderItem
                                {
                                    Name = item["Name"].ToString(),
                                    Id = item["Id"].ToString(),
                                    Messages = GenerateMessages(item["Messages"].ToString())
                                });
                            }
                            return folderItems;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Add Error : " + ex.Message.ToString());
                    }
                }
            }
            return folderItems;
        }

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
                        CreatedDateTime = (DateTimeOffset)mItem["createdDateTime"],
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Add Error : " + ex.Message.ToString());
            }
            return messageItem;
        }

        public void StoreFolder(FolderItem folder)
        {
            string jsonFile = System.Web.Hosting.HostingEnvironment.MapPath("~/mail.json");
            try
            {
                var json = System.IO.File.ReadAllText(jsonFile);
                var folderObject = JObject.Parse(json);
                var folderArrary = folderObject.GetValue("folders") as JArray;

                if (folderArrary == null)
                    folderArrary = new JArray();

                if (!folderArrary.Any(obj => obj["Id"].Value<string>() == folder.Id))
                {
                    folderArrary.Add(JObject.Parse(JsonConvert.SerializeObject(folder)));
                }

                folderObject["folders"] = folderArrary;
                string newFolderContents = JsonConvert.SerializeObject(folderObject, Formatting.Indented);
                System.IO.File.WriteAllText(jsonFile, newFolderContents);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Add Error : " + ex.Message.ToString());
            }
        }

        public void StoreMessage(List<MessageItem> messages, string folderId)
        {
            string jsonFile = System.Web.Hosting.HostingEnvironment.MapPath("~/mail.json");
            try
            {
                var mailBox = JObject.Parse(System.IO.File.ReadAllText(jsonFile));
                var folders = mailBox.GetValue("folders") as JArray;
                if (folders != null)
                {
                    if (!string.IsNullOrEmpty(folderId))
                    {
                        var folder = folders.Where(obj => obj["Id"].Value<string>() == folderId);
                        List<FolderItem> folderItems = new List<FolderItem>();
                        foreach (var item in folder)
                        {
                            var name = item["Name"].ToString();
                            var newFolderItem = new FolderItem
                            {
                                Name = item["Name"].ToString(),
                                Id = item["Id"].ToString(),
                                Messages = GenerateMessages(item["Messages"].ToString())
                            };
                            newFolderItem.Messages.AddRange(messages);
                            newFolderItem.Messages = newFolderItem.Messages.Distinct().ToList();
                            StoreFolder(newFolderItem);
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
                Console.WriteLine("Add Error : " + ex.Message.ToString());
            }
        }
    }
}