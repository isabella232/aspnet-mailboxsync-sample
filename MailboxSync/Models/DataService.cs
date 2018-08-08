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
                                folderItems.Add(new FolderItem
                                {
                                    Name = item["Name"].ToString(),
                                    Id = item["Id"].ToString(),
                                    Messages = JsonConvert.DeserializeObject<List<MessageItem>>(item["Messages"].ToString())
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
                    folderArrary.Add(JObject.Parse(JsonConvert.SerializeObject(folder)));

                folderObject["folders"] = folderArrary;
                string newFolderContents = JsonConvert.SerializeObject(folderObject, Formatting.Indented);
                System.IO.File.WriteAllText(jsonFile, newFolderContents);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Add Error : " + ex.Message.ToString());
            }
        }

        public void StoreMessage(List<MessageItem> messages, string folder)
        {
            string jsonFile = System.Web.Hosting.HostingEnvironment.MapPath("~/mail.json");
            try
            {
                var json = System.IO.File.ReadAllText(jsonFile);
                var folderObject = JObject.Parse(json);
                var folderArrary = folderObject.GetValue("folders") as JArray;
                if (folderArrary != null)
                {
                    var mailData = JObject.Parse(json);
                    JArray messageObject = JArray.Parse(JsonConvert.SerializeObject(messages));

                    if (!string.IsNullOrEmpty(folder))
                    {
                        foreach (var mailFolder in folderArrary.Where(obj => obj["Id"].Value<string>() == folder))
                        {
                            mailFolder["Messages"] = messageObject;
                        }

                        mailData["folders"] = folderArrary;
                        string output = JsonConvert.SerializeObject(mailData, Formatting.Indented);
                        System.IO.File.WriteAllText(jsonFile, output);
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