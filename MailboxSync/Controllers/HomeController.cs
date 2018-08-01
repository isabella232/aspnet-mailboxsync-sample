/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using MailboxSync.Helpers;
using MailboxSync.Models;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Resources;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace MailboxSync.Controllers
{

    [Authorize]
    public class HomeController : Controller
    {
        MailService mailService = new MailService();

        public ActionResult Index()
        {
            return View();
        }

        public List<FolderItem> GetFolders()
        {
            string jsonFile = Server.MapPath("~/mail.json");
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
                                    Id = item["Id"].ToString()
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

        public void AddFolders(FolderItem folder)
        {
            string jsonFile = Server.MapPath("~/mail.json");
            try
            {
                var json = System.IO.File.ReadAllText(jsonFile);
                var folderObject = JObject.Parse(json);
                var folderArrary = folderObject.GetValue("folders") as JArray;

                if (folderArrary == null)
                    folderArrary = new JArray();

                if (!folderArrary.Contains(folder.Id))
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

        public void StoreMessage(List<Message> messages, string folder)
        {
            string jsonFile = Server.MapPath("~/mail.json");
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

        public async Task<ActionResult> AddMessages(string id)
        {
            ResultsViewModel results = new ResultsViewModel();
            try
            {
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();
                results.Items = await mailService.GetMyFolderMessages(graphClient, id);
                var messages = new List<Message>();
                foreach (var item in results.Items)
                {
                    messages.Add(new Message { Id = item.Id, Subject = item.Display });
                }
                StoreMessage(messages, id);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);


            //string jsonFile = Server.MapPath("~/mail.json");
            //try
            //{
            //    var json = System.IO.File.ReadAllText(jsonFile);
            //    var folderObject = JObject.Parse(json);
            //    var folderArrary = folderObject.GetValue("folders") as JArray;





            //    folderObject["folders"] = folderArrary;
            //    string newFolderContents = JsonConvert.SerializeObject(folderObject, Formatting.Indented);
            //    System.IO.File.WriteAllText(jsonFile, newFolderContents);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("Add Error : " + ex.Message.ToString());
            //}
        }

        public ActionResult GetFolderDetails()
        {
            ResultsViewModel results = new ResultsViewModel();
            var folders = GetFolders();
            var resultItems = new List<ResultsItem>();
            foreach (var item in folders)
            {
                resultItems.Add(new ResultsItem
                {
                    Display = item.Name,
                    Id = item.Id
                });
            }
            results.Items = resultItems;
            return View("Index", results);
        }



        // Get messages in all the current user's mail folders.
        public async Task<ActionResult> GetMyMessages()
        {
            ResultsViewModel results = new ResultsViewModel();
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Get the messages.
                results.Items = await mailService.GetMyMessages(graphClient);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        // Get folders in the current user's mail
        public async Task<ActionResult> GetMyMailfolders()
        {
            ResultsViewModel results = new ResultsViewModel();
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Get the folders.
                results.Items = await mailService.GetMyMailFolders(graphClient);

                foreach (var folder in results.Items)
                {
                    AddFolders(new FolderItem
                    {
                        Name = folder.Display,
                        Id = folder.Id
                    });
                }
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        public async Task<ActionResult> SyncFolders()
        {
            ResultsViewModel results = new ResultsViewModel();
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Get the messages.
                results.Items = await mailService.GetMyInboxMessages(graphClient);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        // Get messages in the current user's inbox.
        public async Task<ActionResult> GetMyInboxMessages()
        {
            ResultsViewModel results = new ResultsViewModel();
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Get the messages.
                results.Items = await mailService.GetMyInboxMessages(graphClient);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }


        // Get messages with attachments in the current user's inbox.
        public async Task<ActionResult> GetMyInboxMessagesThatHaveAttachments()
        {
            ResultsViewModel results = new ResultsViewModel();
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Get messages in the Inbox folder that have file attachments.
                results.Items = await mailService.GetMyInboxMessagesThatHaveAttachments(graphClient);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        // Send an email message.
        // This snippet sends a message to the current user on behalf of the current user.
        public async Task<ActionResult> SendMessage()
        {
            ResultsViewModel results = new ResultsViewModel(false);
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Send the message.
                results.Items = await mailService.SendMessage(graphClient);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        // Send an email message with a file attachment.
        // This snippet sends a message to the current user on behalf of the current user.
        public async Task<ActionResult> SendMessageWithAttachment()
        {
            ResultsViewModel results = new ResultsViewModel(false);
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Send the message.
                results.Items = await mailService.SendMessageWithAttachment(graphClient);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        // Get a specified message.
        public async Task<ActionResult> GetMessage(string id)
        {
            ResultsViewModel results = new ResultsViewModel();
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Get the message.
                results.Items = await mailService.GetMessage(graphClient, id);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        // Reply to a specified message.
        public async Task<ActionResult> ReplyToMessage(string id)
        {
            ResultsViewModel results = new ResultsViewModel(false);
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                results.Items = await mailService.ReplyToMessage(graphClient, id);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        // Move a specified message. This creates a new copy of the message in the destination folder.
        // This snippet moves the message to the Drafts folder.
        public async Task<ActionResult> MoveMessage(string id)
        {
            ResultsViewModel results = new ResultsViewModel();
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Move the message.
                results.Items = await mailService.MoveMessage(graphClient, id);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

        // Delete a specified message.
        public async Task<ActionResult> DeleteMessage(string id)
        {
            ResultsViewModel results = new ResultsViewModel(false);
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Delete the message.
                results.Items = await mailService.DeleteMessage(graphClient, id);
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format(Resource.Error_Message, Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }
    }
}