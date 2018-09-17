/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using MailboxSync.Helpers;
using MailboxSync.Models;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Mvc;
using MailboxSync.Services;


namespace MailboxSync.Controllers
{

    [Authorize]
    public class HomeController : Controller
    {
        MailService mailService = new MailService();
        DataService dataService = new DataService();

        public ActionResult Index()
        {
            var results = new FoldersViewModel();
            var folders = dataService.GetFolders();
            var resultItems = new List<FolderItem>();
            resultItems.AddRange(folders);
            results.Items = resultItems;
            return View("Index", results);
        }


        /// <summary>
        /// Get folders in the current user's mail
        /// </summary>
        public async Task<ActionResult> GetMyMailfolders()
        {
            var results = new FoldersViewModel();
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = GraphSdkHelper.GetAuthenticatedClient();

                // Get the folders.
                results.Items = await mailService.GetMyMailFolders(graphClient);

                foreach (var folder in results.Items)
                {
                    if (dataService.FolderExists(folder.Id))
                    {
                        dataService.StoreMessage(folder.MessageItems, folder.Id, folder.SkipToken);
                    }
                    else
                    {
                        dataService.StoreFolder(folder);
                    }

                }
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "AuthenticationFailure")
                {
                    return new EmptyResult();
                }

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format("Error in {0}: {1} {2}", Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return RedirectToAction("Index");
        }

        /// <summary>
        /// Send an email message.
        /// This sends a message to the current user on behalf of the current user.
        /// </summary>
        public async Task<ActionResult> SendMessage()
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = GraphSdkHelper.GetAuthenticatedClient();

                // Send the message.
                await mailService.SendMessage(graphClient);
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "AuthenticationFailure")
                {
                    return new EmptyResult();
                }

                return RedirectToAction("Index", "Error", new { message = string.Format("Error in {0}: {1} {2}", Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return RedirectToAction("Index");
        }

        /// <summary>
        /// Gets the paged messages belonging to a folder using a skip token
        /// </summary>
        /// <param name="folderId">the folder whose message is to be fetched</param>
        /// <param name="skipToken">the skip token that indicates how many items to skip when fetching the messages</param>
        public async Task<ActionResult> GetPagedMessages(string folderId, int? skipToken)
        {
            try
            {
                GraphServiceClient graphClient = GraphSdkHelper.GetAuthenticatedClient();
                var messages = await mailService.GetMyFolderMessages(graphClient, folderId, skipToken);
                if (messages.Messages.Count > 0)
                {
                    dataService.StoreMessage(messages.Messages, folderId, messages.SkipToken);
                }
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "AuthenticationFailure")
                {
                    return new EmptyResult();
                }

                return RedirectToAction("Index", "Error", new { message = string.Format("Error in {0}: {1} {2}", Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return RedirectToAction("Index");
        }


    }
}