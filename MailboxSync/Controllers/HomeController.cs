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
        GraphServiceClient graphClient = GraphServiceClientProvider.GetAuthenticatedClient();

        public ActionResult Index()
        {
            var folderResults = new FoldersViewModel();
            var folders = dataService.GetFolders();
            var resultItems = new List<FolderItem>();
            resultItems.AddRange(folders);
            folderResults.Items = resultItems;
            return View("Index", folderResults);
        }


        /// <summary>
        /// Get folders in the current user's mail
        /// </summary>
        public async Task<ActionResult> GetMyMailfolders()
        {
            try
            {
                // Get the folders.
                var folders = await mailService.GetMyMailFolders(graphClient);

                foreach (var folder in folders)
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

                return RedirectToAction("Index", "Error", new
                {
                    message =
                    $"Error in {Request.RawUrl}: {se.Error.Code} {se.Error.Message}"
                });
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
                // Send the message.
                await mailService.SendMessage(graphClient);
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "AuthenticationFailure")
                {
                    return new EmptyResult();
                }

                return RedirectToAction("Index", "Error", new
                {
                    message =
                    $"Error in {Request.RawUrl}: {se.Error.Code} {se.Error.Message}"
                });
            }
            return RedirectToAction("Index");
        }

        /// <summary>
        /// Gets the paged messages belonging to a folder using a skip token
        /// </summary>
        /// <param name="folderId">the folder whose message is to be fetched</param>
        /// <param name="skipToken">the skip token that indicates how many items to skip </param>
        public async Task<ActionResult> GetPagedMessages(string folderId, int? skipToken)
        {
            try
            {
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

                return RedirectToAction("Index", "Error", new
                {
                    message =
                    $"Error in {Request.RawUrl}: {se.Error.Code} {se.Error.Message}"
                });
            }
            return RedirectToAction("Index");
        }


    }
}