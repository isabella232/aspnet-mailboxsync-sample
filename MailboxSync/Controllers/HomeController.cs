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


namespace MailboxSync.Controllers
{

    [Authorize]
    public class HomeController : Controller
    {
        MailService mailService = new MailService();

        public async Task<ActionResult> AddMessages(string id)
        {
            var results = new FoldersViewModel();
            var dataService = new DataService();
            try
            {
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();
                var messages = await mailService.GetMyFolderMessages(graphClient, id);
                if (messages.Count > 0)
                {
                    dataService.StoreMessage(messages, id);
                }
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == "Caller needs to authenticate.")
                {
                    return new EmptyResult();
                }

                return RedirectToAction("Index", "Error", new { message = string.Format("Error in {0}: {1} {2}", Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);

        }

        public ActionResult Index()
        {
            var results = new FoldersViewModel();
            var dataService = new DataService();
            var folders = dataService.GetFolders();
            var resultItems = new List<FolderItem>();
            resultItems.AddRange(folders);
            results.Items = resultItems;
            return View("Index", results);
        }


        // Get folders in the current user's mail
        public async Task<ActionResult> GetMyMailfolders()
        {
            var results = new FoldersViewModel();
            var dataService = new DataService();
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient();

                // Get the folders.
                results.Items = await mailService.GetMyMailFolders(graphClient);

                foreach (var folder in results.Items)
                {
                    dataService.StoreFolder(folder);
                }
            }
            catch (ServiceException se)
            {
                if (se.Error.Message == "Caller needs to authenticate.")
                {
                    return new EmptyResult();
                }

                // Personal accounts that aren't enabled for the Outlook REST API get a "MailboxNotEnabledForRESTAPI" or "MailboxNotSupportedForRESTAPI" error.
                return RedirectToAction("Index", "Error", new { message = string.Format("Error in {0}: {1} {2}", Request.RawUrl, se.Error.Code, se.Error.Message) });
            }
            return View("Index", results);
        }

    }
}