/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using MailboxSync.Models.Subscription;
using MailBoxSync.Models.Subscription;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Graph;

namespace MailboxSync.Models
{

    public class FolderItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ParentId { get; set; }
        public List<MessageItem> Messages { get; set; }
        public int? SkipToken { get; set; }

    }

    // An entity, such as a user, group, or message.
    public class ResultItem
    {

        // The ID and display name for the entity's radio button.
        public string Id { get; set; }
        public string Display { get; set; }

        // The properties of an entity that display in the UI.
        public Dictionary<string, object> Properties;

        public ResultItem()
        {
            Properties = new Dictionary<string, object>();
        }
    }

    // View model to display a collection of one or more entities returned from the Microsoft Graph. 
    public class ResultsViewModel
    {

        // Set to false if you don't want to display radio buttons with the results.
        public bool Selectable { get; set; }

        // The list of entities to display.
        public IEnumerable<FolderItem> Items { get; set; }
        public ResultsViewModel(bool selectable = true)
        {

            // Indicates whether the results should display radio buttons.
            // This is how an entity ID is passed to methods that require it.
            Selectable = selectable;

            Items = Enumerable.Empty<FolderItem>();
        }
    }

    public class FoldersViewModel
    {

        // Set to false if you don't want to display radio buttons with the results.
        public bool Selectable { get; set; }

        // The list of entities to display.
        public IEnumerable<FolderItem> Items { get; set; }
        public FoldersViewModel(bool selectable = true)
        {

            // Indicates whether the results should display radio buttons.
            // This is how an entity ID is passed to methods that require it.
            Selectable = selectable;

            Items = Enumerable.Empty<FolderItem>();
        }
    }

    public class FolderMessage
    {
        public FolderMessage()
        {
            Messages = new List<MessageItem>();
        }
        public List<MessageItem> Messages { get; set; }


        // This token to be used for pagination
        // It keeps record of how many items to skip in order to get the next set of items
        public int? SkipToken { get; set; }
    }
}