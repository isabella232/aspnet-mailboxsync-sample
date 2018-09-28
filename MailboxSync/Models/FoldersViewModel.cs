/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System.Collections.Generic;
using System.Linq;

namespace MailboxSync.Models
{

    public class FolderMessages
    {
        public FolderMessages()
        {
            Messages = new List<MessageItem>();
        }
        public List<MessageItem> Messages { get; set; }


        // This token to be used for pagination
        // It keeps record of how many items to skip in order to get the next set of items
        public int? SkipToken { get; set; }
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


}