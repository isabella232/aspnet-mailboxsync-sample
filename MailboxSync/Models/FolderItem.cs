/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

using System.Collections.Generic;

namespace MailboxSync.Models
{
    /// <summary>
    /// An Outlook folder (partial representation) with extra properties
    /// There are a lot of properties returned about the folder from the requests that you might not need all of them. Pick only the things that you need.
    /// See https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/mailfolder
    /// </summary>
    public class FolderItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ParentId { get; set; }
        public List<MessageItem> MessageItems { get; set; }
        public int? SkipToken { get; set; }

        // used to flag a folder so that it can be displayed first in the list
        public bool StartupFolder { get; set; }

    }
}