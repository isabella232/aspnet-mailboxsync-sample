/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 *  See LICENSE in the source repository root for complete license information.
 */

using System;
using Newtonsoft.Json;

namespace MailboxSync.Models
{
    /// <summary>
    /// An Outlook mail message (partial representation). 
    /// There are a lot of properties returned about the message from the requests. You do not need all of them. Pick only the things that you need
    /// See https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/message
    /// </summary>
    public class MessageItem
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "subject")]
        public string Subject { get; set; }

        [JsonProperty(PropertyName = "bodyPreview")]
        public string BodyPreview { get; set; }

        [JsonProperty(PropertyName = "createdDateTime")]
        public DateTimeOffset CreatedDateTime { get; set; }

        [JsonProperty(PropertyName = "isRead")]
        public Boolean IsRead { get; set; }

        [JsonProperty(PropertyName = "conversationId")]
        public string ConversationId { get; set; }

        [JsonProperty(PropertyName = "changeKey")]
        public string ChangeKey { get; set; }
    }

}

