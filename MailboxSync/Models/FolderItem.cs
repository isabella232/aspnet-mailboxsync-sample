using System.Collections.Generic;

namespace MailboxSync.Models
{
    public class FolderItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ParentId { get; set; }
        public List<MessageItem> MessageItems { get; set; }
        public int? SkipToken { get; set; }

    }
}