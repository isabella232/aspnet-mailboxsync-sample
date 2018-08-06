using MailBoxSync.Models.Subscription;

namespace MailboxSync.Models.Subscription
{
    // The data that displays in the Notification view.
    public class MessageViewModel
    {
        public MessageItem Message { get; set; }

        // The ID of the user associated with the subscription.
        // Used to filter messages to display in the client.
        public string SubscribedUser { get; set; }

        public MessageViewModel(MessageItem message, string subscribedUserId)
        {
            Message = message;
            SubscribedUser = subscribedUserId;
        }

    }
}