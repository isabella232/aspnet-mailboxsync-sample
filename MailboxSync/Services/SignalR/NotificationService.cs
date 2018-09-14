using Microsoft.AspNet.SignalR;

namespace MailboxSync.Services.SignalR
{
    /// <summary>
    /// Links the notifications to the UI using signalR
    /// </summary>
    public class NotificationService
    {
        /// <summary>
        /// Fires the notification to the client
        /// </summary>
        /// <param name="notificationsCount">Number of notifications</param>
        public void SendNotificationToClient(int notificationsCount)
        {
            var hubContext = GlobalHost.ConnectionManager.GetHubContext<NotificationHub>();
            if (hubContext != null)
            {
                hubContext.Clients.All.showNotification(notificationsCount);
            }
        }
    }
}