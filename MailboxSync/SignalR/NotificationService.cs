using Microsoft.AspNet.SignalR;

namespace MailboxSync.SignalR
{
    public class NotificationService
    {
        public void SendNotificationToClient(int p0)
        {
            var hubContext = GlobalHost.ConnectionManager.GetHubContext<NotificationHub>();
            if (hubContext != null)
            {
                hubContext.Clients.All.showNotification(p0);
            }
        }
    }
}