using System;
using System.Collections;
using System.Collections.Generic;
using System.Web;

namespace MailboxSync.Services
{
    /// <summary>
    /// Stores details of the subscription in a cache for use later
    /// This info is required so the NotificationController can retrieve an access token from the cache and validate the subscription
    /// Production apps typically use some method of persistent storage
    /// </summary>
    public class SubscriptionStore
    {
        public string SubscriptionId { get; set; }
        public string ClientState { get; set; }
        public string UserId { get; set; }
        public string TenantId { get; set; }

        private SubscriptionStore(string subscriptionId, Tuple<string, string, string> parameters)
        {
            SubscriptionId = subscriptionId;
            ClientState = parameters.Item1;
            UserId = parameters.Item2;
            TenantId = parameters.Item3;
        }


        /// <summary>
        /// This sample temporarily stores the current subscription ID, client state, user object ID, and tenant ID.
        /// </summary>
        /// <param name="subscriptionId">received after the subscription has been created successfully</param>
        /// <param name="clientState">the state of the client at the time of creating the subscription</param>
        /// <param name="userId">the signed in user whose details authenticated the creation of the subscription</param>
        /// <param name="tenantId">the tenant id of the signed in user</param>
        public static void SaveSubscriptionInfo(string subscriptionId, string clientState, string userId, string tenantId)
        {
            HttpRuntime.Cache.Insert("subscriptionId_" + subscriptionId,
                Tuple.Create(clientState, userId, tenantId),
                null, DateTime.MaxValue, new TimeSpan(24, 0, 0), System.Web.Caching.CacheItemPriority.NotRemovable, null);
        }


        /// <summary>
        /// Retrieves the subscription from the cache bsed on the id of the subscription
        /// </summary>
        /// <param name="subscriptionId"></param>
        /// <returns></returns>
        public static SubscriptionStore GetSubscriptionInfo(string subscriptionId)
        {
            Tuple<string, string, string> subscriptionParams = HttpRuntime.Cache.Get("subscriptionId_" + subscriptionId) as Tuple<string, string, string>;
            return new SubscriptionStore(subscriptionId, subscriptionParams);
        }

        /// <summary>
        /// Delete all subscriptions
        /// </summary>
        public static void DeleteSubscriptionInfo()
        {
            List<string> keys = new List<string>();
            IDictionaryEnumerator enumerator = HttpRuntime.Cache.GetEnumerator();

            while (enumerator.MoveNext())
                if (enumerator.Key != null)
                    keys.Add(enumerator.Key.ToString());

            foreach (var item in keys)
                HttpRuntime.Cache.Remove(item);
        }

    }
}