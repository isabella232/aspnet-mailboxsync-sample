/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System.Runtime.Caching;
using System.Threading;
using Microsoft.Identity.Client;

namespace MailboxSync.TokenStorage
{
    /// <summary>
    /// The SessionTokenCache is an example demonstration for managing tokens. 
    /// You may want to implement a token cache that conforms to the security policies of your organization.
    /// </summary>
    public class SessionTokenCache
    {
        private static ReaderWriterLockSlim sessionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);
        private readonly string _cacheId;
        private static ObjectCache cache = MemoryCache.Default;
        private static CacheItemPolicy defaultPolicy = new CacheItemPolicy();

        TokenCache tokenCache = new TokenCache();

        public SessionTokenCache(string userId)
        {
            _cacheId = userId + "_TokenCache";
            Load();
        }

        /// <summary>
        /// Reads information from the cache
        /// </summary>
        /// <returns>TokenCache object</returns>
        public TokenCache GetMsalCacheInstance()
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
            Load();
            return tokenCache;
        }

        /// <summary>
        /// check whether the cache has information
        /// </summary>
        /// <returns></returns>
        public bool HasData()
        {
            return (cache[_cacheId] != null && ((byte[])cache[_cacheId]).Length > 0);
        }

        /// <summary>
        /// clear the cache
        /// </summary>
        public void Clear()
        {
            cache.Remove(_cacheId);
        }


        private void Load()
        {
            sessionLock.EnterReadLock();
            var item = cache.GetCacheItem(_cacheId);
            if (item != null)
            {
                tokenCache.Deserialize((byte[])item.Value);
            }
            sessionLock.ExitReadLock();
        }

        /// <summary>
        /// Saves the cache information
        /// </summary>
        private void Persist()
        {
            sessionLock.EnterWriteLock();

            // Optimistically set HasStateChanged to false. 
            // We need to do it early to avoid losing changes made by a concurrent thread.
            tokenCache.HasStateChanged = false;

            cache.Set(new CacheItem(_cacheId, tokenCache.Serialize()), defaultPolicy);

            sessionLock.ExitWriteLock();
        }

        /// <summary>
        /// Triggered right before MSAL needs to access the cache. 
        /// </summary>
        /// <param name="tokenCaaCacheNotificationArgs"></param>
        private void BeforeAccessNotification(TokenCacheNotificationArgs tokenCaaCacheNotificationArgs)
        {
            // Reload the cache from the persistent store in case it changed since the last access. 
            Load();
        }

        /// <summary>
        /// Triggered right after MSAL accessed the cache.
        /// </summary>
        /// <param name="tokenCaaCacheNotificationArgs"></param>
        private void AfterAccessNotification(TokenCacheNotificationArgs tokenCaaCacheNotificationArgs)
        {
            // if the access operation resulted in a cache update
            if (tokenCache.HasStateChanged)
            {
                Persist();
            }
        }
    }
}