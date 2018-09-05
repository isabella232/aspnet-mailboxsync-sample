using System.Runtime.Caching;
using System.Threading;
using Microsoft.Identity.Client;

namespace MailboxSync.TokenStorage
{
    public class SessionTokenCache
    {
        private static ReaderWriterLockSlim sessionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);
        private string _cacheId = string.Empty;
        private static ObjectCache cache = MemoryCache.Default;
        private static CacheItemPolicy defaultPolicy = new CacheItemPolicy();

        TokenCache tokenCache = new TokenCache();

        public SessionTokenCache(string userId)
        {
            _cacheId = userId + "_TokenCache";
            Load();
        }

        public TokenCache GetMsalCacheInstance()
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
            Load();
            return tokenCache;
        }

        public bool HasData()
        {
            return (cache[_cacheId] != null && ((byte[])cache[_cacheId]).Length > 0);
        }

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

        private void Persist()
        {
            sessionLock.EnterWriteLock();

            // Optimistically set HasStateChanged to false. 
            // We need to do it early to avoid losing changes made by a concurrent thread.
            tokenCache.HasStateChanged = false;

            cache.Set(new CacheItem(_cacheId, tokenCache.Serialize()), defaultPolicy);

            sessionLock.ExitWriteLock();
        }

        // Triggered right before MSAL needs to access the cache. 
        private void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            // Reload the cache from the persistent store in case it changed since the last access. 
            Load();
        }

        // Triggered right after MSAL accessed the cache.
        private void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (tokenCache.HasStateChanged)
            {
                Persist();
            }
        }
    }
}