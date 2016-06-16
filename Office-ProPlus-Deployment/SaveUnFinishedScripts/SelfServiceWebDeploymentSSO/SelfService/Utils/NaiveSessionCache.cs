#region

using System.Web;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

#endregion

namespace SelfService.Utils
{
    public class NaiveSessionCache : TokenCache
    {
        private static readonly object FileLock = new object();
        private readonly string CacheId = string.Empty;
        private string UserObjectId = string.Empty;

        public NaiveSessionCache(string userId)
        {
            UserObjectId = userId;
            CacheId = UserObjectId + "_TokenCache";

            AfterAccess = AfterAccessNotification;
            BeforeAccess = BeforeAccessNotification;
            Load();
        }

        public void Load()
        {
            lock (FileLock)
            {
                if (HttpContext.Current != null)
                {
                    Deserialize((byte[]) HttpContext.Current.Session[CacheId]);
                }
            }
        }

        public void Persist()
        {
            lock (FileLock)
            {
                // reflect changes in the persistent store
                HttpContext.Current.Session[CacheId] = Serialize();
                // once the write operation took place, restore the HasStateChanged bit to false
                HasStateChanged = false;
            }
        }

        // Empties the persistent store.
        public override void Clear()
        {
            base.Clear();
            HttpContext.Current.Session.Remove(CacheId);
        }

        public override void DeleteItem(TokenCacheItem item)
        {
            base.DeleteItem(item);
            Persist();
        }

        // Triggered right before ADAL needs to access the cache.
        // Reload the cache from the persistent store in case it changed since the last access.
        private void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            Load();
        }

        // Triggered right after ADAL accessed the cache.
        private void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (HasStateChanged)
            {
                Persist();
            }
        }
    }
}