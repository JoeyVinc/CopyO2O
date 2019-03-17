using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using System.IO;
using Microsoft.Win32;
using Commonfunctions.Cryptography;

namespace CopyO2O
{
    static class TokenCacheHelper
    {
        /// <summary>
        /// Path to the token cache
        /// </summary>
        private static string CacheFilePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\UserToken.cache";

        private static string _cipher = null;
        private static string GetCipher()
        {
            if (_cipher == null)
            {
                _cipher = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                    .OpenSubKey(@"SOFTWARE\Microsoft\Cryptography").GetValue("MachineGuid").ToString() + CacheFilePath;
            }
            return _cipher;
        }

        /// <summary>
        /// Get the user token cache
        /// </summary>
        /// <returns></returns>
        public static TokenCache GetUserCache()
        {
            if (usertokenCache == null)
            {
                usertokenCache = new TokenCache();
                usertokenCache.SetBeforeAccess(BeforeAccessNotification);
                usertokenCache.SetAfterAccess(AfterAccessNotification);
            }
            return usertokenCache;
        }

        static TokenCache usertokenCache;

        private static readonly object FileLock = new object();

        public static void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (FileLock)
            {
                if (File.Exists(CacheFilePath))
                {
                    byte[] decryptedbytes = SimpleAES.DecryptString(File.ReadAllBytes(CacheFilePath), GetCipher());
                    args.TokenCache.Deserialize(decryptedbytes);
                }
                else args.TokenCache.Deserialize(null);
            }
        }

        public static void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (args.HasStateChanged)
            {
                lock (FileLock)
                {
                    byte[] encryptedargs = SimpleAES.EncryptString(args.TokenCache.Serialize(), GetCipher());

                    // reflect changes in the persistent store
                    File.WriteAllBytes(CacheFilePath, encryptedargs);
                    // once the write operation takes place restore the HasStateChanged bit to false
                    //args.TokenCache.HasStateChanged = false;
                }
            }
        }
    }
}
