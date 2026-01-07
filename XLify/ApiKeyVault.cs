using System;
using System.Security.Cryptography;
using System.Text;

namespace XLify
{
    internal static class ApiKeyVault
    {
        public static bool Has()
        {
            var b64 = Properties.Settings.Default.ApiKeyCipher;
            return !string.IsNullOrEmpty(b64);
        }

        public static void Save(string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey)) throw new ArgumentException("Empty key");
            var cipher = Protect(apiKey);
            Properties.Settings.Default.ApiKeyCipher = Convert.ToBase64String(cipher);
            Properties.Settings.Default.Save();
        }

        public static void Clear()
        {
            Properties.Settings.Default.ApiKeyCipher = string.Empty;
            Properties.Settings.Default.Save();
        }

        public static string Get()
        {
            var b64 = Properties.Settings.Default.ApiKeyCipher;
            if (string.IsNullOrEmpty(b64)) return null;
            var bytes = Convert.FromBase64String(b64);
            return Unprotect(bytes);
        }

        private static byte[] Protect(string plaintext)
        {
            var data = Encoding.UTF8.GetBytes(plaintext);
            return ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
        }

        private static string Unprotect(byte[] cipher)
        {
            var data = ProtectedData.Unprotect(cipher, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(data);
        }
    }
}
