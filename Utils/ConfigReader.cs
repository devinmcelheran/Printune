using System;
using System.IO;
using System.Linq;
using System.Net.Http;

namespace Printune
{
    public static class ConfigReader
    {
        public static string Read(string ConfigPath)
        {
            ConfigPath = ExpandIfRelative(ConfigPath);

            var configUri = new Uri(ConfigPath);
            switch (configUri.Scheme)
            {
                case "http":
                    return ReadHttp(configUri);
                case "https":
                    return ReadHttp(configUri);
                case "file":
                    return ReadFile(configUri);
                default:
                    throw new Invocation.InvalidNameOrPathException("Invalid invocation: The provided configuration file path is not a valid HTTP, HTTPS, or file path.");
            }
        }
        public static byte[] ReadAsBytes(string ConfigPath)
        {
            ConfigPath = ExpandIfRelative(ConfigPath);

            var configUri = new Uri(ConfigPath);
            // Allow filesystem paths as well because they could be
            // be on shared drives with UNC paths.
            switch (configUri.Scheme)
            {
                case "http":
                    return ReadHttpAsBytes(configUri);
                case "https":
                    return ReadHttpAsBytes(configUri);
                case "file":
                    return ReadFileAsBytes(configUri);
                default:
                    throw new Invocation.InvalidNameOrPathException("Invalid invocation: The provided configuration file path is not a valid HTTP, HTTPS, or file path.");
            }
        }
        /// <summary>
        /// Copies the configuration file to a local temporary file and returns the local file path.
        /// If it is already a local file, the original file path is returned, but fully qualified.
        /// </summary>
        /// <param name="ConfigPath"></param>
        /// <returns>Fully qualified local file path.</returns>
        public static string CopyToLocal(string ConfigPath)
        {
            ConfigPath = ExpandIfRelative(ConfigPath);
            var configUri = new Uri(ConfigPath);

            if (configUri.IsFile && !configUri.IsUnc)
                return configUri.LocalPath;

            var filePath = Path.GetTempFileName();
            File.WriteAllBytes(filePath, ReadAsBytes(ConfigPath));
            return filePath;
        }
        public static string ExpandIfRelative(string ConfigPath)
        {
            if (File.Exists(ConfigPath))
                return Path.GetFullPath(ConfigPath);
            else
                return ConfigPath;
        }

        private static string ReadFile(Uri ConfigPath)
        {
            return File.ReadAllText(ConfigPath.OriginalString);
        }
        private static string ReadHttp(Uri ConfigPath)
        {
            var bytes = ReadHttpAsBytes(ConfigPath);
            return System.Text.Encoding.UTF8.GetString(bytes);
        }
        private static byte[] ReadHttpAsBytes(Uri ConfigPath)
        {
            using (var client = new HttpClient())
            {
                var response = client.GetAsync(ConfigPath).Result;

                if (response.IsSuccessStatusCode)
                {
                    return response.Content.ReadAsByteArrayAsync().Result;
                }
                else
                {
                    throw new HttpRequestException($"HTTP request of \"{ConfigPath.OriginalString}\" failed with HTTP status code {response.StatusCode}.");
                }
            }
        }
        private static byte[] ReadFileAsBytes(Uri ConfigPath)
        {
            return File.ReadAllBytes(ConfigPath.OriginalString);
        }
    }
}