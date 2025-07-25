using System;
using System.IO;
using System.Net.Http;

namespace Printune
{
    public static class ConfigReader
    {
        public static string Read(string ConfigPath)
        {
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

        private static string ReadFile(Uri ConfigPath)
        {
            return File.ReadAllText(ConfigPath.OriginalString);
        }
        private static string ReadHttp(Uri ConfigPath)
        {
            using (var client = new HttpClient())
            {
                var response = client.GetAsync(ConfigPath).Result;

                if (response.IsSuccessStatusCode)
                {
                    return response.Content.ReadAsStringAsync().Result;
                }
                else
                {
                    if (response.Content == null)
                        throw new HttpRequestException($"HTTP request of \"{ConfigPath.OriginalString}\" failed with HTTP status code {response.StatusCode}.");
                    else
                        throw new HttpRequestException($"HTTP request of \"{ConfigPath.OriginalString}\" failed with HTTP status code {response.StatusCode}"
                                                        + $" and reponse of \n \"{response.Content.ReadAsStringAsync()}\"");
                }
            }
        }
    }
}