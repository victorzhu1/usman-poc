using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Excel_POC.Services
    {
    public class SaveDataApi
        {
        private static readonly HttpClient client = new HttpClient();
        private static string  url = "https://your-api-endpoint.com/upload";
        /// <summary>
        /// Sends CSV data to a specified API endpoint and returns the response.
        /// </summary>
        /// <param name="csvData">The CSV data to send.</param>
        /// <returns>The API response as a string.</returns>
        public static async Task<string> SendDataToApiAsync(string csvData)
            {
            var content = new StringContent(csvData, Encoding.UTF8, "text/csv");

            HttpResponseMessage response = await client.PostAsync(url, content);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsStringAsync();
            }
        }
    }