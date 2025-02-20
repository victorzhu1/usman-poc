using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_POC.APIServices
    {
    private static readonly HttpClient client = new HttpClient();
    public class SaveData
        {
        }
   

    /// <summary>
    /// Sends CSV data to a specified API endpoint and returns the response.
    /// </summary>
    public static async Task<string> SendDataToApiAsync(string csvData)
        {
        var content = new StringContent(csvData, System.Text.Encoding.UTF8, "text/csv");

        HttpResponseMessage response = await client.PostAsync("https://your-api-endpoint.com/upload", content);
        response.EnsureSuccessStatusCode();

        return await response.Content.ReadAsStringAsync();
        }
    }
