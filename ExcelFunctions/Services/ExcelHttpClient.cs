using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelFunctions.Services
{
    internal class ExcelHttpClient
    {
        private readonly HttpClient _httpClient;
    
        public ExcelHttpClient(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<string> Get(string uri)
        {
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, uri);
            HttpResponseMessage response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();

            string responseText = "";
            using (StreamReader reader = new StreamReader(await response.Content.ReadAsStreamAsync()))
            {
                responseText = reader.ReadToEnd();
                return responseText;
            }
        }
    }
}
