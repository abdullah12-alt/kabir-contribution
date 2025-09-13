using System.Net.Http.Json;

namespace Client.Services
{
    public class LookupApiService
    {
        private readonly HttpClient _http;
        private readonly string _apiBaseUrl;
        public LookupApiService(HttpClient http, ConfigurationService configService)
        {
            _http = http;
            _apiBaseUrl = configService.GetApiBaseUrl();
        }

        public async Task<List<string>> GetIncomeSourceTypesAsync()
        {
            var result = await _http.GetFromJsonAsync<List<string>>($"{_apiBaseUrl}/api/lookup/income-source-types");
            return result ?? new List<string>();
        }
    }
}
