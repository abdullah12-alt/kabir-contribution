using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using Client.Models;
using Microsoft.Extensions.Logging;
using Radzen;
namespace Client.Services
{
    public class DDConfigService
    {
        private readonly HttpClient _http;
        private readonly NotificationService _notificationService;
        private readonly string _apiBaseUrl;
        private readonly ILogger<DDConfigService> _logger;
        public DDConfigService(HttpClient http,
                               NotificationService notificationService,
                               ConfigurationService configService,
                               ILogger<DDConfigService> logger)
        {
            _http = http;
            _notificationService = notificationService;
            _apiBaseUrl = configService.GetApiBaseUrl();
            _logger = logger;
        }
        public async Task<DDConfigInfo?> GetConfigAsync()
        {
            try
            {
                var result = await _http.GetFromJsonAsync<DDConfigInfo>($"{_apiBaseUrl}/api/ddconfig");
                _logger.LogInformation("Fetched DDConfigInfo");
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching DDConfigInfo");
                _notificationService.Notify(NotificationSeverity.Error, "Configuration", "Failed to load configuration.", 4000);
                return null;
            }
        }

        public async Task<bool> UpdateConfigAsync(DDConfigInfo model, string updateMode)
        {
            try
            {
                var response = await _http.PutAsJsonAsync($"{_apiBaseUrl}/api/ddconfig?updateMode={updateMode}", model);

             
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating DDConfigInfo");
                _notificationService.Notify(NotificationSeverity.Error, "Configuration", $"Error: {ex.Message}", 5000);
                return false;
            }
        }
    

}
}
