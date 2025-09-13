using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using Client.Models;
using Microsoft.Extensions.Logging;
using Radzen;

namespace Client.Services
{
    public class IncomeSourceTypeService
    {
        private readonly HttpClient _http;
        private readonly NotificationService _notificationService;
        private readonly string _apiBaseUrl;
        private readonly ILogger<IncomeSourceTypeService> _logger;

        public IncomeSourceTypeService(HttpClient http,
                                       NotificationService notificationService,
                                       ConfigurationService configService,
                                       ILogger<IncomeSourceTypeService> logger)
        {
            _http = http;
            _notificationService = notificationService;
            _apiBaseUrl = configService.GetApiBaseUrl();
            _logger = logger;
        }

        public async Task<List<IncomeSourceType>> GetAllAsync()
        {
            try
            {
                var result = await _http.GetFromJsonAsync<List<IncomeSourceType>>($"{_apiBaseUrl}/api/incomesourcetypes");
                _logger.LogInformation("Fetched income source types");
                return result ?? new();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching income source types");
                _notificationService.Notify(NotificationSeverity.Error, "Income Source Types", "Failed to load data.", 4000);
                return new();
            }
        }

        public async Task<IncomeSourceType?> GetByIdAsync(int id)
        {
            try
            {
                var result = await _http.GetFromJsonAsync<IncomeSourceType>($"{_apiBaseUrl}/api/incomesourcetypes/{id}");
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error fetching income source type ID={id}");
                _notificationService.Notify(NotificationSeverity.Error, "Income Source Types", "Failed to load detail.", 4000);
                return null;
            }
        }

        public async Task<bool> CreateOrUpdateAsync(IncomeSourceType model, string updateMode, string userId)
        {
            try
            {
                var response = await _http.PostAsJsonAsync(
                    $"{_apiBaseUrl}/api/incomesourcetypes?updateMode={updateMode}&userId={userId}", model);

                if (response.IsSuccessStatusCode)
                {
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                _notificationService.Notify(NotificationSeverity.Error, "Income Source Types", $"Error: {ex.Message}", 5000);
                return false;
            }
        }

        public async Task<bool> DeleteAsync(long id, string userId)
        {
            try
            {
                var response = await _http.DeleteAsync($"{_apiBaseUrl}/api/incomesourcetypes/{id}?userId={userId}");
                if (response.IsSuccessStatusCode)
                {
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error deleting income source type ID={id}");
                _notificationService.Notify(NotificationSeverity.Error, "Income Source Types", $"Error: {ex.Message}", 5000);
                return false;
            }
        }
    }
}
