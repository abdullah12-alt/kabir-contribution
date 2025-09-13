using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using Client.Models;
using Microsoft.Extensions.Logging;
using Radzen;

namespace Client.Services
{
    public class InstitutionService
    {
        private readonly HttpClient _http;
        private readonly NotificationService _notificationService;
        private readonly string _apiBaseUrl;
        private readonly ILogger<InstitutionService> _logger;

        public InstitutionService(HttpClient http,
                                  NotificationService notificationService,
                                  ConfigurationService configService,
                                  ILogger<InstitutionService> logger)
        {
            _http = http;
            _notificationService = notificationService;
            _apiBaseUrl = configService.GetApiBaseUrl();
            _logger = logger;
        }

        public async Task<List<Institution>> GetAllAsync()
        {
            try
            {
                var result = await _http.GetFromJsonAsync<List<Institution>>($"{_apiBaseUrl}/api/institutions");
                _logger.LogInformation("Fetched institutions");
                return result ?? new();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching institutions");
                _notificationService.Notify(NotificationSeverity.Error, "Institutions", "Failed to load data.", 4000);
                return new();
            }
        }

        public async Task<Institution?> GetByIdAsync(long id)
        {
            try
            {
                var result = await _http.GetFromJsonAsync<Institution>($"{_apiBaseUrl}/api/institutions/{id}");
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error fetching institution ID={id}");
                _notificationService.Notify(NotificationSeverity.Error, "Institutions", "Failed to load detail.", 4000);
                return null;
            }
        }

        public async Task<bool> CreateOrUpdateAsync(Institution model, string updateMode, string userId)
        {
            try
            {

                var response = await _http.PostAsJsonAsync(
                    $"{_apiBaseUrl}/api/institutions?updateMode={updateMode}&userId={userId}", model);

                if (response.IsSuccessStatusCode)
                {
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving institution");
                _notificationService.Notify(NotificationSeverity.Error, "Institutions", $"Error: {ex.Message}", 5000);
                return false;
            }
        }

        public async Task<bool> DeleteAsync(long id, string userId)
        {
            try
            {
                var response = await _http.DeleteAsync($"{_apiBaseUrl}/api/institutions/{id}?userId={userId}");
                if (response.IsSuccessStatusCode)
                {
                    _notificationService.Notify(NotificationSeverity.Success, "Institutions", "Deleted successfully.", 3000);
                    return true;
                }

                _notificationService.Notify(NotificationSeverity.Warning, "Institutions", "Delete failed.", 4000);
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error deleting institution ID={id}");
                _notificationService.Notify(NotificationSeverity.Error, "Institutions", $"Error: {ex.Message}", 5000);
                return false;
            }
        }
    }
}
