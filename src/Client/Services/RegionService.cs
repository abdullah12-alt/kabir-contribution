using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using Client.Models;
using Microsoft.Extensions.Logging;
using Radzen;

namespace Client.Services
{
    public class RegionService
    {
        private readonly HttpClient _http;
        private readonly NotificationService _notificationService;
        private readonly string _apiBaseUrl;
        private readonly ILogger<RegionService> _logger;

        public RegionService(HttpClient http,
                             NotificationService notificationService,
                             ConfigurationService configService,
                             ILogger<RegionService> logger)
        {
            _http = http;
            _notificationService = notificationService;
            _apiBaseUrl = configService.GetApiBaseUrl();
            _logger = logger;
        }

        public async Task<List<RegionDto>> GetAllAsync()
        {
            try
            {
                var result = await _http.GetFromJsonAsync<List<RegionDto>>($"{_apiBaseUrl}/api/region");
                _logger.LogInformation("Fetched regions");
                return result ?? new();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching regions");
                _notificationService.Notify(NotificationSeverity.Error, "Regions", "Failed to load data.", 4000);
                return new();
            }
        }

        public async Task<RegionDto?> GetByIdAsync(long id)
        {
            try
            {
                return await _http.GetFromJsonAsync<RegionDto>($"{_apiBaseUrl}/api/region/{id}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error fetching region ID={id}");
                _notificationService.Notify(NotificationSeverity.Error, "Regions", "Failed to load detail.", 4000);
                return null;
            }
        }

        public async Task<bool> CreateAsync(RegionDto model)
        {
            try
            {
                var response = await _http.PostAsJsonAsync($"{_apiBaseUrl}/api/region", model);
                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating region");
                _notificationService.Notify(NotificationSeverity.Error, "Regions", $"Error: {ex.Message}", 5000);
                return false;
            }
        }

        public async Task<bool> UpdateAsync(RegionDto model)
        {
            try
            {
                var response = await _http.PutAsJsonAsync($"{_apiBaseUrl}/api/region/{model.RegionId}", model);
                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error updating region ID={model.RegionId}");
                _notificationService.Notify(NotificationSeverity.Error, "Regions", $"Error: {ex.Message}", 5000);
                return false;
            }
        }

        public async Task<bool> DeleteAsync(long id)
        {
            try
            {
                var response = await _http.DeleteAsync($"{_apiBaseUrl}/api/region/{id}");
                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error deleting region ID={id}");
                _notificationService.Notify(NotificationSeverity.Error, "Regions", $"Error: {ex.Message}", 5000);
                return false;
            }
        }
    }
}
