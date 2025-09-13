using Client.Models;
using Microsoft.AspNetCore.Components;
using Microsoft.Extensions.Logging;
using Microsoft.JSInterop;
using Radzen;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;


namespace Client.Services
{
    public class ValidationService
    {
        private readonly HttpClient _httpClient;
        private readonly NotificationService _notificationService;
        private readonly string _apiBaseUrl;
        private readonly ILogger<ValidationService> _logger;

        public ValidationService(HttpClient httpClient, NotificationService notificationService, ConfigurationService configService, ILogger<ValidationService> logger)
        {
            _httpClient = httpClient;
            _notificationService = notificationService;
            _apiBaseUrl = configService.GetApiBaseUrl();
            _logger = logger;


        }
     

        public async Task<ValidationResponse?> ValidateBankFileAsync()
        {
            string baiFileId=null;

            if (string.IsNullOrEmpty(baiFileId))
            {
                _notificationService.Notify(NotificationSeverity.Warning, "Validation", "No BAI File ID found.", 4000);
                return new ValidationResponse { Message = "No BAI File ID found.", Errors = new List<string> { "Missing baiFileId" } };
            }

            string apiUrl = $"{_apiBaseUrl}/api/validation/{baiFileId}";
            _logger.LogInformation($"Starting validation request to: {apiUrl}");
            try
            {
                var response = await _httpClient.PostAsJsonAsync(apiUrl, new { });

                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadFromJsonAsync<ValidationResponse>();

                    _notificationService.Notify(NotificationSeverity.Success, "Validation", "Validation completed successfully.", 4000);

                    return result;
                }
                else
                {
                    var errorResult = await response.Content.ReadFromJsonAsync<ValidationResponse>();

                    _notificationService.Notify(NotificationSeverity.Error, "Validation", "Validation failed.", 5000);

                    return errorResult;
                }
            }
            catch (Exception ex)
            {
                _notificationService.Notify(NotificationSeverity.Error, "Validation", $"Error: {ex.Message}", 5000);
                return new ValidationResponse { Message = $"Error: {ex.Message}", Errors = new List<string> { ex.Message } };
            }
        }
    }
}
