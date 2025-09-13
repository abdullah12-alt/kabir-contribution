using Client.Models;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;

namespace Client.Services
{
    public class ExistsResponse
    {
        public bool Exists { get; set; }
    }
    public class FileloadService
    {
        private readonly HttpClient _httpClient;
        private readonly NotificationService _notificationService;
        private readonly string _apiBaseUrl;

        public FileloadService(HttpClient httpClient, NotificationService notificationService, ConfigurationService configService)
        {
            _httpClient = httpClient;
            _notificationService = notificationService;
            _apiBaseUrl = configService.GetApiBaseUrl();
        }
        public async Task<bool> LoadFilesAsync(Radzen.FileInfo baiFile, Radzen.FileInfo detailFile)
        {
            Console.WriteLine($"[FileloadService] Uploading: BaiFile={baiFile?.Name}, DetailFile={detailFile?.Name}");
            if (baiFile == null || detailFile == null)
            {
                _notificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No File Selected",
                    Detail = "Please upload a file before proceeding.",
                    Duration = 5000
                });

                return false;
            }

            try
            {
                var body = new MultipartFormDataContent();
                var baiStream = baiFile.OpenReadStream(5 * 1024 * 1024);
                var baiContent = new StreamContent(baiStream);
                baiContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

                var detailStream = detailFile.OpenReadStream(5 * 1024 * 1024);
                var detailContent = new StreamContent(detailStream);
                detailContent.Headers.ContentType = new MediaTypeHeaderValue("text/csv");

                body.Add(baiContent, "BaiFile", baiFile.Name);
                body.Add(detailContent, "DetailFile", detailFile.Name);


                var apiUrl = $"{_apiBaseUrl}/api/bank-files/load";
                Console.WriteLine($"[FileloadService] Uploading: BaiFile={baiFile?.Name}, DetailFile={detailFile?.Name}");

                var response = await _httpClient.PostAsync(apiUrl, body);



                var jsonResponse = await response.Content.ReadAsStringAsync();

                var parsedResponse = JsonSerializer.Deserialize<ApiResponse>(jsonResponse, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                string message = parsedResponse?.Message ?? "Unknown error occurred.";

                if (response.IsSuccessStatusCode && parsedResponse?.Success == true)
                {
                    var dataJson = parsedResponse.Data?.ToString();
                    if (!string.IsNullOrWhiteSpace(dataJson))
                    {
                        var dataElement = JsonSerializer.Deserialize<JsonElement>(dataJson);
                        if (dataElement.TryGetProperty("FileName", out var baiFileIdElement))
                        {
                            var baiFileId = baiFileIdElement.GetString();
                        
                        }
                    }
                    _notificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Success",
                        Detail = message,
                        Duration = 8000
                    });


                  

                    return true;
                }
                else
                {
                    _notificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "load Failed",
                        Detail = message,
                        Duration = 10000
                    });

                    return false;
                }
            }
            catch (Exception ex)
            {
                _notificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Error uploading files: {ex.Message}",
                    Duration = 10000
                });

                return false;
            }
        }
        public async Task<bool> CheckIfWorkFileExistsAsync()
        {
            try
            {
                var apiUrl = $"{_apiBaseUrl}/api/bank-files/workfile-exists";

                var result = await _httpClient.GetFromJsonAsync<ExistsResponse>(apiUrl);
                return result?.Exists ?? false;
            }
            catch
            {
                return false; // fallback if there's an error
            }
        }



        public async Task<bool> ValidateUnvalidatedRecordsAsync()
        {
            try
            {
                var apiUrl = $"{_apiBaseUrl}/api/validation";
                var response = await _httpClient.PostAsync(apiUrl, null);

                var jsonResponse = await response.Content.ReadAsStringAsync();
                var parsedResponse = JsonSerializer.Deserialize<ApiResponse>(jsonResponse, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                string message = parsedResponse?.Message ?? "Unknown validation result.";

                if (response.IsSuccessStatusCode)
                {
                    _notificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Validation Successful",
                        Detail = message,
                        Duration = 8000
                    });

                    return true;
                }
                else
                {
                    _notificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Validation Failed",
                        Detail = message,
                        Duration = 8000
                    });

                    return false;
                }
            }
            catch (Exception ex)
            {
                _notificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Validation failed: {ex.Message}",
                    Duration = 10000
                });

                return false;
            }
        }


    }
}
