using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
namespace Client.Services
{
    public class AuthenticationService
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;
        public AuthenticationService(HttpClient httpClient, ConfigurationService configService)
        {
            _httpClient = httpClient;
            _apiBaseUrl = configService.GetApiBaseUrl();
        }
        public async Task<bool> LoginAsync(LoginRequest loginRequest)
        {
          var apiUrl = $"{_apiBaseUrl}/api/Authentication/login";
            try
            {
                var response = await _httpClient.PostAsJsonAsync(apiUrl, loginRequest);

                if (response.IsSuccessStatusCode)
                {
                    return true; // Login successful
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Login failed: {error}");
                    return false; // Login failed
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during login request: {ex.Message}");
                return false;
            }
        }

    }
}
