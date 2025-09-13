using Microsoft.Extensions.Configuration;

namespace Client.Services
{
    public class ConfigurationService
    {
        private readonly IConfiguration _configuration;

        public ConfigurationService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public string GetApiBaseUrl()
        {
            return _configuration["ApiBaseUrl"] ?? "https://localhost:7045"; 
        }
    }
}
