namespace Server.Services;
using Server.Models;
using Server.Repositories;


public interface IDDConfigService
{
        Task<DDConfigInfo?> GetConfigAsync();
        Task<int> UpdateConfigAsync(DDConfigInfo model, string updateMode);
}
    public class DDConfigService: IDDConfigService
{
        private readonly IDDConfigRepository _repo;
        private readonly ILogger<DDConfigService> _logger;

        public DDConfigService(IDDConfigRepository repo, ILogger<DDConfigService> logger)
        {
            _repo = repo;
            _logger = logger;
        }
        public async Task<DDConfigInfo?> GetConfigAsync()
        {
            try
            {
                var config = await _repo.GetConfigAsync();
                _logger.LogInformation("Fetched DD configuration settings");
                return config;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching DD configuration");
                throw;
            }
        }
        public async Task<int> UpdateConfigAsync(DDConfigInfo model, string updateMode)
        {
            try
            {
                var result = await _repo.UpdateConfigAsync(model, updateMode);
                _logger.LogInformation("Updated DD configuration (ID={Id}) with mode={Mode}", model.CONFIG_ID, updateMode);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating DD configuration (ID={Id})", model.CONFIG_ID);
                throw;
            }
        }
    }

