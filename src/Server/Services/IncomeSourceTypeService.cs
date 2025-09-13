using Server.Models;
using Server.Repositories;
using Microsoft.Extensions.Logging;

namespace Server.Services
{
    public interface IIncomeSourceTypeService
    {
        Task<IEnumerable<IncomeSourceType>> GetAllAsync();
        Task<IncomeSourceType> GetByIdAsync(int id);
        Task<int> InsertOrUpdateAsync(IncomeSourceType model, string updateMode, string userId);
        Task<int> DeleteAsync(long id, string userId);
    }
    public class IncomeSourceTypeService: IIncomeSourceTypeService
    {
        private readonly IncomeSourceTypeRepository _repo;
        private readonly ILogger<IncomeSourceTypeService> _logger;

        public IncomeSourceTypeService(IncomeSourceTypeRepository repo, ILogger<IncomeSourceTypeService> logger)
        {
            _repo = repo;
            _logger = logger;
        }

        public async Task<IEnumerable<IncomeSourceType>> GetAllAsync()
        {
            try
            {
                var items = await _repo.GetAllAsync();
                _logger.LogInformation("Fetched {Count} income source types", items.Count());
                return items;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching income source types");
                throw;
            }
        }

        public async Task<IncomeSourceType> GetByIdAsync(int id)
        {
            try
            {
                var item = await _repo.GetByIdAsync(id);
                if (item == null)
                    _logger.LogWarning("Income source type ID {Id} not found", id);
                return item;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching income source type ID {Id}", id);
                throw;
            }
        }

        public async Task<int> InsertOrUpdateAsync(IncomeSourceType model, string updateMode, string userId)
        {
            try
            {
                var result = await _repo.InsertOrUpdateAsync(model, updateMode, userId);
                _logger.LogInformation("InsertOrUpdate result={Result} for IncomeSourceTypeId={Id} by {User}", result, model.INCOME_SOURCE_TYPE_ID, userId);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting/updating income source type ID {Id}", model.INCOME_SOURCE_TYPE_ID);
                throw;
            }
        }

        public async Task<int> DeleteAsync(long id, string userId)
        {
            try
            {
                var result = await _repo.DeleteAsync(id, userId);
                _logger.LogInformation("Deleted IncomeSourceTypeId={Id} by {User}, result={Result}", id, userId, result);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting income source type ID {Id}", id);
                throw;
            }
        }
    }

}
