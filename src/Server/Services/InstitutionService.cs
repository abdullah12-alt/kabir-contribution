
using Server.Models;
using Server.Repositories;

namespace Server.Services
{

    public interface IInstitutionService
    {
        Task<IEnumerable<Institution>> GetAllAsync();
        Task<Institution?> GetByIdAsync(long id);
        Task<int> InsertOrUpdateAsync(Institution model, string updateMode, string userId);
        //Task<int> DeleteAsync(long id, string userId);
    }
    public class InstitutionService : IInstitutionService
    {
        private readonly IInstitutionRepository _repo;
        private readonly ILogger<InstitutionService> _logger;
        public InstitutionService(IInstitutionRepository repo, ILogger<InstitutionService> logger)
        {
            _repo = repo;
            _logger = logger;
        }

        public async Task<IEnumerable<Institution>> GetAllAsync()
        {
            try
            {
                var results = await _repo.GetAllAsync();
                _logger.LogInformation("Fetched {Count} institutions", results.Count());
                return results;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching institution list");
                throw;
            }
        }

        public async Task<Institution?> GetByIdAsync(long id)
        {
            try
            {
                var item = await _repo.GetByIdAsync(id);
                if (item == null)
                    _logger.LogWarning("Institution ID {Id} not found", id);

                return item;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Institution ID={Id}", id);
                throw;
            }
        }

        public async Task<int> InsertOrUpdateAsync(Institution model, string updateMode, string userId)
        {
            try
            {
                var result = await _repo.InsertOrUpdateAsync(model, updateMode, userId);
                _logger.LogInformation("Insert/Update result={Result} for InstitutionId={Id} by {User}", result, model.INSTITUTION_ID, userId);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting/updating Institution ID={Id}", model.INSTITUTION_ID);
                throw;
            }
        }

        //public async Task<int> DeleteAsync(long id, string userId)
        //{
        //    try
        //    {
        //        var result = await _repo.DeleteAsync(id, userId);
        //        _logger.LogInformation("Deleted Institution ID={Id} by {User}, Result={Result}", id, userId, result);
        //        return result;
        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "Error deleting Institution ID={Id}", id);
        //        throw;
        //    }
        //}
    }
}
