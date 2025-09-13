using Server.Models;
using Server.Repositories;

namespace Server.Services
{
    public interface IPreEditService
    {
        Task<IEnumerable<InvalidRecord>> GetInvalidRecordsByIdAsync(int baiFileId);
        Task<IEnumerable<InvalidRecord>> GetAllInvalidRecordsAsync();
        Task UpdateInvalidRecordAsync(InvalidRecord record);
        Task MoveRecordToValidAsync(long id);
        Task RecoupAsync(long creditId, long debitId, string userId);
        Task DeleteRecordAsync(long id);
        Task UndeleteRecordAsync(long id, string userId);
        Task<IEnumerable<InvalidRecord>> GetHiddeTransactionsAsync();

        Task<int> HidePreEditRecordAsync(long invalidRecordId, string recordStatus, string userId);

    }

    public class PreEditService : IPreEditService
    {
        private readonly IInvalidRecordRepository _repo;
        private readonly ILogger<PreEditService> _logger;

        public PreEditService(IInvalidRecordRepository repo, ILogger<PreEditService> logger)
        {
            _repo = repo;
            _logger = logger;
        }

        public async Task<IEnumerable<InvalidRecord>> GetInvalidRecordsByIdAsync(int baiFileId)
        {
            try
            {
                var records = await _repo.GetAllByIdAsync(baiFileId);
                _logger.LogInformation("Fetched {Count} invalid records", records.Count());
                return records;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to fetch invalid records");
                throw;
            }
        }

        public async Task<IEnumerable<InvalidRecord>> GetAllInvalidRecordsAsync()
        {
            try
            {
                var records = await _repo.GetAllAsync();
                _logger.LogInformation("Fetched {Count} invalid records", records.Count());
                return records;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to fetch invalid records");
                throw;
            }
        }
        public async Task UpdateInvalidRecordAsync(InvalidRecord record)
        {
            try
            {
                await _repo.UpdateAsync(record);
                _logger.LogInformation("Updated InvalidRecordId={Id} by {User}", record.INVALID_RECORD_ID, record.LAST_MOD_BY);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating InvalidRecordId={Id}", record.INVALID_RECORD_ID);
                throw;
            }
        }

        public async Task MoveRecordToValidAsync(long id)
        {
            try
            {
                await _repo.MoveToValidAsync(id);
                _logger.LogInformation("Moved InvalidRecordId={Id}", id);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error moving InvalidRecordId={Id} to valid", id);
                throw;
            }
        }

        public async Task RecoupAsync(long creditId, long debitId, string userId)
        {
            try
            {
                await _repo.RecoupAsync(creditId, debitId, userId);
                _logger.LogInformation("Recouped CreditId={Credit}, DebitId={Debit} by {User}", creditId, debitId, userId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error recouping CreditId={Credit} and DebitId={Debit}", creditId, debitId);
                throw;
            }
        }

        public async Task DeleteRecordAsync(long id)
        {
            try
            {
                await _repo.DeleteAsync(id);
                _logger.LogInformation("Soft-deleted InvalidRecordId={Id}", id);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting InvalidRecordId={Id}", id);
                throw;
            }
        }

        public async Task UndeleteRecordAsync(long id, string userId)
        {
            try
            {
                await _repo.UndeleteAsync(id, userId);
                _logger.LogInformation("Restored InvalidRecordId={Id} by {User}", id, userId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error restoring InvalidRecordId={Id}", id);
                throw;
            }
        }

        public async Task<int> HidePreEditRecordAsync(long invalidRecordId, string recordStatus, string userId)
        {
            return await _repo.HidePreEditRecordAsync(invalidRecordId, recordStatus, userId);
        }

        public async Task<IEnumerable<InvalidRecord>> GetHiddeTransactionsAsync()
        {
            return await _repo.GetHiddeTransactionsAsync();
        }
    }

}
