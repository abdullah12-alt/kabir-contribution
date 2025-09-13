using Server.Models;
using Server.Repositories;
using System.Collections.Generic;
using System.Threading.Tasks;
namespace Server.Services
{
    public interface IBalanceService
    {
        Task<BalanceSummaryDto> GetBalanceSummaryAsync();
        Task<IList<InvalidRecord>> GetInvalidRecordsAsync();
        Task<int> InsertBalanceAsync(BalanceInsertDto dto);
        Task<IList<SummaryRecordDto>> GetSummaryRecordsAsync();
    }
    public class BalanceService : IBalanceService
    {

        private readonly IBalanceRepository _repo;

        public BalanceService(IBalanceRepository repo)
        {
            _repo = repo;
        }

        public Task<BalanceSummaryDto> GetBalanceSummaryAsync() => _repo.GetBalanceSummaryAsync();
        public Task<IList<InvalidRecord>> GetInvalidRecordsAsync() => _repo.GetInvalidRecordsAsync();
        public Task<int> InsertBalanceAsync(BalanceInsertDto dto) => _repo.InsertBalanceAsync(dto);
        public Task<IList<SummaryRecordDto>> GetSummaryRecordsAsync() => _repo.GetSummaryRecordsAsync();
    }
}
