using Server.Repositories;

namespace Server.Services
{
    public interface ILookupService
    {
        Task<IEnumerable<string>> GetIncomeSourceTypesAsync();
    }
    public class LookupService : ILookupService
    {
        private readonly ILookups _lookupsRepository;

        public LookupService(ILookups lookupsRepository)
        {
            _lookupsRepository = lookupsRepository;
        }

        public async Task<IEnumerable<string>> GetIncomeSourceTypesAsync()
        {
            return await _lookupsRepository.GetAllAsync();
        }
    }
}
