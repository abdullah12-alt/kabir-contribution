using Dapper;
using DDS.API.Data;
using Microsoft.EntityFrameworkCore;
using Server.Services;

namespace Server.Repositories
{

     public interface ILookups
    {
        Task<IEnumerable<string>> GetAllAsync();

    }
    public class LookupsRepository : ILookups
    {
        private readonly DapperDbContext _dbContext;
        private readonly ILogger<LookupsRepository> _logger;

        public LookupsRepository(DapperDbContext dbContext, ILogger<LookupsRepository> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }
        public async Task<IEnumerable<string>> GetAllAsync()
        {
            const string sql = " SELECT FUNB_INCOME_SRC_TYPE from DD_INCOME_SOURCE_TYPE ";

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema");
                var results = await connection.QueryAsync<string>(sql);
                
                return results;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching income source types");
                throw;
            }
        }
    }
}
