using Server.Models;
using Dapper;
using Microsoft.Data.SqlClient;
using Server.Infrastructure.Logging;
using DDS.API.Data;
using System.Data;

namespace Server.Repositories
{
    public interface IRegionRepository
    {
        Task<List<RegionDto>> GetAllRegionsAsync();
        Task<RegionDto?> GetRegionByIdAsync(long regionId);
        Task<long> AddRegionAsync(RegionDto region);
        Task<bool> UpdateRegionAsync(RegionDto region);
        Task<bool> DeleteRegionAsync(long regionId);
    }
    public class RegionRepository : IRegionRepository
    {
        private readonly DapperDbContext _dbContext;
        public RegionRepository(DapperDbContext dbContext)
        {
            _dbContext = dbContext;
        }


        public async Task<List<RegionDto>> GetAllRegionsAsync()
        {
            const string query = @"
            SELECT 
                REGION_ID AS RegionId,
                REGION AS Region,
                EMAIL_RECIPIENTS_TO AS EmailRecipientsTo,
                EMAIL_RECIPIENTS_CC AS EmailRecipientsCc,
                LAST_MOD_BY AS LastModBy,
                LAST_MOD_DATETIME AS LastModDatetime
            FROM DD_REGION
            ORDER BY REGION";

            using var connection = _dbContext.CreateConnection("dds_schema");
            var result = await connection.QueryAsync<RegionDto>(query);
            return result.ToList();
        }

        public async Task<RegionDto?> GetRegionByIdAsync(long regionId)
        {
            const string query = @"
        SELECT 
            REGION_ID AS RegionId,
            REGION AS Region,
            EMAIL_RECIPIENTS_TO AS EmailRecipientsTo,
            EMAIL_RECIPIENTS_CC AS EmailRecipientsCc,
            LAST_MOD_BY AS LastModBy,
            LAST_MOD_DATETIME AS LastModDatetime
        FROM DD_REGION
        WHERE REGION_ID = @regionId";

            using var connection = _dbContext.CreateConnection("dds_schema");
            return await connection.QueryFirstOrDefaultAsync<RegionDto>(query, new { regionId });
        }

        public async Task<long> AddRegionAsync(RegionDto region)
        {
            region.LastModBy = "system";
            const string proc = "up_iud_Regions";
            using var connection = _dbContext.CreateConnection("dds_schema");
            var parameters = new DynamicParameters();
            parameters.Add("region", region.Region);
            parameters.Add("email_recipients_to", region.EmailRecipientsTo);
            parameters.Add("email_recipients_cc", region.EmailRecipientsCc);
            parameters.Add("user_id", region.LastModBy);
            parameters.Add("update_mode", "I");
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

            await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
            return parameters.Get<int>("RETURN_VALUE");
        }
        public async Task<bool> UpdateRegionAsync(RegionDto region)
        {
            region.LastModBy = "system";

            const string proc = "up_iud_Regions";
            using var connection = _dbContext.CreateConnection("dds_schema");
            var parameters = new DynamicParameters();
            parameters.Add("region_id", region.RegionId);
            parameters.Add("region", region.Region);
            parameters.Add("email_recipients_to", region.EmailRecipientsTo);
            parameters.Add("email_recipients_cc", region.EmailRecipientsCc);
            parameters.Add("user_id", region.LastModBy);
            parameters.Add("update_mode", "U");
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

            await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
            return parameters.Get<int>("RETURN_VALUE") == 0;
        }


        public async Task<bool> DeleteRegionAsync(long regionId)
        {

            const string proc = "up_iud_Regions";
            using var connection = _dbContext.CreateConnection("dds_schema");
            var parameters = new DynamicParameters();
            parameters.Add("region_id", regionId);
            parameters.Add("user_id", "system"); 
            parameters.Add("update_mode", "D");
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

            await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
            return parameters.Get<int>("RETURN_VALUE") == 0;
        }
    }
    }
