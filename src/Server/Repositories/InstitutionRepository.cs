using Dapper;
using DDS.API.Data;
using Server.Models;
using System.Data;

namespace Server.Repositories
{
    public interface IInstitutionRepository
    {
        Task<IEnumerable<Institution>> GetAllAsync();
        Task<Institution?> GetByIdAsync(long id);
        Task<int> InsertOrUpdateAsync(Institution model, string updateMode, string userId);
        //Task<int> DeleteAsync(long id, string userId);
    }

    public class InstitutionRepository : IInstitutionRepository
    {
        private readonly DapperDbContext _dbContext;
        private readonly ILogger<InstitutionRepository> _logger;

        public InstitutionRepository(DapperDbContext dbContext, ILogger<InstitutionRepository> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }


        public async Task<IEnumerable<Institution>> GetAllAsync()
        {
            const string sql = @"
               SELECT 
                INSTITUTION_ID, 
                ISNULL(INSTITUTION_CODE_3, INSTITUTION_CODE) AS INSTITUTION_CODE, 
                INSTITUTION_NAME, 
                DD_VENDOR_ID_NUM, 
                AFFINITY_DB_NAME, 
                DD_SEND_REPORT_TO, 
                CREATED_BY, 
                CREATED_DATETIME, 
                LAST_MOD_BY, 
                LAST_MOD_DATE, 
                RECORD_STATUS
            FROM PF_INSTITUTION
            WHERE RECORD_STATUS = 'A'
            ORDER BY INSTITUTION_CODE;
            ";

            try
            {
                using var connection = _dbContext.CreateConnection("pfs_schema");
                var results = await connection.QueryAsync<Institution>(sql);
                _logger.LogInformation("Fetched {Count} institutions", results.Count());
                return results;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching institutions");
                throw;
            }
        }
        public async Task<Institution?> GetByIdAsync(long id)
        {
            const string sql = @"
                SELECT INSTITUTION_ID, INSTITUTION_CODE, INSTITUTION_NAME, 
                       DD_VENDOR_ID_NUM, AFFINITY_DB_NAME, DD_SEND_REPORT_TO, 
                       CREATED_BY, CREATED_DATETIME, LAST_MOD_BY, LAST_MOD_DATE, RECORD_STATUS
                FROM PF_INSTITUTION
                WHERE INSTITUTION_ID = @id";

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema");
                return await connection.QueryFirstOrDefaultAsync<Institution>(sql, new { id });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching institution by ID={Id}", id);
                throw;
            }
        }

        public async Task<int> InsertOrUpdateAsync(Institution model, string updateMode, string userId)
        {
            const string proc = "up_u_Institution_DDS";

            var parameters = new DynamicParameters();
            parameters.Add("institution_code", model.INSTITUTION_CODE);
            parameters.Add("institution_name", model.INSTITUTION_NAME);
            parameters.Add("dd_vendor_id_num", model.DD_VENDOR_ID_NUM);
            parameters.Add("affinity_db_name", model.AFFINITY_DB_NAME);
            parameters.Add("dd_send_report_to", model.DD_SEND_REPORT_TO);
            parameters.Add("user_id", userId);
            parameters.Add("update_mode", updateMode);
            parameters.Add("institution_id", model.INSTITUTION_ID);
            parameters.Add("called_from_another_proc", "N");
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue); 

            try
            {
                using var connection = _dbContext.CreateConnection("pfs_schema");
                await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
                int result = parameters.Get<int>("RETURN_VALUE");
                _logger.LogInformation("Executed {Proc} for InstitutionId={Id} by {User}, Mode={Mode}", proc, model.INSTITUTION_ID, userId, updateMode);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing {Proc} for InstitutionId={Id}", proc, model.INSTITUTION_ID);
                throw;
            }
        }
        //public async Task<int> DeleteAsync(long id, string userId)
        //{
        //    const string proc = "up_d_Institution_DDS";

        //    var parameters = new DynamicParameters();
        //    parameters.Add("institution_id", id);
        //    parameters.Add("user_id", userId);
        //    parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

        //    try
        //    {
        //        using var connection = _dbContext.CreateConnection("pfs_schema");
        //        await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
        //        int result = parameters.Get<int>("RETURN_VALUE");
        //        _logger.LogInformation("Deleted InstitutionId={Id} by {User}", id, userId);
        //        return result;
        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "Error deleting InstitutionId={Id}", id);
        //        throw;
        //    }
        //}

    }

}
