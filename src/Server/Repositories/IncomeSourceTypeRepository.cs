namespace Server.Services;

using Dapper;
using DDS.API.Data;
using Server.Models;
using System.Data;

public interface IIncomeSourceTypeRepository
{
        Task<IEnumerable<IncomeSourceType>> GetAllAsync();
        Task<IncomeSourceType?> GetByIdAsync(int id);
        Task<int> InsertOrUpdateAsync(IncomeSourceType model, string updateMode, string userId);
}

public class IncomeSourceTypeRepository:IIncomeSourceTypeRepository
{
    
        private readonly DapperDbContext _dbContext;
        private readonly ILogger<IncomeSourceTypeRepository> _logger;

        public IncomeSourceTypeRepository(DapperDbContext dbContext, ILogger<IncomeSourceTypeRepository> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }
    public async Task<IEnumerable<IncomeSourceType>> GetAllAsync()
    {
        const string sql = "SELECT * FROM DD_INCOME_SOURCE_TYPE WHERE RECORD_STATUS = 'A' ORDER BY FUNB_INCOME_SRC_TYPE";

        try
        {
            using var connection = _dbContext.CreateConnection("dds_schema");
            var results = await connection.QueryAsync<IncomeSourceType>(sql);
            _logger.LogInformation("Fetched {Count} income source types", results.Count());
            return results;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching income source types");
            throw;
        }
    }
    public async Task<IncomeSourceType?> GetByIdAsync(int id)
    {
        const string sql = "SELECT * FROM DD_INCOME_SOURCE_TYPE WHERE INCOME_SOURCE_TYPE_ID = @id";

        try
        {
            using var connection = _dbContext.CreateConnection("dds_schema");
            return await connection.QueryFirstOrDefaultAsync<IncomeSourceType>(sql, new { id });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching income source type by ID={Id}", id);
            throw;
        }
    }

    public async Task<int> InsertOrUpdateAsync(IncomeSourceType model, string updateMode, string userId)
    {
        model.NCAS_ACCOUNT = "44200011";
        const string proc = "up_iu_Income_Source_Type";

        var parameters = new DynamicParameters();
        parameters.Add("FUNB_INCOME_SRC_TYPE", model.FUNB_INCOME_SRC_TYPE);
        parameters.Add("INCOME_SRC_TYPE_DESCR", model.INCOME_SRC_TYPE_DESCR);
        parameters.Add("PA_INCOME_SRC_TYPE", model.PA_INCOME_SRC_TYPE);
        parameters.Add("PA_PMT_CODE", model.PA_PMT_CODE);
        parameters.Add("PA_PMT_REV_CODE", model.PA_PMT_REV_CODE);
        parameters.Add("PF_DEP_TRANS_CODE", model.PF_DEP_TRANS_CODE);
        parameters.Add("PF_DEP_REV_TRANS_CODE", model.PF_DEP_REV_TRANS_CODE);
        parameters.Add("NCAS_ACCOUNT", model.NCAS_ACCOUNT);
        parameters.Add("user_id", userId);
        parameters.Add("start_pos",6);
        parameters.Add("length", 9);
        parameters.Add("update_mode", updateMode);
        parameters.Add("called_from_another_proc", "N");
        parameters.Add("INCOME_SOURCE_TYPE_ID", model.INCOME_SOURCE_TYPE_ID);

        // Output return value
        parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

        try
        {
            using var connection = _dbContext.CreateConnection("dds_schema");
            await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);

            int result = parameters.Get<int>("RETURN_VALUE");
            _logger.LogInformation("Executed {Proc} for IncomeSourceTypeId={Id} by {User}, Mode={Mode}", proc, model.INCOME_SOURCE_TYPE_ID, userId, updateMode);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error executing {Proc} for IncomeSourceTypeId={Id}", proc, model.INCOME_SOURCE_TYPE_ID);
            throw;
        }
    }

    public async Task<int> DeleteAsync(long id, string userId)
    {
        const string proc = "up_d_Income_Source_Type";

        var parameters = new DynamicParameters();
        parameters.Add("INCOME_SOURCE_TYPE_ID", id);
        parameters.Add("user_id", userId);
        parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);


        try
        {
            using var connection = _dbContext.CreateConnection("dds_schema");
            await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
            int result = parameters.Get<int>("RETURN_VALUE");
            _logger.LogInformation("Deleted IncomeSourceTypeId={Id} by {User}", id, userId);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting IncomeSourceTypeId={Id}", id);
            throw;
        }
    }




}

