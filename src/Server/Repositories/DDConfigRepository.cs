namespace Server.Repositories;
using Dapper;
using DDS.API.Data;
using Server.Models;
using System.Data;





    public interface IDDConfigRepository
    {
        Task<DDConfigInfo?> GetConfigAsync();
        Task<int> UpdateConfigAsync(DDConfigInfo model, string updateMode);
    }

    public class DDConfigRepository:IDDConfigRepository
    {
        private readonly DapperDbContext _dbContext;
        private readonly ILogger<DDConfigRepository> _logger;

        public DDConfigRepository(DapperDbContext dbContext, ILogger<DDConfigRepository> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }

        public async Task<DDConfigInfo?> GetConfigAsync()
        {
            const string sql = "SELECT TOP 1 * FROM DD_CONFIG_INFO";

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema");
                var config = await connection.QueryFirstOrDefaultAsync<DDConfigInfo>(sql);
                _logger.LogInformation("Fetched DD config info");
                return config;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching DD config info");
                throw;
            }
        }

        public async Task<int> UpdateConfigAsync(DDConfigInfo model, string updateMode)
        {
            const string proc = "up_iu_Config";

            var parameters = new DynamicParameters();
            parameters.Add("config_id", model.CONFIG_ID);
            parameters.Add("pa_vendor_id_num", model.PA_VENDOR_ID_NUM);
            parameters.Add("sender_id", model.SENDER_ID);
            parameters.Add("receiver_id", model.RECEIVER_ID);
            parameters.Add("st_treas_email_to_addr", model.ST_TREAS_EMAIL_TO_ADDR);
            parameters.Add("ft1_insurance_code", model.FT1_INSURANCE_CODE);
            parameters.Add("patcode_entering_area", model.PATCODE_ENTERING_AREA);
            parameters.Add("st_treas_email_cc_addr", model.ST_TREAS_EMAIL_CC_ADDR);
            parameters.Add("st_treas_email_text", model.ST_TREAS_EMAIL_TEXT);
            parameters.Add("st_treas_email_subj", model.ST_TREAS_EMAIL_SUBJ);
            parameters.Add("data_refresh_rate", model.DATA_REFRESH_RATE);
            parameters.Add("data_lookback", model.DATA_LOOKBACK);
            parameters.Add("pa_batch_name", model.PA_BATCH_NAME);
            parameters.Add("pf_batch_name", model.PF_BATCH_NAME);
            parameters.Add("called_from_another_proc", "N");
            parameters.Add("update_mode", updateMode);
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema");
                await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
                int result = parameters.Get<int>("RETURN_VALUE");

                _logger.LogInformation("Executed {Proc} for ConfigId={Id}, Mode={Mode}", proc, model.CONFIG_ID, updateMode);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing {Proc} for ConfigId={Id}", proc, model.CONFIG_ID);
                throw;
            }
        }
    
}

