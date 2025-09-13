using System.Data;
using Dapper;
using Microsoft.Data.SqlClient;
using Server.Infrastructure.Data;
using Server.Infrastructure.Logging;
using Server.Models;

namespace Server.Services;

public interface IStateTreasurerRepository
{
    Task<IReadOnlyList<DsnItem>> GetPriorDsnsAsync();
    Task<bool> HasTransactionsOnDateAsync(DateTime postedDate);
    Task<bool> HasIncompleteSendOnDateAsync(DateTime postedDate);
    Task<bool> IsDsnRequiredAsync(DateTime postedDate);
    Task<bool> DsnExistsWithinSixMonthsAsync(string depSeqNum);
    Task<int> InsertDsnAsync(string depSeqNum, DateTime processDate, string createdBy);
    Task<int> MarkSentToTreasurerAsync(DateTime postedDate, string? depSeqNum, string sentIndicator, string userId);
    Task<IReadOnlyList<InstitutionTotal>> GetInstitutionTotalsAsync(DateTime postedDate);
    Task<DailyTotals> GetDailyTotalsAsync(DateTime postedDate);
    Task<IReadOnlyList<AlogRow>> GetAlogRowsAsync(DateTime postedDate);
    Task<ConfigInfo> GetConfigInfoAsync();
    Task<IReadOnlyList<RegionInfo>> GetRegionsAsync();
}

public class StateTreasurerRepository : IStateTreasurerRepository
{
    private readonly DapperDbContext _dbContext;
    private readonly IAppLogger<StateTreasurerRepository> _logger;

    public StateTreasurerRepository(DapperDbContext dbContext, IAppLogger<StateTreasurerRepository> logger)
    {
        _dbContext = dbContext;
        _logger = logger;
    }

    public async Task<IReadOnlyList<DsnItem>> GetPriorDsnsAsync()
    {
        const string sql = @"SELECT DEP_SEQ_NUM AS DepSeqNum, CREATED_DATETIME AS PostedDate
                              FROM DD_DEP_SEQ_NO
                              ORDER BY CREATED_DATETIME DESC, DEP_SEQ_NUM";
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var rows = await conn.QueryAsync<DsnItem>(sql);
            return rows.ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching prior DSNs");
            throw;
        }
    }

    public async Task<bool> HasTransactionsOnDateAsync(DateTime postedDate)
    {
        const string sql = @"SELECT TOP 1 1
                             FROM DD_POSTING_HISTORY
                             WHERE POSTED_DATETIME >= @start AND POSTED_DATETIME < @end";
        var (start, end) = GetDateBounds(postedDate);
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var exists = await conn.ExecuteScalarAsync<int?>(sql, new { start, end });
            return exists.HasValue;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking transactions on date {Date}", postedDate);
            throw;
        }
    }

    public async Task<bool> HasIncompleteSendOnDateAsync(DateTime postedDate)
    {
        const string sql = @"SELECT TOP 1 1
                             FROM DD_POSTING_HISTORY
                             WHERE SENT_TO_ST_TREAS_IND = 'N'
                               AND POSTED_DATETIME >= @start AND POSTED_DATETIME < @end";
        var (start, end) = GetDateBounds(postedDate);
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var exists = await conn.ExecuteScalarAsync<int?>(sql, new { start, end });
            return exists.HasValue;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking incomplete send on date {Date}", postedDate);
            throw;
        }
    }

    public async Task<bool> IsDsnRequiredAsync(DateTime postedDate)
    {
        const string sql = @"SELECT TOP 1 1
                             FROM DD_POSTING_HISTORY
                             WHERE PA_DISTRIBUTION_AMT > 0
                               AND POSTED_DATETIME >= @start AND POSTED_DATETIME < @end";
        var (start, end) = GetDateBounds(postedDate);
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var exists = await conn.ExecuteScalarAsync<int?>(sql, new { start, end });
            return exists.HasValue;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking DSN requirement on date {Date}", postedDate);
            throw;
        }
    }

    public async Task<bool> DsnExistsWithinSixMonthsAsync(string depSeqNum)
    {
        const string sql = @"SELECT TOP 1 1
                             FROM DD_DEP_SEQ_NO
                             WHERE DEP_SEQ_NUM = @depSeqNum
                               AND CREATED_DATETIME > DATEADD(month, -6, GETDATE())";
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var exists = await conn.ExecuteScalarAsync<int?>(sql, new { depSeqNum });
            return exists.HasValue;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking DSN exists {DepSeqNum}", depSeqNum);
            throw;
        }
    }

    public async Task<int> InsertDsnAsync(string depSeqNum, DateTime processDate, string createdBy)
    {
        const string proc = "up_i_DSN";
        var parameters = new DynamicParameters();
        parameters.Add("CREATED_BY", createdBy);
        parameters.Add("DEP_SEQ_NUM", depSeqNum);
        parameters.Add("PROCESS_DATE", processDate);
        parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

        try
        {
            using var connection = _dbContext.CreateConnection("dds_schema");
            await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
            var result = parameters.Get<int>("RETURN_VALUE");
            _logger.LogInformation("Inserted DSN {DepSeqNum} by {User}", depSeqNum, createdBy);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error inserting DSN {DepSeqNum}", depSeqNum);
            throw;
        }
    }

    public async Task<int> MarkSentToTreasurerAsync(DateTime postedDate, string? depSeqNum, string sentIndicator, string userId)
    {
        const string proc = "up_u_Sent_To_St_Treas";
        var parameters = new DynamicParameters();
        parameters.Add("SENT_TO_ST_TREAS_IND", sentIndicator);
        if (!string.IsNullOrWhiteSpace(depSeqNum))
        {
            parameters.Add("DEP_SEQ_NUM", depSeqNum);
        }
        parameters.Add("POSTED_DATETIME", postedDate.Date);
        parameters.Add("user_id", userId);
        parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

        try
        {
            using var connection = _dbContext.CreateConnection("dds_schema");
            await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);
            var result = parameters.Get<int>("RETURN_VALUE");
            _logger.LogInformation("Marked sent to treasurer for {Date} depSeq={DepSeq} by {User}", postedDate, depSeqNum ?? "", userId);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error marking sent to treasurer for {Date}", postedDate);
            throw;
        }
    }

    public async Task<IReadOnlyList<InstitutionTotal>> GetInstitutionTotalsAsync(DateTime postedDate)
    {
        const string sql = @"SELECT ph.INSTITUTION_CODE AS InstitutionCode,
                                    SUM(CASE WHEN ph.DR_CR_FLAG = 'DR' THEN -1 ELSE 1 END * ph.PA_DISTRIBUTION_AMT) AS TotalPAAmt,
                                    SUM(CASE WHEN ph.DR_CR_FLAG = 'DR' THEN -1 ELSE 1 END * ph.PF_DISTRIBUTION_AMT) AS TotalPFAmt,
                                    inst.DD_VENDOR_ID_NUM AS VendorId,
                                    inst.INSTITUTION_NAME AS InstitutionName
                             FROM DD_POSTING_HISTORY ph
                             LEFT JOIN PF_INSTITUTION inst ON inst.INSTITUTION_CODE = ph.INSTITUTION_CODE
                             WHERE ph.POSTED_DATETIME >= @start AND ph.POSTED_DATETIME < @end
                             GROUP BY ph.INSTITUTION_CODE, inst.DD_VENDOR_ID_NUM, inst.INSTITUTION_NAME
                             ORDER BY ph.INSTITUTION_CODE";

        var (start, end) = GetDateBounds(postedDate);
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var rows = await conn.QueryAsync<InstitutionTotal>(sql, new { start, end });
            return rows.ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching institution totals for {Date}", postedDate);
            throw;
        }
    }

    public async Task<DailyTotals> GetDailyTotalsAsync(DateTime postedDate)
    {
        const string paSql = @"SELECT SUM(CASE WHEN DR_CR_FLAG = 'DR' THEN -1 ELSE 1 END * PA_DISTRIBUTION_AMT) AS Total
                               FROM DD_POSTING_HISTORY
                               WHERE POSTED_DATETIME >= @start AND POSTED_DATETIME < @end";
        const string pfSql = @"SELECT SUM(CASE WHEN DR_CR_FLAG = 'DR' THEN -1 ELSE 1 END * PF_DISTRIBUTION_AMT) AS Total
                               FROM DD_POSTING_HISTORY
                               WHERE POSTED_DATETIME >= @start AND POSTED_DATETIME < @end";
        var (start, end) = GetDateBounds(postedDate);
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var pa = await conn.ExecuteScalarAsync<decimal?>(paSql, new { start, end }) ?? 0m;
            var pf = await conn.ExecuteScalarAsync<decimal?>(pfSql, new { start, end }) ?? 0m;
            return new DailyTotals
            {
                PostedDate = postedDate.Date,
                TotalPAAmt = pa,
                TotalPFAmt = pf
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching daily totals for {Date}", postedDate);
            throw;
        }
    }

    public async Task<IReadOnlyList<AlogRow>> GetAlogRowsAsync(DateTime postedDate)
    {
        const string sql = @"SELECT ph.INSTITUTION_CODE AS InstitutionCode,
                                    ist.NCAS_ACCOUNT AS NcasAccount,
                                    ist.PA_INCOME_SRC_TYPE AS PaIncomeSourceType,
                                    SUM(CASE WHEN ph.DR_CR_FLAG = 'DR' THEN -1 ELSE 1 END * ph.PA_DISTRIBUTION_AMT) AS TotalPAAmt,
                                    SUM(CASE WHEN ph.DR_CR_FLAG = 'DR' THEN -1 ELSE 1 END * ph.PF_DISTRIBUTION_AMT) AS TotalPFAmt
                             FROM DD_POSTING_HISTORY ph
                             INNER JOIN DD_INCOME_SOURCE_TYPE ist ON ist.INCOME_SOURCE_TYPE_ID = ph.INCOME_SOURCE_TYPE_ID
                             WHERE ph.POSTED_DATETIME >= @start AND ph.POSTED_DATETIME < @end
                             GROUP BY ph.INSTITUTION_CODE, ist.NCAS_ACCOUNT, ist.PA_INCOME_SRC_TYPE
                             ORDER BY ph.INSTITUTION_CODE, ist.NCAS_ACCOUNT";
        var (start, end) = GetDateBounds(postedDate);
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var rows = await conn.QueryAsync<AlogRow>(sql, new { start, end });
            return rows.ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching ALOG rows for {Date}", postedDate);
            throw;
        }
    }

    public async Task<ConfigInfo> GetConfigInfoAsync()
    {
        const string sql = @"SELECT PA_BATCH_NAME, PF_BATCH_NAME, PA_VENDOR_ID_NUM,
                                   ST_TREAS_EMAIL_TO_ADDR, ST_TREAS_EMAIL_CC_ADDR, 
                                   ST_TREAS_EMAIL_SUBJ, ST_TREAS_EMAIL_TEXT
                            FROM DD_CONFIG_INFO";
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var config = await conn.QueryFirstOrDefaultAsync<ConfigInfo>(sql);
            if (config == null)
            {
                throw new InvalidOperationException("Configuration record not found");
            }
            return config;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching configuration info");
            throw;
        }
    }

    public async Task<IReadOnlyList<RegionInfo>> GetRegionsAsync()
    {
        const string sql = @"SELECT REGION, EMAIL_RECIPIENTS_TO, EMAIL_RECIPIENTS_CC
                            FROM DD_REGION
                            WHERE EMAIL_RECIPIENTS_TO IS NOT NULL";
        try
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            var regions = await conn.QueryAsync<RegionInfo>(sql);
            return regions.ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching regions");
            throw;
        }
    }

    private static (DateTime start, DateTime end) GetDateBounds(DateTime date)
    {
        var start = date.Date;
        var end = start.AddDays(1);
        return (start, end);
    }
}

