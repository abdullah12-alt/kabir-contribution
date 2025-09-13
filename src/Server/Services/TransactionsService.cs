using DDS.API.Data;
using Microsoft.EntityFrameworkCore;
using Server.Services;
using Microsoft.Data.SqlClient;

namespace Server.Services;

using Dapper;
using Server.Infrastructure.Logging;
using Server.Models;


    public interface ITransactionService
    {
    Task<List<WorkTransactionDto>> GetAllTransactionsAsync();
    Task<List<InvalidTransactionDto>> GetInvalidTransactionsAsync();
    Task<List<ValidTransactionDto>> GetValidTransactionsAsync();

}

public class TransactionsService : ITransactionService
{
    private readonly DapperDbContext _dbContext;
    private readonly IAppLogger<TransactionsService> _logger;

    public TransactionsService(DapperDbContext dbContext, IAppLogger<TransactionsService> logger)
    {
        _dbContext = dbContext;
        _logger = logger;
    }

    public async Task<List<WorkTransactionDto>> GetAllTransactionsAsync()
    {
        _logger.LogInformation($"Getting all transactions ");

        using var ddsConnection = (SqlConnection)_dbContext.CreateConnection("dds_schema");
        await ddsConnection.OpenAsync();
        const string query = @"SELECT * FROM dbo.DD_WORK_FILE;";
        var result = await ddsConnection.QueryAsync<WorkTransactionDto>(query);
        _logger.LogInformation($"Retrieved {result.Count()} all transactions");

        return result.ToList();
    }



    public async Task<List<InvalidTransactionDto>> GetInvalidTransactionsAsync()
    {
        _logger.LogInformation("Getting invalid transactions...");

        const string query = @"
            SELECT 
                dir.*, 
                dire.*,
                ISNULL(pi.INSTITUTION_CODE_3, pi.INSTITUTION_CODE) AS INSTITUTION_CODE_3,
                pi.INSTITUTION_NAME
            FROM 
                dds.dbo.DD_INVALID_REC dir
            JOIN 
                dds.dbo.DD_INVALID_REC_ERROR dire 
                ON dir.INVALID_RECORD_ID = dire.INVALID_RECORD_ID
            JOIN 
                pfs.dbo.PF_INSTITUTION pi 
                ON pi.INSTITUTION_CODE = dir.INSTITUTION_CODE
            WHERE 
                pi.RECORD_STATUS = 'A'
        ";

        using var connection = _dbContext.CreateConnection("dds_schema");
        var result = await connection.QueryAsync<InvalidTransactionDto>(query);

        var list = result.ToList();
        _logger.LogInformation("Retrieved {Count} invalid transactions.", list.Count);

        return list;
    }

    public async Task<List<ValidTransactionDto>> GetValidTransactionsAsync()
    {
        _logger.LogInformation("Getting valid transactions...");

        const string query = @"
            SELECT 
                dvr.*, 
                dist.*,
                ISNULL(pi.INSTITUTION_CODE_3, pi.INSTITUTION_CODE) AS INSTITUTION_CODE_3
            FROM 
                dds.dbo.DD_VALID_REC dvr
            JOIN 
                dds.dbo.DD_INCOME_SOURCE_TYPE dist 
                ON dist.INCOME_SOURCE_TYPE_ID = dvr.INCOME_SOURCE_TYPE_ID
            JOIN 
                pfs.dbo.PF_INSTITUTION pi 
                ON pi.INSTITUTION_CODE = dvr.INSTITUTION_CODE
            WHERE 
                pi.RECORD_STATUS = 'A'
        ";

        using var connection = _dbContext.CreateConnection("dds_schema");
        var result = await connection.QueryAsync<ValidTransactionDto>(query);

        var list = result.ToList();
        _logger.LogInformation("Retrieved {Count} valid transactions.", list.Count);

        return list;
    }




}
