using Dapper;
using DDS.API.Data;
using Microsoft.Extensions.Logging;
using Server.Models;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace Server.Repositories
{
    public interface IPostDepositsRepository
    {
        Task CleanUpPartialPostsAsync(string postingStatus, string userId);
        Task<PostCounts> GetPostCountsAsync();
        Task<ConfigInfo> GetConfigInfoAsync();
        Task<List<ValidRecord>> GetRecordsToPostToAffinityAsync();
        Task<string> CreateHL7FlatFileAsync(ValidRecord record, string directory, ConfigInfo config, string processingId, DateTime transactionDate);
        Task DeleteAllFilesInDirectoryAsync(string directory);
        Task EnsureDirectoryExistsAsync(string directory);
        Task<IEnumerable<PostingAck>> GetAllPostingAcksAsync();
        Task<string> GetTransactionGroupNumberAsync();
        Task<bool> MarkAsSentForPostingAsync(long validRecordId);
        Task<(bool success, string pfPostingStatus)> PostPFTransactionAsync(long validRecordId, string transGroupNum, string userId, string paPostingStatus);
    }

    public class PostDepositsRepository : IPostDepositsRepository
    {
        private readonly DapperDbContext _dbContext;
        private readonly ILogger<PostDepositsRepository> _logger;

        public PostDepositsRepository(DapperDbContext dbContext, ILogger<PostDepositsRepository> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }

        public async Task CleanUpPartialPostsAsync(string postingStatus, string userId)
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            // This should call a stored procedure or update as in CleanUpValidRec
            await conn.ExecuteAsync("up_id_Cleanup_Valid_Rec", new { PostingStatus = postingStatus, UserId = userId }, commandType: CommandType.StoredProcedure);
        }

        public async Task<PostCounts> GetPostCountsAsync()
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            _logger.LogInformation("Fetching post counts from DD_VALID_REC");

            var total = await conn.ExecuteScalarAsync<int>("SELECT COUNT(VALID_RECORD_ID) FROM DD_VALID_REC");
            var pa = await conn.ExecuteScalarAsync<int>("SELECT COUNT(VALID_RECORD_ID) FROM DD_VALID_REC WHERE PA_DISTRIBUTION_AMT > 0");
            var pf = await conn.ExecuteScalarAsync<int>("SELECT COUNT(VALID_RECORD_ID) FROM DD_VALID_REC WHERE PF_DISTRIBUTION_AMT > 0");

            _logger.LogInformation("Fetched post counts - Total: {Total}, PA: {PA}, PF: {PF}", total, pa, pf);

            return new PostCounts { TotalRecords = total, TotalPATrans = pa, TotalPFTrans = pf };
        }


        public async Task<ConfigInfo> GetConfigInfoAsync()
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            _logger.LogInformation("Fetching config info from DD_CONFIG_INFO");

            var sql = "SELECT TOP 1 FT1_INSURANCE_CODE, PATCODE_ENTERING_AREA FROM DD_CONFIG_INFO";
         
            return await conn.QueryFirstOrDefaultAsync<ConfigInfo>(sql);
        }

        public async Task<List<ValidRecord>> GetRecordsToPostToAffinityAsync()
        {
            using var conn = _dbContext.CreateConnection("dds_schema");
            _logger.LogInformation("Fetching records to post to Affinity");

            var sql = @"SELECT v.*, i.PA_PMT_CODE 
                        FROM DD_VALID_REC v
                        JOIN DD_INCOME_SOURCE_TYPE i ON v.INCOME_SOURCE_TYPE_ID = i.INCOME_SOURCE_TYPE_ID
                        WHERE v.PA_DISTRIBUTION_AMT > 0 AND v.SENT_FOR_POSTING_DATETIME IS NULL";
            var result = await conn.QueryAsync<ValidRecord>(sql);
            _logger.LogInformation("Fetched {Count} records to post to Affinity", result.AsList().Count);

            return result.AsList();
        }

        public async Task<string> CreateHL7FlatFileAsync(ValidRecord record, string directory, ConfigInfo config, string processingId, DateTime transactionDate)
        {
            _logger.LogInformation("Creating HL7 flat file for VALID_RECORD_ID={RecordId}", record.VALID_RECORD_ID);

            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);


            // Define the parameters
            var sendingApp = "DDS";
            var sendingFacility = "";
            var receivingApp = "";
            var receivingFacility = record.INSTITUTION_CODE;
            var messageControlId = record.VALID_RECORD_ID.ToString();
            var recordedDateTime = transactionDate;
            var patientName = record.PATIENT_NAME;
            var patientIdInternal = record.MEDICAL_RECORD_NUM.PadLeft(7, '0');
            var patientAccountNumber = record.AFFINITY_ACCT_NUM.PadLeft(8, '0');
            var transactionType = "PY";
            var transactionCode = record.PA_PMT_CODE;
            var transactionQuantity = "1";
            var transactionAmountExtended = record.PA_DISTRIBUTION_AMT;
            var departmentCode = config.PATCODE_ENTERING_AREA;
            var insurancePlanId = config.FT1_INSURANCE_CODE;

            // Log all HL7 parameters
            _logger.LogInformation("HL7 Parameters: " +
                "\n SendingApp: {SendingApp}" +
                "\n SendingFacility: {SendingFacility}" +
                "\n ReceivingApp: {ReceivingApp}" +
                "\n ReceivingFacility: {ReceivingFacility}" +
                "\n MessageControlId: {MessageControlId}" +
                "\n RecordedDateTime: {RecordedDateTime}" +
                "\n PatientName: {PatientName}" +
                "\n PatientIdInternal: {PatientIdInternal}" +
                "\n PatientAccountNumber: {PatientAccountNumber}" +
                "\n ProcessingId: {ProcessingId}" +
                "\n TransactionDate: {TransactionDate}" +
                "\n TransactionType: {TransactionType}" +
                "\n TransactionCode: {TransactionCode}" +
                "\n TransactionQuantity: {TransactionQuantity}" +
                "\n TransactionAmountExtended: {TransactionAmountExtended}" +
                "\n DepartmentCode: {DepartmentCode}" +
                "\n InsurancePlanId: {InsurancePlanId}",
                sendingApp,
                sendingFacility,
                receivingApp,
                receivingFacility,
                messageControlId,
                recordedDateTime,
                patientName,
                patientIdInternal,
                patientAccountNumber,
                processingId,
                transactionDate,
                transactionType,
                transactionCode,
                transactionQuantity,
                transactionAmountExtended,
                departmentCode,
                insurancePlanId
            );

            var fileName = Path.Combine(directory, $"{record.INSTITUTION_CODE}_{record.VALID_RECORD_ID}.txt");
            var hl7Message = HL7P03MessageBuilder.Build(
                sendingApp: "DDS",
                sendingFacility: "",
                receivingApp: "",
                receivingFacility: record.INSTITUTION_CODE,
                messageControlId: record.VALID_RECORD_ID.ToString(),
                recordedDateTime: transactionDate,
                patientName: record.PATIENT_NAME,
                patientIdInternal: record.MEDICAL_RECORD_NUM.PadLeft(7, '0'),
                patientAccountNumber: record.AFFINITY_ACCT_NUM.PadLeft(8, '0'),
                processingId: processingId,
                transactionDate: transactionDate,
                transactionType: "PY",
                transactionCode: record.PA_PMT_CODE,
                transactionQuantity: "1",
                transactionAmountExtended: record.PA_DISTRIBUTION_AMT,
                departmentCode: config.PATCODE_ENTERING_AREA,
                insurancePlanId: config.FT1_INSURANCE_CODE
            );
            await File.WriteAllTextAsync(fileName, hl7Message);
            _logger.LogInformation("Created HL7 flat file: {FileName}", fileName);
            return fileName;
        }

        public async Task<(bool success, string pfPostingStatus)> PostPFTransactionAsync(long validRecordId, string transGroupNum, string userId, string paPostingStatus)
        {
            const string proc = "up_id_Post_PFS_Trans";

            var parameters = new DynamicParameters();
            parameters.Add("dValidRecordID", validRecordId);
            parameters.Add("sTransGroupNum", transGroupNum);
            parameters.Add("sUserID", userId);
            parameters.Add("sPAPostingStatus", paPostingStatus);
            parameters.Add("sPFPostingStatus_OUTPUT", dbType: DbType.String, direction: ParameterDirection.Output, size: 50);
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema"); // Adjust schema if needed
                await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);

                int returnValue = parameters.Get<int>("RETURN_VALUE");
                string pfPostingStatus = parameters.Get<string>("sPFPostingStatus_OUTPUT");

                if (returnValue != 0)
                {
                    _logger.LogError("Executed {Proc} for ValidRecordId={ValidId} failed with ReturnValue={ReturnValue}", proc, validRecordId, returnValue);
                    return (false, "Post Fails");
                }

                _logger.LogInformation("Executed {Proc} for ValidRecordId={ValidId} by {User}. PFPostingStatus={Status}", proc, validRecordId, userId, pfPostingStatus);
                return (pfPostingStatus == "Posted", pfPostingStatus);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing {Proc} for ValidRecordId={ValidId}", proc, validRecordId);
                return (false, "Error");
            }
        }
        public async Task<string> GetTransactionGroupNumberAsync()
        {
            const string proc = "up_i_GetTransGroupNum";

            var parameters = new DynamicParameters();
            parameters.Add("sTransGroupNum_OUTPUT", dbType: DbType.String, direction: ParameterDirection.Output, size: 50);

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema"); 
                await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);

                string result = parameters.Get<string>("sTransGroupNum_OUTPUT");
                _logger.LogInformation("Executed {Proc}, generated transaction group number: {GroupNum}", proc, result);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing {Proc} to generate transaction group number", proc);
                throw;
            }
        }

        public async Task<IEnumerable<PostingAck>> GetAllPostingAcksAsync()
        {
            const string sql = "SELECT * FROM dbo.DD_POSTING_ACK";

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema"); // Change schema if needed
                var result = await connection.QueryAsync<PostingAck>(sql);
                _logger.LogInformation("Executed {Sql} to retrieve all posting acks", sql);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing {Sql} to retrieve all posting acks", sql);
                throw;
            }
        }


        public async Task DeleteAllFilesInDirectoryAsync(string directory)
        {
            if (Directory.Exists(directory))
            {
                foreach (var file in Directory.GetFiles(directory, "*.txt"))
                {
                    File.Delete(file);
                }
            }
            await Task.CompletedTask;
        }

        public async Task EnsureDirectoryExistsAsync(string directory)
        {
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            await Task.CompletedTask;
        }

        public async Task<bool> MarkAsSentForPostingAsync(long validRecordId)
        {
            const string proc = "up_u_Sent_for_Posting";

            var parameters = new DynamicParameters();
            parameters.Add("dValidRecordID", validRecordId);
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema"); // Adjust schema name if needed
                await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);

                int returnValue = parameters.Get<int>("RETURN_VALUE");

                if (returnValue != 0)
                {
                    _logger.LogError("Stored procedure {Proc} failed for ValidRecordId={ValidId} with ReturnValue={ReturnValue}", proc, validRecordId, returnValue);
                    return false;
                }

                _logger.LogInformation("Marked ValidRecordId={ValidId} as sent for posting using {Proc}", validRecordId, proc);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing {Proc} for ValidRecordId={ValidId}", proc, validRecordId);
                return false;
            }
        }

    }



    public class ConfigInfo
    {
        public string FT1_INSURANCE_CODE { get; set; } =string.Empty;
        public string PATCODE_ENTERING_AREA { get; set; } = string.Empty;
    }

    public class ValidRecord
    {
        public long VALID_RECORD_ID { get; set; } 
        public string MEDICAL_RECORD_NUM { get; set; } = string.Empty;
        public string AFFINITY_ACCT_NUM { get; set; }= string.Empty;
        public string PATIENT_NAME { get; set; }=   string.Empty;
        public decimal PA_DISTRIBUTION_AMT { get; set; }
        public string INSTITUTION_CODE { get; set; }=string.Empty;
        public string PA_PMT_CODE { get; set; } = string.Empty;
    }
}