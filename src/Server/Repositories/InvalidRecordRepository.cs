using Dapper;
using DDS.API.Data;
using Server.Models;
using Server.Services;
using System.Data;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace Server.Repositories
{
    public interface IInvalidRecordRepository
    {
        Task<IEnumerable<InvalidRecord>> GetAllByIdAsync(int baiFileIDd);
        Task<IEnumerable<InvalidRecord>> GetAllAsync();
        Task UpdateAsync(InvalidRecord record);
        Task MoveToValidAsync(long invalidRecordId);
        Task RecoupAsync(long creditId, long debitId, string userId);
        Task DeleteAsync(long invalidRecordId);
        Task UndeleteAsync(long invalidRecordId, string userId);
        Task<int> HidePreEditRecordAsync(long invalidRecordId, string recordStatus, string userId);
        Task<IEnumerable<InvalidRecord>> GetHiddeTransactionsAsync();
    }

    public class InvalidRecordRepository : IInvalidRecordRepository
    {
        private readonly DapperDbContext _dbContext;
        private readonly ILogger<LoadBankService> _logger;

        public InvalidRecordRepository(DapperDbContext dbContext, ILogger<LoadBankService> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }

        public async Task<IEnumerable<InvalidRecord>> GetAllByIdAsync(int baiFileIDd)
        {

         
            const string sql = "SELECT * FROM DD_INVALID_REC WHERE BAI_FILE_ID = @baiFileId and RECORD_STATUS = 'A'";
            //const string sql = "SELECT * FROM DD_INVALID_REC WHERE BAI_FILE_ID = @baiFileId";

            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");
                var results = await ddsConnection.QueryAsync<InvalidRecord>(sql, new { baiFileId =baiFileIDd });
                _logger.LogInformation("Fetched {Count} invalid records", results.Count());
                return results;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching invalid records");
                throw;
            }
        }


        public async Task<IEnumerable<InvalidRecord>> GetAllAsync()
        {


            //const string sql = "SELECT * FROM DD_INVALID_REC dir join DD_INVALID_REC_ERROR dire on dir.INVALID_RECORD_ID = dire.INVALID_RECORD_ID WHERE  RECORD_STATUS = 'A'";
            const string sql = @"SELECT 
                                    dir.*,
                                    STRING_AGG(dire.INVALID_REC_ERR_MSG, ', ') AS ERROR_MESSAGES
                                FROM 
                                    DD_INVALID_REC dir
                                JOIN 
                                    DD_INVALID_REC_ERROR dire ON dir.INVALID_RECORD_ID = dire.INVALID_RECORD_ID
                                GROUP BY 
                                    dir.INVALID_RECORD_ID,
                                    dir.BAI_FILE_ID,
                                    dir.TOT_FUNB_BENEFIT_AMT,
                                    dir.DR_CR_FLAG,
                                    dir.AS_OF_DATETIME,
                                    dir.DECEASED_IND,
                                    dir.INCOMPLETE_POSTING_ERR_IND,
                                    dir.CREATED_BY,
                                    dir.CREATED_DATETIME,
                                    dir.RECORD_STATUS,
                                    dir.SHARED_DD_NUM_IND,
                                    dir.DD_NUM,
                                    dir.INSTITUTION_CODE,
                                    dir.AFFINITY_ACCT_NUM,
                                    dir.MEDICAL_RECORD_NUM,
                                    dir.PATIENT_NAME,
                                    dir.DISCHARGE_DATE,
                                    dir.FUNB_INCOME_SRC_TYPE,
                                    dir.LAST_MOD_BY,
                                    dir.LAST_MOD_DATETIME,
                                    dir.COMMENT
";

            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");
                var results = await ddsConnection.QueryAsync<InvalidRecord>(sql);
                _logger.LogInformation("Fetched {Count} invalid records", results.Count());
                return results;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching invalid records");
                throw;
            }
        }

        public async Task UpdateAsyncV1(InvalidRecord record)
        {
            const string sql = @"
        UPDATE DD_INVALID_REC
    SET 
        MEDICAL_RECORD_NUM = @MEDICAL_RECORD_NUM,
        AFFINITY_ACCT_NUM = @AFFINITY_ACCT_NUM,
        FUNB_INCOME_SRC_TYPE = @FUNB_INCOME_SRC_TYPE,
        AS_OF_DATETIME = @AS_OF_DATETIME,
        LAST_MOD_DATETIME = GETDATE()
    WHERE INVALID_RECORD_ID = @INVALID_RECORD_ID";

            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");
                await ddsConnection.ExecuteAsync(sql, record);
                _logger.LogInformation("Updated InvalidRecordId={Id} by {User}", record.INVALID_RECORD_ID, record.LAST_MOD_BY);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating invalid record ID {Id}", record.INVALID_RECORD_ID);
                throw;
            }
        }


        public async Task UpdateAsync(InvalidRecord record)
        {
            //Todo: system will replaced with actual username 
            record.LAST_MOD_BY = "System";


            const string sql = @"
            UPDATE DD_INVALID_REC
            SET 
                MEDICAL_RECORD_NUM = @MEDICAL_RECORD_NUM,
                AFFINITY_ACCT_NUM = @AFFINITY_ACCT_NUM,
                FUNB_INCOME_SRC_TYPE = @FUNB_INCOME_SRC_TYPE,
                AS_OF_DATETIME = @AS_OF_DATETIME
            WHERE INVALID_RECORD_ID = @INVALID_RECORD_ID;";

            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");
                await ddsConnection.ExecuteAsync(sql, new
                {
                    record.INVALID_RECORD_ID,
                    record.MEDICAL_RECORD_NUM,
                    record.AFFINITY_ACCT_NUM,
                    FUNB_INCOME_SRC_TYPE = record.FUNB_INCOME_SRC_TYPE,
                    record.AS_OF_DATETIME,
                    record.LAST_MOD_BY,
                });

            this.MoveInvalidRecordToWorkFile(record.INVALID_RECORD_ID);

                ddsConnection.Execute(
                    "DELETE FROM DD_INVALID_REC_ERROR WHERE INVALID_RECORD_ID = @Id",
                    new { Id = record.INVALID_RECORD_ID });


                ddsConnection.Execute(
                   "DELETE FROM DD_INVALID_REC WHERE INVALID_RECORD_ID = @Id",
                   new { Id = record.INVALID_RECORD_ID }
                   );
                ddsConnection.Execute(
                    "UPDATE DD_WORK_FILE SET VALIDATED = 'N' WHERE INVALID_RECORD_ID = @Id",
                    new { Id = record.INVALID_RECORD_ID }
                );

                _logger.LogInformation("Deleted InvalidRecordId={Id} and updated DD_WORK_FILE.RecordId={RecordId} by {User}",
                    record.INVALID_RECORD_ID, record.LAST_MOD_BY);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing InvalidRecordId={Id}", record.INVALID_RECORD_ID);
                throw;
            }
        }

        public async Task MoveToValidAsync(long invalidRecordId)
        {
            string userId = "System";
            string postingMode = "Deposit";

            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");

                var parameters = new DynamicParameters();
                parameters.Add("@dInvalidRecordID", invalidRecordId);
                parameters.Add("@sUserID", userId);
                parameters.Add("@sPostingMode", postingMode);
                parameters.Add("@ReturnVal", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

                await ddsConnection.ExecuteAsync(
                    "up_id_Move_Invalid_Rec",
                    parameters,
                    commandType: CommandType.StoredProcedure
                );

                int result = parameters.Get<int>("@ReturnVal");

                if (result == 0)
                {
                    _logger.LogInformation("Successfully moved InvalidRecordId={Id} to valid by {User} with PostingMode={Mode}",
                        invalidRecordId, userId, postingMode);
                }
                else
                {
                    _logger.LogWarning("Stored procedure returned code {Code} for InvalidRecordId={Id}", result, invalidRecordId);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Exception occurred while moving record ID {Id} to valid", invalidRecordId);
                throw;
            }
        }



        public async Task RecoupAsync(long creditId, long debitId, string userId)
        {
            const string sql = "EXEC up_recoup_DeceasedRecords @creditId, @debitId, @userId";

            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");
                await ddsConnection.ExecuteAsync(sql, new { creditId, debitId, userId });
                _logger.LogInformation("Recouped records CreditId={Credit}, DebitId={Debit}, User={User}", creditId, debitId, userId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error recouping CreditId={Credit} and DebitId={Debit}", creditId, debitId);
                throw;
            }
        }


        public void MoveInvalidRecordToWorkFile(long invalidRecordId)
        {
            using var ddsConnection = _dbContext.CreateConnection("dds_schema");

            var record = ddsConnection.QueryFirstOrDefault<InvalidRecord>(
                    "SELECT * FROM DD_INVALID_REC WHERE INVALID_RECORD_ID = @Id",
                    new { Id = invalidRecordId });
            if (record == null)
            {
                Console.WriteLine($"No record found for ID {invalidRecordId}");
                return;
            }
            // Step 2: INSERT into DD_WORK_FILE
            var insertQuery = @"
                INSERT INTO DD_WORK_FILE (
                    INVALID_RECORD_ID, BAI_FILE_ID, TOT_FUNB_BENEFIT_AMT, DR_CR_FLAG,
                    AS_OF_DATETIME, CREATED_BY, CREATED_DATETIME, RECORD_STATUS,
                    SHARED_DD_NUM_IND, DD_NUM, INSTITUTION_CODE, AFFINITY_ACCT_NUM,
                    MEDICAL_RECORD_NUM, NAME, DECEASED_IND, DISCHARGE_DATE,
                    COMMENT, VALIDATED
                ) VALUES (
                    @INVALID_RECORD_ID, @BAI_FILE_ID, @TOT_FUNB_BENEFIT_AMT, @DR_CR_FLAG,
                    @AS_OF_DATETIME, @CREATED_BY, @CREATED_DATETIME, @RECORD_STATUS,
                    @SHARED_DD_NUM_IND, @DD_NUM, @INSTITUTION_CODE, @AFFINITY_ACCT_NUM,
                    @MEDICAL_RECORD_NUM, @NAME, @DECEASED_IND, @DISCHARGE_DATE,
                    @COMMENT, @VALIDATED
                )";

            ddsConnection.Execute(insertQuery, new
            {
                record.INVALID_RECORD_ID,
                record.BAI_FILE_ID,
                record.TOT_FUNB_BENEFIT_AMT,
                record.DR_CR_FLAG,
                record.AS_OF_DATETIME,
                record.CREATED_BY,
                record.CREATED_DATETIME,
                RECORD_STATUS = record.RECORD_STATUS ?? "A",
                SHARED_DD_NUM_IND = record.SHARED_DD_NUM_IND ?? "N",
                record.DD_NUM,
                record.INSTITUTION_CODE,
                record.AFFINITY_ACCT_NUM,
                record.MEDICAL_RECORD_NUM,
                NAME = record.PATIENT_NAME,
                DECEASED_IND = record.DECEASED_IND ?? "N",
                record.DISCHARGE_DATE,
                record.COMMENT,
                VALIDATED = "N"
            });


            Console.WriteLine($"Record ID {invalidRecordId} inserted into DD_WORK_FILE.");
        }

        public async Task DeleteAsync(long invalidRecordId)
        {

            const string sql = @"



            Delete from DD_INVALID_REC_ERROR
            WHERE INVALID_RECORD_ID = @invalidRecordId;

            Delete from DD_INVALID_REC
            WHERE INVALID_RECORD_ID = @invalidRecordId;
   
            DELETE FROM DD_WORK_FILE
            WHERE INVALID_RECORD_ID = @invalidRecordId;";

            
            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");
                await ddsConnection.ExecuteAsync(sql, new { invalidRecordId ,userId=1});
                _logger.LogInformation("Soft-deleted InvalidRecordId={Id}", invalidRecordId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting InvalidRecordId={Id}", invalidRecordId);
                throw;
            }
        }

        public async Task UndeleteAsync(long invalidRecordId, string userId)
        {
            const string sql = @"
            UPDATE DD_INVALID_REC
            SET RECORD_STATUS = 'A',
                LAST_MOD_BY = @userId,
                LAST_MOD_DATETIME = GETDATE()
            WHERE INVALID_RECORD_ID = @invalidRecordId";

            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");
                await ddsConnection.ExecuteAsync(sql, new { invalidRecordId, userId });
                _logger.LogInformation("Restored InvalidRecordId={Id} by {User}", invalidRecordId, userId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error restoring InvalidRecordId={Id}", invalidRecordId);
                throw;
            }
        }




        public async Task<IEnumerable<InvalidRecord>> GetHiddeTransactionsAsync()
        {
            //const string sql = "SELECT * FROM DD_INVALID_REC WHERE BAI_FILE_ID = @baiFileId and RECORD_STATUS = 'A'";
            const string sql = "SELECT * FROM DD_INVALID_REC WHERE RECORD_STATUS = 'I'";

            try
            {
                using var ddsConnection = _dbContext.CreateConnection("dds_schema");
                var results = await ddsConnection.QueryAsync<InvalidRecord>(sql);

                if (_logger != null)
                    _logger.LogInformation("Fetched {Count} invalid records", results.Count());

                return results;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error fetching invalid records");
                throw;
            }

        }

        public async Task<int> HidePreEditRecordAsync(long invalidRecordId, string recordStatus, string userId)
        {
            _logger.LogInformation("Calling stored procedure up_u_Hide_PreEdit for record ID {RecordId}", invalidRecordId);

            const string procedureName = "up_u_Hide_PreEdit";
            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema");

                var parameters = new DynamicParameters();
                parameters.Add("@dInvalidRecordID", invalidRecordId);
                parameters.Add("@sRecordStatus", recordStatus);
                parameters.Add("@sUserID", userId);
                parameters.Add("@RETURN_VALUE", dbType: System.Data.DbType.Int32, direction: System.Data.ParameterDirection.ReturnValue);

                await connection.ExecuteAsync(procedureName, parameters, commandType: CommandType.StoredProcedure);

                int returnValue = parameters.Get<int>("@RETURN_VALUE");
                _logger.LogInformation("Stored procedure returned {ReturnValue}", returnValue);

                return returnValue;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing stored procedure up_u_Hide_PreEdit");
                throw;
            }
        }


    }



}
