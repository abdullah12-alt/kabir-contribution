using Dapper;
using DDS.API.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Server.Models;
using Server.Repositories;
using System.Data;
namespace Server.Services
{
    public interface IOverrideService
    {
        Task<OverrideResponse> ValidateAccountAsync(long invalidRecordId, OverrideRequest request);
        Task MoveRecordToValidAsync(long invalidRecordId, OverrideRequest request, string userId);
        Task<InvalidRecord?> GetInvalidRecordAsync(long invalidRecordId);
    }
    public class OverrideService : IOverrideService
    {
        private readonly DapperDbContext _dbContext;
        private readonly ILogger<OverrideService> _logger;

        public OverrideService(DapperDbContext dbContext, ILogger<OverrideService> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }
      

        public async Task<OverrideResponse> ValidateAccountAsync(long invalidRecordId, OverrideRequest request)
        {
            try
            {
                // Get the invalid record
                var invalidRecord = await GetInvalidRecordAsync(invalidRecordId);
                if (invalidRecord != null)
                //if (invalidRecord == null)
                {
                    return new OverrideResponse
                    {
                        Success = false,
                        Message = "Invalid record not found"
                    };
                }

                // Validate ATP/PML
                if (request.AtpPml != "A" && request.AtpPml != "P")
                {
                    return new OverrideResponse
                    {
                        Success = false,
                        Message = "ATP/PML must be either 'A' or 'P'"
                    };
                }

                // Validate amounts
                if (request.AccountAmount <= 0)
                {
                    return new OverrideResponse
                    {
                        Success = false,
                        Message = "Account amount must be greater than 0"
                    };
                }

                if (request.PersonalFundsAmount < 0)
                {
                    return new OverrideResponse
                    {
                        Success = false,
                        Message = "Personal funds amount must be 0 or greater"
                    };
                }

                // Validate total amount
                if (request.AccountAmount + request.PersonalFundsAmount != invalidRecord!.TOT_FUNB_BENEFIT_AMT)
                {
                    return new OverrideResponse
                    {
                        Success = false,
                        Message = "The Account Amount + PF Amount must equal the FUNB Total Amount"
                    };
                }

                // Validate account number format
                if (request.AccountNumber.Length != 8 || !request.AccountNumber.All(char.IsDigit))
                {
                    return new OverrideResponse
                    {
                        Success = false,
                        Message = "Account number must be eight digits"
                    };
                }

                // Determine hospital by account number
                var (hospitalCode, hospitalDsn) = DetermineHospitalByAccount(request.AccountNumber);
                if (string.IsNullOrEmpty(hospitalCode))
                {
                    return new OverrideResponse
                    {
                        Success = false,
                        Message = "Cannot determine the institution"
                    };
                }

                // Validate account in Affinity
                var affinityInfo = await ValidateAccountInAffinityAsync(request.AccountNumber, hospitalDsn);
                if (!affinityInfo.Success)
                {
                    return new OverrideResponse
                    {
                        Success = false,
                        Message = affinityInfo.Message
                    };
                }

                // Validate Personal Funds account if needed
                if (request.PersonalFundsAmount > 0)
                {
                    var pfValidation = await ValidatePersonalFundsAccountAsync(affinityInfo.MedicalRecordNumber);
                    if (!pfValidation)
                    {
                        return new OverrideResponse
                        {
                            Success = false,
                            Message = "Personal Funds Account not established"
                        };
                    }
                }

                // Return successful validation with record details
                return new OverrideResponse
                {
                    Success = true,
                    Message = "The record is now ready to be forced",
                    Record = new OverrideRecord
                    {
                        DdNumber = invalidRecord.DD_NUM ?? "",
                        IncomeSource = invalidRecord.FUNB_INCOME_SRC_TYPE ?? "",
                        PatientName = affinityInfo.PatientName,
                        InstitutionCode = hospitalCode,
                        AffinityAccountNumber = affinityInfo.AccountNumber,
                        MedicalRecordNumber = affinityInfo.MedicalRecordNumber,
                        DebitCreditFlag = invalidRecord.DR_CR_FLAG ?? "",
                        AtpPmlFlag = request.AtpPml,
                        FunbAsOfDate = invalidRecord.AS_OF_DATETIME,
                        CreatedDate = invalidRecord.CREATED_DATETIME,
                        DeceasedIndicator = affinityInfo.DeathFlag,
                        FunbAmount = invalidRecord.TOT_FUNB_BENEFIT_AMT,
                        AccountAmount = request.AccountAmount,
                        PersonalFundsAmount = request.PersonalFundsAmount,
                        VisitId = affinityInfo.VisitId
                    }
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating account for InvalidRecordId={Id}", invalidRecordId);
                return new OverrideResponse
                {
                    Success = false,
                    Message = "Error validating account"
                };
            }
        }

        public async Task MoveRecordToValidAsync(long invalidRecordId, OverrideRequest request, string userId)
        {
            const string proc = "up_iu_ValidRecords";

            var parameters = new DynamicParameters();
            parameters.Add("work_file_record_id", invalidRecordId);
            parameters.Add("valid_record_id", null);
            parameters.Add("income_source_type_id", null);
            parameters.Add("bai_file_id", null);
            parameters.Add("institution_code", null);
            parameters.Add("affinity_acct_num", request.AccountNumber);
            parameters.Add("affinity_visit_id", null);
            parameters.Add("medical_record_num", null);
            parameters.Add("dd_num", null);
            parameters.Add("atp_pml_flag", request.AtpPml);
            parameters.Add("affinity_atp_rate_id", 0);
            parameters.Add("tot_funb_benefit_amt", null);
            parameters.Add("dr_cr_flag", null);
            parameters.Add("as_of_datetime", null);
            parameters.Add("patient_name", null);
            parameters.Add("pf_distribution_amt", request.PersonalFundsAmount);
            parameters.Add("pa_distribution_amt", request.AccountAmount);
            parameters.Add("deceased_ind", null);
            parameters.Add("created_by", userId);
            parameters.Add("tot_days_inhouse", 0);
            parameters.Add("spec_proc_cond_hash_tot", 0);
            parameters.Add("sent_for_posting_datetime", null);
            parameters.Add("update_status", "I");
            parameters.Add("posted_to_affinity", null);
            parameters.Add("override", "Y");
            parameters.Add("RETURN_VALUE", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema");
                await connection.ExecuteAsync(proc, parameters, commandType: CommandType.StoredProcedure);

                int result = parameters.Get<int>("RETURN_VALUE");
                if (result != 0)
                {
                    throw new Exception($"Error forcing record. Return value: {result}");
                }

                _logger.LogInformation("Executed {Proc} for InvalidRecordId={Id}, User={User}", proc, invalidRecordId, userId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing {Proc} for InvalidRecordId={Id}", proc, invalidRecordId);
                throw;
            }
        }
        public async Task<InvalidRecord?> GetInvalidRecordAsync(long invalidRecordId)
        {
            const string sql = @"
            SELECT BAI_FILE_ID, DD_NUM, DD_INVALID_REC.FUNB_INCOME_SRC_TYPE, 
                   DD_INCOME_SOURCE_TYPE.INCOME_SOURCE_TYPE_ID, PATIENT_NAME, 
                   INSTITUTION_CODE, AFFINITY_ACCT_NUM, MEDICAL_RECORD_NUM, 
                   DR_CR_FLAG, COMMENT, TOT_FUNB_BENEFIT_AMT, AS_OF_DATETIME, 
                   DD_INVALID_REC.CREATED_DATETIME, DECEASED_IND, SHARED_DD_NUM_IND, 
                   INVALID_RECORD_ID, RECORD_STATUS, CREATED_BY, LAST_MOD_BY, 
                   LAST_MOD_DATETIME, INVALID_REC_ERR_MSG_ID, INVALID_REC_ERR_MSG, 
                   DISCHARGE_DATE
            FROM DD_INVALID_REC 
            INNER JOIN DD_INCOME_SOURCE_TYPE 
                ON DD_INVALID_REC.FUNB_INCOME_SRC_TYPE = DD_INCOME_SOURCE_TYPE.FUNB_INCOME_SRC_TYPE
            WHERE INVALID_RECORD_ID = @InvalidRecordId";

            try
            {
                using var connection = _dbContext.CreateConnection("dds_schema");
                var result = await connection.QueryFirstOrDefaultAsync<InvalidRecord>(sql, new { InvalidRecordId = invalidRecordId });

                if (result == null)
                {
                    _logger.LogWarning("No invalid record found for ID={Id}", invalidRecordId);
                }
                else
                {
                    _logger.LogInformation("Fetched invalid record ID={Id}", invalidRecordId);
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting invalid record ID={Id}", invalidRecordId);
                return null;
            }
        }



        private async Task<bool> ValidatePersonalFundsAccountAsync(string medicalRecordNumber)
        {
            const string sql = "SELECT PATIENT_ID FROM PF_PATIENT WHERE MEDICAL_RECORD_NUM = @MedicalRecordNumber";

            try
            {
                using var connection = _dbContext.CreateConnection("pfs_schema");
                var result = await connection.QueryFirstOrDefaultAsync(sql, new { MedicalRecordNumber = medicalRecordNumber });

                bool exists = result != null;
                _logger.LogInformation("PF account {Status} for MRUN={MRUN}",
                    exists ? "found" : "not found", medicalRecordNumber);

                return exists;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating Personal Funds account for MRUN={MRUN}", medicalRecordNumber);
                return false;
            }
        }

        private (string hospitalCode, string hospitalDsn) DetermineHospitalByAccount(string accountNumber)
        {
            var firstDigit = accountNumber[0];
            var secondDigit = accountNumber.Length > 1 ? accountNumber[1] : ' ';

            return firstDigit switch
            {
                '0' => ("0", "JUH"),
                '1' => ("1", "CHERRY"),
                '2' => ("2", "BROUGHT"),
                '3' => ("3", "DIX"),
                '4' => ("4", "MURDOCH"),
                '5' => ("5", "OBERRY"),
                '6' => ("6", "CASWELL"),
                '7' => ("7", "WCAROLINA"),
                '8' => ("9", "NCSCC"),
                '9' => secondDigit switch
                {
                    '0' => ("E", "BLACKMNT"),
                    '6' => ("H", "JFKADATC"),
                    '2' => ("Q", "WBJADATC"),
                    _ => ("", "")
                },
                _ => ("", "")
            };
        }

        private async Task<(bool Success, string Message, string AccountNumber, string MedicalRecordNumber, string PatientName, string DeathFlag, string VisitId)> ValidateAccountInAffinityAsync(string accountNumber, string hospitalDsn)
        {
            try
            {
          
                var connectionString = $"DRIVER={{InterSystems IRIS ODBC35}};SERVER=hes001.dhr.state.nc.us;PORT=1972;DATABASE={hospitalDsn};STATIC CURSORS=1;AUTHENTICATION METHOD=0;UID=dhhsHearts;PWD=apollo30;";

                using var connection = new SqlConnection(connectionString);
                var sql = @"
                    SELECT VISIT.PATIENT_ACCOUNT_NUMBER, VISIT.VISIT_ID, PATIENT.MRUN, 
                           PATIENT.NAME, PATIENT.DEATH_FLAG
                    FROM VISIT 
                    INNER JOIN PATIENT ON VISIT.PATIENT_ID = PATIENT.PATIENT_ID
                    WHERE VISIT.PATIENT_ACCOUNT_NUMBER = @AccountNumber";

                var result = await connection.QueryFirstOrDefaultAsync(sql, new { AccountNumber = accountNumber });

                if (result == null)
                {
                    return (false, "Account was not found in Affinity", "", "", "", "", "");
                }

                return (true, "",
                    result.PATIENT_ACCOUNT_NUMBER.ToString().PadLeft(8, '0'),
                    result.MRUN.ToString().PadLeft(7, '0'),
                    result.NAME?.ToString() ?? "",
                    result.DEATH_FLAG?.ToString() ?? "N",
                    result.VISIT_ID?.ToString() ?? "");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating account in Affinity for account {AccountNumber}", accountNumber);
                return (false, "Error validating account in Affinity", "", "", "", "", "");
            }
        }


     
    }
}
