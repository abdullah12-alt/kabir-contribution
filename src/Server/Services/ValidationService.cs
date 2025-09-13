using Dapper;
using DDS.API.Data;
using Server.Infrastructure.Logging;
using System.Data;
using static Server.Shared.Constants;

namespace Server.Services
{
    public interface IValidationService
    {
        Task<(bool success, List<string> errors)> ValidateRecordsAsync();
    }
    public class ValidationService : IValidationService
    {

        private readonly DapperDbContext _dbContext;
        private readonly IAppLogger<ValidationService> _logger;

        public ValidationService(DapperDbContext dbContext, IAppLogger<ValidationService> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }

        public async Task<(bool success, List<string> errors)> ValidateRecordsAsync()
        {
            _logger.LogInformation("Starting validation");

            using var ddsConnection = _dbContext.CreateConnection("dds_schema");
            using var pfsConnection = _dbContext.CreateConnection("pfs_schema");

            ddsConnection.Open();
            pfsConnection.Open();
            using var transaction = ddsConnection.BeginTransaction();
            try
            {
                _logger.LogInformation("Fetching unvalidated records from DD_WORK_FILE table");

                var unvalidatedRecords = await ddsConnection.QueryAsync<dynamic>(
                      "SELECT * FROM DD_WORK_FILE WHERE VALIDATED = 'N'",
                      null,
                      transaction
                  );



                if (unvalidatedRecords == null || !unvalidatedRecords.Any())
                {

                    _logger.LogWarning("No unvalidated records found");

                    return (false, new List<string> { Messages.NoUnvalidatedRecords });
                }
                _logger.LogInformation("Fetched unvalidated records", new { RecordsCount = unvalidatedRecords.Count() });


                var errorsList = new List<string>();


                foreach (var record in unvalidatedRecords)
                {

                    var errors = new List<string>();
                    string? patientName = null;
                    char multiMRUN = 'N';
                    bool specialAmount = false;
                    decimal toPatientAcct = 0;
                    decimal toPersonalFunds = 0;
                    bool allToPF = false;
                    //decimal benefitDiff = 0;
                    string? atp_pml_flag = null;
                    decimal? personalFundsAmount = null;
                    decimal? netAmount = null;

                    int specialConditionCode = 0;

                    string incomeSourceType = record?.INCOME_SOURCE_TYPE;

                    if (!string.IsNullOrEmpty(incomeSourceType))
                    {
                        if (incomeSourceType.Length >= 2 && incomeSourceType.Substring(0, 2) == "XX")
                        {
                            incomeSourceType = incomeSourceType.Substring(2);
                        }
                        else if (incomeSourceType == "VA BENEF")
                        {
                            incomeSourceType = "VA BENEFIT";
                        }
                    }
                    else
                    {
                        // Optionally log or handle missing income source type
                        incomeSourceType = string.Empty;
                    }


                    var incomeSourceTypeId = await ddsConnection.QuerySingleOrDefaultAsync<long?>(
                    "SELECT INCOME_SOURCE_TYPE_ID FROM DD_INCOME_SOURCE_TYPE WHERE FUNB_INCOME_SRC_TYPE = @incSourceType",
                    new { incSourceType = incomeSourceType }, transaction);

                    if (incomeSourceTypeId == null )
                    {
                        errors.Add($"{Messages.IncomeSourceMismatch} {incomeSourceType}");
                    }

                    // Business rule: Institution must exist in PFS

                    var atp_pml_info = await ddsConnection.QuerySingleOrDefaultAsync<dynamic>(
                              "SELECT TOP(1) * FROM DD_ATP_PML_INFO WHERE DD_NUMBER = @DirectDepositNumber",
                              new { DirectDepositNumber = record.DD_NUM }, transaction);

                    string? institutionCode = null;
                    if (atp_pml_info == null)
                    {
                        errors.Add($"{Messages.INVALID_DD_NUM} {record.DD_NUM}");
                    }
                    else
                    {
                        institutionCode = atp_pml_info.INSTITUTION_CODE;
                        atp_pml_flag = atp_pml_info.ATP_PML_FLAG;
                        personalFundsAmount = atp_pml_info.PERSONAL_FUNDS_AMOUNT;
                        netAmount = atp_pml_info.NET_AMOUNT;
                    }
                    if (atp_pml_info != null && record.AS_OF_DATETIME < atp_pml_info!.START_DATE)
                    {
                        errors.Add(Messages.NO_VALID_ATP_PML_RECORD_FOR_DATE);
                        _logger.LogWarning("Validation failed - No valid ATP/PML rate for date", new { DD_NUM = record.DD_NUM });
                    }


                    _logger.LogInformation("Fetched ATP_PML info", new { record.DD_NUM });

                    if (institutionCode == null)
                    {
                        string msg = $"{Messages.InvalidInstitutionCode} {record.INSTITUTION_CODE}";

                        errors.Add(msg);

                        _logger.LogWarning("Validation failed - Invalid Institution", new { msg });

                    }
                    if (record.DD_NUM == null)
                    {

                        string msg = $"{Messages.DDNumberBlank} {record.INSTITUTION_CODE}";
                        errors.Add(msg);
                        _logger.LogWarning("Validation failed - DD Number is Blank", new { msg });

                    }

                    // Business rule: Patient must exist in PFS

                    var patientMRUN = await ddsConnection.QuerySingleOrDefaultAsync<long?>(
                        "SELECT TOP(1) MRUN FROM DD_ATP_PML_INFO WHERE DD_NUMBER = @DirectDepositNumber",
                        new { DirectDepositNumber = record.DD_NUM }, transaction);


                    var patient = await pfsConnection.QueryFirstOrDefaultAsync<dynamic>(
                       "SELECT * FROM PF_PATIENT WHERE MEDICAL_RECORD_NUM = @medicalRecordNumber",
                       new { MedicalRecordNumber = patientMRUN });


                    if (patient == null)
                    {
                        string msg = $"{Messages.NoMatchingPatient} {record.MEDICAL_RECORD_NUM}";
                        errors.Add(msg);
                        _logger.LogWarning("Validation failed - Patient not found", new { msg });
                    }
                    else
                    {
                        patientName = $"{patient?.PATIENT_LAST_NAME}, {patient?.PATIENT_FIRST_NAME}";

                        _logger.LogWarning("Validation Logic - Patient Name", patientName);


                    }
                    // Business Rule: Patients with multiple MRUNS must go to Pre-Edit Report

                    var visitInfo = await ddsConnection.QueryAsync<dynamic>(
                        "SELECT * FROM DD_VISIT_INFO WHERE INSTITUTION_CODE = @InstCode AND MRUN = @MedicalRecordNumber ORDER BY ADMIT_ARRIVE_DATE DESC, DISCHARGE_DISPOSITION_DATE DESC",
                        new { InstCode = institutionCode, MedicalRecordNumber = patientMRUN }, transaction);

                    var checkForMultiPatientId = visitInfo.GroupBy(v => v.PATIENT_ID)
                                                          .Where(g => g.Count() > 1)
                                                          .Select(g => g.Key).ToList();
                    _logger.LogInformation("Checked for multiple patient IDs", new { checkForMultiPatientId });



                    if (checkForMultiPatientId.Count > 1)
                    {
                        multiMRUN = 'Y';
                        errors.Add($" {Messages.TwoOrMorePatientsWithSameMRUN} {patientMRUN}");
                    }
                    _logger.LogInformation("Fetched Visit Info", new { MRUN = patientMRUN, VisitCount = visitInfo.Count() });

                    // Business Rule: Affinity Account Number cannot be blank
                    string? accountNumber = null;
                    DateTime? dischargeDate = null;
                    if (visitInfo.Any())
                    {
                        accountNumber = visitInfo.First().ACCOUNT_NUMBER;
                        dischargeDate = visitInfo.First().DISCHARGE_DISPOSITION_DATE;
                    }

                    if (visitInfo == null || !visitInfo.Any())
                    {
                        errors.Add($"{Messages.NO_VISITS_FOR_PATIENT} {patientMRUN}");
                        _logger.LogWarning("Validation failed - No visits found for patient", new { patientMRUN });
                    }



                    if (accountNumber == null)
                    {
                        errors.Add($" {Messages.ACCOUNT_NUMBER_BLANK} {accountNumber}");
                    }

                    // Business rule: All debits need to go to the Pre-Edit Report
                    if (record.DR_CR_FLAG == "DR")
                    {
                        string msg = $" {Messages.DebitAmount} {record.DR_CR_FLAG}";
                        errors.Add(msg);
                        _logger.LogWarning("Debit-only record sent to pre-edit", new { msg });

                    }




                    // Business rule: Records of deceased patients must go to the Deceased Patient report 
                    var deceasedCheck = await ddsConnection.QuerySingleOrDefaultAsync<int>(
                            "SELECT COUNT(1) FROM DD_ATP_PML_INFO WHERE DD_NUMBER = @DirectDepositNumber AND DEATH_FLAG = 'Y'",
                            new { DirectDepositNumber = record.DD_NUM }, transaction);
                    _logger.LogInformation("Checked deceased status", new { IsDeceased = deceasedCheck > 0 });

                    var deceasedInd = 'N';
                    if (deceasedCheck > 0)
                    {
                        errors.Add(Messages.PATIENT_DECEASED);
                        deceasedInd = 'Y';
                    }


                    // If total benefit amount is $250 and from Social Security or VA, it could be a stimulus payment
                    if (record.TOT_FUNB_BENEFIT_AMT == 250 && new[] { "SSA ERP", "SOC SEC", "VA BENEFIT", "SUPP SEC" }.Contains(incomeSourceType))
                    {
                        specialAmount = true;
                        errors.Add($" {Messages.POSSIBLE_STIMULUS_AMOUNT}");
                    }

                    if (deceasedInd == 'Y' || incomeSourceType == "SUPP SEC" || specialAmount == true)
                    {
                        // Put all money into personal funds
                        toPatientAcct = 0;
                        toPersonalFunds = record.TOT_FUNB_BENEFIT_AMT;
                        allToPF = true;
                    }


                    if (specialAmount == false && dischargeDate != null)
                    {
                        if (atp_pml_flag == "A")
                        {
                            if ((dischargeDate - record.AS_OF_DATETIME).TotalDays > 60)
                            {
                                toPatientAcct = 0;
                                toPersonalFunds = record.TOT_FUNB_BENEFIT_AMT;
                                allToPF = true;
                            }
                        }
                        else if ((dischargeDate - record.AS_OF_DATETIME).TotalDays > 30)
                        {
                            toPatientAcct = 0;
                            toPersonalFunds = record.TOT_FUNB_BENEFIT_AMT;
                            allToPF = true;
                        }
                    }
                    // --- If special case (stimulus, deceased, etc.) or discharge timing triggers all-to-PF ---
                    if (allToPF)
                    {
                        toPatientAcct = 0;
                        toPersonalFunds = record.TOT_FUNB_BENEFIT_AMT;
                    }
                    else
                    {
                        decimal benefitAmt = record.TOT_FUNB_BENEFIT_AMT ?? 0;
                        decimal pfAmt = personalFundsAmount ?? 0;
                        decimal paAmt = netAmount ?? 0;
                        decimal atpPmlTotal = pfAmt + paAmt;

                        // Default assignments to be adjusted
                        toPersonalFunds = 0;
                        toPatientAcct = 0;

                        if (benefitAmt > atpPmlTotal)
                        {
                            // More money than expected — apply to PA first
                            toPatientAcct = paAmt;
                            toPersonalFunds = pfAmt;

                            decimal extra = benefitAmt - atpPmlTotal;
                            toPatientAcct += extra;

                            specialConditionCode = 1; // DDS logic: Net > ATP/PML
                        }
                        else if (benefitAmt < atpPmlTotal)
                        {
                            // Less money than expected — apply to PF first
                            toPersonalFunds = pfAmt;
                            toPatientAcct = paAmt;

                            decimal shortfall = atpPmlTotal - benefitAmt;

                            if (shortfall >= toPatientAcct)
                            {
                                toPersonalFunds = benefitAmt;
                                toPatientAcct = 0;
                            }
                            else
                            {
                                toPatientAcct -= shortfall;
                            }

                            specialConditionCode = 2; // DDS logic: Net < ATP/PML
                        }
                        else
                        {
                            // Exact match — no adjustment needed
                            toPatientAcct = paAmt;
                            toPersonalFunds = pfAmt;
                            specialConditionCode = 0;
                        }

                        // Validation check
                        if (toPatientAcct == 0 && toPersonalFunds == 0)
                            errors.Add($"{Messages.DISTRIBUTION_AMTS_EQUAL_ZERO}");

                        if (pfAmt == 0)
                        {
                            errors.Add(Messages.PERSONAL_FUNDS_ACCT_NOT_ESTABLISHED);
                            _logger.LogWarning("Validation failed - Personal Funds account not established", new { DD_NUM = record.DD_NUM });
                        }
                    }


                    if (errors.Any())
                    {
                        errorsList.AddRange(errors);
                        _logger.LogInformation("Inserting invalid record", new { recordId = record.RECORD_ID });
                        ///////////////////////////////////////////////////////////////////
                        ///parameters that are passing to up_iu_InvalidRecords proc///////
                        /////////////////////////////////////////////////////////////////
                        _logger.LogInformation("Calling stored procedure up_iu_InvalidRecords", new
                        {
                            Parameters = new
                            {
                                record.RECORD_ID,
                                record.BAI_FILE_ID,
                                record.TOT_FUNB_BENEFIT_AMT,
                                record.DR_CR_FLAG,
                                DeceasedInd = deceasedInd,
                                IncompletePostingErrInd = "N",
                                CreatedBy = "System",
                                RecordStatus = "A",
                                SharedDDNumInd = multiMRUN,
                                DDNum = record.DD_NUM,
                                InstitutionCode = record.INSTITUTION_CODE,
                                MedicalRecordNum = patientMRUN,
                                PatientName = $"{patient?.PATIENT_LAST_NAME}, @{patient?.PATIENT_FIRST_NAME}",
                                IncomeSource = record.INCOME_SOURCE_TYPE,
                                LastModBy = "System",
                                Comment = record.COMMENT,
                                UpdateStatus = "I"
                            }
                        });




                        DynamicParameters invalidRecordParams = new DynamicParameters();
                        invalidRecordParams.Add("@work_file_record_id", record.RECORD_ID);
                        invalidRecordParams.Add("@bai_file_id", record.BAI_FILE_ID);
                        invalidRecordParams.Add("@tot_funb_benefit_amt", record.TOT_FUNB_BENEFIT_AMT);
                        invalidRecordParams.Add("@dr_cr_flag", record.DR_CR_FLAG);
                        invalidRecordParams.Add("@as_of_datetime", record.AS_OF_DATETIME);
                        invalidRecordParams.Add("@deceased_ind", deceasedInd);
                        invalidRecordParams.Add("@incomplete_posting_err_ind", "N"); // Y = Yes, N = No
                        invalidRecordParams.Add("@created_by", "System");
                        invalidRecordParams.Add("@record_status", "A"); // A = Active, I = Inactive
                        invalidRecordParams.Add("@shared_dd_num_ind", multiMRUN);
                        invalidRecordParams.Add("@dd_num", record.DD_NUM);
                        invalidRecordParams.Add("@institution_code", institutionCode);
                        invalidRecordParams.Add("@affinity_acct_num", accountNumber);
                        invalidRecordParams.Add("@medical_record_num", patientMRUN);
                        invalidRecordParams.Add("@patient_name", patientName);
                        invalidRecordParams.Add("@discharge_date", dischargeDate);
                        invalidRecordParams.Add("@funb_income_src_type", incomeSourceType);
                        invalidRecordParams.Add("@last_mod_by", "Sysyem");
                        invalidRecordParams.Add("@comment", record.COMMENT);
                        invalidRecordParams.Add("@update_status", "I"); // Insert Mode
                        invalidRecordParams.Add("@invalid_record_id_OUTPUT", dbType: DbType.String, size: 14, direction: ParameterDirection.Output);
                        invalidRecordParams.Add("@proc_result", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

                        await ddsConnection.ExecuteAsync(
                             "up_iu_InvalidRecords",
                             invalidRecordParams,
                             commandType: CommandType.StoredProcedure,
                             transaction: transaction
                         );

                        var procResult = invalidRecordParams.Get<int>("@proc_result");
                        var insertedInvalidRecordId = invalidRecordParams.Get<string>("@invalid_record_id_OUTPUT");


                        await ddsConnection.ExecuteAsync(
                            @"UPDATE DD_WORK_FILE
                              SET INVALID_RECORD_ID = @invalidRecordId
                              WHERE RECORD_ID = @recordId",
                            new
                            {
                                invalidRecordId = insertedInvalidRecordId, // from OUTPUT param earlier
                                recordId = record.RECORD_ID                // original DD_WORK_FILE record
                            },
                            transaction: transaction
                        );



                        foreach (var errorMsg in errors)
                        {
                            await ddsConnection.ExecuteAsync(
                                @"INSERT INTO dbo.DD_INVALID_REC_ERROR (
                                      INVALID_RECORD_ID,
                                      INVALID_REC_ERR_MSG
                                  )
                                  VALUES (
                                      @invalid_record_id,
                                      @invalid_rec_err_msg
                                  );",
                                new
                                {
                                    invalid_record_id = insertedInvalidRecordId,
                                    invalid_rec_err_msg = errorMsg
                                },
                                transaction: transaction
                            );
                        }



                    }
                    else
                    {
                        ///////////////////////////////////////////////////////////////////
                        ///parameters that are passing to DD_VALID_REC proc///////
                        /////////////////////////////////////////////////////////////////

                        _logger.LogInformation("Inserting validated record into DD_VALID_REC", new
                        {
                            baiFileId = record.BAI_FILE_ID,
                            incomeSourceTypeId = incomeSourceTypeId,
                            institutionCode = institutionCode,
                            affinityAcctNum = visitInfo!.First().ACCOUNT_NUMBER,
                            medicalRecordNum = patientMRUN,
                            ddNum = record.DD_NUM,
                            totFunbBenefitAmt = record.TOT_FUNB_BENEFIT_AMT,
                            drCrFlag = record.DR_CR_FLAG,
                            asOfDatetime = record.AS_OF_DATETIME,
                            patientName = patientName,
                            pfDistributionAmt = toPersonalFunds,
                            paDistributionAmt = toPatientAcct,
                            deceasedInd = deceasedInd,
                            createdBy = "System",
                            createdDateTime = DateTime.UtcNow,
                            totDaysInhouse = visitInfo!.First().DAYS_IN_HOUSE,
                            specProcCondHashTot = specialConditionCode
                        });


                        //  Insert validated record into DDS DD_VALID_REC table
                        await ddsConnection.ExecuteAsync(
                                 @"INSERT INTO DD_VALID_REC 
                                (BAI_FILE_ID, INCOME_SOURCE_TYPE_ID, INSTITUTION_CODE, AFFINITY_ACCT_NUM, 
                                 MEDICAL_RECORD_NUM, DD_NUM, TOT_FUNB_BENEFIT_AMT, DR_CR_FLAG, AS_OF_DATETIME, 
                                 PATIENT_NAME, PF_DISTRIBUTION_AMT, PA_DISTRIBUTION_AMT, DECEASED_IND, 
                                 CREATED_BY, CREATED_DATETIME, TOT_DAYS_INHOUSE, SPEC_PROC_COND_HASH_TOT)
                              VALUES 
                                (@baiFileId, @incomeSourceTypeId, @institutionCode, @affinityAcctNum, 
                                 @medicalRecordNum, @ddNum, @totFunbBenefitAmt, @drCrFlag, @asOfDatetime, 
                                 @patientName, @pfDistributionAmt, @paDistributionAmt, @deceasedInd, 
                                 @createdBy, @createdDateTime, @totDaysInhouse, @specProcCondHashTot);",
                                 new
                                 {
                                     baiFileId = record.BAI_FILE_ID,
                                     incomeSourceTypeId = incomeSourceTypeId,
                                     institutionCode = institutionCode,
                                     affinityAcctNum = visitInfo!.First().ACCOUNT_NUMBER,
                                     medicalRecordNum = patientMRUN,
                                     ddNum = record.DD_NUM,
                                     totFunbBenefitAmt = record.TOT_FUNB_BENEFIT_AMT,
                                     drCrFlag = record.DR_CR_FLAG,
                                     asOfDatetime = record.AS_OF_DATETIME,
                                     patientName = patientName,
                                     pfDistributionAmt = toPersonalFunds,
                                     paDistributionAmt = toPatientAcct,
                                     deceasedInd = deceasedInd,
                                     createdBy = "System",
                                     createdDateTime = DateTime.UtcNow,
                                     totDaysInhouse = visitInfo!.First().DAYS_IN_HOUSE,
                                     specProcCondHashTot = specialConditionCode
                                 },
                                 transaction);
                    }
                }


                // Business logic: Mark all processed records as validated

                await ddsConnection.ExecuteAsync(
                  "UPDATE DD_WORK_FILE SET VALIDATED = 'Y'",
                  null,
                  transaction);

                transaction.Commit();
                // Final outcome based on presence of validation errors

                if (errorsList.Any())
                {
                    _logger.LogWarning("Validation completed with errors", new { errors = errorsList });

                    return (false, new List<string> { $"{Messages.ValidationFailed} {string.Join(", ", errorsList)}" });
                }
                _logger.LogInformation("Validation completed successfully");

                return (true, errorsList);
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError("Exception occurred during validation", ex, new { });

                return (false, new List<string> { Messages.ValidationError + ex.Message });
            }
        }




    }
}




