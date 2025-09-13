using CsvHelper;
using Server.Exceptions;
using Server.Mappings;
using Server.Models;
using Server.Repositories;
using Server.Shared;
using System.Globalization;
using static Server.Shared.Constants;

namespace Server.Services
{

    public interface ILoadBankService
    {
        Task<(bool Success, string Message, string? BaiFileId)> ProcessFiles(IFormFile baiFile, IFormFile detailFile, string userId);
        Task<bool> AnyWorkFileRecordsAsync();
    }
    public class LoadBankService : ILoadBankService
    {
        private readonly ILoadBankRepository _repository;
        private readonly ILogger<LoadBankService> _logger;
        public LoadBankService(IConfiguration configuration, ILoadBankRepository repository, ILogger<LoadBankService> logger)
        {
            _repository = repository;
            _logger = logger;
        }
        public async Task<(bool Success, string Message, string? BaiFileId)> ProcessFiles(IFormFile baiFile, IFormFile detailFile, string userId)
        {
            _logger.LogInformation("Starting ProcessFiles: {BaiFileName}, {DetailFileName}, {User}", baiFile?.FileName, detailFile?.FileName, userId);
           

            try
            {
                _logger.LogInformation("Parsing detail file: {DetailFileName}", detailFile?.FileName);

                // Parse detail file (CSV)
                var detailRecords = await ParseDetailFile(detailFile!);

                _logger.LogInformation("Calling LoadBankFile. BaiFile: {BaiFile}, RecordCount: {RecordCount}, User: {User}",
                    baiFile?.FileName,
                    detailRecords.Count,
                    userId);


                // Process BAI file and get result
                var (success, baiFileId) = await LoadBankFile(baiFile!, detailRecords, userId);

                var message = success ? "Files processed successfully" : "Error processing files";

                _logger.LogInformation("File processing result: Success = {Success}, BaiFileId = {BaiFileId}, Message = {Message}",success, baiFileId, message);

                return (success, message, success ? baiFileId : null);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Exception occurred during ProcessFiles | BaiFile: {BaiFile}, DetailFile: {DetailFile}, User: {UserId}",
                    baiFile?.FileName,
                    detailFile?.FileName,
                    userId);


                return (false, $"Error processing files: {ex.Message}", null);
            }
        }

        private Task<List<DetailRecord>> ParseDetailFile(IFormFile detailFile)
        {
            using var stream = detailFile.OpenReadStream();
            using var reader = new StreamReader(stream);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            csv.Context.RegisterClassMap<DetailRecordMap>();
            var records = csv.GetRecords<DetailRecord>().ToList();
            return Task.FromResult(records);
        }

        private async Task<(bool Success, string? BaiFileId)> LoadBankFile(IFormFile baiFile, List<DetailRecord> detailRecords, string userId)
        {
            _logger.LogInformation("Starting file processing. BaiFile: {BaiFile}, DetailFileRecordCount: {DetailFileRecordCount}, User: {User}",
                baiFile?.FileName,
                detailRecords.Count,
                userId);

            var lines = await ReadBaiFileLines(baiFile!);
            var ctx = new BaiProcessingContext();

            foreach (var record in lines)
            {
                var fields = record.Split(',');
                _logger.LogInformation("Processing record type: {RecordType}", fields[0]);

                switch (fields[0])
                {
                    case RecordTypes.FILE_HDR_REC:
                        await HandleFileHeaderAsync(fields, userId, ctx);
                        break;
                    case RecordTypes.GROUP_HDR_REC:
                        HandleGroupHeader(fields, ctx);
                        break;
                    case RecordTypes.ACCOUNT_HDR_REC:
                        HandleAccountHeader(fields, ctx);
                        break;
                    case RecordTypes.TRANS_DTL_REC:
                        HandleTransDetail(fields, ctx);
                        break;
                    case RecordTypes.CONTINUE_REC:
                        HandleContinue(fields, ctx);
                        break;
                    case RecordTypes.ACCOUNT_TRLR_REC:
                        HandleAccountTrailer(ctx);
                        break;
                    case RecordTypes.GROUP_TRLR_REC:
                        await HandleGroupTrailerAsync(detailRecords, userId, ctx);
                        break;
                    case RecordTypes.FILE_TRLR_REC:
                        HandleFileTrailer(ctx);
                        break;
                }
            }

            return (true, ctx.BaiFileId);
        }

        private async Task<List<string>> ReadBaiFileLines(IFormFile baiFile)
        {
            var lines = new List<string>();
            try
            {
                using var stream = baiFile.OpenReadStream();
                using var reader = new StreamReader(stream);
                string? line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    if (!string.IsNullOrEmpty(line))
                        lines.Add(line);
                }

                _logger.LogInformation("Successfully read BAI file. Line Count: {Count}, File Name: {FileName}", lines.Count, baiFile.FileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to read BAI file: {FileName}", baiFile.FileName);
                throw;
            }
            return lines;
        }

        private async Task HandleFileHeaderAsync(string[] fields, string userId, BaiProcessingContext ctx)
        {
            _logger.LogInformation("Processing FILE_HDR_REC record");

            if (ctx.InFile)
                throw new BaiFileException(BaiFileErrorCode.UnexpectedFileHdr);

            if (fields[0] != RecordTypes.FILE_HDR_REC)
                throw new BaiFileException(BaiFileErrorCode.NoHeaderRecord);

            ctx.LastRecord = RecordTypes.FILE_HDR_REC;

            var config = await _repository.GetConfigAsync();
            if (config == null)
                throw new BaiFileException(BaiFileErrorCode.CouldNotGetConfigInfo);

            if (fields[1].Trim() != config.SENDER_ID)
                throw new BaiFileException(BaiFileErrorCode.SenderIdRcvdInvalid);

            if (fields[2].Trim() != config.RECEIVER_ID)
                throw new BaiFileException(BaiFileErrorCode.ReceiverIdRcvdInvalid);

            var rawDate = fields[3].Trim();
            try
            {
                ctx.BaiFileDate = rawDate.Length == 6
                    ? DateTime.ParseExact(rawDate, "yyMMdd", CultureInfo.InvariantCulture)
                    : DateTime.ParseExact(rawDate, "yyyyMMdd", CultureInfo.InvariantCulture);

                _logger.LogInformation("Parsed BAI file date: {Date}", ctx.BaiFileDate);
            }
            catch (FormatException)
            {
                throw new BaiFileException(BaiFileErrorCode.InvalidDateFormat);
            }

            ctx.FileNumber = int.Parse(ctx.BaiFileDate!.Value.ToString("yyyyMMdd"));

            var (procResult, tempBaiFileId) = await _repository.InsertBaiFileSummaryAsync(ctx.BaiFileDate.Value, ctx.FileNumber, userId);
            if (procResult == (int)LoadErrors.FILE_PROCESSED_ALREADY)
                throw new Exception($"File ID {ctx.FileNumber} has already been processed previously.");
            else if (procResult != 0)
                throw new Exception($"Insert failed for file ID {ctx.FileNumber}");

            ctx.BaiFileId = tempBaiFileId;
            ctx.InFile = true;
        }

        private void HandleGroupHeader(string[] fields, BaiProcessingContext ctx)
        {
            _logger.LogInformation("Processing GROUP_HDR_REC record");

            if (ctx.InGroup)
                throw new BaiFileException(BaiFileErrorCode.UnexpectedGroupHdr);

            ctx.InGroup = true;
            ctx.MiscDebits = 0;
            ctx.MiscCredits = 0;
            ctx.AsOfDate = fields[4].Trim();
            var asOfTime = fields.ElementAtOrDefault(5)?.Trim();
            _logger.LogDebug("Parsed GROUP_HDR_REC fields | Context: {@Context}", new { AsOfDate = ctx.AsOfDate, AsOfTime = asOfTime, CustomerNumber = fields.ElementAtOrDefault(1) });

            DateTime fileDate;
            try
            {
                fileDate = ctx.AsOfDate!.Length == 6
                    ? DateTime.ParseExact(ctx.AsOfDate, "yyMMdd", CultureInfo.InvariantCulture)
                    : DateTime.ParseExact(ctx.AsOfDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                _logger.LogInformation("Parsed GROUP_HDR_REC as-of date: {Date}", fileDate.ToString("yyyy-MM-dd"));
            }
            catch
            {
                throw new BaiFileException(BaiFileErrorCode.InvalidDateFormat);
            }

            LogBaiFileSummary(ctx.BaiFileId!, ctx.FileNumber, 0m, 0m, 0m, 0m, fileDate);
            ctx.LastRecord = RecordTypes.GROUP_HDR_REC;
        }

        private void HandleAccountHeader(string[] fields, BaiProcessingContext ctx)
        {
            if (ctx.InAccount)
            {
                _logger.LogError("At ACCOUNT_HDR_REC --> Uexpected Account Header");
                throw new BaiFileException(BaiFileErrorCode.UnexpectedAccountHdr);
            }

            if (!ctx.InGroup)
            {
                _logger.LogError("At ACCOUNT_HDR_REC --> No Group Header Received");
                throw new BaiFileException(BaiFileErrorCode.NoGroupHdrReceived);
            }

            _logger.LogInformation("Processing ACCOUNT_HDR_REC for account {CustomerAccountNumber}", ctx.CustomerAccountNumber);

            ctx.InAccount = true;
            ctx.CheckDebits = 0;
            ctx.CheckCredits = 0;
            ctx.CustomerAccountNumber = fields.ElementAtOrDefault(1)?.Trim();

            for (int i = 3; i < fields.Length; i += 4)
            {
                var typeCode = fields.ElementAtOrDefault(i)?.Trim();
                var amountStr = fields.ElementAtOrDefault(i + 1)?.Trim();
                if (!string.IsNullOrEmpty(typeCode) && !string.IsNullOrEmpty(amountStr))
                {
                    if (!decimal.TryParse(amountStr, out var parsedAmount))
                    {
                        _logger.LogWarning("Failed to parse amount in ACCOUNT_HDR_REC | TypeCode: {TypeCode}, AmountStr: {AmountStr}", typeCode, amountStr);
                        continue;
                    }

                    parsedAmount /= 100;
                    switch (typeCode)
                    {
                        case BaiTypeCodes.TotalCredits: ctx.TotalCredits = parsedAmount; break;
                        case BaiTypeCodes.TotalDebits: ctx.TotalDebits = parsedAmount; break;
                        case BaiTypeCodes.OpeningLedger: ctx.OpeningLedger = parsedAmount; break;
                        case BaiTypeCodes.ClosingLedger: ctx.ClosingLedger = parsedAmount; break;
                        case BaiTypeCodes.OpeningAvailable: ctx.OpeningAvailable = parsedAmount; break;
                        case BaiTypeCodes.ClosingAvailable: ctx.ClosingAvailable = parsedAmount; break;
                    }
                    _logger.LogDebug("Processed account header field | Context: {@Context}", new { TypeCode = typeCode, Amount = parsedAmount });
                }
            }
            ctx.LastRecord = RecordTypes.ACCOUNT_HDR_REC;
        }

        private void HandleTransDetail(string[] fields, BaiProcessingContext ctx)
        {
            _logger.LogInformation("Processing TRANS_DTL_REC record");
            ctx.LastRecord = RecordTypes.TRANS_DTL_REC;

            var typeCode = fields.ElementAtOrDefault(1)?.Trim();
            var amountStr = fields.ElementAtOrDefault(2)?.Trim();

            if (!int.TryParse(typeCode, out int transCode))
            {
                _logger.LogError("Invalid transaction code encountered | TypeCode: {TypeCode}", typeCode);
                throw new BaiFileException(BaiFileErrorCode.InvalidTransactionCode, typeCode);
            }

            if (!decimal.TryParse(amountStr, out decimal rawAmount))
                throw new BaiFileException(BaiFileErrorCode.InvalidAmount, amountStr);

            var amount = Math.Round(rawAmount / 100, 2);
            if ((transCode >= 200 && transCode < 300) || (transCode >= 400 && transCode < 500) || (transCode >= 600 && transCode < 700))
                ctx.MiscDebits += amount;

            if ((transCode >= 300 && transCode < 400) || (transCode >= 500 && transCode < 600))
                ctx.MiscCredits += amount;

            var type = (transCode >= 200 && transCode < 700) ? "Debit/Credit" : "Other";
            _logger.LogInformation("Processed transaction detail: TransactionCode= {Code}, Amount= {Amount} ,Type={Type}", transCode, amount, type);
        }

        private void HandleContinue(string[] fields, BaiProcessingContext ctx)
        {
            switch (ctx.LastRecord)
            {
                case RecordTypes.ACCOUNT_HDR_REC:
                    _logger.LogInformation("Processed continuation record (ACCOUNT_HDR_REC) | Summary: {@Summary}", new
                    {
                        ctx.TotalDebits,
                        ctx.TotalCredits,
                        ctx.OpeningLedger,
                        ctx.ClosingLedger,
                        ctx.OpeningAvailable,
                        ctx.ClosingAvailable
                    });

                    for (int i = 1; i <= 21; i += 4)
                    {
                        var typeCode = fields.ElementAtOrDefault(i)?.Trim();
                        var amountStr = fields.ElementAtOrDefault(i + 1)?.Trim();

                        if (!string.IsNullOrEmpty(typeCode) && !string.IsNullOrEmpty(amountStr) &&
                            decimal.TryParse(amountStr, out var rawAmount))
                        {
                            var amount = rawAmount / 100;
                            switch (typeCode)
                            {
                                case BaiTypeCodes.OpeningLedger: ctx.OpeningLedger = amount; break;
                                case BaiTypeCodes.ClosingLedger: ctx.ClosingLedger = amount; break;
                                case BaiTypeCodes.OpeningAvailable: ctx.OpeningAvailable = amount; break;
                                case BaiTypeCodes.ClosingAvailable: ctx.ClosingAvailable = amount; break;
                                case BaiTypeCodes.TotalCredits: ctx.TotalCredits = amount; break;
                                case BaiTypeCodes.TotalDebits: ctx.TotalDebits = amount; break;
                                default:
                                    break;
                            }
                        }
                    }
                    break;
            }
        }

        private void HandleAccountTrailer(BaiProcessingContext ctx)
        {
            _logger.LogInformation("Processing ACCOUNT_TRLR_REC record");
            if (!ctx.InGroup || !ctx.InAccount)
            {
                _logger.LogError("Unexpected ACCOUNT_TRLR_REC: Invalid state | Context: {@Context}", new { ctx.InGroup, ctx.InAccount });
                throw new BaiFileException(BaiFileErrorCode.UnexpectedAccountTrlr);
            }
            ctx.LastRecord = RecordTypes.ACCOUNT_TRLR_REC;
            ctx.InAccount = false;
        }

        private async Task HandleGroupTrailerAsync(List<DetailRecord> detailRecords, string userId, BaiProcessingContext ctx)
        {
            _logger.LogInformation("Processing GROUP_TRLR_REC record");
            if (!ctx.InGroup)
            {
                _logger.LogError("Unexpected GROUP_TRLR_REC: Not in group | Context: {@Context}", new { ctx.LastRecord });
                throw new BaiFileException(BaiFileErrorCode.UnexpectedGroupTrlr);
            }
            ctx.InGroup = false;
            ctx.LastRecord = RecordTypes.GROUP_TRLR_REC;

            var fileDate = DateTime.ParseExact(ctx.AsOfDate!, "yyMMdd", CultureInfo.InvariantCulture);
            _logger.LogInformation("Parsed file date for group trailer | Context: {@Context}", new { FileDate = fileDate });

            var filteredDetails = detailRecords.Where(r => r.AsOfDate.Date == fileDate.Date).ToList();
            _logger.LogInformation("Filtered detail records | Context: {@Context}", new { DetailRecordCount = filteredDetails.Count, AsOfDate = fileDate });

            ctx.CheckDebits = filteredDetails.Sum(r => r.DebitAmount);
            ctx.CheckCredits = filteredDetails.Sum(r => r.CreditAmount);
            _logger.LogInformation("Calculated check debits and credits | Context: {@Context}", new { ctx.CheckDebits, ctx.CheckCredits });

            bool isEmpty = ctx.CheckDebits == 0 && ctx.CheckCredits == 0;
            bool hasDetails = filteredDetails.Any();
            if (!isEmpty && hasDetails)
            {
                var miscDebitsMissing = false;
                var miscCreditsMissing = false;

                if (ctx.CheckDebits != ctx.TotalDebits)
                {
                    if (ctx.CheckDebits != ctx.TotalDebits - ctx.MiscDebits)
                    {
                        _logger.LogError("Debit mismatch detected: {CheckDebits} != {TotalDebits} | Context: {@Context}", ctx.CheckDebits, ctx.TotalDebits, new { MiscDebits = ctx.MiscDebits });
                        throw new BaiFileException(BaiFileErrorCode.AccountDebitsDontMatch);
                    }
                    miscDebitsMissing = true;
                }

                if (ctx.CheckCredits != ctx.TotalCredits)
                {
                    if (ctx.CheckCredits != ctx.TotalCredits - ctx.MiscCredits)
                    {
                        _logger.LogError("Credit mismatch detected: {CheckCredits} != {TotalCredits} | Context: {@Context}", ctx.CheckCredits, ctx.TotalCredits, new { MiscCredits = ctx.MiscCredits });
                        throw new BaiFileException(BaiFileErrorCode.AccountCreditsDontMatch);
                    }
                    miscCreditsMissing = true;
                }

                _logger.LogInformation("Detected missing misc credits | Context: {@Context}", new { ctx.MiscCredits });

                foreach (var detail in filteredDetails)
                {
                    string incomeSourceType = NormalizeIncomeSource(detail.EntryDescription!);
                    string ddNum = ExtractDDNumber(detail.RecipientID!, incomeSourceType);
                    string comment = $"{detail.RecipientID}\n{detail.FirstAddenda}\n{detail.RecipientName}";

                    _logger.LogInformation("Processing detail record: {RecipientID} | Context: {@Context}", detail.RecipientID, new { IncomeSourceType = incomeSourceType, DDNum = ddNum });
                    LogDetailRecordInfo(detail, comment);

                    await _repository.InsertWorkFileAsync(new WorkFileInsert
                    {
                        BaiFileId = decimal.Parse(ctx.BaiFileId!),
                        TotalFunbBenefitAmount = detail.CreditAmount > 0 ? detail.CreditAmount : detail.DebitAmount,
                        DrCrFlag = detail.CreditAmount > 0 ? "CR" : "DR",
                        AsOfDate = fileDate,
                        CreatedBy = userId,
                        DdNumber = ddNum,
                        InstitutionCode = null,
                        AffinityAccountNumber = null,
                        MedicalRecordNumber = null,
                        IncomeSourceType = incomeSourceType,
                        Name = null,
                        DischargeDate = null,
                        Comment = comment,
                        PaPostingStatus = null,
                        PfPostingStatus = null,
                        PaErrCode = null,
                        PfErrCode = null
                    });
                }

                if (miscDebitsMissing)
                {
                    await _repository.InsertWorkFileAsync(new WorkFileInsert
                    {
                        BaiFileId = decimal.Parse(ctx.BaiFileId!),
                        TotalFunbBenefitAmount = ctx.MiscDebits,
                        DrCrFlag = "DR",
                        AsOfDate = fileDate,
                        CreatedBy = userId,
                        DdNumber = "",
                        InstitutionCode = null,
                        AffinityAccountNumber = null,
                        MedicalRecordNumber = null,
                        IncomeSourceType = "BANK DEBIT",
                        Name = null,
                        DischargeDate = null,
                        Comment = "Total Misc Debits",
                        PaPostingStatus = null,
                        PfPostingStatus = null,
                        PaErrCode = null,
                        PfErrCode = null
                    });
                }

                if (miscCreditsMissing)
                {
                    await _repository.InsertWorkFileAsync(new WorkFileInsert
                    {
                        BaiFileId = decimal.Parse(ctx.BaiFileId!),
                        TotalFunbBenefitAmount = ctx.MiscCredits,
                        DrCrFlag = "CR",
                        AsOfDate = fileDate,
                        CreatedBy = userId,
                        DdNumber = "",
                        InstitutionCode = null,
                        AffinityAccountNumber = null,
                        MedicalRecordNumber = null,
                        IncomeSourceType = "BANK CREDT",
                        Name = null,
                        DischargeDate = null,
                        Comment = "Total Misc Credits",
                        PaPostingStatus = null,
                        PfPostingStatus = null,
                        PaErrCode = null,
                        PfErrCode = null
                    });
                }
            }

            _logger.LogInformation("Calling stored procedure up_iud_BAI_File_Summary (Update) | Context: {@Context}", new
            {
                Parameters = new
                {
                    ctx.BaiFileId,
                    FileIdNum = ctx.FileNumber,
                    AvailableBal = ctx.ClosingAvailable,
                    CollectedBal = 0m,
                    FunbTotalCredits = ctx.TotalCredits,
                    FunbTotalDebits = ctx.TotalDebits,
                    LedgerBal = ctx.ClosingLedger,
                    BaiFileDate = fileDate.ToString("yyyy-MM-dd"),
                    UpdateStatus = "U"
                }
            });

            await _repository.UpdateBaiFileSummaryAsync(ctx.FileNumber, ctx.BaiFileId!, ctx.ClosingAvailable, 0m, ctx.TotalCredits, ctx.TotalDebits, ctx.ClosingLedger, fileDate);
            LogBaiFileSummary(ctx.BaiFileId!, ctx.FileNumber, ctx.ClosingAvailable, ctx.TotalCredits, ctx.TotalDebits, ctx.ClosingLedger, fileDate);
        }

        private void HandleFileTrailer(BaiProcessingContext ctx)
        {
            _logger.LogInformation("Processing FILE_TRLR_REC record");
            if (!ctx.InFile || ctx.InGroup || ctx.InAccount)
                throw new BaiFileException(BaiFileErrorCode.UnexpectedFileTrlrRec);
            ctx.LastRecord = RecordTypes.FILE_TRLR_REC;
            ctx.InFile = false;
            ctx.CheckCredits = 0;
            ctx.CheckDebits = 0;
        }

        private class BaiProcessingContext
        {
            public bool InFile { get; set; }
            public bool InGroup { get; set; }
            public bool InAccount { get; set; }
            public string? BaiFileId { get; set; }
            public int FileNumber { get; set; }
            public string LastRecord { get; set; } = string.Empty;
            public decimal TotalDebits { get; set; }
            public decimal TotalCredits { get; set; }
            public decimal MiscDebits { get; set; }
            public decimal MiscCredits { get; set; }
            public string? AsOfDate { get; set; }
            public string? CustomerAccountNumber { get; set; }
            public DateTime? BaiFileDate { get; set; }
            public decimal CheckDebits { get; set; }
            public decimal CheckCredits { get; set; }
            public decimal OpeningLedger { get; set; }
            public decimal ClosingLedger { get; set; }
            public decimal OpeningAvailable { get; set; }
            public decimal ClosingAvailable { get; set; }
        }
        private string NormalizeIncomeSource(string source)
        {
            if (source.StartsWith("XX")) source = source[2..];
            return source == "VA BENEF" ? "VA BENEFIT" : source;
        }

        private string ExtractDDNumber(string recipientId, string incomeSourceType)
        {
            var parts = recipientId.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (incomeSourceType == "CIV SERV" && parts.Length > 1)
                return parts[1];
            return parts[0];
        }

        public void LogDetailRecordInfo(DetailRecord detail, string comment)
        {
            _logger.LogInformation("Inserting record into DDWorkFile | Context: \n" +
                $"As Of Date: {detail.AsOfDate:yyyy-MM-dd}\n" +
                $"As Of Time: {detail.AsOfTime}\n" +
                $"Debit Amount: {detail.DebitAmount:C}\n" +
                $"Credit Amount: {detail.CreditAmount:C}\n" +
                $"Entry Description: {detail.EntryDescription}\n" +
                $"Recipient ID: {detail.RecipientID}\n" +
                $"First Addenda: {detail.FirstAddenda}\n" +
                $"Recipient Name: {detail.RecipientName}\n" +
                $"Bank ID: {detail.BankID}\n" +
                $"Bank Name: {detail.BankName}\n" +
                $"Account Number: {detail.AccountNumber}\n" +
                $"Account Type: {detail.AccountType}\n" +
                $"Currency: {detail.Currency}\n" +
                $"Sending Company ID: {detail.SendingCompanyID}\n" +
                $"Sending Company Name: {detail.SendingCompanyName}\n" +
                $"Trace Number: {detail.TraceNumber}\n" +
                $"Entry Class Code: {detail.EntryClassCode}\n" +
                $"Comment: {comment}\n"
            );
        }
        
        public void LogBaiFileSummary(string baiFileId, long fileNumber, decimal closingAvailable, decimal totalCredits, decimal totalDebits, decimal closingLedger, DateTime fileDate)
        {
            _logger.LogInformation($"Updating BAI File Summary | Context:\n" +
                $"Bai File ID: {baiFileId}\n" +
                $"File Number: {fileNumber}\n" +
                $"Available Balance: {closingAvailable:C}\n" +
                $"Collected Balance: 0.00\n" +
                $"Total Credits: {totalCredits:C}\n" +
                $"Total Debits: {totalDebits:C}\n" +
                $"Ledger Balance: {closingLedger:C}\n" +
                $"BAI File Date: {fileDate:yyyy-MM-dd}\n" +
                $"Update Status: U\n"
            );
        }


        public async Task<bool> AnyWorkFileRecordsAsync()
        {
            return await _repository.AnyWorkFileRecordsAsync();
        }




    }


}



