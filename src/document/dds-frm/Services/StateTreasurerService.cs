using Server.Infrastructure.Logging;
using Server.Models;
using Server.Repositories;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace Server.Services;

public interface IStateTreasurerService
{
    Task<IReadOnlyList<DsnItem>> GetPriorDsnsAsync();
    Task<StateTreasurerStatusDto> GetStatusAsync(DateTime postedDate);
    Task<bool> DsnExistsWithinSixMonthsAsync(string depSeqNum);
    Task<int> InsertDsnAsync(InsertDsnRequest request, DateTime postedDate, string createdBy);
    Task<int> MarkSentToTreasurerAsync(MarkSentRequest request, string userId);
    Task<DailyTotals> GetTotalsAsync(DateTime postedDate);
    Task<IReadOnlyList<InstitutionTotal>> GetInstitutionTotalsAsync(DateTime postedDate);
    Task<AlogResponse> BuildAlogAsync(DateTime postedDate, string? sequenceNum = null, DateTime? processDate = null);
    Task<FileGenerationResponse> GenerateFilesAsync(FileGenerationRequest request);
    Task<EmailResponse> SendEmailAsync(EmailRequest request, string userId);
    Task<ProcessResponse> ProcessAsync(ProcessRequest request);
}

public class StateTreasurerService : IStateTreasurerService
{
    private readonly IStateTreasurerRepository _repository;
    private readonly ILogger<StateTreasurerService> _logger;
    private readonly IConfiguration _configuration;

    public StateTreasurerService(IStateTreasurerRepository repository, ILogger<StateTreasurerService> logger, IConfiguration configuration)
    {
        _repository = repository;
        _logger = logger;
        _configuration = configuration;
    }

    public Task<IReadOnlyList<DsnItem>> GetPriorDsnsAsync() => _repository.GetPriorDsnsAsync();

    public async Task<StateTreasurerStatusDto> GetStatusAsync(DateTime postedDate)
    {
        var hasTx = await _repository.HasTransactionsOnDateAsync(postedDate);
        var hasIncomplete = await _repository.HasIncompleteSendOnDateAsync(postedDate);
        var dsnRequired = await _repository.IsDsnRequiredAsync(postedDate);
        return new StateTreasurerStatusDto
        {
            HasTransactions = hasTx,
            HasIncompleteSend = hasIncomplete,
            IsDsnRequired = dsnRequired
        };
    }

    public Task<bool> DsnExistsWithinSixMonthsAsync(string depSeqNum) => _repository.DsnExistsWithinSixMonthsAsync(depSeqNum);

    public async Task<int> InsertDsnAsync(InsertDsnRequest request, DateTime postedDate, string createdBy)
    {
        if (request is null)
            throw new ArgumentNullException(nameof(request));

        var dsnRequired = await _repository.IsDsnRequiredAsync(postedDate);
        if (!dsnRequired)
        {
            // DSN not required → skip insert, return 0
            return 0;
        }

        // Validate DSN input
        if (string.IsNullOrWhiteSpace(request.DepSeqNum) ||
            !Regex.IsMatch(request.DepSeqNum, "^[A-Za-z0-9]+$"))
        {
            throw new ArgumentException("DSN must be alpha-numeric.", nameof(request.DepSeqNum));
        }

        // Duplicate check: last 6 months
        if (await _repository.DsnExistsWithinSixMonthsAsync(request.DepSeqNum))
        {
            throw new InvalidOperationException("This DSN exists within the last 6 months. Enter a different DSN.");
        }

        if (request.ProcessDate == default)
        {
            throw new ArgumentException("Process date is required.", nameof(request.ProcessDate));
        }

        // Insert DSN
        return await _repository.InsertDsnAsync(request.DepSeqNum, request.ProcessDate, createdBy);
    }

    public async Task<int> MarkSentToTreasurerAsync(MarkSentRequest request, string userId)
    {
        var result = await _repository.MarkSentToTreasurerAsync(request.PostedDate, request.DepSeqNum, "Y", userId);
        if (result != 0)
        {
            _logger.LogWarning("Mark sent returned non-zero result {Result} for date {Date}", result, request.PostedDate);
        }
        return result;
    }

    public Task<DailyTotals> GetTotalsAsync(DateTime postedDate) => _repository.GetDailyTotalsAsync(postedDate);

    public Task<IReadOnlyList<InstitutionTotal>> GetInstitutionTotalsAsync(DateTime postedDate) => _repository.GetInstitutionTotalsAsync(postedDate);

    public async Task<AlogResponse> BuildAlogAsync(DateTime postedDate, string? sequenceNum = null, DateTime? processDate = null)
    {
        var rows = await _repository.GetAlogRowsAsync(postedDate);
        var totals = await _repository.GetDailyTotalsAsync(postedDate);
        return new AlogResponse
        {
            PostedDate = postedDate.Date,
            SequenceNumber = sequenceNum,
            ProcessDate = processDate,
            Rows = rows.ToList(),
            Totals = totals
        };
    }

    public async Task<FileGenerationResponse> GenerateFilesAsync(FileGenerationRequest request)
    {
        try
        {
            var config = await _repository.GetConfigInfoAsync();
            var institutionTotals = await _repository.GetInstitutionTotalsAsync(request.PostedDate);
            var dailyTotals = await _repository.GetDailyTotalsAsync(request.PostedDate);

            // Generate PA file content
            var paContent = GeneratePaFileContent(config, dailyTotals, request);

            // Generate PF file content
            var pfContent = GeneratePfFileContent(config, institutionTotals, dailyTotals, request);

            return new FileGenerationResponse
            {
                PaFileName = "OSTMHPA.txt",
                PfFileName = "OSTMHLT.txt",
                PaFileContent = paContent,
                PfFileContent = pfContent,
                Totals = dailyTotals
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating files for {Date}", request.PostedDate);
            throw;
        }
    }

    private string GeneratePaFileContent(ConfigInfoTreasure config, DailyTotals totals, FileGenerationRequest request)
    {
        var lines = new List<string>();

        // Header record: HRT,OSTMHPA,O,MM/dd/yyyy,MM/dd/yyyy,amount000000000000000,69 spaces
        var paAmount = (totals.TotalPAAmt * 100).ToString("F0").PadLeft(15, '0');
        var headerLine = $"HRT,OSTMHPA,O,{request.PostedDate:MM/dd/yyyy},{request.ProcessDate:MM/dd/yyyy},{paAmount}{new string(' ', 69)}";
        lines.Add(headerLine);

        // Detail record: vendor_id(20),payee(40),stifno(20),amount(11),26 spaces
        var vendorId = config.PaVendorIdNum.PadRight(20);
        var payee = "Mental Health Summary".PadRight(40);
        var stifno = new string(' ', 20);
        var detailAmount = (totals.TotalPAAmt * 100).ToString("F0").PadLeft(11, '0');
        var detailLine = $"{vendorId},{payee},{stifno},{detailAmount}{new string(' ', 26)}";
        lines.Add(detailLine);
        return string.Join(Environment.NewLine, lines);
    }

    private string GeneratePfFileContent(ConfigInfoTreasure config, IReadOnlyList<InstitutionTotal> institutionTotals,
        DailyTotals totals, FileGenerationRequest request)
    {
        var lines = new List<string>();

        // Header record: HRT,OSTMHLT,O,MM/dd/yyyy,MM/dd/yyyy,amount000000000000000,69 spaces
        var pfAmount = (totals.TotalPFAmt * 100).ToString("F0").PadLeft(15, '0');
        var headerLine = $"HRT,OSTMHLT,O,{request.PostedDate:MM/dd/yyyy},{request.ProcessDate:MM/dd/yyyy},{pfAmount}{new string(' ', 69)}";
        lines.Add(headerLine);

        // Detail records for each institution with PF amounts
        foreach (var inst in institutionTotals.Where(x => x.TotalPFAmt != 0))
        {
            if (string.IsNullOrEmpty(inst.VendorId) || inst.VendorId == "Unknown")
                continue;

            var vendorId = inst.VendorId.PadRight(20);
            var payee = (inst.InstitutionName ?? "").PadRight(40);
            var stifno = new string(' ', 20);
            var detailAmount = (inst.TotalPFAmt * 100).ToString("F0").PadLeft(11, '0');
            var detailLine = $"{vendorId},{payee},{stifno},{detailAmount}{new string(' ', 26)}";
            lines.Add(detailLine);
        }

        return string.Join(Environment.NewLine, lines);
    }

    public async Task<EmailResponse> SendEmailAsync(EmailRequest request, string userId)
    {
        try
        {
            var config = await _repository.GetConfigInfoAsync();
            var response = new EmailResponse();

            // 1. Generate PA & PF files once
            var institutionTotals = await _repository.GetInstitutionTotalsAsync(request.PostedDate);
            var dailyTotals = await _repository.GetDailyTotalsAsync(request.PostedDate);
            var fileReq = new FileGenerationRequest
            {
                PostedDate = request.PostedDate,
                ProcessDate = DateTime.Now,
                DepSeqNum = request.DepSeqNum
            };

            var paContent = GeneratePaFileContent(config, dailyTotals, fileReq);
            var pfContent = GeneratePfFileContent(config, institutionTotals, dailyTotals, fileReq);

            // 2. TreasurerFiles folder
            var folderPath = Path.Combine(AppContext.BaseDirectory, "TreasurerFiles");
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            var paPath = Path.Combine(folderPath, "OSTMHPA.txt");
            var pfPath = Path.Combine(folderPath, "OSTMHLT.txt");

            if (File.Exists(paPath)) File.Delete(paPath);
            if (File.Exists(pfPath)) File.Delete(pfPath);

            await File.WriteAllTextAsync(paPath, paContent);
            await File.WriteAllTextAsync(pfPath, pfContent);

            // 2. Build subject and body
            var subject = string.IsNullOrEmpty(request.DepSeqNum)
                ? $"{config.StTreasEmailSubj} for {request.PostedDate:MM/dd/yyyy}"
                : $"{config.StTreasEmailSubj} for {request.PostedDate:MM/dd/yyyy} DSN # {request.DepSeqNum} on {DateTime.Now:MM/dd/yyyy HH:mm:ss}";

            var body = $"Attached are the OSTMHPA and OSTMHLT files for {request.PostedDate:MM/dd/yyyy}.";

            var smtpSection = _configuration.GetSection("Smtp");

            // 3. Create a helper method to send with attachments
            async Task sendToAsync(string recipient)
            {
                using var smtp = new SmtpClient
                {
                    Host = smtpSection["Host"],
                    Port = int.Parse(smtpSection["Port"]),
                    EnableSsl = bool.Parse(smtpSection["EnableSsl"]),
                    Credentials = new NetworkCredential(
                        smtpSection["UserName"],
                        smtpSection["Password"])
                };

                using var mail = new MailMessage
                {
                    From = new MailAddress(smtpSection["FromEmail"], smtpSection["FromName"]),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = false
                };

                mail.To.Add(recipient);

                // Correct attachment with MIME type
                mail.Attachments.Add(new Attachment(paPath, System.Net.Mime.MediaTypeNames.Text.Plain));
                mail.Attachments.Add(new Attachment(pfPath, System.Net.Mime.MediaTypeNames.Text.Plain));

                await smtp.SendMailAsync(mail);

                _logger.LogInformation("Email sent to {Recipient} with Treasurer files", recipient);
                response.SentTo.Add(recipient);
            }

            // 4. Send to Treasurer email (if configured)
            if (!string.IsNullOrEmpty(config.StTreasEmailToAddr))
            {
                await sendToAsync(config.StTreasEmailToAddr);
            }

            // 5. Send to each region email
            var regions = await _repository.GetRegionsAsync();
            foreach (var region in regions)
            {
                if (!string.IsNullOrEmpty(region.EmailRecipientsTo))
                {
                    await sendToAsync(region.EmailRecipientsTo);
                }
            }

            // 6. Cleanup
            System.IO.File.Delete(paPath);
            System.IO.File.Delete(pfPath);

            response.Attachments.Add("OSTMHPA.txt");
            response.Attachments.Add("OSTMHLT.txt");
            response.Success = true;
            response.Message = "Emails sent successfully";
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error sending State Treasurer emails for {Date}", request.PostedDate);
            return new EmailResponse
            {
                Success = false,
                Message = $"Error: {ex.Message}"
            };
        }
    }

    public async Task<ProcessResponse> ProcessAsync(ProcessRequest request)
    {
        try
        {
            var createdBy = "SYSTEM"; // TODO: replace with authenticated user
            var userId = "SYSTEM";
            
            // 1. Insert DSN
            var dsnResult = await InsertDsnAsync(request.Dsn, request.Dsn.PostedDate, createdBy);
            if (dsnResult != 0)
                return new ProcessResponse
                {
                    Success = false,
                    Message = $"Insert DSN failed with code {dsnResult}"
                };

            // 2. Generate files
            var fileReq = new FileGenerationRequest
            {
                PostedDate = request.Dsn.PostedDate,
                ProcessDate = DateTime.Now,
                DepSeqNum = request.Dsn.DepSeqNum
            };
            var files = await GenerateFilesAsync(fileReq);

            // 3. Send email
            var emailReq = new EmailRequest
            {
                PostedDate = request.Dsn.PostedDate,
                DepSeqNum = request.Dsn.DepSeqNum,
                SendAlogs = request.SendAlogs
            };
            var emailResp = await SendEmailAsync(emailReq, userId);

            return new ProcessResponse
            {
                Success = emailResp.Success,
                Message = emailResp.Success
                    ? "Process completed successfully"
                    : $"Email step failed: {emailResp.Message}",
                Files = files,
                EmailResponse = emailResp
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error running State Treasurer process");
            return new ProcessResponse
            {
                Success = false,
                Message = $"Process failed: {ex.Message}"
            };
        }
    }
}



