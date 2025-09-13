using System.ComponentModel.DataAnnotations;

namespace Server.Models;

public class ProcessRequest
{
    [Required]
    public InsertDsnRequest Dsn { get; set; } = new();
    public bool SendAlogs { get; set; } = false;
}

public class ProcessResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public FileGenerationResponse Files { get; set; } = new();
    public EmailResponse EmailResponse { get; set; } = new();
}

public class DsnItem
{
    public string DepSeqNum { get; set; } = string.Empty;
    public DateTime PostedDate { get; set; }
}

public class InstitutionTotal
{
    public string InstitutionCode { get; set; } = string.Empty;
    public string? VendorId { get; set; }
    public string? InstitutionName { get; set; }
    public decimal TotalPAAmt { get; set; }
    public decimal TotalPFAmt { get; set; }
}

public class DailyTotals
{
    public DateTime PostedDate { get; set; }
    public decimal TotalPAAmt { get; set; }
    public decimal TotalPFAmt { get; set; }
}

public class AlogRow
{
    public string InstitutionCode { get; set; } = string.Empty;
    public string NcasAccount { get; set; } = string.Empty;
    public string PaIncomeSourceType { get; set; } = string.Empty;
    public decimal TotalPAAmt { get; set; }
    public decimal TotalPFAmt { get; set; }
}

public class StateTreasurerStatusDto
{
    public bool HasTransactions { get; set; }
    public bool HasIncompleteSend { get; set; }
    public bool IsDsnRequired { get; set; }
}

public class InsertDsnRequest
{
    [Required]
    [MaxLength(20)]
    public string DepSeqNum { get; set; } = string.Empty;

    [Required]
    public DateTime ProcessDate { get; set; }

    [Required]
    public DateTime PostedDate { get; set; }
}

public class MarkSentRequest
{
    [Required]
    public DateTime PostedDate { get; set; }
    public string? DepSeqNum { get; set; }
    public string? UserId { get; set; }
}

public class AlogResponse
{
    public DateTime PostedDate { get; set; }
    public string? SequenceNumber { get; set; }
    public DateTime? ProcessDate { get; set; }
    public DailyTotals Totals { get; set; } = new();
    public List<AlogRow> Rows { get; set; } = new();
}

// Configuration models
public class ConfigInfoTreasure
{
    public string PaBatchName { get; set; } = string.Empty;
    public string PfBatchName { get; set; } = string.Empty;
    public string PaVendorIdNum { get; set; } = string.Empty;
    public string? StTreasEmailToAddr { get; set; }
    public string? StTreasEmailCcAddr { get; set; }
    public string? StTreasEmailSubj { get; set; }
    public string? StTreasEmailText { get; set; }
}

public class RegionInfo
{
    public string Region { get; set; } = string.Empty;
    public string? EmailRecipientsTo { get; set; }
    public string? EmailRecipientsCc { get; set; }
}

// Repository models
public class ConfigInfo
{
    public string PaBatchName { get; set; } = string.Empty;
    public string PfBatchName { get; set; } = string.Empty;
    public string PaVendorIdNum { get; set; } = string.Empty;
    public string? StTreasEmailToAddr { get; set; }
    public string? StTreasEmailCcAddr { get; set; }
    public string? StTreasEmailSubj { get; set; }
    public string? StTreasEmailText { get; set; }
}

// File generation models
public class FileGenerationRequest
{
    [Required]
    public DateTime PostedDate { get; set; }
    public DateTime ProcessDate { get; set; }
    public string? DepSeqNum { get; set; }
}

public class FileGenerationResponse
{
    public string PaFileName { get; set; } = string.Empty;
    public string PfFileName { get; set; } = string.Empty;
    public string PaFileContent { get; set; } = string.Empty;
    public string PfFileContent { get; set; } = string.Empty;
    public DailyTotals Totals { get; set; } = new();
}

public class EmailRequest
{
    [Required]
    public DateTime PostedDate { get; set; }
    public string? DepSeqNum { get; set; }
    public string? CustomMessage { get; set; }
    public bool SendAlogs { get; set; } = true;
}

public class EmailResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<string> SentTo { get; set; } = new();
    public List<string> Attachments { get; set; } = new();
}

