namespace Server.Models
{
    public class DDWorkFile
    {
        public int Id { get; set; }
        public int BaiFileId { get; set; }
        public decimal TotFunbBenefitAmt { get; set; }
        public string? DrCrFlag { get; set; }
        public DateTime AsOfDateTime { get; set; }
        public string? IncomeSourceType { get; set; }
        public string? Comment { get; set; }
        public string? DDNum { get; set; }
        public string? CreatedBy { get; set; }
        public string? SharedDDNumInd { get; set; }
        public string? DeceasedInd { get; set; }
        public string? RecordStatus { get; set; }
        public string? Validated { get; set; }
        public string? UpdateStatus { get; set; }
    }
}
