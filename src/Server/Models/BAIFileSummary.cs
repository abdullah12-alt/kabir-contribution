namespace Server.Models
{
    public class BAIFileSummary
    {
        public int BaiFileId { get; set; }
        public DateTime BaiFileDateTime { get; set; }
        public string? FileIdNum { get; set; }
        public decimal AvailableBal { get; set; }
        public decimal CollectedBal { get; set; }
        public decimal FunbTotalCredits { get; set; }
        public decimal FunbTotalDebits { get; set; }
        public decimal LedgerBal { get; set; }
        public string? CreatedBy { get; set; }
        public string? UpdateStatus { get; set; }
    }
}
