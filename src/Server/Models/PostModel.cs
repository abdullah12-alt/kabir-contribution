namespace Server.Models
{
    public class PostingAck
    {
        public long POSTING_ACK_ID { get; set; }
        public long INVALID_RECORD_ID { get; set; }
        public DateTime CREATED_DATETIME { get; set; }
        public long MSG_CONTROL_ID { get; set; }
        public string? PA_POSTING_STATUS { get; set; }
        public string? PF_POSTING_STATUS { get; set; }
        public int? PA_ERR_CODE { get; set; }
        public int? PF_ERR_CODE { get; set; }
    }
    public class PostCounts
    {
        public int TotalRecords { get; set; }
        public int TotalPATrans { get; set; }
        public int TotalPFTrans { get; set; }
    }

}
