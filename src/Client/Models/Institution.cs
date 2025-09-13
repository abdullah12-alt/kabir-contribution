namespace Client.Models
{
    // Models/Institution.cs
    public class Institution
    {
        public long INSTITUTION_ID { get; set; }
        public string? INSTITUTION_CODE { get; set; }
        public string? INSTITUTION_NAME { get; set; }
        public string? DD_VENDOR_ID_NUM { get; set; }
        public string? AFFINITY_DB_NAME { get; set; }
        public string? DD_SEND_REPORT_TO { get; set; }
        public string? CREATED_BY { get; set; }
        public DateTime CREATED_DATETIME { get; set; }
        public string? LAST_MOD_BY { get; set; }
        public DateTime? LAST_MOD_DATE { get; set; }
        public string? RECORD_STATUS { get; set; }
    }
}
