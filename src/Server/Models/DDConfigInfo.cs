namespace Server.Models
{
    // Models/DDConfigInfo.cs
    public class DDConfigInfo
    {
        public int CONFIG_ID { get; set; }
        public string? PA_VENDOR_ID_NUM { get; set; }
        public string? SENDER_ID { get; set; }
        public string? RECEIVER_ID { get; set; }
        public string? ST_TREAS_EMAIL_TO_ADDR { get; set; }
        public string? FT1_INSURANCE_CODE { get; set; }
        public string? PATCODE_ENTERING_AREA { get; set; }
        public string? ST_TREAS_EMAIL_CC_ADDR { get; set; }
        public string? ST_TREAS_EMAIL_TEXT { get; set; }
        public string? ST_TREAS_EMAIL_SUBJ { get; set; }
        public int? DATA_REFRESH_RATE { get; set; }
        public int? DATA_LOOKBACK { get; set; }
        public string? PA_BATCH_NAME { get; set; }
        public string? PF_BATCH_NAME { get; set; }
    }
}
