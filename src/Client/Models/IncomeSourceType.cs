namespace Client.Models
{
    // Models/IncomeSourceType.cs
    public class IncomeSourceType
    {
        public long INCOME_SOURCE_TYPE_ID { get; set; }
        public string? FUNB_INCOME_SRC_TYPE { get; set; }
        public string? INCOME_SRC_TYPE_DESCR { get; set; }
        public string? PA_INCOME_SRC_TYPE { get; set; }
        public string? PA_PMT_CODE { get; set; }
        public string? PA_PMT_REV_CODE { get; set; }
        public string? PF_DEP_TRANS_CODE { get; set; }
        public string? PF_DEP_REV_TRANS_CODE { get; set; }
        public string? NCAS_ACCOUNT { get; set; }
        public int? START_POS { get; set; }
        public int? LENGTH { get; set; }
        public string? CREATED_BY { get; set; }
        public DateTime CREATED_DATETIME { get; set; }
        public string? LAST_MOD_BY { get; set; }
        public DateTime? LAST_MOD_DATETIME { get; set; }
        public string? RECORD_STATUS { get; set; }
    }
}
