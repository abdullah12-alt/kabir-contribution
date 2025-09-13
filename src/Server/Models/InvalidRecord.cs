namespace Server.Models
{
    public class InvalidRecord
    {
        public long INVALID_RECORD_ID { get; set; }
        public long BAI_FILE_ID { get; set; }
        public decimal TOT_FUNB_BENEFIT_AMT { get; set; }
        public string? DR_CR_FLAG { get; set; }
        public DateTime? AS_OF_DATETIME { get; set; }
        public string? DECEASED_IND { get; set; }
        public string? INCOMPLETE_POSTING_ERR_IND { get; set; }
        public string? CREATED_BY { get; set; }
        public DateTime CREATED_DATETIME { get; set; }
        public string? RECORD_STATUS { get; set; }
        public string? SHARED_DD_NUM_IND { get; set; }
        public string? DD_NUM { get; set; }
        public string? INSTITUTION_CODE { get; set; }
        public string? AFFINITY_ACCT_NUM { get; set; }
        public string? MEDICAL_RECORD_NUM { get; set; }
        public string? PATIENT_NAME { get; set; }
        public DateTime? DISCHARGE_DATE { get; set; }
        public string? FUNB_INCOME_SRC_TYPE { get; set; }
        public string? LAST_MOD_BY { get; set; }
        public DateTime? LAST_MOD_DATETIME { get; set; }
        public long? INVALID_REC_ERR_MSG_ID { get; set; }
        public string?INVALID_REC_ERR_MSG { get; set; }
        public string? COMMENT { get; set; }
    }

}
