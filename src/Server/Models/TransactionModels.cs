namespace Server.Models
{
    public class WorkTransactionDto
    {
        public long RECORD_ID { get; set; }
        public long? INVALID_RECORD_ID { get; set; }
        public long BAI_FILE_ID { get; set; }
        public decimal TOT_FUNB_BENEFIT_AMT { get; set; }
        public string? DR_CR_FLAG { get; set; }
        public DateTime? AS_OF_DATETIME { get; set; }
        public string? CREATED_BY { get; set; }
        public DateTime? CREATED_DATETIME { get; set; }
        public string? RECORD_STATUS { get; set; }
        public string? SHARED_DD_NUM_IND { get; set; }
        public string? DD_NUM { get; set; }
        public string? INSTITUTION_CODE { get; set; }
        public string? AFFINITY_ACCT_NUM { get; set; }
        public string? MEDICAL_RECORD_NUM { get; set; }
        public string? INCOME_SOURCE_TYPE { get; set; }
        public string? NAME { get; set; }
        public string? DECEASED_IND { get; set; }
        public DateTime? DISCHARGE_DATE { get; set; }
        public string? COMMENT { get; set; }
        public string? MSG_CONTROL_ID { get; set; }
        public string? PA_POSTING_STATUS { get; set; }
        public string? PF_POSTING_STATUS { get; set; }
        public int? PA_ERR_CODE { get; set; }
        public int? PF_ERR_CODE { get; set; }
        public string? VALIDATED { get; set; }
    }
    public class InvalidTransactionDto
    {
        public long INVALID_RECORD_ID { get; set; }
        public long BAI_FILE_ID { get; set; }
        public decimal TOT_FUNB_BENEFIT_AMT { get; set; }
        public string? DR_CR_FLAG { get; set; }
        public string? INSTITUTION_CODE_3 { get; set; }
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
        public string? INVALID_REC_ERR_MSG { get; set; }
        public string? COMMENT { get; set; }
    }
    public class ValidTransactionDto
    {
        public long VALID_RECORD_ID { get; set; }
        public string? INCOME_SOURCE_TYPE_ID { get; set; }
        public string? INSTITUTION_CODE_3 { get; set; }
        public long BAI_FILE_ID { get; set; }
        public string? INSTITUTION_CODE { get; set; }
        public string? AFFINITY_ACCT_NUM { get; set; }
        public string? MEDICAL_RECORD_NUM { get; set; }
        public string? DD_NUM { get; set; }
        public string? FUNB_INCOME_SRC_TYPE { get; set; }
        public decimal TOT_FUNB_BENEFIT_AMT { get; set; }
        public string? DR_CR_FLAG { get; set; }
        public DateTime AS_OF_DATETIME { get; set; }
        public string? PATIENT_NAME { get; set; }
        public decimal PF_DISTRIBUTION_AMT { get; set; }
        public decimal PA_DISTRIBUTION_AMT { get; set; }
        public string? DECEASED_IND { get; set; }
        public string? CREATED_BY { get; set; }

        public DateTime CREATED_DATETIME { get; set; }
        public int TOT_DAYS_INHOUSE { get; set; }
        public decimal SPEC_PROC_COND_HASH_TOT { get; set; }
        public DateTime SENT_FOR_POSTING_DATETIME { get; set; }
        public string? ATP_PML_FLAG { get; set; }
        public string? AFFINITY_VISIT_ID { get; set; }
        public string? AFFINITY_ATP_RATE_ID { get; set; }
        public string? POSTED_TO_AFFINITY { get; set; }
        public string? OVERRIDE { get; set; }
    }

}
