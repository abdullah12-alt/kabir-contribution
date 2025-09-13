namespace Client.Models
{
    public class BalanceSummaryDto
    {
        public decimal CarryOverBal { get; set; }
        public DateTime LastBalanceDate { get; set; }
        public decimal HiddenAdjustments { get; set; }
        public decimal PostedBenefits { get; set; }
        public decimal LedgerBalance { get; set; }
        public decimal InvalidTotal { get; set; }
        public decimal DeceasedTotal { get; set; }
        public decimal BeginningBalance { get; set; }
        public decimal EndingBalance { get; set; }
        public decimal AdjustedEndingBalance { get; set; }
        public decimal Difference { get; set; }

    }
    public class BalanceInsertDto
    {
        public decimal BeginningBalance { get; set; }
        public decimal CarryOverBalance { get; set; }
        public decimal TotalPosted { get; set; }
        public decimal EndingBalance { get; set; }
        public decimal InvalidTotal { get; set; }
        public decimal DeceasedTotal { get; set; }
        public decimal AdjustedEndingBalance { get; set; }
        public decimal LedgerBalance { get; set; }
        public string CreatedBy { get; set; }= string.Empty;
        public decimal Adjustments { get; set; }
    }
    public class SummaryRecordDto
    {
        public DateTime BAI_FILE_DATETIME { get; set; }
        public decimal LEDGER_BAL { get; set; }
        public decimal AVAILABLE_BAL { get; set; }
        public decimal COLLECTED_BAL { get; set; }
        public decimal FUNB_TOTAL_CREDITS { get; set; }
        public decimal FUNB_TOTAL_DEBITS { get; set; }
        public DateTime CREATED_DATETIME { get; set; }
        public string CREATED_BY { get; set; } = string.Empty;
    }
    public class BalanceRecord
    {
        public int BALANCE_ID { get; set; }
        public decimal? BEGINNING_BAL { get; set; }
        public decimal? CARRYOVER_BAL { get; set; }
        public decimal? TOT_CR_DR_POSTED { get; set; }
        public decimal? ENDING_BAL { get; set; }
        public decimal? TOT_CR_DR_PRE_EDIT_RPT { get; set; }
        public decimal? TOT_CR_DR_DECEASED_EXCEPT { get; set; }
        public decimal? ADJ_ENDING_BAL { get; set; }
        public decimal? LEDGER_BAL { get; set; }
        public string CREATED_BY { get; set; } = string.Empty;
        public DateTime? CREATED_DATETIME { get; set; }
        public decimal? TOT_CR_DR_ADJUSTMENTS { get; set; }
    }
    public class BalanceSummaryRowDto
    {
        public string Label { get; set; } = string.Empty;
        public decimal Value { get; set; }
    }
}
