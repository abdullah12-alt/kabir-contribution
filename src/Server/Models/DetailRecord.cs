namespace Server.Models
{
    public class DetailRecord
    {
        //public DateTime Date { get; set; }              // As Of Date + As of Time
        public DateTime AsOfDate { get; set; }           // As Of Date 
        public string? AsOfTime { get; set; }              // As of Time
        public decimal DebitAmount { get; set; }        // Debit Amount
        public decimal CreditAmount { get; set; }       // Credit Amount
        public string? EntryDescription { get; set; }    // Entry Description
        public string? RecipientID { get; set; }         // Recipient ID
        public string? FirstAddenda { get; set; }        // First Addenda
        public string? RecipientName { get; set; }       // Recipient Name
        public string? BankID { get; set; }              // Bank ID
        public string? BankName { get; set; }            // Bank Name
        public string? AccountNumber { get; set; }       // Account Number
        public string? AccountType { get; set; }         // Account Type
        public string? Currency { get; set; }            // Currency
        public string? SendingCompanyID { get; set; }    // Sending Company ID
        public string? SendingCompanyName { get; set; }  // Sending Company Name
        public string? TraceNumber { get; set; }         // Trace Number
        public string? EntryClassCode { get; set; }      // Entry Class Code
    }
}
