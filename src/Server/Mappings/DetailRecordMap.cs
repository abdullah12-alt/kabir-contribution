using CsvHelper.Configuration;
using Server.Models;

namespace Server.Mappings
{
    public sealed class DetailRecordMap : ClassMap<DetailRecord>
    {
        public DetailRecordMap()
        {
            // Map CSV columns to properties
            Map(m => m.AsOfDate).Name("As Of Date");
            Map(m => m.AsOfTime).Name("As of Time");
            Map(m => m.DebitAmount).Name("Debit Amount");
            Map(m => m.CreditAmount).Name("Credit Amount");
            Map(m => m.EntryDescription).Name("Entry Description");
            Map(m => m.RecipientID).Name("Recipient ID");
            Map(m => m.FirstAddenda).Name("First Addenda");
            Map(m => m.RecipientName).Name("Recipient Name");
            Map(m => m.BankID).Name("Bank ID");
            Map(m => m.BankName).Name("Bank Name");
            Map(m => m.AccountNumber).Name("Account Number");
            Map(m => m.AccountType).Name("Account Type");
            Map(m => m.Currency).Name("Currency");
            Map(m => m.SendingCompanyID).Name("Sending Company ID");
            Map(m => m.SendingCompanyName).Name("Sending Company Name");
            Map(m => m.TraceNumber).Name("Trace Number");
            Map(m => m.EntryClassCode).Name("Entry Class Code");
        }
    }
}
