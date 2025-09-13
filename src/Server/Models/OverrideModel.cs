using System.ComponentModel.DataAnnotations;

namespace Server.Models
{
    public class OverrideRequest
    {
    

        public string AccountNumber { get; set; } = string.Empty;

        

        public decimal AccountAmount { get; set; }



        public decimal PersonalFundsAmount { get; set; }



        public string AtpPml { get; set; } = string.Empty;
    }

    public class OverrideResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public OverrideRecord? Record { get; set; }
    }

    public class OverrideRecord
    {
        public string DdNumber { get; set; } = string.Empty;
        public string IncomeSource { get; set; } = string.Empty;
        public string PatientName { get; set; } = string.Empty;
        public string InstitutionCode { get; set; } = string.Empty;
        public string AffinityAccountNumber { get; set; } = string.Empty;
        public string MedicalRecordNumber { get; set; } = string.Empty;
        public string DebitCreditFlag { get; set; } = string.Empty;
        public string AtpPmlFlag { get; set; } = string.Empty;
        public DateTime? FunbAsOfDate { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string DeceasedIndicator { get; set; } = string.Empty;
        public decimal FunbAmount { get; set; }
        public decimal AccountAmount { get; set; }
        public decimal PersonalFundsAmount { get; set; }
        public string VisitId { get; set; } = string.Empty;
    }
}
