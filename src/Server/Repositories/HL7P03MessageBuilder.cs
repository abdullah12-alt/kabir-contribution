using System;
using System.Text;

namespace Server.Repositories
{
    public static class HL7P03MessageBuilder
    {
        public static string Build(
            string sendingApp,
            string sendingFacility,
            string receivingApp,
            string receivingFacility,
            string messageControlId,
            DateTime recordedDateTime,
            string patientName,
            string patientIdInternal,
            string patientAccountNumber,
            string processingId,
            DateTime transactionDate,
            string transactionType,
            string transactionCode,
            string transactionQuantity,
            decimal transactionAmountExtended,
            string departmentCode,
            string insurancePlanId)
        {
            var hl7 = new StringBuilder();

            hl7.AppendLine($"MSH|^~\\&|{sendingApp}|{sendingFacility}|{receivingApp}|{receivingFacility}|{recordedDateTime:yyyyMMddHHmmss}||DFT^P03|{messageControlId}|{processingId}|2.3");
            hl7.AppendLine($"EVN|P03|{recordedDateTime:yyyyMMddHHmmss}");
            hl7.AppendLine($"PID|||{patientIdInternal}||{patientName}|||");
            hl7.AppendLine($"PV1||I|{receivingFacility}||||");
            hl7.AppendLine($"FT1|||{transactionDate:yyyyMMddHHmmss}||{transactionType}|{transactionCode}|{transactionQuantity}|{transactionAmountExtended:0.00}|{departmentCode}|{insurancePlanId}");
            hl7.AppendLine($"ZIN|{patientAccountNumber}");

            return hl7.ToString();
        }
    }
}
