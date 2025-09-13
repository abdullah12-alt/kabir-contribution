namespace Server.Shared
{
    public class Constants {
        public static class Messages {

            public const string ValidationSuccess = "Validation completed successfully.";
            public const string ValidationFailed = "Validation failed.";
            public const string InvalidInstitutionCode = "Invalid Institution Code:";
            public const string NoUnvalidatedRecords = "No unvalidated records found for this BAI file.";
            public const string NoMatchingPatient = "No Matching Patient:";
            public const string ValidationError = "Validation failed due to an error: ";
            public const string ErrorValidatingRecords = "Error validating records for BAI File ID";
            public const string TransactionForDeceasedPatient = "Transaction for deceased patient.";
            public const string UserInactive = "User is inactive.";
            public const string InvalidCredentials = "Invalid credentials.";
            public const string UserNotExist = "User does not exist.";
            public const string LoginSuccessful = "Login successful.";
            public const string PasswordChange = "Password changed successfully.";
            public const string InvalidFileId = "Invalid file bai Id";
            public const string IncomeSourceMismatch = "No ATP/PML record matches the income source type coming in";
            public const string DebitAmount = "Benefit has a debit amount";
            public const string DDNumberBlank = "DD Number is blank";
            public const string INVALID_DD_NUM = "Bank DD Number does not match an existing DD Number.";
            public const string TwoOrMorePatientsWithSameMRUN = "Two or more patients setup with the same MRUN";
            public const string ACCOUNT_NUMBER_BLANK = "Visit found has a blank Account Number";
            public const string DISTRIBUTION_AMTS_EQUAL_ZERO = "Both Personal Funds and Patient Distribution Amounts equal zero";
            public const string POSSIBLE_STIMULUS_AMOUNT = "Potential Stimulus Amount Received";
            
            public const string WORK_FILE_STORED_PROC_FAILED = "work file stored proc failed";
            public const string PATIENT_DECEASED = "Patient is deceased";
            public const string NO_VISITS_FOR_PATIENT = "Could not locate patient visits";
            public const string NO_VALID_ATP_PML_RECORD_FOR_DATE = "No valid ATP/PML Rates setup for the required date";
            public const string PERSONAL_FUNDS_ACCT_NOT_ESTABLISHED = "Personal Funds Account has not been set up for this account";
            public const string NO_VALID_RESPONSIBLE_PAYEE = "No valid responsible payee set up for this DD Number";


        }

        public static class RecordTypes
        {
            public const string FILE_HDR_REC = "01";
            public const string GROUP_HDR_REC = "02";
            public const string ACCOUNT_HDR_REC = "03";
            public const string TRANS_DTL_REC = "16";
            public const string ACCOUNT_TRLR_REC = "49";
            public const string CONTINUE_REC = "88";
            public const string GROUP_TRLR_REC = "98";
            public const string FILE_TRLR_REC = "99";
            public const string MODULE = "Load FUNB Transaction File - ";
        }
        public static class BaiTypeCodes
        {
            public const string TotalCredits = "100";
            public const string TotalDebits = "400";
            public const string OpeningLedger = "010";
            public const string ClosingLedger = "015";
            public const string OpeningAvailable = "040";
            public const string ClosingAvailable = "045";
        }

    }
}
