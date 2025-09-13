using Server.Shared;

namespace Server.Exceptions
{
    public class BaiFileException: Exception
    {
        public BaiFileErrorCode ErrorCode { get; }
        public BaiFileException(BaiFileErrorCode errorCode, string? detail = null) : base(GetMessageForCode(errorCode, detail))
        {
            ErrorCode = errorCode;
        }

        private static string GetMessageForCode(BaiFileErrorCode code, string? detail = null)
        {
            return code switch
            {
                BaiFileErrorCode.NoHeaderRecord => "Invalid Bank file. No header record was found.",
                BaiFileErrorCode.SenderIdRcvdInvalid => "Invalid Bank file. Sender ID does not match DDS Configuration.",
                BaiFileErrorCode.ReceiverIdRcvdInvalid => "Invalid Bank file. Receiver ID does not match DDS Configuration.",
                BaiFileErrorCode.UnexpectedFileHdr => "Invalid Bank file. Unexpected File Header record.",
                BaiFileErrorCode.UnexpectedGroupHdr => "Invalid Bank file. Unexpected Group Header record.",
                BaiFileErrorCode.InvalidDateFormat => "Invalid BAI file. File date format is incorrect.",
                BaiFileErrorCode.CouldNotGetConfigInfo => "Could not access configuration information.",
                BaiFileErrorCode.FileProcessedAlready => "Bank File was previously processed.",
                BaiFileErrorCode.MissingFileHeaderDate => "The File Header date was not captured before the Group Header.",
                BaiFileErrorCode.UnexpectedAccountHdr => "Invalid Bank file. Unexpected Account Header record.",
                BaiFileErrorCode.NoGroupHdrReceived => "Invalid Bank file. No group header received.",
                BaiFileErrorCode.InvalidTransactionCode => $"Invalid transaction code: {detail}",
                BaiFileErrorCode.InvalidAmount => $"Invalid transaction amount: {detail}",
                BaiFileErrorCode.UnexpectedAccountTrlr => "Invalid Bank file. Unexpected Account Trailer record.",
                BaiFileErrorCode.UnexpectedGroupTrlr => "Invalid Bank file. Unexpected Group Trailer record.",
                BaiFileErrorCode.AccountCreditsDontMatch => "Invalid Bank file. Total credits in account summary do not equal the detail total credits.",
                BaiFileErrorCode.AccountDebitsDontMatch => "Invalid Bank file. Total debits in account summary do not equal the detail total debits.",
                BaiFileErrorCode.UnexpectedFileTrlrRec => "Invalid Bank file. Unexpected File Trailer record.",

                _ => "An unknown BAI file error occurred."
            };
        }
    }

}
