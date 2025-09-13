namespace Server.Shared
{
    public enum LoadErrors
    {
        NO_HEADER_RECORD = 2500,
        SENDER_ID_RCVD_INVALID = 2501,
        RECEIVER_ID_RCVD_INVALID = 2502,
        UNEXPECTED_FILE_HDR = 2503,
        UNEXPECTED_GROUP_HDR = 2504,
        UNEXPECTED_ACCOUNT_HDR = 2505,
        NO_GROUP_REC_RECEIVED = 2506,
        UNEXPECTED_TRANS_DTL = 2507,
        UNEXPECTED_ACCOUNT_TRLR = 2508,
        UNEXPECTED_GROUP_TRLR = 2509,
        UNEXPECTED_FILE_TRLR_REC = 2510,
        BAI_FILE_STORED_PROC_FAILED = 2511,
        FUNB_SUMMARY_STORED_PROC_FAILED = 2512,
        ACCOUNT_CREDITS_DONT_MATCH = 2513,
        ACCOUNT_DEBITS_DONT_MATCH = 2514,
        FILE_PROCESSED_ALREADY = 6014,
        NO_VALID_BAI_FILES_RECEIVED = 6015,
        COULD_NOT_GET_CONFIG_INFO = 6016,
        ERROR_CREATING_WORK_RECORD = 6017,
        WORK_FILE_STORED_PROC_FAILED = 6018,
        NO_FILE_TRAILER_RECEIVED = 6019,
        DEBITS_ON_DETAIL_NOT_MATCH = 6020,
        CREDITS_ON_DETAIL_NOT_MATCH = 6021
    }
    public enum BAIRecordType
    {
        FILE_HDR_REC = 1,    // 01 - File Header
        GROUP_HDR_REC = 2,   // 02 - Group Header
        ACCOUNT_HDR_REC = 3, // 03 - Account Header
        TRANS_DTL_REC = 16,  // 16 - Transaction Detail
        ACCOUNT_TRLR_REC = 49, // 49 - Account Trailer
        CONTINUE_REC = 88,    // 88 - Continuation Record
        GROUP_TRLR_REC = 98,  // 98 - Group Trailer
        FILE_TRLR_REC = 99    // 99 - File Trailer
    }
    public enum BaiFileErrorCode
    {
        None,
        NoHeaderRecord,
        SenderIdRcvdInvalid,
        ReceiverIdRcvdInvalid,
        UnexpectedFileHdr,
        UnexpectedGroupHdr,
        FileProcessedAlready,
        CouldNotGetConfigInfo,
        InvalidDateFormat,
        MissingFileHeaderDate,
        UnexpectedAccountHdr,    
        NoGroupHdrReceived,
        InvalidTransactionCode,   
        InvalidAmount,
        UnexpectedAccountTrlr,
        UnexpectedGroupTrlr,
        AccountCreditsDontMatch,  
        AccountDebitsDontMatch,
        UnexpectedFileTrlrRec
    }


}
