namespace Client.Models
{
    public class HidePreEditRequestDto
    {
        public long InvalidRecordId { get; set; }
        public string RecordStatus { get; set; } = string.Empty;
        public string UserId { get; set; } = string.Empty;
    }
}
