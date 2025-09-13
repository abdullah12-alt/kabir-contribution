// Models/RegionDto.cs
namespace Client.Models
{
    public class RegionDto
    {
        public long RegionId { get; set; }
        public string? Region { get; set; }
        public string? EmailRecipientsTo { get; set; }
        public string? EmailRecipientsCc { get; set; }
        public string? LastModBy { get; set; }
        public DateTime? LastModDatetime { get; set; }
    }
}
