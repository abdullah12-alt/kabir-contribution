namespace Server.Models
{
    public class BaiFileContext
    {
        public string BaiFileId { get; set; }
        public DateTime SetTime { get; set; } = DateTime.UtcNow;
    }

}
