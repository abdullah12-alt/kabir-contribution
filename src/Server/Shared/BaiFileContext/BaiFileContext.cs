namespace Server.Shared.BaiFileContext
{
    public interface IBaiFileContext
    {
        string? BaiFileId { get; set; }
    }
    public class BaiFileContext : IBaiFileContext
    {
        public string? BaiFileId { get; set; }
    }
}
