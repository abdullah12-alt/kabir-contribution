namespace Server.Models
{
    public class LoadFileRequest
    {
        public IFormFile? BaiFile { get; set; }
        public IFormFile? DetailFile { get; set; }
    }
}
