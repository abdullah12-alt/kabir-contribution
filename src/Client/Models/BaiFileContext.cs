namespace Client.Models
{
    using System.Text.Json.Serialization;

    public class BaiFileContext
    {
        [JsonPropertyName("baiFileId")]
        public string? BaiFileId { get; set; }

        [JsonPropertyName("setTime")]
        public DateTime SetTime { get; set; }
    }


}
