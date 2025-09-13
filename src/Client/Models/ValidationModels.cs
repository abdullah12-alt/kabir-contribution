 namespace Client.Models
{

 public class ValidationResponse
    {
        public string? Message { get; set; }
        public List<string>? Errors { get; set; }
    }

    public class LogMessage
    {
        public string Text { get; }
        public bool IsError { get; }

        public LogMessage(string text, bool isError)
        {
            Text = text;
            IsError = isError;
        }
    }
}
    