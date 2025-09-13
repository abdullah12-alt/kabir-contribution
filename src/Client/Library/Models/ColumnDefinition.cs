namespace Client.Library
{
    public class ColumnDefinition
    {
        public string Property { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string? Width { get; set; }
        public string? Format { get; set; } // e.g., "{0:yyyy-MM-dd}"
        public string? Align { get; set; } // e.g., "center", "left", "right"
        public int? MaxCharLength { get; set; } = 200; // Optional char limit

    }

}
