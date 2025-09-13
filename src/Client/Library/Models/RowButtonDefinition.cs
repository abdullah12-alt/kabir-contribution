using Radzen;
namespace Client.Library
{
    public class RowButtonDefinition<TItem>
    {
        public string Icon { get; set; } = "edit";
        public string? Label { get; set; }
        public string? Tooltip { get; set; }
        public string BackgroundColor { get; set; } = ButtonColors.Secondary; 
        public string TextColor { get; set; } = ButtonColors.White;                      
        public string? Css { get; set; } = null;
        public Func<TItem, Task>? Callback { get; set; }  
    }
}
