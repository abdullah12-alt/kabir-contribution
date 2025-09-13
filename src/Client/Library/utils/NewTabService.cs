using Microsoft.JSInterop;
namespace Client.Library
{
    public interface INewTabService
    {
        Task OpenUrlInNewTab(string url);

    }



    public class NewTabService : INewTabService
    {
        private readonly IJSRuntime _jsRuntime;
        private IJSObjectReference? _jsModule;

        public NewTabService(IJSRuntime jsRuntime)
        {
            _jsRuntime = jsRuntime;
        }

        private async Task EnsureModuleLoaded()
        {
            if (_jsModule == null)
            {
                _jsModule = await _jsRuntime.InvokeAsync<IJSObjectReference>(
                    "import", "./js/browser-utils.js");
            }
        }

        public async Task OpenUrlInNewTab(string url)
        {
            await EnsureModuleLoaded();
            await _jsModule.InvokeVoidAsync("openUrlInNewTab", url);
        }
    }

}
