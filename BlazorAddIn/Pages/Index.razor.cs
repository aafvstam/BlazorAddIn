using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    public partial class Index
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        public IJSObjectReference JSModule { get; set; } = default!;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Index.razor.js");
            }
        }

        private async Task Setup() =>
            await JSModule.InvokeVoidAsync("setupDocument");

        private async Task InsertContentControls() =>
            await JSModule.InvokeVoidAsync("insertContentControls");

        private async Task ModifyContentControls() =>
            await JSModule.InvokeVoidAsync("modifyContentControls");

        void MoveToPage(string page)
        {
            NavigationManager.NavigateTo(page, true);
        }
    }
}