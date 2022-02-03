using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    public partial class Page2
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;
        public IJSObjectReference JSModule { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;


        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Page2.razor.js");
            }
        }

        private async Task InsertParagraph() =>
            await JSModule.InvokeVoidAsync("insertParagraph");

        void MoveToPage(string page)
        {
            NavigationManager.NavigateTo(page, true);
        }
    }
}