using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Threading.Tasks;

namespace BlazorAddIn.Pages
{
    public partial class Index
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; }

        private IJSObjectReference _jsModule;

        protected override async Task OnInitializedAsync()
        {
            _jsModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./scripts/jsExamples.js");
        }

        private async Task ShowAlertWindow() =>
            await _jsModule.InvokeVoidAsync("showAlert", new { Name = "John", Age = 35 });

        private async Task InsertParagraph() =>
            await _jsModule.InvokeVoidAsync("insertParagraph");

        private async Task Setup() =>
            await _jsModule.InvokeVoidAsync("setupDocument");
    }
}
