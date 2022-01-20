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

        private async Task InsertParagraph() =>
            await _jsModule.InvokeVoidAsync("insertParagraph");

        private async Task Setup() =>
            await _jsModule.InvokeVoidAsync("setupDocument");

        private async Task InsertContentControls() =>
            await _jsModule.InvokeVoidAsync("insertContentControls");

        private async Task ModifyContentControls() =>
            await _jsModule.InvokeVoidAsync("modifyContentControls");
    }
}