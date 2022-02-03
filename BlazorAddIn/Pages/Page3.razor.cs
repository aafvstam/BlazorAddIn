using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    public partial class Page3
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        public IJSObjectReference JSModule { get; set; } = default!;

        public string ParagraphCount { get; set; } = "<undefined>";

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Page3.razor.js");
            }
        }

        private async Task Setup()
        {
            // VSTO Way of Working
            // Get ActiveDocument
            // Clear ActiveDocument (Get Document Body, Select All, Delete?)

            await Clear();

            // Get ActiveDocument
            // ActiveDocument (Get Document Start, Insert Paragraph at specified location?)

            await InsertParagraph("One more paragraph.", "Start");
            await InsertParagraph("Co-locating Index.razor.js Demo.", "Start");
            await InsertParagraph("Inserting another paragraph.", "Start");
            await InsertParagraph("Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.", "Start");
            await ReplaceParagraph("To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.");
        }

        private async Task InsertParagraph(string text, string location)
        {
            await JSModule.InvokeVoidAsync("insertParagraph", text, location);
        }

        private async Task ReplaceParagraph(string text)
        {
            await JSModule.InvokeVoidAsync("replaceParagraph", text);
        }

        private async Task Clear()
        {
            await JSModule.InvokeVoidAsync("clearDocument");
        }

        //private async Task CountParagraps()
        //{
        //    ParagraphCount = await JSModule.InvokeAsync<string>("paragraphCount");

        //    Console.WriteLine("Test");
        //    Console.WriteLine(ParagraphCount);
        //}

        void MoveToPage(string page)
        {
            NavigationManager.NavigateTo(page, true);
        }
    }
}