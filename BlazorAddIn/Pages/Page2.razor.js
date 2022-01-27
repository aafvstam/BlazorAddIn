﻿export function insertParagraph() {

    console.log("Hello JavaScript in Blazor!?!?!?");

    return Word.run((context) => {

        // insert a paragraph at the start of the document.
        const paragraph = context.document.body.insertParagraph("Hello World from Page2.razor.js", Word.InsertLocation.start);

        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}
