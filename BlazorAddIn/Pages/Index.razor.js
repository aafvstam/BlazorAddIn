

async function tagAndAddEventHandlersContentControls() {
    // Traverses each content control of the document and wraps a content control on each with either a even or odd tags.
    await Word.run(async (context) => {
        let currentdocument = context.document;
        currentdocument.load("$all");

        await context.sync();

        let contentcontrolsTagged = 0;

        let contentcontrols = currentdocument.contentControls;
        context.load(contentcontrols);
        await context.sync();

        for (let i = 0; i < contentcontrols.items.length; i++) {
            let contentControl = contentcontrols.items[i];
            // For even, tag "even".
            if (i % 2 === 0) {
                // Tag
                contentControl.tag = "even";
                console.log("Content Control Tagged Even!");
            } else {
                contentControl.tag = "odd";
                console.log("Content Control Tagged Odd!");
            }

            contentControl.onDeleted.add(handleContentControlDeleted);
            contentControl.onSelectionChanged.add(handleSelectionChanged);

            contentcontrolsTagged++;
        }

        await context.sync();
        console.log("Content controls tagged and handled: " + contentcontrolsTagged);
    });
}

async function deleteEvenContentControls() {
    // Traverses each content control of the document and deletes the even content controls
    await Word.run(async (context) => {
        let currentdocument = context.document;
        currentdocument.load("$all");

        await context.sync();

        let contentcontrols = currentdocument.contentControls;
        context.load(contentcontrols);

        await context.sync();

        let contentcontrolsRemaining = contentcontrols.items.length;

        for (let i = 0; i < contentcontrols.items.length; i++) {
            let contentControl = contentcontrols.items[i];

            // This will reinstate the handler but it should have been persisted from the prev. function
            // ------------------------------------------------------------------------------------------
            // contentControl.onDeleted.add(handleContentControlDeleted);
            // await context.sync();

            // delete even cc
            if (i % 2 === 0) {
                contentControl.delete(true);
                contentcontrolsRemaining--;
            }
        }

        await context.sync();
        console.log("Content controls remaining: " + contentcontrolsRemaining);
    });
}

export async function modifyContentControls() {
    // Adds title and colors to odd and even content controls and changes their appearance.
    await Word.run(async (context) => {
        // Gets the complete sentence (as range) associated with the insertion point.
        let evenContentControls = context.document.contentControls.getByTag("even");
        let oddContentControls = context.document.contentControls.getByTag("odd");
        evenContentControls.load("length");
        oddContentControls.load("length");

        await context.sync();

        for (let i = 0; i < evenContentControls.items.length; i++) {
            // Change a few properties and append a paragraph
            evenContentControls.items[i].set({
                color: "red",
                title: "Odd ContentControl #" + (i + 1),
                appearance: "Tags"
            });
            evenContentControls.items[i].insertParagraph("This is an odd content control", "End");
        }

        for (let j = 0; j < oddContentControls.items.length; j++) {
            // Change a few properties and append a paragraph
            oddContentControls.items[j].set({
                color: "green",
                title: "Even ContentControl #" + (j + 1),
                appearance: "Tags"
            });
            oddContentControls.items[j].insertHtml("This is an <b>even</b> content control", "End");
        }

        await context.sync();
    });
}


export async function setupDocument() {
    await Word.run(async (context) => {
        context.document.body.clear();
        context.document.body.insertParagraph("One more paragraph. ", "Start");
        context.document.body.insertParagraph("Co-locating Index.razor.js Demo", "Start");
        context.document.body.insertParagraph("Inserting another paragraph. ", "Start");
        context.document.body.insertParagraph(
            "Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
            "Start"
        );
        context.document.body.paragraphs
            .getLast()
            .insertText(
                "To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries. ",
                "Replace"
            );
    });
}

export async function insertContentControls() {
    // Traverses each paragraph of the document and wraps a content control on each with either a even or odd tags.
    await Word.run(async (context) => {
        let currentdocument = context.document;
        currentdocument.load("$all");

        await context.sync();

        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("$none"); // Don't need any properties; just wrap each paragraph with a content control.

        await context.sync();

        let contentcontrolsinserted = 0;

        for (let i = 0; i < paragraphs.items.length; i++) {
            let contentControl = paragraphs.items[i].insertContentControl();
            contentcontrolsinserted++;
        }

        await context.sync();

        // Tag --------------------------------

        let contentcontrolsTagged = 0;

        let contentcontrols = currentdocument.contentControls;

        context.load(contentcontrols);
        await context.sync();

        for (let i = 0; i < contentcontrols.items.length; i++) {
            let contentControl = contentcontrols.items[i];

            // For even, tag "even".
            if (i % 2 === 0) {
                // Tag
                contentControl.tag = "even";
                console.log("Content Control Tagged Even!");
            } else {
                contentControl.tag = "odd";
                console.log("Content Control Tagged Odd!");
            }

            await context.sync();

            contentcontrolsTagged++;
        }

        await context.sync();

        console.log("Content controls tagged: " + contentcontrolsTagged);
    });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}