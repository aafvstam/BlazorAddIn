
export async function clearDocument() {
    await Word.run(async (context) => {
        context.document.body.clear();
    });
}

export async function insertParagraph(text, location) {
    await Word.run(async (context) => {
        context.document.body.insertParagraph(text, location);
    });
}

export async function replaceParagraph(text) {
    await Word.run(async (context) => {
        context.document.body.paragraphs
            .getLast()
            .insertText(
                text,
                "Replace"
            );
    });
}

export async function paragraphCount() {

    let numberofParagraphs = 0;

    await Word.run(async (context) => {
        let currentdocument = context.document;
        currentdocument.load("$all");

        await context.sync();

        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("$none"); // Don't need any properties;

        await context.sync();

        numberofParagraphs = paragraphs.items.length;

        console.log("Number of paragraphs");
        console.log(numberofParagraphs);

    });
}
