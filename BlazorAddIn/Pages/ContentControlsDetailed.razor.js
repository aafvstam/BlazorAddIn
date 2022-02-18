
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

export function emailRegistration(message) {
    const result = message;
    if (result === '' || result === null)
        return 'Please provide an email'
    const returnMessage = 'Hi ' + result.split('@')[0] + ' your email: ' + result + ' has been accepted.';
    return returnMessage;
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

        return numberofParagraphs.toString() ;
}
