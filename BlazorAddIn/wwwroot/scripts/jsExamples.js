export function showAlert(obj) {
    const message = 'Name is ' + obj.name + ' Age is ' + obj.age;
    alert(message);
}

export function insertParagraph() {

    console.log("Hello JavaScript in Blazor!?!?!?");

    return Word.run((context) => {

        // insert a paragraph at the start of the document.
        const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.start);

        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}