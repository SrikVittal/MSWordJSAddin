Office.onReady(() => {
    document.getElementById("sayHello").addEventListener("click", () => {
        Word.run((context) => {
            const range = context.document.getSelection();
            range.insertText("Hello World", Word.InsertLocation.replace);
            return context.sync();
        }).catch((error) => {
            console.error("Error: " + error);
        });
    });
});
