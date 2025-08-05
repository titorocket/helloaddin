function insertText() {
  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText("Hello Word", Word.InsertLocation.replace);
    await context.sync();
  });
}
