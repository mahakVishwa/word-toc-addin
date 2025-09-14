async function generateTOC() {
  await Word.run(async (context) => {
    // 1. Get all paragraphs
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text, style");
    await context.sync();

    // 2. Filter only Headings
    let headings = paragraphs.items.filter(p => 
      p.style.includes("Heading")
    );

    // 3. Build TOC string
    let toc = "Table of Contents\n\n";
    headings.forEach((h) => {
      if (h.style.includes("Heading 1")) {
        toc += h.text + "\n";
      } else if (h.style.includes("Heading 2")) {
        toc += "    " + h.text + "\n";
      } else if (h.style.includes("Heading 3")) {
        toc += "        " + h.text + "\n";
      }
    });

    // 4. Insert TOC at top of document
    let body = context.document.body;
    body.insertParagraph(toc, Word.InsertLocation.start);

    await context.sync();
    console.log("TOC generated!");
  });
}