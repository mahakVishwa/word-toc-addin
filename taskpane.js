Office.onReady(() => {
    console.log("✅ Office.js is ready!");
    const btn = document.getElementById("generate-btn");
    if (btn) {
        btn.disabled = false;
        btn.onclick = generateTOC;
    } else {
        console.warn("⚠️ generate-btn not found in DOM");
    }
});

async function generateTOC() {
    console.log("🚀 generateTOC() triggered...");
    try {
        await Word.run(async (context) => {
            const body = context.document.body;

            console.log("📄 Loading paragraphs...");
            const paragraphs = body.paragraphs;
            paragraphs.load("text, style");
            await context.sync();
            console.log("✅ Paragraphs loaded:", paragraphs.items.length);

            // Remove old TOC if the first paragraph is a TOC title
            const firstPara = paragraphs.items[0];
            if (firstPara && firstPara.text.toLowerCase().startsWith("table of contents")) {
                console.log("🗑 Clearing old TOC...");
                firstPara.clear();
            }

            // Insert TOC field at the start
            console.log("📌 Inserting native Table of Contents field...");
            const tocRange = body.getRange("start");
            tocRange.insertParagraph("Table of Contents", Word.InsertLocation.start).style = "Heading 1";

            // Insert the native TOC field
            // Options: 'Classic' style, show levels 1-3, include hyperlinks
            tocRange.insertTableOfContents("Classic", true, 1, 3);

            await context.sync();
            console.log("🎉 TOC generated successfully with clickable entries!");
        });
    } catch (err) {
        console.error("💥 Error in generateTOC():", err);
    }
}
