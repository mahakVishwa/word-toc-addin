Office.onReady(() => {
    console.log("Office.js is ready!");
    const btn = document.getElementById("generate-btn");
    btn.disabled = false;
    btn.onclick = generateTOC;
});

async function generateTOC() {
    try {
        console.log("TOC generation started...");
        await Word.run(async (context) => {
            const body = context.document.body;
            const paragraphs = body.paragraphs;
            paragraphs.load("text, style, font, hyperlink, id");
            await context.sync();

            console.log("Paragraphs loaded:", paragraphs.items.length);
            // Filter headings
            paragraphs.items.forEach(p => {
                console.log("Paragraph:", p.text, "| Style:", p.style);
            });
            const headings = paragraphs.items.filter(p => p.style && p.style.toLowerCase().includes("heading"));
            console.log("Headings found:", headings.length);

            if (headings.length === 0) {
                console.log("No headings found!");
                return;
            }

            // Remove old TOC if it exists (optional)
            const firstPara = paragraphs.items[0];
            if (firstPara.text.startsWith("Table of Contents")) {
                firstPara.clear();
            }

            // Insert TOC title
            const tocTitle = body.insertParagraph("Table of Contents", Word.InsertLocation.start);
            tocTitle.style = "Heading 1";

            // Add TOC entries
            headings.forEach(h => {
                let indent = 0;
                if (h.style.includes("Heading 2")) indent = 1;
                else if (h.style.includes("Heading 3")) indent = 2;

                const para = body.insertParagraph(h.text, Word.InsertLocation.start);
                para.style = "Normal";
                para.leftIndent = 36 * indent; // 36 points per level
                para.insertHyperlink(h.text, h.id, "End"); // clickable link
            });

            await context.sync();
            console.log("TOC generated successfully!");
        });
    } catch (error) {
        console.error("Error generating TOC:", error);
    }
}





