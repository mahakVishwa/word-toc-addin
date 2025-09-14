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

            // Remove old TOC if it exists at the top
            const firstPara = paragraphs.items[0];
            if (firstPara && firstPara.text.toLowerCase().startsWith("table of contents")) {
                console.log("🗑 Clearing old TOC...");
                firstPara.clear();
            }

            // Collect headings
            const headings = paragraphs.items.filter(p =>
                p.style && p.style.toLowerCase().includes("heading")
            );
            console.log("🔍 Headings found:", headings.length);
            headings.forEach((h, i) => console.log(`   [${i}] "${h.text.trim()}" style="${h.style}"`));

            // Build TOC entries array
            const tocEntries = headings.map((h, idx) => {
                let level = 0;
                const styleName = h.style.toLowerCase();
                if (styleName.includes("heading 2")) level = 1;
                else if (styleName.includes("heading 3")) level = 2;

                return { text: h.text.trim(), level };
            });

            // Insert TOC title at the very top
            const tocTitle = body.insertParagraph("Table of Contents", Word.InsertLocation.start);
            tocTitle.style = "Heading 1";

            // Insert TOC entries **after the title** to keep correct order
            for (let i = tocEntries.length - 1; i >= 0; i--) {
                const entry = tocEntries[i];
                const para = body.insertParagraph(entry.text, Word.InsertLocation.start);
                para.style = "Normal";
                para.leftIndent = 36 * entry.level;
            }

            await context.sync();
            console.log("🎉 TOC generated successfully!");
        });
    } catch (err) {
        console.error("💥 Error in generateTOC():", err);
    }
}
