Office.onReady(() => {
    console.log("✅ Office.js is ready!");
    const btn = document.getElementById("generate-btn");
    btn.disabled = false;
    btn.onclick = generateTOC;
});

async function generateTOC() {
    console.log("🚀 TOC generation triggered...");
    try {
        await Word.run(async (context) => {
            console.log("📄 Loading paragraphs...");
            const body = context.document.body;
            const paragraphs = body.paragraphs;
            paragraphs.load("text, style");
            await context.sync();
            console.log("✅ Paragraphs loaded:", paragraphs.items.length);

            // Filter headings
            const headings = paragraphs.items.filter(p =>
                p.style && p.style.toLowerCase().includes("heading")
            );
            console.log("🔍 Headings found:", headings.length);
            headings.forEach((h, i) => {
                console.log(`   [Heading ${i}] text="${h.text}" style="${h.style}"`);
            });

            if (headings.length === 0) {
                console.log("⚠️ No headings found, aborting TOC generation.");
                return;
            }

            // Remove old TOC if it exists
            const firstPara = paragraphs.items[0];
            console.log("🔎 First paragraph text:", firstPara.text);
            if (firstPara.text.toLowerCase().startsWith("table of contents")) {
                console.log("🗑 Clearing old TOC...");
                firstPara.clear();
            }

            // Insert TOC title
            console.log("📌 Inserting TOC title...");
            const tocTitle = body.insertParagraph("Table of Contents", Word.InsertLocation.start);
            tocTitle.style = "Heading 1";

            // Add TOC entries
            headings.forEach((h, idx) => {
                try {
                    const styleName = h.style ? h.style.toLowerCase() : "";
                    console.log(`➡️ Processing heading ${idx}: "${h.text}" style="${h.style}"`);

                    let level = 0;
                    if (styleName.includes("heading 2")) level = 1;
                    else if (styleName.includes("heading 3")) level = 2;

                    console.log(`   ↳ Indent level = ${level}`);

                    // Unique bookmark name
                    const bookmarkName = `toc_heading_${idx}`;
                    console.log(`   ↳ Bookmark name: ${bookmarkName}`);

                    // Add bookmark at heading
                    try {
                        const headingRange = h.getRange();
                        headingRange.insertBookmark(bookmarkName);
                        console.log("   ✅ Bookmark inserted");
                    } catch (bmErr) {
                        console.error("   ❌ Failed to insert bookmark:", bmErr);
                    }

                    // Insert TOC entry
                    const entry = body.insertParagraph(h.text, Word.InsertLocation.start);
                    entry.style = "Normal";
                    entry.leftIndent = 36 * level;
                    console.log("   ✅ TOC entry created");

                    // Insert hyperlink to bookmark
                    try {
                        entry.insertHyperlink(h.text, bookmarkName, "Replace");
                        console.log("   🔗 Hyperlink inserted");
                    } catch (linkErr) {
                        console.error("   ❌ Failed inserting hyperlink:", linkErr);
                    }
                } catch (headingErr) {
                    console.error(`❌ Error while processing heading ${idx}:`, headingErr);
                }
            });

            await context.sync();
            console.log("🎉 TOC generated successfully!");
        });
    } catch (error) {
        console.error("💥 Error in generateTOC():", error);
    }
}
