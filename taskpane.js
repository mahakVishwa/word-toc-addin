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
            const body = context.document.body;

            console.log("📄 Loading paragraphs...");
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
            if (firstPara.text.toLowerCase().startsWith("table of contents")) {
                console.log("🗑 Clearing old TOC...");
                firstPara.clear();
            }

            // We'll build all TOC entries first (so no upside-down list)
            let tocEntries = [];

            headings.forEach((h, idx) => {
                try {
                    const styleName = h.style ? h.style.toLowerCase() : "";
                    let level = 0;
                    if (styleName.includes("heading 2")) level = 1;
                    else if (styleName.includes("heading 3")) level = 2;

                    // Unique bookmark name
                    const bookmarkName = `toc_heading_${idx}`;

                    // Add temporary bookmark
                    try {
                        const headingRange = h.getRange();
                        headingRange.insertBookmark(bookmarkName);
                        console.log(`   ✅ Bookmark inserted: ${bookmarkName}`);
                    } catch (bmErr) {
                        console.error("   ❌ Failed to insert bookmark:", bmErr);
                    }

                    // Save entry details for later insertion
                    tocEntries.push({ text: h.text, level, bookmarkName });
                } catch (headingErr) {
                    console.error(`❌ Error while processing heading ${idx}:`, headingErr);
                }
            });

            // Insert TOC title first
            const tocTitle = body.insertParagraph("Table of Contents", Word.InsertLocation.start);
            tocTitle.style = "Heading 1";

            // Insert TOC entries in correct order (reverse array so top->bottom matches doc)
            tocEntries.reverse().forEach(entry => {
                const para = body.insertParagraph(entry.text, Word.InsertLocation.start);
                para.style = "Normal";
                para.leftIndent = 36 * entry.level;
                try {
                    para.insertHyperlink(entry.text, entry.bookmarkName, "Replace");
                    console.log(`   🔗 Hyperlink inserted for: ${entry.text}`);
                } catch (linkErr) {
                    console.error("   ❌ Failed inserting hyperlink:", linkErr);
                }
            });

            await context.sync();
            console.log("🎉 TOC generated successfully!");

            // OPTIONAL: cleanup bookmarks after sync (keeps doc clean)
            headings.forEach((h, idx) => {
                try {
                    const range = h.getRange();
                    range.deleteBookmark(`toc_heading_${idx}`);
                    console.log(`🧹 Removed bookmark: toc_heading_${idx}`);
                } catch (cleanupErr) {
                    console.warn("⚠️ Failed cleaning bookmark:", cleanupErr);
                }
            });

            await context.sync();
            console.log("✨ Bookmarks cleaned, doc stays tidy!");
        });
    } catch (error) {
        console.error("💥 Error in generateTOC():", error);
    }
}
