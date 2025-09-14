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

            // detect headings
            const headings = paragraphs.items.filter(p =>
                p.style && p.style.toLowerCase().includes("heading")
            );
            console.log("🔍 Headings found:", headings.length);
            headings.forEach((h, i) => console.log(`   [${i}] "${h.text.trim()}" style="${h.style}"`));

            if (headings.length === 0) {
                console.log("⚠️ No headings found - aborting.");
                return;
            }

            // If the first paragraph is an old TOC title, clear it (clean start)
            try {
                const firstPara = paragraphs.items[0];
                console.log("🔎 First paragraph text:", firstPara.text.trim());
                if (firstPara.text && firstPara.text.toLowerCase().startsWith("table of contents")) {
                    console.log("🗑 Clearing old TOC start paragraph...");
                    firstPara.clear();
                }
            } catch (eFirst) {
                console.warn("⚠️ Could not check/clear first paragraph:", eFirst);
            }

            // Build entries list (so we can insert in correct order)
            let tocEntries = [];
            for (let i = 0; i < headings.length; i++) {
                const h = headings[i];
                try {
                    const styleName = (h.style || "").toLowerCase();
                    let level = 0;
                    if (styleName.includes("heading 2")) level = 1;
                    else if (styleName.includes("heading 3")) level = 2;

                    // unique bookmark name (add timestamp to avoid collisions)
                    const bookmarkName = `toc_heading_${i}_${Date.now()}`;

                    // insert temporary bookmark at the heading range
                    try {
                        const headingRange = h.getRange();
                        headingRange.insertBookmark(bookmarkName);
                        console.log(`   ✅ Inserted bookmark "${bookmarkName}" at heading idx ${i}`);
                    } catch (bmErr) {
                        console.error(`   ❌ Failed to insert bookmark for heading idx ${i}:`, bmErr);
                    }

                    tocEntries.push({
                        text: h.text.replace(/\r?\n/g, "").trim(),
                        level,
                        bookmarkName,
                        headingIndex: i
                    });
                } catch (entryErr) {
                    console.error(`   ❌ Error building entry for heading idx ${i}:`, entryErr);
                }
            }

            // Insert TOC title at top
            console.log("📌 Inserting TOC title at document start...");
            const tocTitle = body.insertParagraph("Table of Contents", Word.InsertLocation.start);
            tocTitle.style = "Heading 1";

            // Insert entries in the correct top->bottom order:
            // tocEntries was built in doc order (top→bottom). Because we insert each item at start,
            // we must iterate in reverse so they appear top→bottom in the final doc.
            console.log("🔧 Inserting TOC entries (in reverse to keep order correct)...");
            for (let i = tocEntries.length - 1; i >= 0; i--) {
                const e = tocEntries[i];
                try {
                    const para = body.insertParagraph(e.text, Word.InsertLocation.start);
                    para.style = "Normal";
                    para.leftIndent = 36 * e.level; // adjust indent per level
                    console.log(`   ✅ Inserted TOC line: "${e.text}" (level ${e.level})`);

                    // Attempt to make it clickable by linking to the temporary bookmark.
                    // We try a couple of approaches with try/catch because some clients
                    // expose hyperlink insertion differently.
                    try {
                        // Preferred: use paragraph's range to insert hyperlink (if available)
                        const paraRange = para.getRange();
                        // Try with displayText, anchorName, "Replace"
                        try {
                            paraRange.insertHyperlink(e.text, e.bookmarkName, "Replace");
                            console.log(`   🔗 insertHyperlink (range) queued for "${e.text}" -> ${e.bookmarkName}`);
                        } catch (innerErr1) {
                            console.warn("   ⚠ paraRange.insertHyperlink failed:", innerErr1, "Trying paragraph.insertHyperlink...");
                            // fallback: some hosts may have paragraph.insertHyperlink
                            try {
                                if (typeof para.insertHyperlink === "function") {
                                    para.insertHyperlink(e.text, e.bookmarkName, "Replace");
                                    console.log(`   🔗 paragraph.insertHyperlink queued for "${e.text}"`);
                                } else {
                                    throw new Error("paragraph.insertHyperlink not available");
                                }
                            } catch (innerErr2) {
                                console.warn("   ⚠ paragraph.insertHyperlink failed or not available:", innerErr2);
                                console.log("   ℹ️ Falling back to plain text TOC entry (no link).");
                            }
                        }
                    } catch (linkOuterErr) {
                        console.error("   ❌ Unexpected error when trying to insert hyperlink:", linkOuterErr);
                    }
                } catch (insErr) {
                    console.error("   ❌ Failed to insert TOC entry paragraph:", insErr);
                }
            }

            // Commit all queued ops (bookmarks + TOC entries + hyperlink attempts)
            await context.sync();
            console.log("🔁 context.sync() done - TOC + bookmarks queued.");

            // CLEANUP: remove temporary bookmarks we created (keep document tidy)
            console.log("🧹 Attempting to clean up temporary bookmarks...");
            for (let i = 0; i < tocEntries.length; i++) {
                const name = tocEntries[i].bookmarkName;
                try {
                    // Try deleting bookmark via heading range first
                    try {
                        const headingRange = headings[tocEntries[i].headingIndex].getRange();
                        headingRange.deleteBookmark(name);
                        console.log(`   🗑 Deleted bookmark (via range): ${name}`);
                    } catch (rangeDelErr) {
                        console.warn(`   ⚠ range.deleteBookmark failed for ${name}:`, rangeDelErr);
                        // fallback: try document.bookmarks (not guaranteed across clients)
                        try {
                            if (context.document && context.document.bookmarks && typeof context.document.bookmarks.getItem === "function") {
                                const bm = context.document.bookmarks.getItem(name);
                                bm.delete();
                                console.log(`   🗑 Deleted bookmark (via document.bookmarks): ${name}`);
                            } else {
                                throw new Error("document.bookmarks API not available");
                            }
                        } catch (docBmErr) {
                            console.warn(`   ⚠ Could not delete bookmark ${name} via fallback:`, docBmErr);
                        }
                    }
                } catch (cleanupErr) {
                    console.warn(`   ⚠ Final cleanup attempt failed for ${name}:`, cleanupErr);
                }
            }

            // final sync for bookmark deletions
            try {
                await context.sync();
                console.log("✨ Bookmark cleanup sync complete.");
            } catch (finalSyncErr) {
                console.warn("⚠ Final sync after bookmark cleanup failed:", finalSyncErr);
            }

            console.log("🎉 generateTOC finished!");
        }); // end Word.run
    } catch (err) {
        console.error("💥 generateTOC outer catch:", err);
    }
}
