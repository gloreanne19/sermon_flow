const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

// Load bible-gateway-api dynamically (ESM)
let BibleGatewayAPI;
(async () => {
    try {
        const mod = await import('bible-gateway-api');
        BibleGatewayAPI = mod.BibleGatewayAPI;
        console.log("BibleGatewayAPI library loaded successfully.");
    } catch (e) {
        console.error("Failed to load bible-gateway-api, will use fallback scraping.", e);
    }
})();

// Utility to clean and split verses
const processVerseData = ($, reference) => {
    // Remove unwanted elements
    $('.footnote, .crossreference, .full-text-link, .passage-display-caption, .publisher-info-bottom').remove();
    
    let verses = [];
    
    // Bible Gateway structure: .text contains .versenum and the text
    $('.passage-content .text').each((i, el) => {
        const verseNum = $(el).find('.versenum').text().trim();
        const chapterNum = $(el).find('.chapternum').text().trim();
        
        // Clone to remove the verse number from the text content
        const entry = $(el).clone();
        entry.find('.versenum, .chapternum').remove();
        let text = entry.text().replace(/\s+/g, ' ').trim();
        
        if (text) {
            // If we found a verse number, start a new verse
            if (verseNum || chapterNum) {
                verses.push({
                    num: verseNum || chapterNum || "1",
                    text: text
                });
            } else if (verses.length > 0) {
                // If no verse number, append to the last verse (for segments split by tags)
                verses[verses.length - 1].text += " " + text;
            }
        }
    });

    // Clean up extra spaces
    return verses.map(v => ({
        num: v.num,
        text: v.text.replace(/\s+/g, ' ').trim()
    })).filter(v => v.text.length > 0);
};

// Improved scraping withverse splitting
const fetchVerseScraping = async (reference, version) => {
    try {
        const url = `https://www.biblegateway.com/passage/?search=${encodeURIComponent(reference)}&version=${version}`;
        const { data } = await axios.get(url, {
            headers: { 
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36' 
            }
        });
        const $ = cheerio.load(data);
        return processVerseData($, reference);
    } catch (error) {
        console.error(`Error fetching ${reference} (${version}):`, error.message);
        return [];
    }
};

app.get('/api/bible', async (req, res) => {
    const { ref } = req.query;
    if (!ref) return res.status(400).json({ error: "Reference is required" });

    console.log(`Processing multi-verse request for: ${ref}...`);

    const [kjvVerses, mbbVerses] = await Promise.all([
        fetchVerseScraping(ref, 'KJV'),
        fetchVerseScraping(ref, 'MBBTAG')
    ]);

    if (kjvVerses.length > 0 || mbbVerses.length > 0) {
        res.json({
            reference: ref,
            KJV: kjvVerses,
            MBBTAG: mbbVerses
        });
    } else {
        res.status(404).json({ error: "Verses not found" });
    }
});

app.listen(PORT, () => {
    console.log(`Bible API Server running on http://localhost:${PORT}`);
    console.log(`Using Robust Scraper with specific content selectors.`);
});

