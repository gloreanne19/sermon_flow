const axios = require('axios');
const cheerio = require('cheerio');

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
            if (verseNum || chapterNum) {
                verses.push({
                    num: verseNum || chapterNum || "1",
                    text: text
                });
            } else if (verses.length > 0) {
                verses[verses.length - 1].text += " " + text;
            }
        }
    });

    return verses.map(v => ({
        num: v.num,
        text: v.text.replace(/\s+/g, ' ').trim()
    })).filter(v => v.text.length > 0);
};

// Scraping function
const fetchVerseScraping = async (reference, version) => {
    try {
        const url = `https://www.biblegateway.com/passage/?search=${encodeURIComponent(reference)}&version=${version}`;
        const { data } = await axios.get(url, {
            headers: { 
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36' 
            },
            timeout: 10000
        });
        const $ = cheerio.load(data);
        return processVerseData($, reference);
    } catch (error) {
        console.error(`Error fetching ${reference} (${version}):`, error.message);
        return [];
    }
};

// Vercel Serverless Function Handler
module.exports = async (req, res) => {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    const { ref } = req.query;
    if (!ref) {
        return res.status(400).json({ error: "Reference is required" });
    }

    console.log(`Processing Vercel Function request for: ${ref}...`);

    try {
        const [kjvVerses, mbbVerses] = await Promise.all([
            fetchVerseScraping(ref, 'KJV'),
            fetchVerseScraping(ref, 'MBBTAG')
        ]);

        if (kjvVerses.length > 0 || mbbVerses.length > 0) {
            return res.status(200).json({
                reference: ref,
                KJV: kjvVerses,
                MBBTAG: mbbVerses
            });
        } else {
            return res.status(404).json({ error: "Verses not found" });
        }
    } catch (error) {
        return res.status(500).json({ error: error.message });
    }
};
