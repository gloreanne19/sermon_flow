const axios = require('axios');
const cheerio = require('cheerio');

// Utility to clean and split verses from BibleGateway HTML
const processVerseData = ($, reference) => {
    $('.footnote, .crossreference, .full-text-link, .passage-display-caption, .publisher-info-bottom, .passage-other-trans').remove();
    let verses = [];
    $('.passage-content .text').each((i, el) => {
        const verseNum = $(el).find('.versenum').text().trim();
        const chapterNum = $(el).find('.chapternum').text().trim();
        const entry = $(el).clone();
        entry.find('.versenum, .chapternum').remove();
        let text = entry.text().replace(/\s+/g, ' ').trim();
        if (text) {
            if (verseNum || chapterNum) {
                verses.push({ num: verseNum || chapterNum || "1", text: text });
            } else if (verses.length > 0) {
                verses[verses.length - 1].text += " " + text;
            }
        }
    });
    return verses.map(v => ({ num: v.num, text: v.text.replace(/\s+/g, ' ').trim() })).filter(v => v.text.length > 0);
};

// Robust Scraper with Browser-like Headers
const fetchVerseScraping = async (reference, version) => {
    try {
        const url = `https://www.biblegateway.com/passage/?search=${encodeURIComponent(reference)}&version=${version}`;
        const { data } = await axios.get(url, {
            headers: { 
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.9',
                'Cache-Control': 'no-cache',
                'Pragma': 'no-cache'
            },
            timeout: 15000
        });
        const $ = cheerio.load(data);
        const result = processVerseData($, reference);
        return result;
    } catch (error) {
        console.error(`BibleGateway Error (${version}): ${error.message}`);
        return [];
    }
};

module.exports = async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    if (req.method === 'OPTIONS') return res.status(200).end();

    const { ref } = req.query;
    if (!ref) return res.status(400).json({ error: "Reference required" });

    console.log(`BIBLEGATEWAY: Scraping ${ref}`);

    try {
        const [kjvV, mbbV] = await Promise.all([
            fetchVerseScraping(ref, 'KJV'),
            fetchVerseScraping(ref, 'MBBTAG')
        ]);

        if (kjvV.length > 0 || mbbV.length > 0) {
            return res.status(200).json({ reference: ref, KJV: kjvV, MBBTAG: mbbV });
        } else {
            return res.status(404).json({ error: "Verse not found on BibleGateway." });
        }
    } catch (e) {
        return res.status(500).json({ error: e.message });
    }
};
