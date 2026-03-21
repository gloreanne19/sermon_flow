const axios = require('axios');
const cheerio = require('cheerio');

const processVerseData = ($, reference) => {
    $('.footnote, .crossreference, .full-text-link, .passage-display-caption, .publisher-info-bottom').remove();
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

const fetchVerseScraping = async (reference, version) => {
    try {
        const url = `https://www.biblegateway.com/passage/?search=${encodeURIComponent(reference)}&version=${version}`;
        const { data } = await axios.get(url, {
            headers: { 
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.9'
            },
            timeout: 15000
        });
        const $ = cheerio.load(data);
        return processVerseData($, reference);
    } catch (error) {
        console.error(`Error fetching ${reference} (${version}):`, error.message);
        return [];
    }
};

exports.handler = async (event, context) => {
    const headers = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'GET, OPTIONS'
    };

    if (event.httpMethod === 'OPTIONS') {
        return { statusCode: 200, headers, body: '' };
    }

    const { ref } = event.queryStringParameters || {};
    if (!ref) {
        return { statusCode: 400, headers, body: JSON.stringify({ error: "Reference is required" }) };
    }

    console.log(`Netlify Function: Scraping BibleGateway for: ${ref}...`);

    try {
        const [kjvVerses, mbbVerses] = await Promise.all([
            fetchVerseScraping(ref, 'KJV'),
            fetchVerseScraping(ref, 'MBBTAG')
        ]);

        if (kjvVerses.length > 0 || mbbVerses.length > 0) {
            return {
                statusCode: 200,
                headers,
                body: JSON.stringify({ reference: ref, KJV: kjvVerses, MBBTAG: mbbVerses })
            };
        } else {
            return {
                statusCode: 404,
                headers,
                body: JSON.stringify({ error: "Verses not found on BibleGateway." })
            };
        }
    } catch (error) {
        return {
            statusCode: 500,
            headers,
            body: JSON.stringify({ error: error.message })
        };
    }
};
