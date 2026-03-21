const axios = require('axios');

// Normalize a Bible reference for bible-api.com (e.g. "Heb. 9:14" → "hebrews 9:14")
const BOOK_ABBREVS = {
  'gen': 'genesis', 'ex': 'exodus', 'exo': 'exodus', 'lev': 'leviticus', 'num': 'numbers',
  'deut': 'deuteronomy', 'deu': 'deuteronomy', 'josh': 'joshua', 'jos': 'joshua',
  'judg': 'judges', 'jdg': 'judges', 'ruth': 'ruth', '1 sam': '1 samuel', '2 sam': '2 samuel',
  '1 ki': '1 kings', '1 kgs': '1 kings', '2 ki': '2 kings', '2 kgs': '2 kings',
  '1 chr': '1 chronicles', '2 chr': '2 chronicles', 'ezra': 'ezra', 'neh': 'nehemiah',
  'esth': 'esther', 'est': 'esther', 'ps': 'psalms', 'psa': 'psalms', 'prov': 'proverbs',
  'pro': 'proverbs', 'eccl': 'ecclesiastes', 'ecc': 'ecclesiastes', 'isa': 'isaiah',
  'jer': 'jeremiah', 'lam': 'lamentations', 'ezek': 'ezekiel', 'eze': 'ezekiel',
  'dan': 'daniel', 'hos': 'hosea', 'joel': 'joel', 'amos': 'amos', 'obad': 'obadiah',
  'jon': 'jonah', 'mic': 'micah', 'nah': 'nahum', 'hab': 'habakkuk', 'zeph': 'zephaniah',
  'hag': 'haggai', 'zech': 'zechariah', 'zec': 'zechariah', 'mal': 'malachi',
  'matt': 'matthew', 'mat': 'matthew', 'mk': 'mark', 'lk': 'luke', 'jn': 'john',
  'acts': 'acts', 'rom': 'romans', '1 cor': '1 corinthians', '2 cor': '2 corinthians',
  'gal': 'galatians', 'eph': 'ephesians', 'phil': 'philippians', 'col': 'colossians',
  '1 thess': '1 thessalonians', '2 thess': '2 thessalonians', '1 tim': '1 timothy',
  '2 tim': '2 timothy', 'tit': 'titus', 'phm': 'philemon', 'heb': 'hebrews',
  'jas': 'james', 'jms': 'james', '1 pet': '1 peter', '1 pe': '1 peter',
  '2 pet': '2 peter', '2 pe': '2 peter', '1 jn': '1 john', '2 jn': '2 john',
  '3 jn': '3 john', 'jude': 'jude', 'rev': 'revelation',
};

const normalizeRef = (ref) => {
  // Remove trailing period, clean up spacing
  let r = ref.trim().replace(/\.\s*/g, ' ').replace(/\s+/g, ' ').trim();

  // Split book from chapter:verse
  const match = r.match(/^([\d\s]*[a-zA-Z]+(?:\s+[a-zA-Z]+)*)\s+(\d+.*)/);
  if (!match) return encodeURIComponent(r);

  let book = match[1].trim().toLowerCase();
  const chapVerse = match[2].trim();

  // Try to expand abbreviation
  const expanded = BOOK_ABBREVS[book] || book;
  return encodeURIComponent(`${expanded} ${chapVerse}`);
};

// Fetch from bible-api.com (KJV, free, no auth)
const fetchKJV = async (reference) => {
  try {
    const url = `https://bible-api.com/${normalizeRef(reference)}?translation=kjv`;
    const { data } = await axios.get(url, { timeout: 10000 });

    if (!data || !data.verses || data.verses.length === 0) return [];

    return data.verses.map(v => ({
      num: String(v.verse),
      text: v.text.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim()
    }));
  } catch (e) {
    console.error('KJV fetch error:', e.message);
    return [];
  }
};

// Fetch Tagalog from helloao.org (large free Bible API, no auth)
const fetchMBBTAG = async (reference) => {
  try {
    // helloao.org uses "MBBTAG" as a valid translation ID
    const url = `https://bible.helloao.org/api/MBBTAG/${normalizeRef(reference)}.json`;
    const { data } = await axios.get(url, { timeout: 10000 });

    if (!data || !data.verses || data.verses.length === 0) return [];

    return data.verses.map(v => ({
      num: String(v.number),
      text: v.content.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim()
    }));
  } catch (e) {
    console.error('MBBTAG fetch error:', e.message);
    return [];
  }
};

// Vercel Serverless Function Handler
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

  if (req.method === 'OPTIONS') return res.status(200).end();

  const { ref } = req.query;
  if (!ref) return res.status(400).json({ error: 'Reference is required' });

  console.log(`Fetching via REST API: ${ref}`);

  try {
    const [kjvVerses, mbbVerses] = await Promise.all([
      fetchKJV(ref),
      fetchMBBTAG(ref),
    ]);

    if (kjvVerses.length > 0 || mbbVerses.length > 0) {
      return res.status(200).json({ reference: ref, KJV: kjvVerses, MBBTAG: mbbVerses });
    } else {
      return res.status(404).json({ error: 'Verses not found' });
    }
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
};
