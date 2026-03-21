const express = require('express');
const axios = require('axios');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

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
  let r = ref.trim().replace(/\.\s*/g, ' ').replace(/\s+/g, ' ').trim();
  const match = r.match(/^([\d\s]*[a-zA-Z]+(?:\s+[a-zA-Z]+)*)\s+(\d+.*)/);
  if (!match) return encodeURIComponent(r);
  let book = match[1].trim().toLowerCase();
  const chapVerse = match[2].trim();
  const expanded = BOOK_ABBREVS[book] || book;
  return encodeURIComponent(`${expanded} ${chapVerse}`);
};

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

const fetchMBBTAG = async (reference) => {
  try {
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

app.get('/api/bible', async (req, res) => {
  const { ref } = req.query;
  if (!ref) return res.status(400).json({ error: 'Reference is required' });

  console.log(`Local server fetching: ${ref}`);

  const [kjvVerses, mbbVerses] = await Promise.all([
    fetchKJV(ref),
    fetchMBBTAG(ref),
  ]);

  if (kjvVerses.length > 0 || mbbVerses.length > 0) {
    return res.json({ reference: ref, KJV: kjvVerses, MBBTAG: mbbVerses });
  } else {
    return res.status(404).json({ error: 'Verses not found' });
  }
});

app.listen(PORT, () => {
  console.log(`Bible API Server running on http://localhost:${PORT}`);
  console.log(`Using bible-api.com (KJV) + bible.helloao.org (MBBTAG)`);
});
