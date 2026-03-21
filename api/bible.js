const axios = require('axios');

const BOOK_MAP = {
  'gen': 'GEN', 'genesis': 'GEN', 'ex': 'EXO', 'exodus': 'EXO', 'exo': 'EXO', 'lev': 'LEV', 'leviticus': 'LEV',
  'num': 'NUM', 'numbers': 'NUM', 'deut': 'DEU', 'deuteronomy': 'DEU', 'deu': 'DEU',
  'josh': 'JOS', 'joshua': 'JOS', 'jos': 'JOS', 'judg': 'JDG', 'judges': 'JDG', 'jdg': 'JDG',
  'ruth': 'RUT', 'ruth': 'RUT', '1 sam': '1SA', '1 samuel': '1SA', '2 sam': '2SA', '2 samuel': '2SA',
  '1 ki': '1KI', '1 kings': '1KI', '1 kgs': '1KI', '2 ki': '2KI', '2 kings': '2KI', '2 kgs': '2KI',
  '1 chr': '1CH', '1 chronicles': '1CH', '2 chr': '2CH', '2 chronicles': '2CH', 'ezra': 'EZR', 'ezra': 'EZR',
  'neh': 'NEH', 'nehemiah': 'NEH', 'esth': 'EST', 'esther': 'EST', 'est': 'EST', 'job': 'JOB',
  'ps': 'PSA', 'psalms': 'PSA', 'psa': 'PSA', 'prov': 'PRO', 'proverbs': 'PRO', 'pro': 'PRO',
  'eccl': 'ECC', 'ecclesiastes': 'ECC', 'ecc': 'ECC', 'song': 'SNG', 'isa': 'ISA', 'isaiah': 'ISA',
  'jer': 'JER', 'jeremiah': 'JER', 'lam': 'LAM', 'lamentations': 'LAM', 'ezek': 'EZK', 'ezekiel': 'EZK', 'eze': 'EZK',
  'dan': 'DAN', 'daniel': 'DAN', 'hos': 'HOS', 'hosea': 'HOS', 'joel': 'JOL', 'joel': 'JOL', 'amos': 'AMO', 'amos': 'AMO',
  'obad': 'OBA', 'obadiah': 'OBA', 'jon': 'JON', 'jonah': 'JON', 'mic': 'MIC', 'micah': 'MIC', 'nah': 'NAM', 'nahum': 'NAM',
  'hab': 'HAB', 'habakkuk': 'HAB', 'zeph': 'ZEP', 'zephaniah': 'ZEP', 'hag': 'HAG', 'haggai': 'HAG', 'zech': 'ZEC',
  'zechariah': 'ZEC', 'zec': 'ZEC', 'mal': 'MAL', 'malachi': 'MAL', 'matt': 'MAT', 'matthew': 'MAT', 'mat': 'MAT',
  'mk': 'MRK', 'mark': 'MRK', 'lk': 'LUK', 'luke': 'LUK', 'jn': 'JHN', 'john': 'JHN', 'acts': 'ACT', 'acts': 'ACT',
  'rom': 'ROM', 'romans': 'ROM', '1 cor': '1CO', '1 corinthians': '1CO', '2 cor': '2CO', '2 corinthians': '2CO',
  'gal': 'GAL', 'galatians': 'GAL', 'eph': 'EPH', 'ephesians': 'EPH', 'phil': 'PHP', 'philippians': 'PHP',
  'col': 'COL', 'colossians': 'COL', '1 thess': '1TH', '1 thessalonians': '1TH', '2 thess': '2TH', '2 thessalonians': '2TH',
  '1 tim': '1TI', '1 timothy': '1TI', '2 tim': '2TI', '2 timothy': '2TI', 'tit': 'TIT', 'titus': 'TIT', 'phm': 'PHM',
  'philemon': 'PHM', 'heb': 'HEB', 'hebrews': 'HEB', 'jas': 'JAS', 'james': 'JAS', 'jms': 'JAS',
  '1 pet': '1PE', '1 peter': '1PE', '1 pe': '1PE', '2 pet': '2PE', '2 peter': '2PE', '2 pe': '2PE',
  '1 jn': '1JN', '1 john': '1JN', '2 jn': '2JN', '2 john': '2JN', '3 jn': '3JN', '3 john': '3JN', 'jude': 'JUD', 'jude': 'JUD',
  'rev': 'REV', 'revelation': 'REV'
};

const HELLO_AO_BOOKS = {
  'GEN': 'genesis', 'EXO': 'exodus', 'LEV': 'leviticus', 'NUM': 'numbers', 'DEU': 'deuteronomy',
  'JOS': 'joshua', 'JDG': 'judges', 'RUT': 'ruth', '1SA': '1samuel', '2SA': '2samuel',
  '1KI': '1kings', '2KI': '2kings', '1CH': '1chronicles', '2CH': '2chronicles', 'EZR': 'ezra',
  'NEH': 'nehemiah', 'EST': 'esther', 'JOB': 'job', 'PSA': 'psalms', 'PRO': 'proverbs',
  'ECC': 'ecclesiastes', 'SNG': 'songofsolomon', 'ISA': 'isaiah', 'JER': 'jeremiah', 'LAM': 'lamentations',
  'EZK': 'ezekiel', 'DAN': 'daniel', 'HOS': 'hosea', 'JOL': 'joel', 'AMO': 'amos', 'OBA': 'obadiah',
  'JON': 'jonah', 'MIC': 'micah', 'NAM': 'nahum', 'HAB': 'habakkuk', 'ZEP': 'zephaniah', 'HAG': 'haggai',
  'ZEC': 'zechariah', 'MAL': 'malachi', 'MAT': 'matthew', 'MRK': 'mark', 'LUK': 'luke', 'JHN': 'john',
  'ACT': 'acts', 'ROM': 'romans', '1CO': '1corinthians', '2CO': '2corinthians', 'GAL': 'galatians',
  'EPH': 'ephesians', 'PHP': 'philippians', 'COL': 'colossians', '1TH': '1thessalonians', '2TH': '2thessalonians',
  '1TI': '1timothy', '2TI': '2timothy', 'TIT': 'titus', 'PHM': 'philemon', 'HEB': 'hebrews', 'JAS': 'james',
  '1PE': '1peter', '2PE': '2peter', '1JN': '1john', '2JN': '2john', '3JN': '3john', 'JUD': 'jude', 'REV': 'revelation'
};

const parseRef = (ref) => {
  const r = ref.trim().replace(/\.\s*/g, ' ').replace(/\s+/g, ' ').trim();
  const match = r.match(/^([\d\s]*[a-zA-Z]+(?:\s+[a-zA-Z]+)*)\s+(\d+):?(\d+)?(?:-?(\d+))?/i);
  if (!match) return null;
  const bookStr = match[1].trim().toLowerCase();
  const chapter = match[2];
  const verseStart = match[3] || "1";
  const verseEnd = match[4] || verseStart;
  const usfm = BOOK_MAP[bookStr] || null;
  return { bookStr, usfm, chapter, verseStart, verseEnd, fullRef: r };
};

const fetchKJV = async (refObj) => {
  try {
    const url = `https://bible-api.com/${encodeURIComponent(refObj.fullRef)}?translation=kjv`;
    const { data } = await axios.get(url, { timeout: 15000 });
    if (!data || !data.verses) return [];
    return data.verses.map(v => ({ num: String(v.verse), text: v.text.trim() }));
  } catch (e) { console.error(`KJV Error: ${e.message}`); return []; }
};

const fetchMBBTAG = async (refObj) => {
  if (!refObj.usfm || !HELLO_AO_BOOKS[refObj.usfm]) return [];
  const bookPath = HELLO_AO_BOOKS[refObj.usfm];
  
  // Try MBBTAG12 (Modern) and MBBTAG (Legacy) as fallbacks
  for (const trans of ['MBBTAG12', 'MBBTAG']) {
    try {
      const url = `https://bible.helloao.org/api/${trans}/${bookPath}/${refObj.chapter}.json`;
      const { data } = await axios.get(url, { timeout: 15000 });
      if (data && data.chapter && data.chapter.verses) {
        const start = parseInt(refObj.verseStart);
        const end = parseInt(refObj.verseEnd);
        return data.chapter.verses
          .filter(v => {
            const vNum = parseInt(v.number);
            return vNum >= start && vNum <= end;
          })
          .map(v => ({
            num: String(v.number),
            text: v.content.map(c => typeof c === 'string' ? c : (c.text || '')).join(' ').replace(/\s+/g, ' ').trim()
          }));
      }
    } catch (e) { /* continue try other translation ID */ }
  }
  return [];
};

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { ref } = req.query;
  if (!ref) return res.status(400).json({ error: "Reference required" });

  const refObj = parseRef(ref);
  if (!refObj) return res.status(400).json({ error: "Invalid format" });

  try {
    const [kjvV, mbbV] = await Promise.all([fetchKJV(refObj), fetchMBBTAG(refObj)]);
    if (kjvV.length > 0 || mbbV.length > 0) {
      return res.status(200).json({ reference: ref, KJV: kjvV, MBBTAG: mbbV });
    } else {
      return res.status(404).json({ error: "Verse not found in cloud databases" });
    }
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
};
