import React, { useState, useEffect, useRef } from 'react';
import pptxgen from "pptxgenjs";
import { Download, Play, FileText, CheckCircle2, Loader2, Sparkles, RefreshCw } from 'lucide-react';
import localBibleData from './bibleData.json';

const DEFAULT_SERMON = `Text: Genesis 1:1, John 3:16
Subject: PASTE PASTOR SERMON HERE AS IS

How to use this tool:

1. Setup Your Title Slide
    - Use "Text:" for your Bible references.
    - Use "Subject:" for your sermon title.

2. Creating Content Slides
    - Every line you type becomes a new slide.
    - Use headers like "Introduction" or "Conclusion".

3. Automatic Bible Verses
    - Just type a reference and it will appear:
    • Psalm 23:1
    • Matthew 5:3-5

4. Smart Indention
    - Indented lines keep the header at the top.
    
    A. This is a sub-point
    B. The header "Smart Indention" stays visible!

5. Theme Customization (Right Side)
    - Click the THEME tab to change colors and fonts.

6. Export as PowerPoint
    - Click EXPORT PPTX to download your presentation.`;

const DEFAULT_LYRICS = `How to use Lyrics Mode:

▶ BLANK LINE = NEW SLIDE
   Separate stanzas or chorus sections with a blank line.
   Each block of text becomes its own slide.

▶ NO BLANK LINES = ONE LINE PER SLIDE
   If your song has no blank lines, each
   individual line gets its own slide.

▶ SECTION LABELS (auto-detected)
   Start a block with a label like Verse, Chorus,
   Bridge, Intro, Outro — it appears at the top
   of the slide in accent color.
`;



// Initial Bible cache from local storage
const getInitialCache = () => {
  const saved = localStorage.getItem('bible_cache');
  return saved ? JSON.parse(saved) : {};
};

const LABEL_REGEX = /^(verse|chorus|bridge|pre-chorus|tag|outro|intro|refrain|hook|coda)\s*[IVXivx\d]*\.?$/i;

const cleanColor = (color) => {
  if (!color) return '#1a1a1a';
  if (typeof color !== 'string') return color;
  // Strip alpha channel (#RRGGBBAA -> #RRGGBB)
  if (color.startsWith('#') && color.length === 9) return color.substring(0, 7);
  return color;
};

const toPptxColor = (hex) => {
  if (!hex) return '000000';
  let cleaned = hex.startsWith('#') ? hex.substring(1) : hex;
  if (cleaned.length > 6) cleaned = cleaned.substring(0, 6);
  return cleaned.toUpperCase();
};

function App() {
  const [sermonText, setSermonText] = useState(DEFAULT_SERMON);
  const [lyricsText, setLyricsText] = useState(DEFAULT_LYRICS);
  const [slides, setSlides] = useState([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [status, setStatus] = useState('');
  const [bibleCache, setBibleCache] = useState(getInitialCache());
  const [activeSlideIndex, setActiveSlideIndex] = useState(0);
  const screenRef = useRef(null);
  const projectorWindowRef = useRef(null);
  const [isBlackout, setIsBlackout] = useState(false);
  const [mode, setMode] = useState('sermon'); // 'sermon' | 'lyrics'
  const [activeTab, setActiveTab] = useState('script'); // 'script' or 'theme'
  const [theme, setTheme] = useState({
    bg: '#111111',
    card: '#F5F5F5',
    text: '#1A1A1A',
    titleColor: '#000000',
    subtitleColor: '#666666',
    accent: '#f13f48ff',
    fontFace: 'Arial',
    sizeMultiplier: 1.0,
    uppercase: true,
    italic: true,
    bold: true,
    bgImage: null,
    overlayOpacity: 0.8
  });

  useEffect(() => {
    localStorage.setItem('bible_cache', JSON.stringify(bibleCache));
  }, [bibleCache]);

  // Broadcast channel for sync
  useEffect(() => {
    const channel = new BroadcastChannel('sermon_flow_sync');
    channel.postMessage({
      type: 'update',
      slide: slides[activeSlideIndex],
      theme,
      isBlackout,
      index: activeSlideIndex
    });
    return () => channel.close();
  }, [activeSlideIndex, slides, theme, isBlackout]);

  useEffect(() => {
    const handleKeyDown = (e) => {
      if (document.activeElement.tagName === 'TEXTAREA') return;

      if (e.key === 'ArrowDown' || e.key === ' ' || e.key === 'ArrowRight') {
        setActiveSlideIndex(prev => Math.min(prev + 1, slides.length - 1));
        e.preventDefault();
      } else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') {
        setActiveSlideIndex(prev => Math.max(prev - 1, 0));
        e.preventDefault();
      } else if (e.key.toLowerCase() === 'b') {
        setIsBlackout(prev => !prev);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [slides]);

  const clearCache = () => {
    setBibleCache({});
    localStorage.removeItem('bible_cache');
    setStatus('Cache cleared! Re-fetching verses...');
    setTimeout(() => setStatus(''), 3000);
  };

  const toggleFullscreen = () => {
    if (!document.fullscreenElement) {
      screenRef.current?.requestFullscreen().catch(err => {
        console.error(`Error attempting to enable full-screen mode: ${err.message}`);
      });
    } else {
      document.exitFullscreen();
    }
  };

  const openProjectorWindow = () => {
    if (projectorWindowRef.current && !projectorWindowRef.current.closed) {
      projectorWindowRef.current.focus();
      return;
    }

    const width = 1280;
    const height = 720;
    const left = window.screen.width; // Try to open on the second monitor
    const top = 0;

    const projectorWindow = window.open(
      '',
      'SermonFlowProjector',
      `width=${width},height=${height},left=${left},top=${top},menubar=no,toolbar=no,location=no,status=no`
    );

    if (projectorWindow) {
      projectorWindow.document.title = "Sermon Flow - Projector";
      projectorWindow.document.body.style.margin = '0';
      projectorWindow.document.body.style.padding = '0';
      projectorWindow.document.body.style.overflow = 'hidden';
      projectorWindow.document.body.style.backgroundColor = '#000';
      projectorWindow.document.body.innerHTML = '<div id="projector-root"></div>';

      // Inject Styles
      const style = projectorWindow.document.createElement('style');
      style.textContent = `
        body { margin: 0; font-family: sans-serif; }
        #projector-root { width: 100vw; height: 100vh; display: flex; align-items: center; justify-content: center; position: relative; overflow: hidden; background: #000; }
        .slide-content { text-align: center; width: 100%; height: 100%; display: flex; flex-direction: column; justify-content: center; align-items: center; padding: 5%; box-sizing: border-box; z-index: 1; transition: opacity 0.5s ease; }
        .overlay { position: absolute; inset: 0; z-index: 0; }
        .background { position: absolute; inset: 0; background-size: cover; background-position: center; z-index: -1; }
      `;
      projectorWindow.document.head.appendChild(style);

      projectorWindowRef.current = projectorWindow;

      // Handle Sync
      const channel = new BroadcastChannel('sermon_flow_sync');

      const updateContent = (data) => {
        const root = projectorWindow.document.getElementById('projector-root');
        if (!root) return;

        if (data.isBlackout) {
          root.innerHTML = '';
          root.style.backgroundColor = '#000';
          return;
        }

        const slide = data.slide;
        const theme = data.theme;
        if (!slide) return;

        // Dynamic Styles for Font Sizes
        const getFontSize = (text, type) => {
          let baseSize = 40;
          const length = text?.length || 0;
          if (type === 'title') baseSize = length > 40 ? 35 : 48;
          else if (type === 'subtitle') baseSize = 22;
          else if (type === 'content') {
            if (length < 30) baseSize = 50;
            else if (length < 60) baseSize = 38;
            else if (length < 120) baseSize = 28;
            else if (length < 200) baseSize = 22;
            else baseSize = 18;
          }
          else if (type === 'scripture') {
            if (length < 60) baseSize = 53;
            else if (length < 100) baseSize = 45;
            else if (length < 160) baseSize = 37;
            else if (length < 250) baseSize = 32;
            else if (length < 380) baseSize = 28;
            else baseSize = 24;
          }
          
          const multiplier = theme.sizeMultiplier;
          return `${baseSize * multiplier * 0.14}vw`;
        };

        const getLyricFontSize = (lines) => {
          const totalChars = lines.join(' ').length || 1;
          const lineCount = lines.length;
          let base;
          if (lineCount <= 1) base = totalChars < 60 ? 52 : totalChars < 120 ? 38 : 28;
          else if (lineCount <= 2) base = totalChars < 100 ? 44 : totalChars < 180 ? 34 : 26;
          else if (lineCount <= 4) base = totalChars < 160 ? 36 : totalChars < 280 ? 28 : 22;
          else base = totalChars < 300 ? 28 : 20;
          return `${base * theme.sizeMultiplier * 0.14}vw`;
        };

        let contentHtml = '';
        const textTransform = theme.uppercase ? 'uppercase' : 'none';
        const fontWeight = theme.bold ? '700' : '400';
        const fontStyle = theme.italic ? 'italic' : 'normal';

        if (slide.type === 'title') {
          contentHtml = `
            <div style="text-align:center; font-weight:${fontWeight}; font-style:${fontStyle}; text-transform:${textTransform}; color:${theme.titleColor};">
              <div style="font-size:${getFontSize(slide.title, 'title')}; line-height:1.1; margin-bottom:1vw;">${slide.title}</div>
              <div style="font-size:${getFontSize(slide.subtitle, 'subtitle')}; color:${theme.subtitleColor}; font-weight:600;">${slide.subtitle || ''}</div>
            </div>
          `;
        } else if (slide.type === 'content') {
          contentHtml = `
            <div style="width: 100%; height: 100%; display: flex; flex-direction: column; text-transform: ${textTransform}; font-weight: ${fontWeight}; font-style: ${fontStyle}; color: ${theme.text};">
              ${slide.mainTitle ? `<div style="flex: 0 0 auto; text-align: center; font-size: 3.5vw; color: ${theme.accent}; font-weight: 800; padding: 0.5vw 0;">${slide.mainTitle}</div>` : ''}
              <div style="flex: 1 1 auto; display: flex; align-items: center; justify-content: center; text-align: center; overflow: hidden; padding: 0 4%;">
                <div style="font-size:${getFontSize(slide.title, 'content')}; line-height:1.2;">${slide.title}</div>
              </div>
            </div>
          `;
        } else if (slide.type === 'scripture') {
          contentHtml = `
            <div style="width: 100%; height: 100%; display: flex; flex-direction: column; text-transform: ${textTransform}; font-weight: ${fontWeight}; font-style: ${fontStyle};">
              <div style="flex: 0 0 auto; text-align: center; font-size: 3.5vw; color: ${theme.accent}; font-weight: 800; padding: 0.5vw 0;">${slide.reference}</div>
              <div style="flex: 1 1 auto; display: flex; align-items: center; justify-content: center; text-align: center; color: ${theme.text}; overflow: hidden; padding: 0 3%;">
                <div style="font-size: ${getFontSize(slide.text, 'scripture')}; line-height: 1.2;">${slide.verseNum ? `<sup style="font-size: 0.6em; opacity: 0.8; margin-right: 0.1em;">${slide.verseNum}</sup>` : ''}&ldquo;${slide.text}&rdquo;</div>
              </div>
            </div>
          `;
        } else if (slide.type === 'lyric') {
          const lyricFontSize = getLyricFontSize(slide.lines || []);
          const linesHtml = (slide.lines || []).map(l => `<div>${l}</div>`).join('');
          contentHtml = `
            <div style="width: 100%; height: 100%; display: flex; flex-direction: column; text-transform: ${textTransform}; font-weight: ${fontWeight}; font-style: ${fontStyle};">
              ${slide.label ? `<div style="flex: 0 0 auto; text-align: center; font-size: 3vw; color: ${theme.accent}; font-weight: 800; padding: 0.5vw 0; letter-spacing: 0.1em;">${slide.label}</div>` : ''}
              <div style="flex: 1 1 auto; display: flex; align-items: center; justify-content: center; text-align: center; color: ${theme.text}; overflow: hidden; padding: 0 4%;">
                <div style="font-size: ${lyricFontSize}; line-height: 1.2;">${linesHtml}</div>
              </div>
            </div>
          `;
        }

        root.innerHTML = `
          <div class="background" style="background-color:${theme.bg}; background-image:${theme.bgImage ? `url(${theme.bgImage})` : 'none'};"></div>
          <div class="overlay" style="background-color:${theme.card}; opacity:${theme.bgImage ? theme.overlayOpacity : 1};"></div>
          <div class="slide-content" style="font-family:${theme.fontFace};">
            ${contentHtml}
          <div class="brand-watermark" style="color:${theme.accent};">By Gloreanne</div>
        `;
      };

      channel.onmessage = (event) => {
        if (event.data.type === 'update') {
          updateContent(event.data);
        }
      };

      // Initial Sync
      updateContent({
        slide: slides[activeSlideIndex],
        theme,
        isBlackout,
        index: activeSlideIndex
      });

      projectorWindow.onbeforeunload = () => {
        channel.close();
      };
    }
  };

  useEffect(() => {
    if (mode === 'lyrics') parseLyrics();
    else parseSermon();
  }, [sermonText, lyricsText, bibleCache, mode]);

  const parseLyrics = () => {
    const text = lyricsText.trim();
    const blocks = text.split(/\n[ \t]*\n/);
    let currentSlides = [];

    if (blocks.length > 1) {
      // Blank-line-delimited: each block = one slide
      blocks.forEach(block => {
        const lines = block.split('\n').map(l => l.trim()).filter(l => l);
        if (lines.length === 0) return;

        let label = '';
        let contentLines = lines;
        if (LABEL_REGEX.test(lines[0])) {
          label = lines[0];
          contentLines = lines.slice(1);
        }
        if (contentLines.length === 0) return;

        currentSlides.push({
          type: 'lyric',
          label,
          lines: contentLines,
          text: contentLines.join('\n')
        });
      });
    } else {
      // No blank lines: each individual line = one slide
      const lines = text.split('\n').map(l => l.trim()).filter(l => l);
      lines.forEach(line => {
        let label = '';
        let content = line;
        if (LABEL_REGEX.test(line)) {
          label = line;
          content = '';
        }
        if (!content && !label) return;
        currentSlides.push({
          type: 'lyric',
          label,
          lines: content ? [content] : [],
          text: content
        });
      });
    }

    setSlides(currentSlides);
  };

  const fetchVerseData = async (ref) => {
    if (bibleCache[ref]) return;

    // First check local JSON file for faster loading/fallback
    if (localBibleData[ref]) {
      setBibleCache(prev => ({
        ...prev,
        [ref]: localBibleData[ref]
      }));
      return;
    }

    try {
      // Logic: 
      // 1. Use VITE_API_URL if provided
      // 2. If locally developing and no URL, try localhost:3001
      // 3. Otherwise, use Netlify Functions path
      const apiUrl = import.meta.env.VITE_API_URL;
      const endpoint = apiUrl
        ? `${apiUrl}/api/bible`
        : (import.meta.env.DEV ? 'http://localhost:3001/api/bible' : '/api/bible');

      console.log(`Fetching from: ${endpoint}?ref=${ref}`);
      const resp = await fetch(`${endpoint}?ref=${encodeURIComponent(ref)}`);

      if (!resp.ok) throw new Error(`Server returned ${resp.status}`);

      const data = await resp.json();
      if (data.KJV && data.MBBTAG) {
        setBibleCache(prev => ({
          ...prev,
          [ref]: data
        }));
      } else {
        throw new Error("Invalid API response format");
      }
    } catch (e) {
      console.error("Failed to fetch verse", ref, e);
      // Mark as error in cache to prevent infinite loading state
      setBibleCache(prev => ({
        ...prev,
        [ref]: { error: true, message: e.message }
      }));
    }
  };

  const handleImageUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setTheme({ ...theme, bgImage: reader.result });
      };
      reader.readAsDataURL(file);
    }
  };

  const getDynamicFontSize = (text, type, isPptx = false) => {
    let baseSize = 40;
    const length = text?.length || 0;

    if (type === 'title') baseSize = length > 40 ? 35 : 48;
    else if (type === 'subtitle') baseSize = 22;
    else if (type === 'content') {
      if (length < 30) baseSize = 50;
      else if (length < 60) baseSize = 38;
      else if (length < 120) baseSize = 28;
      else if (length < 200) baseSize = 22;
      else baseSize = 18;
    }
    else if (type === 'scripture') {
      if (length < 60) baseSize = 53;
      else if (length < 100) baseSize = 45;
      else if (length < 160) baseSize = 37;
      else if (length < 250) baseSize = 32;
      else if (length < 380) baseSize = 28;
      else baseSize = 24;
    }

    const multiplier = theme.sizeMultiplier;
    const scaled = baseSize * multiplier;
    return isPptx ? scaled : `${scaled * 0.14}cqw`;
  };

  const parseSermon = () => {
    const rawLines = sermonText.split('\n');
    let currentSlides = [];

    // Service Header Detection (for filenames, not slides)
    const firstLine = rawLines[0]?.trim()?.toLowerCase();
    const isServiceHeader = (firstLine === 'divine service' || firstLine === 'sunday school');
    const startIndex = isServiceHeader ? 1 : 0;

    // Global Headers for the Title Slide
    const textMatch = sermonText.match(/Text:\s*(.*)/i);
    const subjectMatch = sermonText.match(/Subject:\s*(.*)/i);
    const title = subjectMatch ? subjectMatch[1].trim() : "Sermon";
    const subtitle = textMatch ? textMatch[1].trim() : "";

    currentSlides.push({
      type: 'title',
      title: title.toUpperCase(),
      subtitle: subtitle
    });

    let contextStack = [];

    rawLines.slice(startIndex).forEach(line => {
      const trimmed = line.trim();
      if (!trimmed) return;

      // Skip internal headers
      if (trimmed.toLowerCase().startsWith('text:') || trimmed.toLowerCase().startsWith('subject:')) return;

      const firstCharIndex = line.search(/\S/);
      const currentIndent = firstCharIndex === -1 ? 0 : firstCharIndex;
      const currentIsListItem = trimmed.match(/^([a-z0-9]+\.|[•\-\*])/i);

      while (contextStack.length > 0) {
        const last = contextStack[contextStack.length - 1];
        if (currentIndent < last.indent) {
          contextStack.pop();
          continue;
        }
        if (currentIndent === last.indent) {
          contextStack.pop();
          continue;
        }
        break;
      }

      const parentContext = contextStack.length > 0 ? contextStack[contextStack.length - 1].text : "";
      const scriptureRegex = /([1-9]?\s?[a-zA-Z]+\.?\s\d+:\d+([\-\u2013\u2014]\d+)?)/;
      const scriptureMatch = trimmed.match(scriptureRegex);
      const ref = scriptureMatch ? scriptureMatch[1].trim() : null;

      if (ref) {
        const data = bibleCache[ref];
        if (!data) {
          fetchVerseData(ref);
          currentSlides.push({ type: 'scripture', context: parentContext, reference: ref, version: 'English (KJV)', text: "Loading..." });
        } else if (data.error) {
          currentSlides.push({ type: 'scripture', context: parentContext, reference: ref, version: 'Error', text: "Verse not found." });
        } else {
          const kjvArray = Array.isArray(data.KJV) ? data.KJV : [{ num: '', text: data.KJV }];
          const mbbArray = Array.isArray(data.MBBTAG) ? data.MBBTAG : (data.MBBTAG ? [{ num: '', text: data.MBBTAG }] : []);
          kjvArray.forEach(v => currentSlides.push({ type: 'scripture', context: parentContext, reference: v.num ? `${ref} (v.${v.num})` : ref, version: 'English (KJV)', text: v.text, verseNum: v.num }));
          mbbArray.forEach(v => currentSlides.push({ type: 'scripture', context: parentContext, reference: v.num ? `${ref} (v.${v.num})` : ref, version: 'Tagalog (MBBTAG)', text: v.text, verseNum: v.num }));
        }
      } else {
        currentSlides.push({ type: 'content', mainTitle: parentContext, title: trimmed });
        contextStack.push({ indent: currentIndent, text: trimmed });
      }
    });

    setSlides(currentSlides);
  };

  const getLyricFontSize = (lines, isPptx = false) => {
    const totalChars = lines.join(' ').length;
    const lineCount = lines.length;
    let base;
    if (lineCount <= 1) base = totalChars < 60 ? 52 : totalChars < 120 ? 38 : 28;
    else if (lineCount <= 2) base = totalChars < 100 ? 44 : totalChars < 180 ? 34 : 26;
    else if (lineCount <= 4) base = totalChars < 160 ? 36 : totalChars < 280 ? 28 : 22;
    else base = totalChars < 300 ? 28 : 20;
    const scaled = base * theme.sizeMultiplier;
    return isPptx ? scaled : `${scaled * 0.14}cqw`;
  };

  const generatePPTX = async () => {
    setIsGenerating(true);
    setStatus('Creating PowerPoint...');

    try {
      let pptx = new pptxgen();
      pptx.layout = "LAYOUT_16x9";

      // Helper to ensure PPTX colors are 6-digit hex without alpha or '#'
      for (const slideData of slides) {
        let slide = pptx.addSlide();

        if (theme.bgImage) {
          slide.background = { data: theme.bgImage };
        } else {
          slide.background = { color: toPptxColor(theme.bg) };
        }

        // Card Overlay / Slide Surface
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: {
            color: toPptxColor(theme.card),
            alpha: theme.bgImage ? (theme.overlayOpacity * 100) : 0 // Fully opaque card if no image
          }
        });

        const processText = (txt) => theme.uppercase ? (txt || "").toUpperCase() : (txt || "");

        if (slideData.type === 'title') {
          slide.addText(processText(slideData.title), {
            x: 0.5, y: 1.2, w: 9.0, h: 2.0,
            fontSize: getDynamicFontSize(slideData.title, 'title', true),
            fontFace: theme.fontFace,
            bold: theme.bold, italic: theme.italic, color: toPptxColor(theme.titleColor),
            align: "center", valign: "middle", shrinkText: true
          });
          slide.addText(processText(slideData.subtitle), {
            x: 0.5, y: 3.5, w: 9, h: 1,
            fontSize: getDynamicFontSize(slideData.subtitle, 'subtitle', true),
            fontFace: theme.fontFace,
            bold: theme.bold, italic: theme.italic, color: toPptxColor(theme.subtitleColor),
            align: "center", valign: "top", shrinkText: true
          });
        }
        else if (slideData.type === 'content') {
          if (slideData.mainTitle) {
            slide.addText(processText(slideData.mainTitle), {
              x: 0.5, y: 0.3, w: 9.0, h: 0.8,
              fontSize: 24 * theme.sizeMultiplier,
              fontFace: theme.fontFace,
              bold: true, color: toPptxColor(theme.accent),
              align: "center", valign: "bottom"
            });
            slide.addText(processText(slideData.title), {
              x: 0.5, y: 1.4, w: 9.0, h: 4.0,
              fontSize: getDynamicFontSize(slideData.title, 'content', true),
              fontFace: theme.fontFace,
              bold: theme.bold, italic: theme.italic, color: toPptxColor(theme.text),
              align: "center", valign: "middle", wrap: true, autoFit: true
            });
          } else {
            slide.addText(processText(slideData.title), {
              x: 0.5, y: 0.5, w: 9.0, h: 4.625,
              fontSize: getDynamicFontSize(slideData.title, 'content', true),
              fontFace: theme.fontFace,
              bold: theme.bold, italic: theme.italic, color: toPptxColor(theme.text),
              align: "center", valign: "middle", wrap: true, autoFit: true
            });
          }
        }
        else if (slideData.type === 'lyric') {
          const lyricFontSize = getLyricFontSize(slideData.lines, true);

          if (slideData.label) {
            slide.addText(processText(slideData.label), {
              x: 0.5, y: 0.3, w: 9.0, h: 0.6,
              fontSize: 20 * theme.sizeMultiplier, fontFace: theme.fontFace,
              bold: true, color: toPptxColor(theme.accent),
              align: 'center', valign: 'bottom'
            });
          }

          slide.addText(processText(slideData.text), {
            x: 0.5, y: slideData.label ? 1.2 : 0.5,
            w: 9.0, h: slideData.label ? 3.9 : 4.625,
            fontSize: lyricFontSize,
            fontFace: theme.fontFace,
            bold: theme.bold, italic: theme.italic,
            color: toPptxColor(theme.text),
            align: 'center', valign: 'middle',
            shrinkText: true, breakLine: true
          });
        }
        else if (slideData.type === 'scripture') {
          // Fixed Reference: Clear of the text area
          let displayRef = slideData.reference;

          slide.addText(processText(displayRef), {
            x: 0.5, y: 0.3, w: 9.0, h: 0.8,
            fontSize: 24 * theme.sizeMultiplier, 
            fontFace: theme.fontFace, 
            bold: true, 
            color: toPptxColor(theme.accent),
            align: "center", 
            valign: "bottom"
          });

          // Verse Text: Lower starting point to avoid any collision
          let textRuns = [];
          if (slideData.verseNum) {
            textRuns.push({ text: processText(slideData.verseNum + ""), options: { superscript: true } });
            textRuns.push({ text: processText(` \u201C${slideData.text}\u201D`) });
          } else {
            textRuns.push({ text: processText(`\u201C${slideData.text}\u201D`) });
          }

          slide.addText(textRuns, {
            x: 0.5, y: 1.4, w: 9.0, h: 4.0,
            fontSize: getDynamicFontSize(slideData.text, 'scripture', true),
            fontFace: theme.fontFace,
            bold: theme.bold,
            italic: theme.italic,
            color: toPptxColor(theme.text),
            align: "center",
            valign: "middle",
            wrap: true,
            autoFit: true
          });
        }
      }

      // 1. Dynamic Filename Calculation
      const now = new Date();
      const monthDay = `${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      const sanitize = (fn) => fn.replace(/[/\\?%*:|"<>]/g, '').replace(/\s+/g, '_').substring(0, 50);

      let filename = 'presentation';
      if (mode === 'lyrics') {
        const first = slides[0];
        filename = (first && (first.lines?.[0] || first.title || first.text)) || 'Song_Lyrics';
      } else {
        const firstLine = sermonText.trim().split('\n')[0].toLowerCase();
        let serviceType = 'Sermon';
        if (firstLine.includes('divine service')) serviceType = 'Divine_Service';
        else if (firstLine.includes('sunday school')) serviceType = 'Sunday_School';
        
        filename = `${monthDay}_${serviceType}`;
      }

      // 2. Write File
      await pptx.writeFile({ fileName: `${sanitize(filename)}.pptx` });
      setStatus('Export Ready!');
    } catch (error) {
      console.error(error);
      setStatus('Export Error.');
    } finally {
      setIsGenerating(false);
      setTimeout(() => setStatus(''), 5000);
    }
  };

  return (
    <div className="app-container">
      <header>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
          <img src="/church-logo.png" alt="Church Logo" style={{ height: '40px', width: 'auto' }} />
          <h1>BBC <span>PRESENTATION</span></h1>
        </div>
        <div style={{ textAlign: 'right' }}>
          <p>PRESENTATION CONSOLE</p>
          <p style={{ fontSize: '0.65rem', opacity: 0.8, color: cleanColor(theme.accent), marginTop: '2px', letterSpacing: '0.1em', fontWeight: 'bold' }}>BY GLOREANNE</p>
          <div style={{ fontSize: '0.7rem', color: 'var(--primary)', fontWeight: 'bold', letterSpacing: '0.05em' }}>• {slides.length} SLIDES</div>
        </div>
      </header>

      <main className="main-content">
        <section className="presenter-sidebar">
          <div className="sidebar-tabs">
            <button
              className={`tab-btn ${activeTab === 'script' ? 'active' : ''}`}
              onClick={() => setActiveTab('script')}
            >
              SCRIPT
            </button>
            <button
              className={`tab-btn ${activeTab === 'theme' ? 'active' : ''}`}
              onClick={() => setActiveTab('theme')}
            >
              THEME
            </button>
          </div>

          <div className="sidebar-content">
            {activeTab === 'script' ? (
              <>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                    <div className="section-label" style={{ margin: 0 }}>
                      {mode === 'lyrics' ? '♪ Song Lyrics' : <><FileText size={14} /> Sermon Script</>}
                    </div>
                    <div className="mode-switch">
                      <button
                        onClick={() => { setMode('sermon'); setActiveSlideIndex(0); }}
                        className={mode === 'sermon' ? 'active' : ''}
                      >SERMON</button>
                      <button
                        onClick={() => { setMode('lyrics'); setActiveSlideIndex(0); }}
                        className={mode === 'lyrics' ? 'active' : ''}
                      >LYRICS</button>
                    </div>
                  </div>
                  <textarea
                    value={mode === 'lyrics' ? lyricsText : sermonText}
                    onChange={(e) => mode === 'lyrics' ? setLyricsText(e.target.value) : setSermonText(e.target.value)}
                    placeholder={mode === 'lyrics' ? 'Paste song lyrics here...' : 'Type or paste sermon content here...'}
                    spellCheck="false"
                  />
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
                  <div className="section-label" style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span><Play size={14} /> Active Cues</span>
                    <button
                      onClick={() => setIsBlackout(!isBlackout)}
                      className={`blackout-btn ${isBlackout ? 'active' : ''}`}
                    >
                      {isBlackout ? 'LIVE: BLACKOUT' : 'BLACKOUT (B)'}
                    </button>
                  </div>
                  <div className="cues-list">
                    {slides.map((slide, idx) => (
                      <div
                        key={idx}
                        className={`cue-item ${activeSlideIndex === idx ? 'active-cue' : ''}`}
                        onClick={() => setActiveSlideIndex(idx)}
                      >
                        <span className="cue-icon">
                          {slide.type === 'title' && <Sparkles size={12} />}
                          {slide.type === 'content' && <FileText size={12} />}
                          {slide.type === 'scripture' && <Play size={12} />}
                          {slide.type === 'lyric' && <span style={{ fontSize: '10px' }}>♪</span>}
                        </span>
                        <span className="cue-num">{String(idx + 1).padStart(2, '0')}</span>
                        <span className="cue-text">
                          {slide.type === 'scripture' && slide.reference}
                          {slide.type === 'lyric' && (slide.label ? `[${slide.label}] ${slide.lines[0] || ''}` : (slide.lines[0] || '---'))}
                          {(slide.type === 'title' || slide.type === 'content') && (slide.title || '---')}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              </>
            ) : (
              <div className="theme-editor">
                <div className="section-label">Design Customization</div>

                <div className="theme-field">
                  <label>Background Image</label>
                  <div style={{ display: 'flex', gap: '0.5rem' }}>
                    <input type="file" accept="image/*" onChange={handleImageUpload} />
                    {theme.bgImage && (
                      <button onClick={() => setTheme({ ...theme, bgImage: null })} className="clear-img">X</button>
                    )}
                  </div>
                </div>

                <div className="grid-2">
                  <div className="theme-field">
                    <label>BG Color</label>
                    <input type="color" value={theme.bg} onChange={(e) => setTheme({ ...theme, bg: e.target.value })} />
                  </div>
                  <div className="theme-field">
                    <label>Slide Color</label>
                    <input type="color" value={theme.card} onChange={(e) => setTheme({ ...theme, card: e.target.value })} />
                  </div>
                </div>

                <div className="theme-field">
                  <label>Font Family</label>
                  <select
                    value={theme.fontFace}
                    onChange={(e) => setTheme({ ...theme, fontFace: e.target.value })}
                  >
                    <option value="Arial">Arial</option>
                    <option value="Calibri">Calibri</option>
                    <option value="Impact">Impact</option>
                    <option value="Georgia">Georgia</option>
                    <option value="Verdana">Verdana</option>
                    <option value="Times New Roman">Times New Roman</option>
                  </select>
                </div>

                <div className="grid-2">
                  <div className="theme-field">
                    <label>Title</label>
                    <input type="color" value={theme.titleColor} onChange={(e) => setTheme({ ...theme, titleColor: e.target.value })} />
                  </div>
                  <div className="theme-field">
                    <label>Subtitle</label>
                    <input type="color" value={theme.subtitleColor} onChange={(e) => setTheme({ ...theme, subtitleColor: e.target.value })} />
                  </div>
                </div>

                <div className="grid-2">
                  <div className="theme-field">
                    <label>Body Text</label>
                    <input type="color" value={theme.text} onChange={(e) => setTheme({ ...theme, text: e.target.value })} />
                  </div>
                  <div className="theme-field">
                    <label>Reference</label>
                    <input type="color" value={theme.accent} onChange={(e) => setTheme({ ...theme, accent: e.target.value })} />
                  </div>
                </div>

                <div className="theme-field">
                  <label>Overlay Opacity: {(theme.overlayOpacity * 100).toFixed(0)}%</label>
                  <input type="range" min="0" max="1" step="0.05" value={theme.overlayOpacity} onChange={(e) => setTheme({ ...theme, overlayOpacity: parseFloat(e.target.value) })} />
                </div>

                <div className="theme-field">
                  <label>Font Size: {theme.sizeMultiplier.toFixed(1)}x</label>
                  <input type="range" min="0.5" max="2.0" step="0.1" value={theme.sizeMultiplier} onChange={(e) => setTheme({ ...theme, sizeMultiplier: parseFloat(e.target.value) })} />
                </div>

                <div className="grid-2">
                  <button onClick={() => setTheme({ ...theme, uppercase: !theme.uppercase })} className={theme.uppercase ? 'active' : ''}>UPPER</button>
                  <button onClick={() => setTheme({ ...theme, italic: !theme.italic })} className={theme.italic ? 'active' : ''}>ITALIC</button>
                </div>
              </div>
            )}

            <div style={{ display: 'flex', gap: '0.75rem' }}>
              <button
                className="btn btn-primary"
                onClick={generatePPTX}
                disabled={isGenerating}
                style={{ flex: 1 }}
              >
                {isGenerating ? <Loader2 className="loader" /> : <Download size={18} />}
                EXPORT PPTX
              </button>
              <button
                className="btn refresh-btn"
                onClick={clearCache}
                title="Refresh Bible Cache"
              >
                <RefreshCw size={18} />
              </button>
            </div>
            {status && <div className="status-msg">{status}</div>}
          </div>
        </section>

        <section className="live-projection-area">
          <div className="projection-header">
            <div className="section-label" style={{ margin: 0 }}><Play size={14} /> LIVE PROJECTION</div>
            <div style={{ display: 'flex', gap: '0.5rem' }}>
              <button
                className="btn btn-glass"
                onClick={openProjectorWindow}
              >
                <Download size={12} style={{ transform: 'rotate(-90deg)' }} /> POP-OUT
              </button>
              <button
                className="btn btn-primary"
                onClick={toggleFullscreen}
              >
                <Play size={12} fill="currentColor" /> FULLSCREEN
              </button>
            </div>
          </div>
          <div className="screen-container" style={{ position: 'relative', width: '100%', maxWidth: 'calc((100vh - 12rem) * 16 / 9)', aspectRatio: '16/9', overflow: 'hidden', borderRadius: '8px', border: '1px solid var(--glass-border)', containerType: 'inline-size' }}>
            {slides[activeSlideIndex] && (
              <div className="projection-screen" ref={screenRef} style={{
                visibility: isBlackout ? 'hidden' : 'visible',
                backgroundColor: cleanColor(theme.bg),
                backgroundImage: theme.bgImage ? `url(${theme.bgImage})` : 'none',
                backgroundSize: 'cover',
                backgroundPosition: 'center',
                width: '100%', height: '100%', position: 'absolute', inset: 0
              }}>
                {/* Semi-transparent overlay / Slide Card */}
                <div style={{
                  position: 'absolute',
                  inset: 0,
                  backgroundColor: cleanColor(theme.card),
                  opacity: theme.bgImage ? theme.overlayOpacity : 1, // Becomes the main color if no image
                  zIndex: 0
                }} />

                <div style={{
                  position: 'relative',
                  zIndex: 1,
                  width: '100%',
                  height: '100%',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  padding: '4%',
                  fontFamily: theme.fontFace,
                  boxSizing: 'border-box'
                }}>
                  {slides[activeSlideIndex].type === 'title' && (
                    <div key={activeSlideIndex} className="slide-anim" style={{ textAlign: 'center', fontWeight: theme.bold ? '700' : '400', fontStyle: theme.italic ? 'italic' : 'normal', textTransform: theme.uppercase ? 'uppercase' : 'none', color: cleanColor(theme.titleColor) }}>
                      <div style={{ fontSize: getDynamicFontSize(slides[activeSlideIndex].title, 'title'), letterSpacing: '-0.03em', lineHeight: '1.1', marginBottom: '1vw' }}>{slides[activeSlideIndex].title}</div>
                      <div style={{ fontSize: getDynamicFontSize(slides[activeSlideIndex].subtitle, 'subtitle'), color: cleanColor(theme.subtitleColor), fontWeight: '600' }}>{slides[activeSlideIndex].subtitle}</div>
                    </div>
                  )}
                  {slides[activeSlideIndex].type === 'content' && (
                    <div key={activeSlideIndex} className="slide-anim" style={{ 
                      width: '100%', height: '100%', display: 'flex', flexDirection: 'column',
                      textTransform: theme.uppercase ? 'uppercase' : 'none', fontWeight: theme.bold ? '700' : '400', fontStyle: theme.italic ? 'italic' : 'normal', color: cleanColor(theme.text) 
                    }}>
                      {slides[activeSlideIndex].mainTitle && (
                        <div style={{ flex: '0 0 auto', textAlign: 'center', fontSize: '3.5cqw', color: cleanColor(theme.accent), padding: '0.5cqw 0', fontWeight: '800' }}>
                          {slides[activeSlideIndex].mainTitle}
                        </div>
                      )}
                      <div style={{ flex: '1 1 auto', display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', overflow: 'hidden', padding: '0 4%' }}>
                        <div style={{ fontSize: getDynamicFontSize(slides[activeSlideIndex].title, 'content'), lineHeight: '1.2', letterSpacing: '-0.01em' }}>
                          {slides[activeSlideIndex].title}
                        </div>
                      </div>
                    </div>
                  )}
                  {slides[activeSlideIndex].type === 'scripture' && (
                    <div key={activeSlideIndex} className="slide-anim" style={{ 
                      width: '100%', height: '100%', display: 'flex', flexDirection: 'column',
                      textTransform: theme.uppercase ? 'uppercase' : 'none', fontWeight: theme.bold ? '700' : '400', 
                      fontStyle: theme.italic ? 'italic' : 'normal', color: cleanColor(theme.text) 
                    }}>
                      <div style={{ flex: '0 0 auto', textAlign: 'center', fontSize: '3.5cqw', color: cleanColor(theme.accent), padding: '0.5cqw 0', fontWeight: '800' }}>
                        {slides[activeSlideIndex].reference}
                      </div>
                      <div style={{ flex: '1 1 auto', display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', overflow: 'hidden', padding: '0 4%' }}>
                        <div style={{ fontSize: getDynamicFontSize(slides[activeSlideIndex].text, 'scripture'), lineHeight: '1.25' }}>
                          {slides[activeSlideIndex].verseNum && <sup style={{ fontSize: '0.6em', opacity: 0.8, marginRight: '0.1em' }}>{slides[activeSlideIndex].verseNum}</sup>}
                          &ldquo;{slides[activeSlideIndex].text}&rdquo;
                        </div>
                      </div>
                    </div>
                  )}
                  {slides[activeSlideIndex].type === 'lyric' && (
                    <div key={activeSlideIndex} className="slide-anim" style={{
                      width: '100%', height: '100%', display: 'flex', flexDirection: 'column',
                      textTransform: theme.uppercase ? 'uppercase' : 'none',
                      fontWeight: theme.bold ? '700' : '400',
                      fontStyle: theme.italic ? 'italic' : 'normal',
                      color: cleanColor(theme.text)
                    }}>
                      {slides[activeSlideIndex].label && (
                        <div style={{ flex: '0 0 auto', textAlign: 'center', fontSize: '3cqw', color: cleanColor(theme.accent), padding: '0.5cqw 0', fontWeight: '800', letterSpacing: '0.1em' }}>
                          {slides[activeSlideIndex].label}
                        </div>
                      )}
                      <div style={{ flex: '1 1 auto', display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', overflow: 'hidden', padding: '0 5%' }}>
                        <div style={{ fontSize: getLyricFontSize(slides[activeSlideIndex].lines), lineHeight: '1.2' }}>
                          {slides[activeSlideIndex].lines.map((line, i) => (
                            <div key={i}>{line}</div>
                          ))}
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
          <div className="status-bar">
            <span>COMMAND: {activeSlideIndex + 1} OF {slides.length}</span>
            <span>NAV: ARROWS / SPACE</span>
            <span>SIGNAL: {isBlackout ? 'MUTED' : 'ONLINE'}</span>
          </div>
        </section>
      </main>
    </div>
  );
}

export default App;
