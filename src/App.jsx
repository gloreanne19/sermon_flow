import React, { useState, useEffect, useRef } from 'react';
import pptxgen from "pptxgenjs";
import { Download, Play, FileText, CheckCircle2, Loader2, Sparkles, RefreshCw } from 'lucide-react';
import localBibleData from './bibleData.json';

const DEFAULT_SERMON = `Text: Genesis 1:1, John 3:16
Subject: Welcome to Sermon Flow BBC

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

// Initial Bible cache from local storage
const getInitialCache = () => {
  const saved = localStorage.getItem('bible_cache');
  return saved ? JSON.parse(saved) : {};
};

function App() {
  const [sermonText, setSermonText] = useState(DEFAULT_SERMON);
  const [slides, setSlides] = useState([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [status, setStatus] = useState('');
  const [bibleCache, setBibleCache] = useState(getInitialCache());
  const [activeSlideIndex, setActiveSlideIndex] = useState(0);
  const screenRef = useRef(null);
  const projectorWindowRef = useRef(null);
  const [isBlackout, setIsBlackout] = useState(false);
  const [activeTab, setActiveTab] = useState('script'); // 'script' or 'theme'
  const [theme, setTheme] = useState({
    bg: '#111111',
    card: '#F5F5F5',
    text: '#1A1A1A',
    titleColor: '#000000',
    subtitleColor: '#666666',
    accent: '#6366f1',
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
        body { font-family: sans-serif; }
        #projector-root { width: 100vw; height: 100vh; display: flex; align-items: center; justify-content: center; position: relative; overflow: hidden; }
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
            if (length < 80) baseSize = 40;
            else if (length < 130) baseSize = 34;
            else if (length < 180) baseSize = 28;
            else if (length < 250) baseSize = 24;
            else if (length < 350) baseSize = 20;
            else baseSize = 16;
          }
          return `${baseSize * theme.sizeMultiplier * 0.1}vw`;
        };

        let contentHtml = '';
        const textTransform = theme.uppercase ? 'uppercase' : 'none';
        const fontWeight = theme.bold ? '900' : '400';
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
            <div style="text-align:center; text-transform:${textTransform}; font-weight:${fontWeight}; font-style:${fontStyle}; color:${theme.text}; width:100%;">
              ${slide.mainTitle ? `<div style="font-size:1.8vw; color:${theme.subtitleColor}; margin-bottom:2vw; font-weight:700;">${slide.mainTitle.toUpperCase()}</div>` : ''}
              <div style="font-size:${getFontSize(slide.title, 'content')}; line-height:1.2;">${slide.title}</div>
            </div>
          `;
        } else if (slide.type === 'scripture') {
          contentHtml = `
            <div style="position: absolute; top: 3%; left: 0; width: 100%; text-align: center; text-transform: ${textTransform}; font-weight: ${fontWeight}; font-style: ${fontStyle};">
              <div style="font-size: 2vw; color: ${theme.accent}; font-weight: 800;">${slide.reference}</div>
            </div>
            <div style="width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; text-align: center; text-transform: ${textTransform}; font-weight: ${fontWeight}; font-style: ${fontStyle}; color: ${theme.text}; padding: 8% 5% 2% 5%; box-sizing: border-box;">
              <div style="font-size: ${getFontSize(slide.text, 'scripture')}; line-height: 1.2;">"${slide.text}"</div>
            </div>
          `;
        }

        root.innerHTML = `
          <div class="background" style="background-color:${theme.bg}; background-image:${theme.bgImage ? `url(${theme.bgImage})` : 'none'};"></div>
          <div class="overlay" style="background-color:${theme.card}; opacity:${theme.bgImage ? theme.overlayOpacity : 1};"></div>
          <div class="slide-content" style="font-family:${theme.fontFace};">
            ${contentHtml}
          </div>
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
    parseSermon();
  }, [sermonText, bibleCache]);

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
        : (import.meta.env.DEV ? 'http://localhost:3001/api/bible' : '/.netlify/functions/bible');

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
      if (length < 80) baseSize = 40;
      else if (length < 130) baseSize = 34;
      else if (length < 180) baseSize = 28;
      else if (length < 250) baseSize = 24;
      else if (length < 350) baseSize = 20;
      else baseSize = 16;
    }

    const scaled = baseSize * theme.sizeMultiplier;
    return isPptx ? scaled : `${scaled * 0.1}vw`; // 0.1 factor to convert PPT size to approximate VW
  };

  const parseSermon = () => {
    const rawLines = sermonText.split('\n');
    let currentSlides = [];

    // Extract Global Headers
    const textMatch = sermonText.match(/Text:\s*(.*)/i);
    const subjectMatch = sermonText.match(/Subject:\s*(.*)/i);
    const title = subjectMatch ? subjectMatch[1].trim() : "Sermon";
    const subtitle = textMatch ? textMatch[1].trim() : "";

    currentSlides.push({
      type: 'title',
      title: title,
      subtitle: subtitle
    });

    let contextStack = []; // [{indent: 0, text: "I. ..."}]

    rawLines.forEach(line => {
      const trimmed = line.trim();
      if (!trimmed) return;

      // Skip global title/subject
      if (trimmed.toLowerCase().startsWith('text:') || trimmed.toLowerCase().startsWith('subject:')) return;

      // 1. Calculate Indent
      const firstCharIndex = line.search(/\S/);
      const currentIndent = firstCharIndex === -1 ? 0 : firstCharIndex;

      // 2. Adjust Context Stack (Smarter List Detection)
      const currentIsListItem = trimmed.match(/^([a-z0-9]+\.|[•\-\*])/i);

      while (contextStack.length > 0) {
        const last = contextStack[contextStack.length - 1];

        // If current is less indented, pop
        if (currentIndent < last.indent) {
          contextStack.pop();
          continue;
        }

        // If current is same indent
        if (currentIndent === last.indent) {
          // If the last one was a "Header" (ended with colon) AND current is a list item, 
          // we keep the header for this point.
          if (last.text.endsWith(':') && currentIsListItem) {
            break;
          }
          // Otherwise, it's a sibling, so pop the previous sibling
          contextStack.pop();
          continue;
        }
        break;
      }

      const parentContext = contextStack.length > 0 ? contextStack[contextStack.length - 1].text : "";

      // 3. Detect Scripture (Handle hyphen - and en-dash \u2013)
      const scriptureRegex = /([1-9]?\s?[a-zA-Z]+\.?\s\d+:\d+([\-\u2013\u2014]\d+)?)/;
      const scriptureMatch = trimmed.match(scriptureRegex);

      let ref = null;
      if (scriptureMatch) {
        ref = scriptureMatch[1].trim();
      }

      if (ref) {
        const data = bibleCache[ref];
        if (!data) {
          fetchVerseData(ref);
          currentSlides.push({ type: 'scripture', context: parentContext, reference: ref, version: 'English (KJV)', text: "Loading..." });
        } else if (data.error) {
          currentSlides.push({
            type: 'scripture',
            context: parentContext,
            reference: ref,
            version: 'Error',
            text: "Verse not found or API is offline. Ensure your Bible API server is running."
          });
        } else {
          // Normalize to handle both Array (from API) and String (from local JSON)
          const kjvArray = Array.isArray(data.KJV) ? data.KJV : [{ num: '', text: data.KJV }];
          const mbbArray = Array.isArray(data.MBBTAG) ? data.MBBTAG : (data.MBBTAG ? [{ num: '', text: data.MBBTAG }] : []);

          kjvArray.forEach((verse) => {
            currentSlides.push({
              type: 'scripture',
              context: parentContext,
              reference: verse.num ? `${ref} (v.${verse.num})` : ref,
              version: 'English (KJV)',
              text: verse.text
            });
          });

          mbbArray.forEach((verse) => {
            currentSlides.push({
              type: 'scripture',
              context: parentContext,
              reference: verse.num ? `${ref} (v.${verse.num})` : ref,
              version: 'Tagalog (MBBTAG)',
              text: verse.text
            });
          });
        }
      } else {
        // 4. Content Slide
        currentSlides.push({
          type: 'content',
          mainTitle: parentContext,
          title: trimmed
        });

        // Push current to stack for deeper indents
        contextStack.push({ indent: currentIndent, text: trimmed });
      }
    });

    setSlides(currentSlides);
  };

  const generatePPTX = async () => {
    setIsGenerating(true);
    setStatus('Creating PowerPoint...');

    try {
      let pptx = new pptxgen();
      pptx.layout = "LAYOUT_16x9";

      for (const slideData of slides) {
        let slide = pptx.addSlide();

        if (theme.bgImage) {
          slide.background = { data: theme.bgImage };
        } else {
          slide.background = { color: theme.bg.replace('#', '') };
        }

        // Card Overlay / Slide Surface
        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: {
            color: theme.card.replace('#', ''),
            alpha: theme.bgImage ? (theme.overlayOpacity * 100) : 0 // Fully opaque card if no image
          }
        });

        const processText = (txt) => theme.uppercase ? txt.toUpperCase() : txt;

        if (slideData.type === 'title') {
          slide.addText(processText(slideData.title), {
            x: 0.25, y: 1, w: 9.5, h: 2.5,
            fontSize: getDynamicFontSize(slideData.title, 'title', true),
            fontFace: theme.fontFace,
            bold: theme.bold, italic: theme.italic, color: theme.titleColor.replace('#', ''),
            align: "center", valign: "middle",
            shrinkText: true
          });
          slide.addText(processText(slideData.subtitle || ""), {
            x: 0.5, y: 3.5, w: 9, h: 1,
            fontSize: getDynamicFontSize(slideData.subtitle, 'subtitle', true),
            fontFace: theme.fontFace,
            bold: theme.bold, italic: theme.italic, color: theme.subtitleColor.replace('#', ''),
            align: "center", valign: "top",
            shrinkText: true
          });
        }
        else if (slideData.type === 'content') {
          const textToDisplay = slideData.title || "";
          slide.addText(processText(textToDisplay), {
            x: 0.25, y: 0.25, w: 9.5, h: 5.125,
            fontSize: getDynamicFontSize(textToDisplay, 'content', true),
            fontFace: theme.fontFace,
            bold: theme.bold,
            italic: theme.italic,
            color: theme.text.replace('#', ''),
            align: "center",
            valign: "middle",
            shrinkText: true
          });
        }
        else if (slideData.type === 'scripture') {
          slide.addText(processText(slideData.reference), {
            x: 0.25, y: 0.2, w: 9.5, h: 0.8,
            fontSize: 40 * theme.sizeMultiplier, fontFace: theme.fontFace, bold: true, color: theme.accent.replace('#', ''),
            align: "center", valign: "top", shrinkText: true
          });

          slide.addText(processText(slideData.text), {
            x: 0.15, y: 1.0, w: 9.7, h: 4.4,
            fontSize: 150 * theme.sizeMultiplier, // Extremely huge default so PowerPoint auto-shrinks it down to precisely the exact bounding box size
            fontFace: theme.fontFace,
            bold: theme.bold,
            italic: theme.italic,
            color: theme.text.replace('#', ''),
            align: "center",
            valign: "middle",
            margin: 0,
            breakLine: false,
            fit: "shrink"
          });
        }
      }

      await pptx.writeFile({ fileName: `Sermon_${new Date().getTime()}.pptx` });
      setStatus('Success! Opening download...');
    } catch (error) {
      console.error(error);
      setStatus('Error generating PowerPoint.');
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
          <h1>SERMON FLOW <span>BBC</span></h1>
        </div>
        <div style={{ textAlign: 'right' }}>
          <p>PRESENTATION CONSOLE</p>
          <div style={{ fontSize: '0.7rem', color: 'var(--primary)', fontWeight: 'bold', letterSpacing: '0.05em' }}>• {slides.length} SLIDES</div>
        </div>
      </header>

      <main className="main-content">
        <section className="presenter-sidebar">
          <div className="sidebar-tabs" style={{ display: 'flex', borderBottom: '1px solid var(--glass-border)' }}>
            <button
              className={`tab-btn ${activeTab === 'script' ? 'active' : ''}`}
              onClick={() => setActiveTab('script')}
              style={{ flex: 1, padding: '1rem', background: 'transparent', border: 'none', color: activeTab === 'script' ? 'var(--primary)' : '#888', fontWeight: 'bold', cursor: 'pointer', borderBottom: activeTab === 'script' ? '2px solid var(--primary)' : 'none' }}
            >
              SCRIPT
            </button>
            <button
              className={`tab-btn ${activeTab === 'theme' ? 'active' : ''}`}
              onClick={() => setActiveTab('theme')}
              style={{ flex: 1, padding: '1rem', background: 'transparent', border: 'none', color: activeTab === 'theme' ? 'var(--primary)' : '#888', fontWeight: 'bold', cursor: 'pointer', borderBottom: activeTab === 'theme' ? '2px solid var(--primary)' : 'none' }}
            >
              THEME
            </button>
          </div>

          <div className="sidebar-content">
            {activeTab === 'script' ? (
              <>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                  <div className="section-label"><FileText size={14} /> Sermon Script</div>
                  <textarea
                    value={sermonText}
                    onChange={(e) => setSermonText(e.target.value)}
                    placeholder="Type or paste sermon content here..."
                    spellCheck="false"
                  />
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
                  <div className="section-label" style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span><Play size={14} /> Active Cues</span>
                    <button
                      onClick={() => setIsBlackout(!isBlackout)}
                      style={{
                        fontSize: '0.6rem',
                        padding: '2px 8px',
                        borderRadius: '4px',
                        border: '1px solid var(--glass-border)',
                        background: isBlackout ? '#ef4444' : 'transparent',
                        color: '#fff',
                        cursor: 'pointer'
                      }}
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
                        </span>
                        <span className="cue-num">{String(idx + 1).padStart(2, '0')}</span>
                        <span className="cue-text">
                          {slide.type === 'scripture' ? `${slide.reference}` : (slide.title || "---")}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              </>
            ) : (
              <div className="theme-editor" style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
                <div className="section-label">Design Customization</div>

                <div className="theme-field">
                  <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Background Image</label>
                  <div style={{ display: 'flex', gap: '0.5rem' }}>
                    <input type="file" accept="image/*" onChange={handleImageUpload} style={{ fontSize: '0.7rem', flex: 1 }} />
                    {theme.bgImage && (
                      <button onClick={() => setTheme({ ...theme, bgImage: null })} style={{ padding: '0 8px', background: '#ef4444', border: 'none', borderRadius: '4px', color: '#fff' }}>X</button>
                    )}
                  </div>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem' }}>
                  <div className="theme-field">
                    <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>BG Color (Fallback)</label>
                    <input type="color" value={theme.bg} onChange={(e) => setTheme({ ...theme, bg: e.target.value })} style={{ width: '100%', height: '35px' }} />
                  </div>
                  <div className="theme-field">
                    <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Slide/Card Color</label>
                    <input type="color" value={theme.card} onChange={(e) => setTheme({ ...theme, card: e.target.value })} style={{ width: '100%', height: '35px' }} />
                  </div>
                </div>

                <div className="theme-field">
                  <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Main Font Face</label>
                  <select
                    value={theme.fontFace}
                    onChange={(e) => setTheme({ ...theme, fontFace: e.target.value })}
                    style={{ width: '100%', padding: '0.5rem', background: 'rgba(0,0,0,0.2)', border: '1px solid var(--glass-border)', color: '#fff', borderRadius: '4px' }}
                  >
                    <option value="Arial">Arial (Clean)</option>
                    <option value="Calibri">Calibri (Modern)</option>
                    <option value="Impact">Impact (Strong)</option>
                    <option value="Georgia">Georgia (Classic)</option>
                    <option value="Verdana">Verdana (Wide)</option>
                    <option value="Times New Roman">Times New Roman</option>
                  </select>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem' }}>
                  <div className="theme-field">
                    <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Title Color</label>
                    <input type="color" value={theme.titleColor} onChange={(e) => setTheme({ ...theme, titleColor: e.target.value })} style={{ width: '100%', height: '35px' }} />
                  </div>
                  <div className="theme-field">
                    <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Subtitle Color</label>
                    <input type="color" value={theme.subtitleColor} onChange={(e) => setTheme({ ...theme, subtitleColor: e.target.value })} style={{ width: '100%', height: '35px' }} />
                  </div>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem' }}>
                  <div className="theme-field">
                    <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Body Text Color</label>
                    <input type="color" value={theme.text} onChange={(e) => setTheme({ ...theme, text: e.target.value })} style={{ width: '100%', height: '35px' }} />
                  </div>
                  <div className="theme-field">
                    <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Reference Color</label>
                    <input type="color" value={theme.accent} onChange={(e) => setTheme({ ...theme, accent: e.target.value })} style={{ width: '100%', height: '35px' }} />
                  </div>
                </div>

                <div className="theme-field">
                  <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Overlay Opacity: {(theme.overlayOpacity * 100).toFixed(0)}%</label>
                  <input type="range" min="0" max="1" step="0.05" value={theme.overlayOpacity} onChange={(e) => setTheme({ ...theme, overlayOpacity: parseFloat(e.target.value) })} style={{ width: '100%' }} />
                </div>

                <div className="theme-field">
                  <label style={{ fontSize: '0.8rem', color: '#888', display: 'block', marginBottom: '0.5rem' }}>Font Size Multiplier: {theme.sizeMultiplier.toFixed(1)}x</label>
                  <input type="range" min="0.5" max="2.0" step="0.1" value={theme.sizeMultiplier} onChange={(e) => setTheme({ ...theme, sizeMultiplier: parseFloat(e.target.value) })} style={{ width: '100%' }} />
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.5rem' }}>
                  <button onClick={() => setTheme({ ...theme, uppercase: !theme.uppercase })} style={{ padding: '0.5rem', background: theme.uppercase ? 'var(--primary)' : 'rgba(255,255,255,0.05)', color: '#fff', border: '1px solid var(--glass-border)', borderRadius: '8px', cursor: 'pointer' }}>UPPERCASE</button>
                  <button onClick={() => setTheme({ ...theme, italic: !theme.italic })} style={{ padding: '0.5rem', background: theme.italic ? 'var(--primary)' : 'rgba(255,255,255,0.05)', color: '#fff', border: '1px solid var(--glass-border)', borderRadius: '8px', cursor: 'pointer' }}>ITALIC</button>
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
                style={{ background: 'rgba(255,255,255,0.05)', color: '#fff', width: 'auto', padding: '0 1rem', border: '1px solid var(--glass-border)' }}
              >
                <RefreshCw size={18} />
              </button>
            </div>
            {status && <div style={{ fontSize: '0.75rem', color: 'var(--primary)', textAlign: 'center', opacity: 0.8 }}>{status}</div>}
          </div>
        </section>

        <section className="live-projection-area">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
            <div className="section-label" style={{ margin: 0 }}><Play size={14} /> LIVE PROJECTION</div>
            <div style={{ display: 'flex', gap: '0.5rem' }}>
              <button
                className="btn"
                onClick={openProjectorWindow}
                style={{ padding: '0.4rem 1rem', fontSize: '0.7rem', background: 'rgba(255,255,255,0.05)', color: '#fff', border: '1px solid var(--glass-border)', borderRadius: '4px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '0.5rem' }}
              >
                <Download size={12} style={{ transform: 'rotate(-90deg)' }} /> POP-OUT
              </button>
              <button
                className="btn"
                onClick={toggleFullscreen}
                style={{ padding: '0.4rem 1rem', fontSize: '0.7rem', background: 'var(--primary)', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '0.5rem' }}
              >
                <Play size={12} fill="currentColor" /> FULLSCREEN
              </button>
            </div>
          </div>
          {slides[activeSlideIndex] && (
            <div className="projection-screen" ref={screenRef} style={{
              visibility: isBlackout ? 'hidden' : 'visible',
              backgroundColor: theme.bg,
              backgroundImage: theme.bgImage ? `url(${theme.bgImage})` : 'none',
              backgroundSize: 'cover',
              backgroundPosition: 'center',
              // Add a special behavior for fullscreen mode via pseudo-class is usually better, but inline works for basics
              display: 'flex',
              flexDirection: 'column'
            }}>
              {/* Semi-transparent overlay / Slide Card */}
              <div style={{
                position: 'absolute',
                inset: 0,
                backgroundColor: theme.card,
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
                padding: '2% 3%',
                fontFamily: theme.fontFace,
                boxSizing: 'border-box'
              }}>
                {slides[activeSlideIndex].type === 'title' && (
                  <div key={activeSlideIndex} className="slide-anim" style={{ textAlign: 'center', fontWeight: theme.bold ? '900' : '400', fontStyle: theme.italic ? 'italic' : 'normal', textTransform: theme.uppercase ? 'uppercase' : 'none', color: theme.titleColor }}>
                    <div style={{ fontSize: getDynamicFontSize(slides[activeSlideIndex].title, 'title'), letterSpacing: '-0.03em', lineHeight: '1.1', marginBottom: '1vw' }}>{slides[activeSlideIndex].title}</div>
                    <div style={{ fontSize: getDynamicFontSize(slides[activeSlideIndex].subtitle, 'subtitle'), color: theme.subtitleColor, fontWeight: '600' }}>{slides[activeSlideIndex].subtitle}</div>
                  </div>
                )}
                {slides[activeSlideIndex].type === 'content' && (
                  <div key={activeSlideIndex} className="slide-anim" style={{ textAlign: 'center', textTransform: theme.uppercase ? 'uppercase' : 'none', fontWeight: theme.bold ? '900' : '400', fontStyle: theme.italic ? 'italic' : 'normal', color: theme.text, width: '100%' }}>
                    {slides[activeSlideIndex].mainTitle && (
                      <div style={{ fontSize: '1.5vw', color: theme.subtitleColor, marginBottom: '2vw', fontWeight: '700' }}>
                        {slides[activeSlideIndex].mainTitle.toUpperCase()}
                      </div>
                    )}
                    <div style={{
                      fontSize: getDynamicFontSize(slides[activeSlideIndex].title, 'content'),
                      lineHeight: '1.2',
                      letterSpacing: '-0.01em'
                    }}>
                      {slides[activeSlideIndex].title}
                    </div>
                  </div>
                )}
                {slides[activeSlideIndex].type === 'scripture' && (
                  <div key={activeSlideIndex} className="slide-anim" style={{ 
                    position: 'relative', width: '100%', height: '100%', display: 'flex', flexDirection: 'column', 
                    textTransform: theme.uppercase ? 'uppercase' : 'none', fontWeight: theme.bold ? '900' : '400', 
                    fontStyle: theme.italic ? 'italic' : 'normal', color: theme.text 
                  }}>
                    {/* Fixed Reference at the Top */}
                    <div style={{ flex: '0 0 auto', textAlign: 'center', fontSize: '1.8vw', color: theme.accent, padding: '1vw 0', fontWeight: '800' }}>
                      {slides[activeSlideIndex].reference}
                    </div>
                    {/* Centered Verse Content that shrinks if needed */}
                    <div style={{ flex: '1 1 auto', display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', overflow: 'hidden' }}>
                      <div style={{
                        fontSize: getDynamicFontSize(slides[activeSlideIndex].text, 'scripture'),
                        lineHeight: '1.2',
                        maxWidth: '90%'
                      }}>
                        "{slides[activeSlideIndex].text}"
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

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
