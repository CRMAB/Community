(function () { 
  'use strict';
  
  // Enhanced overlay + drag helpers with toggle, ESC, cleanup, footer support Version 2.0.2
  
  // Theme Detection - Async
  async function getLevelUpTheme() { 
    try { 
      if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) { 
        const result = await chrome.storage.local.get(['levelup-theme-mode']); 
        if (result['levelup-theme-mode']) { 
          return result['levelup-theme-mode'] === 'dark' ? 'dark' : 'light'; 
        } 
      } 
      if (typeof localStorage !== 'undefined') { 
        const saved = localStorage.getItem('levelup-theme-mode'); 
        if (saved && (saved === 'light' || saved === 'dark')) { 
          return saved; 
        } 
      } 
      if (typeof window !== 'undefined' && window.matchMedia) { 
        const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches; 
        return prefersDark ? 'dark' : 'light'; 
      } 
    } catch (error) { 
      console.warn('Fehler beim Lesen des LevelUp Themes:', error); 
    } 
    return 'light'; 
  } 
  function getLevelUpThemeSync() { 
    try { 
      // Versuche localStorage (synchron)
      if (typeof localStorage !== 'undefined') { 
        const saved = localStorage.getItem('levelup-theme-mode'); 
        if (saved === 'light' || saved === 'dark') { 
          return saved; 
        } 
      } 
    } catch (error) { 
      console.warn('Fehler beim Lesen des LevelUp Themes:', error); 
    } 
    return 'light'; 
  } 
  function createOverlay(title, options = {}) { 
    let themeMode = options.theme; 
    if (!themeMode || themeMode === 'auto') { 
      if (options.rootId) { 
        try { 
          const storageKey = 'crmab-overlay-theme-' + options.rootId; 
          const saved = localStorage.getItem(storageKey); 
          if (saved === 'light' || saved === 'dark') { 
            themeMode = saved; 
          } else { 
            themeMode = getLevelUpThemeSync(); 
          } 
        } catch (e) { 
          themeMode = getLevelUpThemeSync(); 
        } 
      } else { themeMode = getLevelUpThemeSync(); } 
    } 
    const opts = { rootId: options.rootId || null, styleId: options.styleId || null, position: options.position || 'top-right', top: options.top, right: options.right, left: options.left, bottom: options.bottom, width: options.width || 560, maxHeight: options.maxHeight || '70vh', theme: themeMode, toggle: options.toggle !== false && options.rootId !== null, escKey: options.escKey !== false, draggable: options.draggable !== false, footer: options.footer !== false, footerText: options.footerText || '', onClose: options.onClose || null, onToggle: options.onToggle || null, customStyles: options.customStyles || null }; 
    if (opts.toggle && opts.rootId) { 
      const existing = document.getElementById(opts.rootId); 
      if (existing) { 
        const handlers = existing._handlers || {}; 
        if (handlers.esc) { document.removeEventListener('keydown', handlers.esc, { capture: true }); } 
        if (handlers.mm) { window.removeEventListener('mousemove', handlers.mm); } 
        if (handlers.mu) { window.removeEventListener('mouseup', handlers.mu); } 
        if (opts.styleId) { const styleEl = document.getElementById(opts.styleId); if (styleEl) styleEl.remove(); } 
        existing.remove(); 
        if (opts.onToggle) opts.onToggle(false); 
        if (opts.onClose) opts.onClose(); 
        return null; 
      } 
    } 
    const themes = { light: { background: '#fff', color: '#323130', border: '#e1e1e1', headerBg: '#f3f2f1', headerColor: '#323130', closeColor: '#605e5c', footerBg: '#f3f2f1', footerBorder: '#ddd' }, dark: { background: '#1f2937', color: '#eef2ff', border: '#374151', headerBg: '#0b1220', headerColor: '#e5e7eb', closeColor: '#cbd5e1', footerBg: '#0b1220', footerBorder: '#374151' } }; 
    const theme = themes[opts.theme] || themes.light; 
    let top = opts.top; let right = opts.right; let left = opts.left; let bottom = opts.bottom; 
    if (!top && !right && !left && !bottom) { switch (opts.position) { case 'top-left': top = 80; left = 40; break; case 'bottom-right': bottom = 40; right = 40; break; case 'bottom-left': bottom = 40; left = 40; break; case 'top-right': default: top = 80; right = 40; break; } } 
    const host = document.createElement('div'); if (opts.rootId) host.id = opts.rootId; host.style.position = 'fixed'; if (top !== undefined) host.style.top = top + 'px'; if (right !== undefined) host.style.right = right + 'px'; if (left !== undefined) host.style.left = left + 'px'; if (bottom !== undefined) host.style.bottom = bottom + 'px'; host.style.zIndex = '999999'; host.style.width = opts.width + 'px'; host.style.maxHeight = opts.maxHeight; host.style.fontFamily = 'Segoe UI, Tahoma, Arial'; host.style.boxShadow = '0 8px 24px rgba(0,0,0,0.2)'; host.style.borderRadius = '8px'; host.style.overflow = 'hidden'; host.style.display = 'flex'; host.style.flexDirection = 'column'; host.style.background = theme.background; host.style.color = theme.color; host.style.border = '1px solid ' + theme.border; 
    const header = document.createElement('div'); header.style.cursor = opts.draggable ? 'move' : 'default'; header.style.padding = '10px 14px'; header.style.background = theme.headerBg; header.style.color = theme.headerColor; header.style.display = 'flex'; header.style.alignItems = 'center'; header.style.justifyContent = 'space-between'; header.style.userSelect = 'none'; header.style.borderBottom = '1px solid ' + theme.border; 
    const titleEl = document.createElement('div'); titleEl.textContent = title; titleEl.style.fontWeight = '600'; titleEl.style.fontSize = '14px'; 
    const close = document.createElement('button'); close.textContent = '×'; close.title = 'Close'; close.style.fontSize = '18px'; close.style.lineHeight = '18px'; close.style.border = 'none'; close.style.background = 'transparent'; close.style.color = theme.closeColor; close.style.cursor = 'pointer'; close.style.padding = '0'; close.style.width = '24px'; close.style.height = '24px'; close.style.display = 'flex'; close.style.alignItems = 'center'; close.style.justifyContent = 'center'; 
    const body = document.createElement('div'); body.style.padding = '10px 14px'; body.style.overflow = 'auto'; body.style.flex = '1'; 
    let footer = null; let footerTextEl = null; let themeSwitchBtn = null; 
    if (opts.footer) { footer = document.createElement('div'); footer.style.display = 'flex'; footer.style.justifyContent = 'space-between'; footer.style.alignItems = 'center'; footer.style.padding = '8px 12px'; footer.style.borderTop = '1px solid ' + theme.footerBorder; footer.style.background = theme.footerBg; footer.style.fontSize = '12px'; footer.style.color = theme.color; footerTextEl = document.createElement('div'); footerTextEl.style.flex = '1'; if (opts.footerText) { footerTextEl.textContent = opts.footerText; } themeSwitchBtn = document.createElement('button'); themeSwitchBtn.title = themeMode === 'dark' ? 'Light mode' : 'Dark mode'; themeSwitchBtn.textContent = themeMode === 'dark' ? '☀️' : '🌙'; themeSwitchBtn.style.cssText = 'background:transparent;border:1px solid ' + theme.border + ';color:' + theme.color + ';padding:4px 8px;border-radius:4px;cursor:pointer;font-size:14px;margin-left:8px;'; themeSwitchBtn.addEventListener('mouseenter', function() { this.style.background = themeMode === 'dark' ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.05)'; }); themeSwitchBtn.addEventListener('mouseleave', function() { this.style.background = 'transparent'; }); themeSwitchBtn.addEventListener('click', function() { const newMode = themeMode === 'dark' ? 'light' : 'dark'; themeMode = newMode; updateTheme(newMode); this.textContent = newMode === 'dark' ? '☀️' : '🌙'; this.title = newMode === 'dark' ? 'Light mode' : 'Dark mode'; if (opts.rootId) { try { const storageKey = 'crmab-overlay-theme-' + opts.rootId; localStorage.setItem(storageKey, newMode); } catch (e) {} } }); footer.appendChild(footerTextEl); footer.appendChild(themeSwitchBtn); } 
    header.appendChild(titleEl); header.appendChild(close); host.appendChild(header); host.appendChild(body); if (footer) host.appendChild(footer); document.body.appendChild(host); 
    const handlers = {}; 
    const updateTheme = (newThemeMode) => { const newTheme = themes[newThemeMode] || themes.light; host.style.background = newTheme.background; host.style.color = newTheme.color; host.style.border = '1px solid ' + newTheme.border; header.style.background = newTheme.headerBg; header.style.color = newTheme.headerColor; header.style.borderBottom = '1px solid ' + newTheme.border; close.style.color = newTheme.closeColor; if (footer) { footer.style.borderTop = '1px solid ' + newTheme.footerBorder; footer.style.background = newTheme.footerBg; footer.style.color = newTheme.color; } if (themeSwitchBtn) { themeSwitchBtn.style.border = '1px solid ' + newTheme.border; themeSwitchBtn.style.color = newTheme.color; } host.setAttribute('data-theme', newThemeMode); }; 
    const closeOverlay = () => { if (handlers.esc) { document.removeEventListener('keydown', handlers.esc, { capture: true }); } if (handlers.mm) { window.removeEventListener('mousemove', handlers.mm); } if (handlers.mu) { window.removeEventListener('mouseup', handlers.mu); } if (opts.styleId) { const styleEl = document.getElementById(opts.styleId); if (styleEl) styleEl.remove(); } host.remove(); if (opts.onClose) opts.onClose(); if (opts.onToggle) opts.onToggle(false); }; 
    close.onclick = closeOverlay; 
    if (opts.escKey) { handlers.esc = (ev) => { if (ev.key === 'Escape' && document.body.contains(host)) { closeOverlay(); } }; document.addEventListener('keydown', handlers.esc, { capture: true }); } 
    if (opts.draggable) { const dragHandlers = makeDraggable(host, header); handlers.mm = dragHandlers.mousemove; handlers.mu = dragHandlers.mouseup; } 
    host._handlers = handlers; 
    if (opts.customStyles && opts.styleId) { const styleEl = document.createElement('style'); styleEl.id = opts.styleId; styleEl.textContent = opts.customStyles; document.head.appendChild(styleEl); } 
    if (opts.onToggle) opts.onToggle(true); 
    return { host, body, header, footer, close, handlers, closeOverlay, updateFooter: (text) => { if (footerTextEl) footerTextEl.textContent = text; }, updateTheme: (newMode) => { if (newMode === 'light' || newMode === 'dark') { themeMode = newMode; updateTheme(newMode); if (themeSwitchBtn) { themeSwitchBtn.textContent = newMode === 'dark' ? '☀️' : '🌙'; themeSwitchBtn.title = newMode === 'dark' ? 'Light mode' : 'Dark mode'; const currentTheme = themes[newMode] || themes.light; themeSwitchBtn.style.border = '1px solid ' + currentTheme.border; themeSwitchBtn.style.color = currentTheme.color; } } } }; 
  } 
  function makeDraggable(container, handle) { let isDown = false; let offset = { x: 0, y: 0 }; const onMouseDown = (e) => { if (e.button !== 0) return; isDown = true; const rect = container.getBoundingClientRect(); offset.x = rect.left - e.clientX; offset.y = rect.top - e.clientY; document.body.style.userSelect = 'none'; }; const onMouseMove = (e) => { if (!isDown) return; const maxL = window.innerWidth - container.offsetWidth - 8; const maxT = window.innerHeight - container.offsetHeight - 8; const newL = Math.max(8, Math.min(maxL, e.clientX + offset.x)); const newT = Math.max(8, Math.min(maxT, e.clientY + offset.y)); container.style.left = newL + 'px'; container.style.top = newT + 'px'; container.style.right = 'auto'; container.style.bottom = 'auto'; }; const onMouseUp = () => { isDown = false; document.body.style.userSelect = ''; }; handle.addEventListener('mousedown', onMouseDown); window.addEventListener('mousemove', onMouseMove); window.addEventListener('mouseup', onMouseUp); return { mousemove: onMouseMove, mouseup: onMouseUp }; } 
  const ce = (tag) => document.createElement(tag); 
  const txt = (el, t) => { el.textContent = t; return el; }; 
  const style = (el, s) => { el.style.cssText = s; return el; }; 
  const clearEl = (el) => { while (el.firstChild) el.removeChild(el.firstChild); }; 
  // Array chunking utility
  const chunk = (arr, size) => { 
    const result = []; 
    for (let i = 0; i < arr.length; i += size) {
      result.push(arr.slice(i, i + size)); 
    }
    return result; 
  }; 
  
  // GUID sanitization
  const sanitizeGuid = (value = '') => { 
    const openBracket = String.fromCharCode(91);
    const closeBracket = String.fromCharCode(93);
    const openBrace = String.fromCharCode(123);
    const closeBrace = String.fromCharCode(125);
    
    const patternParts = [];
    patternParts.push(openBracket);
    patternParts.push('0-9a-fA-F');
    patternParts.push(closeBracket);
    patternParts.push(openBrace + '8' + closeBrace + '-');
    patternParts.push(openBracket);
    patternParts.push('0-9a-fA-F');
    patternParts.push(closeBracket);
    patternParts.push(openBrace + '4' + closeBrace + '-');
    patternParts.push(openBracket);
    patternParts.push('0-9a-fA-F');
    patternParts.push(closeBracket);
    patternParts.push(openBrace + '4' + closeBrace + '-');
    patternParts.push(openBracket);
    patternParts.push('0-9a-fA-F');
    patternParts.push(closeBracket);
    patternParts.push(openBrace + '4' + closeBrace + '-');
    patternParts.push(openBracket);
    patternParts.push('0-9a-fA-F');
    patternParts.push(closeBracket);
    patternParts.push(openBrace + '12' + closeBrace);
    
    const pattern = patternParts.join('');
    const guidRegex = new RegExp(pattern);
    const match = (value || '').match(guidRegex);
    const target = match ? match[0] : value;
    
    const braceChars = [];
    braceChars.push(openBracket);
    braceChars.push(openBrace);
    braceChars.push(closeBrace);
    braceChars.push(closeBracket);
    const bracePattern = braceChars.join('');
    const braceRegex = new RegExp(bracePattern, 'g');
    
    return target.replace(braceRegex, '').toLowerCase(); 
  }; 
  
  const guidPathValue = (value = '') => sanitizeGuid(value); 
  const guidLiteral = (value = '', quote = "'") => quote + sanitizeGuid(value) + quote; 
  const escapeODataValue = (value = '') => { 
    const input = value == null ? '' : String(value); 
    return input.split("'").join("''"); 
  }; 
  
  // File download helper
  const downloadFile = (content, filename, mimeType) => { 
    const contentArray = [];
    contentArray.push(content);
    const blob = new Blob(contentArray, { type: mimeType }); 
    const url = URL.createObjectURL(blob); 
    const anchor = ce('a'); 
    anchor.href = url; 
    anchor.download = filename; 
    anchor.click(); 
    URL.revokeObjectURL(url); 
  }; 
  window.__lvlUpOverlay = window.__lvlUpOverlay || {}; 
  window.__lvlUpOverlay.createOverlay = createOverlay; 
  window.__lvlUpOverlay.makeDraggable = makeDraggable; 
  window.__lvlUpOverlay.getLevelUpTheme = getLevelUpTheme; 
  window.__lvlUpOverlay.getLevelUpThemeSync = getLevelUpThemeSync; 
  window.__lvlUpOverlay.ce = ce; 
  window.__lvlUpOverlay.txt = txt; 
  window.__lvlUpOverlay.style = style; 
  window.__lvlUpOverlay.clearEl = clearEl; 
  window.__lvlUpOverlay.chunk = chunk; 
  window.__lvlUpOverlay.sanitizeGuid = sanitizeGuid; 
  window.__lvlUpOverlay.guidPathValue = guidPathValue; 
  window.__lvlUpOverlay.guidLiteral = guidLiteral;
  window.__lvlUpOverlay.escapeODataValue = escapeODataValue;
  window.__lvlUpOverlay.downloadFile = downloadFile;
  window.__lvlUpOverlay.version = '2.0.2';
})();
