// Personal Settings Helper - Language Switcher (v1.1 - Enhanced with Settings & Logs)
(async function () {
  'use strict';

  const ensureOverlay = async () => {
    try {
      if (window.__ensureOverlayV2 && typeof window.__ensureOverlayV2.ensureOverlayV2 === 'function') {
        return await window.__ensureOverlayV2.ensureOverlayV2();
      }
    } catch (ex) {
      console.warn('ensureOverlayV2 failed:', ex);
    }
    return window.__lvlUpOverlay || null;
  };

  const overlayApi = await ensureOverlay();
  if (!overlayApi || typeof overlayApi.createOverlay !== 'function') {
    alert('Overlay helpers are not available. Please load Helper Overlay V2.');
    return;
  }

  const gc = typeof Xrm !== 'undefined' && Xrm.Utility && typeof Xrm.Utility.getGlobalContext === 'function'
    ? Xrm.Utility.getGlobalContext()
    : null;
  if (!gc) {
    alert('Global context is not available. Please run inside Dynamics 365.');
    return;
  }

  const userSettings = gc.userSettings || {};
  const sanitizeGuid = (id) => {
    if (!id) return '';
    return id.replace(/[{}]/g, '').toLowerCase();
  };
  const userIdRaw = userSettings.userId || '';
  const userKey = sanitizeGuid(userIdRaw);
  if (!userKey) {
    alert('Could not determine current user ID.');
    return;
  }
  const userDisplayName = userSettings.userName || userSettings.userId || 'Unknown user';

  // Color scheme system
  const defaultColors = { 
    light: { 
      primary: '#0078d4', 
      primaryHover: '#106ebe', 
      border: '#ddd', 
      borderLight: '#eee', 
      text: '#323130', 
      textMuted: '#605e5c', 
      textError: '#a4262c', 
      bg: '#fff', 
      bgHover: '#f3f2f1', 
      bgError: '#fde7e9',
      textOnPrimary: '#fff', 
      success: '#107c10', 
      successHover: '#0e6b0e', 
      danger: '#d13438', 
      dangerHover: '#a80000', 
      buttonSecondary: '#605e5c'
    }, 
    dark: { 
      primary: '#60a5fa', 
      primaryHover: '#3b82f6', 
      border: '#374151', 
      borderLight: '#4b5563', 
      text: '#eef2ff', 
      textMuted: '#9ca3af', 
      textError: '#fca5a5', 
      bg: '#1f2937', 
      bgHover: '#374151', 
      bgError: '#7f1d1d',
      textOnPrimary: '#fff', 
      success: '#10b981', 
      successHover: '#059669', 
      danger: '#ef4444', 
      dangerHover: '#dc2626', 
      buttonSecondary: '#6b7280'
    } 
  };
  let customColors = null; 
  try { 
    const stored = localStorage.getItem('crmab-personalsettings-colors'); 
    if (stored) customColors = JSON.parse(stored); 
  } catch (ex) {}
  const getColors = () => { 
    const isDark = ui.host.getAttribute('data-theme') === 'dark'; 
    const mode = isDark ? 'dark' : 'light'; 
    return (customColors && customColors[mode]) || defaultColors[mode]; 
  };

  // Activity log system
  let activityLog = [];
  const addLog = (message, type) => { 
    activityLog.push({ timestamp: new Date(), message, type: type || 'info' }); 
    if (activityLog.length > 100) activityLog.shift(); 
  };

  const ui = overlayApi.createOverlay('Personal Settings', {
    rootId: 'crmab-personal-settings-root',
    width: 460,
    maxHeight: '75vh',
    footer: true,
    footerText: 'Loading...'
  });
  if (!ui) return;

  const { body } = ui;
  if (!body) return;
  body.style.paddingBottom = '20px';

  const ce = (tag) => document.createElement(tag);
  const txt = (el, t) => { el.textContent = t; return el; };
  const style = (el, s) => { el.style.cssText = s; return el; };

  const c = getColors();

  const section = style(ce('div'), 'display:flex;flex-direction:column;gap:8px;margin-bottom:16px;');
  body.appendChild(section);

  const intro = style(ce('div'), `color:${c.textMuted};font-size:13px;line-height:1.5;`);
  const introTitle = txt(ce('strong'), 'Personal Language Settings');
  const introDescription = txt(ce('span'), 'Select UI and help language. After saving, refresh the browser to apply the change.');
  intro.appendChild(introTitle);
  intro.appendChild(ce('br'));
  intro.appendChild(introDescription);
  section.appendChild(intro);

  const userInfo = style(ce('div'), `font-size:12px;color:${c.text};background:${c.bgHover};padding:8px;border-radius:4px;`);
  userInfo.textContent = `Current user: ${userDisplayName}`;
  section.appendChild(userInfo);

  const form = style(ce('div'), 'display:flex;flex-direction:column;gap:12px;');
  section.appendChild(form);

  const createField = (labelText, inputNode) => {
    const wrapper = style(ce('label'), `display:flex;flex-direction:column;gap:4px;font-size:12px;color:${c.text};`);
    const label = style(ce('span'), 'font-weight:600;');
    label.textContent = labelText;
    wrapper.appendChild(label);
    wrapper.appendChild(inputNode);
    return wrapper;
  };

  const langSelect = ce('select');
  langSelect.id = 'crmab-personal-language';
  langSelect.className = 'crmab-select';
  langSelect.style.cssText = `padding:8px;border:1px solid ${c.border};border-radius:4px;background:${c.bg};color:${c.text};`;

  const helpLangSelect = ce('select');
  helpLangSelect.id = 'crmab-personal-help-language';
  helpLangSelect.className = 'crmab-select';
  helpLangSelect.style.cssText = `padding:8px;border:1px solid ${c.border};border-radius:4px;background:${c.bg};color:${c.text};`;

  // Auto-sync Help language when UI language changes
  langSelect.addEventListener('change', () => {
    if (langSelect.value && helpLangSelect.value !== langSelect.value) {
      helpLangSelect.value = langSelect.value;
      addLog(`Help language auto-synced to ${langSelect.options[langSelect.selectedIndex]?.text || langSelect.value}`, 'info');
    }
  });

  const localeInfo = style(ce('div'), `font-size:12px;color:${c.textMuted};`);

  form.appendChild(createField('UI language', langSelect));
  form.appendChild(createField('Help language', helpLangSelect));
  form.appendChild(localeInfo);

  const buttonBar = style(ce('div'), 'display:flex;gap:8px;flex-wrap:wrap;');
  const saveBtn = txt(ce('button'), 'Save');
  saveBtn.className = 'crmab-button';
  saveBtn.style.cssText = `padding:8px 16px;background:${c.primary};color:${c.textOnPrimary};border:none;border-radius:4px;cursor:pointer;font-size:12px;`;
  saveBtn.addEventListener('mouseenter', function(){ this.style.background=c.primaryHover; });
  saveBtn.addEventListener('mouseleave', function(){ this.style.background=c.primary; });
  buttonBar.appendChild(saveBtn);
  form.appendChild(buttonBar);

  const statusDiv = style(ce('div'), `padding:8px;border-radius:4px;background:${c.bgHover};color:${c.textMuted};font-size:12px;`);
  statusDiv.id = 'crmab-personal-status';
  body.appendChild(statusDiv);

  const setStatus = (msg, isError = false) => {
    statusDiv.textContent = msg;
    statusDiv.style.background = isError ? c.bgError : c.bgHover;
    statusDiv.style.color = isError ? c.textError : c.textMuted;
    if (ui.updateFooter) {
      ui.updateFooter(isError ? `Error: ${msg}` : msg);
    }
    addLog(msg, isError ? 'error' : 'info');
    console.log('[PersonalSettings]', msg);
  };

  const apiBase = () => gc.getClientUrl() + '/api/data/v9.2';
  const baseHeaders = () => ({
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    Accept: 'application/json',
    'Content-Type': 'application/json'
  });

  const fetchJson = async (path, options = {}) => {
    const url = path.startsWith('http') ? path : apiBase() + path;
    const res = await fetch(url, {
      ...options,
      headers: { ...baseHeaders(), ...(options.headers || {}) }
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`HTTP ${res.status}: ${text || res.statusText}`);
    }
    if (res.status === 204) return null;
    return res.json();
  };

  const patchJson = async (path, data) => {
    const res = await fetch(apiBase() + path, {
      method: 'PATCH',
      headers: { ...baseHeaders(), 'If-Match': '*' },
      body: JSON.stringify(data)
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`HTTP ${res.status}: ${text || res.statusText}`);
    }
  };

  const optionPlaceholder = (target, text) => {
    const opt = ce('option');
    opt.value = '';
    opt.textContent = text;
    target.appendChild(opt);
  };

  let currentSettings = null;
  let languages = [];

  const populateSelects = () => {
    [langSelect, helpLangSelect].forEach(sel => {
      while (sel.firstChild) {
        sel.removeChild(sel.firstChild);
      }
    });
    optionPlaceholder(langSelect, 'Please choose...');
    optionPlaceholder(helpLangSelect, 'Please choose...');
    languages.forEach(lang => {
      const option = ce('option');
      option.value = String(lang.lcid);
      option.textContent = `${lang.name} (${lang.lcid})`;
      langSelect.appendChild(option.cloneNode(true));
      helpLangSelect.appendChild(option);
    });
  };

  const lcidNameMap = {
    1025: 'Arabic (Saudi Arabia)',
    1026: 'Bulgarian',
    1027: 'Catalan',
    1028: 'Chinese (Traditional)',
    1029: 'Czech',
    1030: 'Danish',
    1031: 'German (Germany)',
    1032: 'Greek',
    1033: 'English (USA)',
    1034: 'Spanish (Spain)',
    1035: 'Finnish',
    1036: 'French',
    1037: 'Hebrew',
    1038: 'Hungarian',
    1040: 'Italian',
    1041: 'Japanese',
    1042: 'Korean',
    1043: 'Dutch',
    1044: 'Norwegian (Bokmål)',
    1045: 'Polish',
    1046: 'Portuguese (Brazil)',
    1049: 'Russian',
    1050: 'Croatian',
    1051: 'Slovak',
    1053: 'Swedish',
    1054: 'Thai',
    1055: 'Turkish',
    1057: 'Indonesian',
    1060: 'Slovenian',
    1061: 'Estonian',
    1062: 'Latvian',
    1063: 'Lithuanian',
    1066: 'Vietnamese',
    1069: 'Basque',
    1081: 'Hindi',
    1094: 'Malayalam',
    1104: 'Mongolian',
    1110: 'Galician',
    2052: 'Chinese (Simplified)',
    2070: 'Portuguese (Portugal)',
    3082: 'Spanish (Mexico)'
  };

  const loadLanguages = async () => {
    setStatus('Loading available languages...');
    languages = [];
    try {
      const data = await fetchJson('/RetrieveProvisionedLanguages()');
      const values = data && data.RetrieveProvisionedLanguages && Array.isArray(data.RetrieveProvisionedLanguages) ? data.RetrieveProvisionedLanguages : [];
      languages = values.map(code => ({
        id: code,
        name: lcidNameMap[code] || `Language ${code}`,
        lcid: code
      }));
    } catch (ex) {
      console.warn('RetrieveProvisionedLanguages not available', ex);
      addLog('RetrieveProvisionedLanguages failed: ' + ex.message, 'error');
    }
    if (languages.length === 0) {
      const fallbackCode = userSettings && userSettings.languageId ? userSettings.languageId : gc.userSettings.uilanguageId;
      const fallbackList = fallbackCode ? [fallbackCode, 1033, 1031] : [1033, 1031];
      const unique = Array.from(new Set(fallbackList));
      languages = unique.map(code => ({
        id: code,
        name: lcidNameMap[code] || `Language ${code}`,
        lcid: code
      }));
      setStatus('Provisioned languages could not be fetched. Fallback list loaded.', true);
    } else {
      setStatus(`Languages loaded (${languages.length}).`);
    }
    populateSelects();
  };

  const loadUserSettings = async () => {
    setStatus('Loading user settings...');
    try {
      const response = await fetchJson(`/usersettingscollection(${userKey})?$select=uilanguageid,helplanguageid,localeid`);
      if (!response) {
        throw new Error('User settings could not be found.');
      }
      currentSettings = {
        usersettingsid: userKey,
        uilanguageid: response.uilanguageid || null,
        helplanguageid: response.helplanguageid || null,
        localeid: response.localeid || null
      };
      langSelect.value = currentSettings.uilanguageid ? String(currentSettings.uilanguageid) : '';
      helpLangSelect.value = currentSettings.helplanguageid ? String(currentSettings.helplanguageid) : '';
      const locale = currentSettings.localeid ? `Format LCID: ${currentSettings.localeid}` : 'No locale ID available';
      localeInfo.textContent = `Current locale: ${locale}`;
      setStatus('User settings fetched.');
    } catch (ex) {
      setStatus(`Load failed: ${ex.message}`, true);
      throw ex;
    }
  };

  const ensureLanguageSelected = (select) => {
    const { value } = select;
    if (!value) return null;
    const parsed = Number(value);
    if (Number.isNaN(parsed)) return null;
    return parsed;
  };

  // Force hard refresh - clears caches and reloads the page
  // This mimics Ctrl+Shift+R behavior without modifying the URL
  const forceHardRefresh = () => {
    // Clear all browser storage caches first
    try {
      sessionStorage.clear();
    } catch (e) {}
    try {
      // Clear any D365-specific cache keys
      const keysToRemove = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && (key.includes('cache') || key.includes('lang') || key.includes('uci'))) {
          keysToRemove.push(key);
        }
      }
      keysToRemove.forEach(k => localStorage.removeItem(k));
    } catch (e) {}
    
    // Simple reload - the browser will reload the current page
    // The true parameter is deprecated but still works in most browsers to bypass cache
    window.location.reload();
  };

  const showRefreshDialog = () => {
    const refreshOverlay = overlayApi.createOverlay('Language Changed', { 
      width: 450, 
      rootId: 'personalsettings-refresh-overlay' 
    });
    if (!refreshOverlay) return;
    
    const colors = getColors();
    const content = ce('div');
    content.style.cssText = 'display:flex;flex-direction:column;gap:16px;';
    
    const message = ce('div');
    message.style.cssText = `font-size:14px;color:${colors.text};line-height:1.6;`;
    
    const titleStrong = style(ce('strong'), 'display:block;margin-bottom:8px;font-size:16px;');
    titleStrong.textContent = '✅ Language settings saved!';
    message.appendChild(titleStrong);
    
    const p1 = style(ce('p'), 'margin:0 0 8px 0;');
    p1.textContent = 'Click the button below to refresh the page. After the refresh, press Ctrl + Shift + R to complete the language change.';
    message.appendChild(p1);
    
    content.appendChild(message);
    
    const btnContainer = ce('div');
    btnContainer.style.cssText = 'display:flex;flex-direction:column;gap:10px;';
    
    // Hard Refresh button
    const hardRefreshBtn = ce('button');
    const hardRefreshIcon = document.createTextNode('🔄 ');
    const hardRefreshStrong = ce('strong');
    hardRefreshStrong.textContent = 'Refresh Page';
    hardRefreshBtn.appendChild(hardRefreshIcon);
    hardRefreshBtn.appendChild(hardRefreshStrong);
    hardRefreshBtn.style.cssText = `padding:12px 16px;background:${colors.primary};color:${colors.textOnPrimary};border:none;border-radius:6px;cursor:pointer;font-size:14px;text-align:center;`;
    hardRefreshBtn.addEventListener('mouseenter', () => { hardRefreshBtn.style.background = colors.primaryHover; });
    hardRefreshBtn.addEventListener('mouseleave', () => { hardRefreshBtn.style.background = colors.primary; });
    hardRefreshBtn.onclick = () => {
      addLog('Performing refresh...', 'info');
      forceHardRefresh();
    };
    
    // Refresh Later button
    const laterBtn = ce('button');
    laterBtn.textContent = 'Refresh Later';
    laterBtn.style.cssText = `padding:10px 16px;background:transparent;color:${colors.textMuted};border:1px solid ${colors.border};border-radius:6px;cursor:pointer;font-size:13px;`;
    laterBtn.onclick = () => {
      refreshOverlay.closeOverlay();
    };
    
    btnContainer.appendChild(hardRefreshBtn);
    btnContainer.appendChild(laterBtn);
    content.appendChild(btnContainer);
    
    // Instructions box
    const infoBox = ce('div');
    infoBox.style.cssText = `margin-top:8px;padding:12px;background:${colors.bgHover};border-radius:4px;font-size:12px;color:${colors.text};line-height:1.6;`;
    
    const stepsTitle = ce('strong');
    stepsTitle.textContent = '📋 Steps to complete:';
    stepsTitle.style.display = 'block';
    stepsTitle.style.marginBottom = '8px';
    infoBox.appendChild(stepsTitle);
    
    const stepsList = ce('ol');
    stepsList.style.cssText = 'margin:0;padding-left:20px;';
    
    const step1 = ce('li');
    step1.textContent = 'Click "Refresh Page" above';
    stepsList.appendChild(step1);
    
    const step2 = ce('li');
    step2.appendChild(document.createTextNode('Press '));
    const kbd = style(ce('kbd'), `background:${colors.bg};padding:2px 6px;border-radius:3px;border:1px solid ${colors.border};font-family:monospace;`);
    kbd.textContent = 'Ctrl + Shift + R';
    step2.appendChild(kbd);
    step2.appendChild(document.createTextNode(' to force reload'));
    stepsList.appendChild(step2);
    
    const step3 = ce('li');
    step3.style.color = colors.textMuted;
    step3.textContent = 'Reload Helper Overlay manually via LevelUp';
    stepsList.appendChild(step3);
    
    infoBox.appendChild(stepsList);
    
    // Warning about overlay
    const warning = ce('div');
    warning.style.cssText = `margin-top:10px;padding:8px;background:${colors.bgError};border-radius:4px;font-size:11px;color:${colors.textError};`;
    warning.appendChild(document.createTextNode('⚠️ '));
    const warningStrong = ce('strong');
    warningStrong.textContent = 'Note:';
    warning.appendChild(warningStrong);
    warning.appendChild(document.createTextNode(' After refresh, Helper Overlay must be reloaded manually via LevelUp.'));
    infoBox.appendChild(warning);
    
    content.appendChild(infoBox);
    
    refreshOverlay.body.appendChild(content);
  };

  const updateSettings = async () => {
    if (!currentSettings || !currentSettings.usersettingsid) {
      setStatus('User settings not loaded.', true);
      return;
    }
    const uiLang = ensureLanguageSelected(langSelect);
    const helpLang = ensureLanguageSelected(helpLangSelect);
    if (!uiLang) {
      setStatus('Please select a UI language.', true);
      langSelect.focus();
      return;
    }
    setStatus('Saving settings...');
    saveBtn.disabled = true;
    try {
      await patchJson(`/usersettingscollection(${currentSettings.usersettingsid})`, {
        uilanguageid: uiLang,
        helplanguageid: helpLang || uiLang
      });
      setStatus('Language settings saved successfully!');
      addLog('Language settings saved: UI=' + uiLang + ', Help=' + (helpLang || uiLang), 'info');
      
      // Show refresh dialog instead of simple alert
      showRefreshDialog();
    } catch (ex) {
      console.error('Update failed', ex);
      setStatus(`Save failed: ${ex.message}`, true);
    } finally {
      saveBtn.disabled = false;
    }
  };

  saveBtn.addEventListener('click', updateSettings);

  // Settings dialog
  function showSettings() { 
    const settingsOverlay = overlayApi.createOverlay('Settings', { width: 500, rootId: 'personalsettings-settings-overlay' }); 
    if (!settingsOverlay) return; 
    let currentMode = ui.host.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'; 
    const colors = getColors(); 
    const settingsContent = ce('div'); 
    const modeTabsDiv = style(ce('div'), 'display:flex;gap:8px;margin-bottom:16px;'); 
    const lightTab = txt(ce('button'), '☀️ Light Mode'); 
    const darkTab = txt(ce('button'), '🌙 Dark Mode'); 
    const tabStyle = (isActive) => `flex:1;padding:8px 16px;border:none;background:${isActive ? colors.primary : colors.bgHover};color:${isActive ? colors.textOnPrimary : colors.text};border-radius:4px;cursor:pointer;font-weight:600;`; 
    lightTab.style.cssText = tabStyle(currentMode === 'light'); 
    darkTab.style.cssText = tabStyle(currentMode === 'dark'); 
    const colorInputsContainer = ce('div'); 
    const colorInputs = {}; 
    const hexInputs = {}; 
    const renderColorInputs = (mode) => { 
      while (colorInputsContainer.firstChild) colorInputsContainer.removeChild(colorInputsContainer.firstChild); 
      colorInputs[mode] = {}; 
      hexInputs[mode] = {}; 
      const modeColors = (customColors && customColors[mode]) || defaultColors[mode]; 
      Object.keys(defaultColors[mode]).forEach(key => { 
        const row = style(ce('div'), 'display:flex;align-items:center;gap:8px;margin-bottom:10px;'); 
        const label = txt(style(ce('label'), 'flex:1;font-size:13px;color:' + colors.text + ';'), key); 
        const colorPicker = ce('input'); 
        colorPicker.type = 'color'; 
        colorPicker.value = modeColors[key]; 
        colorPicker.style.cssText = 'width:60px;height:32px;border:1px solid ' + colors.border + ';border-radius:4px;cursor:pointer;'; 
        colorInputs[mode][key] = colorPicker; 
        const hexInput = ce('input'); 
        hexInput.type = 'text'; 
        hexInput.value = modeColors[key]; 
        hexInput.maxLength = 7; 
        hexInput.placeholder = '#000000'; 
        hexInput.style.cssText = 'width:80px;padding:4px 8px;border:1px solid ' + colors.border + ';border-radius:4px;font-family:monospace;font-size:12px;text-transform:uppercase;background:' + colors.bg + ';color:' + colors.text + ';'; 
        hexInputs[mode][key] = hexInput; 
        const preview = style(ce('div'), `width:40px;height:32px;background:${modeColors[key]};border:1px solid ${colors.border};border-radius:4px;`); 
        const syncColor = (newColor) => { 
          if (/^#[0-9A-Fa-f]{6}$/.test(newColor)) { 
            colorPicker.value = newColor; 
            hexInput.value = newColor.toUpperCase(); 
            preview.style.background = newColor; 
          } 
        }; 
        colorPicker.addEventListener('input', () => { syncColor(colorPicker.value); }); 
        hexInput.addEventListener('input', () => { 
          const val = hexInput.value.trim(); 
          if (val.charAt(0) !== '#' && val.length > 0) { hexInput.value = '#' + val; } 
          syncColor(hexInput.value); 
        }); 
        hexInput.addEventListener('blur', () => { if (!/^#[0-9A-Fa-f]{6}$/.test(hexInput.value)) { hexInput.value = colorPicker.value; } }); 
        row.appendChild(label); 
        row.appendChild(colorPicker); 
        row.appendChild(hexInput); 
        row.appendChild(preview); 
        colorInputsContainer.appendChild(row); 
      }); 
    }; 
    renderColorInputs(currentMode); 
    lightTab.onclick = () => { 
      currentMode = 'light'; 
      lightTab.style.cssText = tabStyle(true); 
      darkTab.style.cssText = tabStyle(false); 
      renderColorInputs('light'); 
    }; 
    darkTab.onclick = () => { 
      currentMode = 'dark'; 
      lightTab.style.cssText = tabStyle(false); 
      darkTab.style.cssText = tabStyle(true); 
      renderColorInputs('dark'); 
    }; 
    modeTabsDiv.appendChild(lightTab); 
    modeTabsDiv.appendChild(darkTab); 
    settingsContent.appendChild(modeTabsDiv); 
    settingsContent.appendChild(colorInputsContainer); 
    const btnRow = style(ce('div'), 'display:flex;gap:8px;margin-top:16px;'); 
    const saveSettingsBtn = txt(ce('button'), '💾 Save'); 
    saveSettingsBtn.style.cssText = `flex:1;padding:8px 16px;background:${colors.primary};color:${colors.textOnPrimary};border:none;border-radius:4px;cursor:pointer;font-weight:600;`; 
    const resetBtn = txt(ce('button'), '🔄 Reset to Defaults'); 
    resetBtn.style.cssText = `flex:1;padding:8px 16px;background:${colors.bgHover};color:${colors.text};border:1px solid ${colors.border};border-radius:4px;cursor:pointer;font-weight:600;`; 
    const cancelBtn = txt(ce('button'), '✖ Cancel'); 
    cancelBtn.style.cssText = `flex:1;padding:8px 16px;background:${colors.bgHover};color:${colors.text};border:1px solid ${colors.border};border-radius:4px;cursor:pointer;font-weight:600;`; 
    saveSettingsBtn.onclick = () => { 
      const newColors = { light: {}, dark: {} }; 
      ['light', 'dark'].forEach(mode => { 
        if (colorInputs[mode]) { 
          Object.keys(colorInputs[mode]).forEach(key => { 
            newColors[mode][key] = colorInputs[mode][key].value; 
          }); 
        } else { 
          newColors[mode] = (customColors && customColors[mode]) || defaultColors[mode]; 
        } 
      }); 
      try { 
        localStorage.setItem('crmab-personalsettings-colors', JSON.stringify(newColors)); 
        customColors = newColors; 
        addLog('Settings saved', 'info'); 
        settingsOverlay.closeOverlay(); 
        ui.closeOverlay(); 
        requestAnimationFrame(() => { location.reload(); }); 
      } catch (ex) { 
        alert('Error saving settings: ' + ex.message); 
      } 
    }; 
    resetBtn.onclick = () => { 
      if (!confirm('Reset all color settings to defaults?')) return; 
      try { 
        localStorage.removeItem('crmab-personalsettings-colors'); 
        customColors = null; 
        addLog('Settings reset to defaults', 'info'); 
        settingsOverlay.closeOverlay(); 
        ui.closeOverlay(); 
        requestAnimationFrame(() => { location.reload(); }); 
      } catch (ex) { 
        alert('Error resetting settings: ' + ex.message); 
      } 
    }; 
    cancelBtn.onclick = () => settingsOverlay.closeOverlay(); 
    btnRow.appendChild(saveSettingsBtn); 
    btnRow.appendChild(resetBtn); 
    btnRow.appendChild(cancelBtn); 
    settingsContent.appendChild(btnRow); 
    settingsOverlay.body.appendChild(settingsContent); 
  }

  // Logs dialog
  function showLogViewer() { 
    const cl = getColors(); 
    const logsOverlay = overlayApi.createOverlay('Activity Logs', { width: 600, rootId: 'personalsettings-logs-overlay' }); 
    if (!logsOverlay) return; 
    const logsContent = ce('div'); 
    if (activityLog.length === 0) { 
      logsContent.appendChild(txt(style(ce('div'), `padding:20px;text-align:center;color:${cl.textMuted};`), 'No activity logged yet.')); 
    } else { 
      activityLog.slice().reverse().forEach(log => { 
        const logRow = style(ce('div'), `padding:8px 12px;border-bottom:1px solid ${cl.borderLight};font-size:12px;color:${log.type === 'error' ? cl.textError : cl.textMuted};`); 
        const time = style(ce('span'), 'font-weight:600;margin-right:8px;'); 
        time.textContent = log.timestamp.toLocaleTimeString(); 
        const msg = ce('span'); 
        msg.textContent = log.message; 
        logRow.appendChild(time); 
        logRow.appendChild(msg); 
        logsContent.appendChild(logRow); 
      }); 
    } 
    const clearLogsBtn = txt(ce('button'), '🗑 Clear Logs'); 
    clearLogsBtn.style.cssText = `width:100%;margin-top:12px;padding:8px;background:${cl.danger};color:${cl.textOnPrimary};border:none;border-radius:4px;cursor:pointer;font-weight:600;`; 
    clearLogsBtn.onclick = () => { 
      if (!confirm('Clear all logs?')) return; 
      activityLog = []; 
      logsOverlay.closeOverlay(); 
      showLogViewer(); 
    }; 
    logsContent.appendChild(clearLogsBtn); 
    logsOverlay.body.appendChild(logsContent); 
  }

  // Footer buttons
  if (ui.footer) { 
    const settingsBtn = txt(ce('button'), String.fromCharCode(9881) + ' Settings'); 
    settingsBtn.style.cssText = `background:transparent;border:1px solid ${c.border};color:${c.text};padding:6px 12px;border-radius:4px;cursor:pointer;font-weight:600;font-size:12px;`; 
    settingsBtn.onclick = showSettings; 
    const logBtn = txt(ce('button'), String.fromCharCode(128276) + ' Logs'); 
    logBtn.style.cssText = `background:transparent;border:1px solid ${c.border};color:${c.text};padding:6px 12px;border-radius:4px;cursor:pointer;font-weight:600;font-size:12px;margin-left:8px;`; 
    logBtn.onclick = showLogViewer; 
    ui.footer.insertBefore(settingsBtn, ui.footer.firstChild); 
    ui.footer.insertBefore(logBtn, ui.footer.children[1]); 
  }

  try {
    await loadLanguages();
    await loadUserSettings();
  } catch (ex) {
    console.warn('Initialization incomplete', ex);
  }
})();
