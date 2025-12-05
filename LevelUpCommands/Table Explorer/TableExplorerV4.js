// Table Explorer V4 - Entity metadata explorer with enhanced JS display and Environment Variables (v1.0)
(async function () {
  const { createOverlay } = window.__lvlUpOverlay || (function(){ alert('Overlay helpers not loaded'); return {}; })();
  if (!createOverlay) return;
  const ui = createOverlay('Table Explorer', { rootId: 'crmab-tableexplorer-root', width: 1000, maxHeight: '88vh', footer: true, footerText: 'Select entity...' });
  if (!ui) return;
  if (ui.body) { ui.body.style.maxHeight = '80vh'; ui.body.style.overflowY = 'auto'; ui.body.style.paddingBottom = '20px'; }
  try {
    const gc = Xrm.Utility.getGlobalContext();
    const ce = (tag) => document.createElement(tag);
    const txt = (el, t) => { el.textContent = t; return el; };
    const style = (el, s) => { el.style.cssText = s; return el; };
    const clearEl = (el) => { while (el.firstChild) el.removeChild(el.firstChild); };
    const apiBase = () => gc.getClientUrl() + '/api/data/v9.2';
    const headers = () => ({ 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', 'Accept': 'application/json', 'Content-Type': 'application/json' });
    const fetchJson = async (path) => { const res = await fetch(apiBase() + path, { headers: headers() }); if (!res.ok) throw new Error('HTTP ' + res.status); return res.json(); };
    const patchJson = async (path, data) => { const res = await fetch(apiBase() + path, { method: 'PATCH', headers: { ...headers(), 'If-Match': '*' }, body: JSON.stringify(data) }); if (!res.ok) throw new Error('HTTP ' + res.status); };
    const postJson = async (path, data) => { const res = await fetch(apiBase() + path, { method: 'POST', headers: headers(), body: JSON.stringify(data) }); if (!res.ok) throw new Error('HTTP ' + res.status); return res.json(); };
    const deleteJson = async (path) => { const res = await fetch(apiBase() + path, { method: 'DELETE', headers: headers() }); if (!res.ok) throw new Error('HTTP ' + res.status); };
    const dispName = (obj) => { try { return (obj && obj.UserLocalizedLabel && obj.UserLocalizedLabel.Label) || ''; } catch(e) { return ''; } };
    const flattenValue = (v) => { if (v == null) return ''; if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v); if (v && typeof v === 'object' && ('Value' in v)) return String(v.Value); try { if (v.UserLocalizedLabel && v.UserLocalizedLabel.Label) return v.UserLocalizedLabel.Label; } catch(e) {} try { return JSON.stringify(v); } catch(e) { return String(v); } };
    
    // Logging system
    const logs = [];
    const addLog = (message, level = 'info') => {
      const timestamp = new Date().toISOString().substr(11, 8);
      logs.push({ timestamp, level, message });
      console.log(`[${timestamp}] [${level.toUpperCase()}] ${message}`);
    };
    
    const showLogViewer = () => {
      const isDark = ui.host.getAttribute('data-theme') === 'dark';
      const theme = isDark ? { bg: '#1f2937', text: '#eef2ff', border: '#374151', codeBg: '#0b1220' } : { bg: '#fff', text: '#323130', border: '#e1e1e1', codeBg: '#f3f2f1' };
      const logOverlay = ce('div');
      logOverlay.style.cssText = `position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:999999;background:${theme.bg};color:${theme.text};border:1px solid ${theme.border};border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.5);width:800px;max-height:70vh;padding:16px;`;
      const titleDiv = txt(ce('div'), 'Activity Log');
      titleDiv.style.cssText = 'font:600 16px Segoe UI;margin-bottom:12px;';
      const logContent = ce('pre');
      logContent.style.cssText = `background:${theme.codeBg};padding:12px;overflow:auto;max-height:500px;font-size:12px;font-family:Consolas,monospace;margin:0;color:${theme.text};border:1px solid ${theme.border};border-radius:4px;`;
      logContent.textContent = logs.map(l => `[${l.timestamp}] [${l.level.toUpperCase()}] ${l.message}`).join('\n');
      const btnDiv = ce('div');
      btnDiv.style.cssText = 'margin-top:12px;display:flex;gap:8px;';
      const copyBtn = txt(ce('button'), 'Copy to Clipboard');
      copyBtn.style.cssText = 'padding:6px 12px;background:#0078d4;color:#fff;border:none;border-radius:4px;cursor:pointer;';
      copyBtn.onclick = () => {
        navigator.clipboard.writeText(logContent.textContent).then(() => alert('Logs copied to clipboard')).catch(() => alert('Copy failed'));
      };
      const closeBtn = txt(ce('button'), 'Close');
      closeBtn.style.cssText = `padding:6px 12px;background:${isDark ? '#374151' : '#e1e1e1'};color:${theme.text};border:none;border-radius:4px;cursor:pointer;`;
      closeBtn.onclick = () => document.body.removeChild(logOverlay);
      btnDiv.appendChild(copyBtn);
      btnDiv.appendChild(closeBtn);
      logOverlay.appendChild(titleDiv);
      logOverlay.appendChild(logContent);
      logOverlay.appendChild(btnDiv);
      document.body.appendChild(logOverlay);
    };
    
    // === Color Scheme Configuration ===
    const defaultColors = {
      light: {
        primary: '#0078d4',
        primaryHover: '#106ebe',
        border: '#ddd',
        borderLight: '#eee',
        text: '#323130',
        textMuted: '#605e5c',
        textError: '#a80000',
        bg: '#fff',
        bgHover: '#f3f2f1'
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
        bgHover: '#374151'
      }
    };

    let customColors = null;
    try {
      const saved = localStorage.getItem('crmab-tableexplorer-colors');
      if (saved) customColors = JSON.parse(saved);
    } catch (e) {
      addLog('Failed to load custom colors from localStorage', 'error');
    }

    const getColors = () => {
      const isDark = ui.host.getAttribute('data-theme') === 'dark';
      const mode = isDark ? 'dark' : 'light';
      return (customColors && customColors[mode]) || defaultColors[mode];
    };

    const showSettings = () => {
      const isDark = ui.host.getAttribute('data-theme') === 'dark';
      const theme = isDark ? { bg: '#1f2937', text: '#eef2ff', border: '#374151', inputBg: '#0b1220', inputBorder: '#4b5563', btnBg: '#374151', btnText: '#eef2ff' } 
                            : { bg: '#fff', text: '#323130', border: '#e1e1e1', inputBg: '#f3f2f1', inputBorder: '#c8c6c4', btnBg: '#e1e1e1', btnText: '#323130' };
      
      const settingsOverlay = ce('div');
      settingsOverlay.style.cssText = `position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:999999;background:${theme.bg};color:${theme.text};border:1px solid ${theme.border};border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.5);width:650px;max-height:85vh;overflow-y:auto;`;
      
      const header = txt(ce('div'), 'Color Scheme Settings');
      header.style.cssText = 'padding:20px 20px 16px 20px;font:600 18px Segoe UI;border-bottom:1px solid ' + theme.border;
      
      const content = ce('div');
      content.style.cssText = 'padding:20px;';
      
      // Mode tabs
      let selectedMode = 'light';
      const modeTabsDiv = ce('div');
      modeTabsDiv.style.cssText = 'display:flex;gap:8px;margin-bottom:20px;';
      
      const lightTab = txt(ce('button'), '☀️ Light Mode');
      const darkTab = txt(ce('button'), '🌙 Dark Mode');
      
      const updateTabStyles = () => {
        lightTab.style.cssText = `padding:8px 16px;border:1px solid ${theme.border};border-radius:4px;cursor:pointer;background:${selectedMode === 'light' ? theme.btnBg : 'transparent'};color:${theme.text};font-weight:${selectedMode === 'light' ? '600' : '400'};`;
        darkTab.style.cssText = `padding:8px 16px;border:1px solid ${theme.border};border-radius:4px;cursor:pointer;background:${selectedMode === 'dark' ? theme.btnBg : 'transparent'};color:${theme.text};font-weight:${selectedMode === 'dark' ? '600' : '400'};`;
      };
      updateTabStyles();
      
      const inputsContainer = ce('div');
      const colorInputs = {};
      
      const createColorInput = (label, key, description) => {
        const wrapper = ce('div');
        wrapper.style.cssText = 'margin-bottom:16px;';
        const labelDiv = txt(ce('div'), label);
        labelDiv.style.cssText = 'margin-bottom:4px;font-size:13px;font-weight:600;';
        const descDiv = txt(ce('div'), description);
        descDiv.style.cssText = 'margin-bottom:6px;font-size:11px;color:' + theme.textMuted;
        const inputWrapper = ce('div');
        inputWrapper.style.cssText = 'display:flex;gap:8px;align-items:center;';
        const input = ce('input');
        input.type = 'text';
        input.dataset.key = key;
        input.style.cssText = `flex:1;padding:8px;background:${theme.inputBg};color:${theme.text};border:1px solid ${theme.inputBorder};border-radius:4px;font-family:Consolas,monospace;font-size:13px;`;
        const preview = ce('div');
        preview.style.cssText = `width:48px;height:48px;border:1px solid ${theme.inputBorder};border-radius:4px;flex-shrink:0;`;
        input.addEventListener('input', () => {
          if (/^#[0-9A-Fa-f]{6}$/.test(input.value)) {
            preview.style.background = input.value;
            input.style.borderColor = theme.inputBorder;
          } else {
            input.style.borderColor = '#dc3545';
          }
        });
        inputWrapper.appendChild(input);
        inputWrapper.appendChild(preview);
        wrapper.appendChild(labelDiv);
        wrapper.appendChild(descDiv);
        wrapper.appendChild(inputWrapper);
        return { wrapper, input, preview };
      };
      
      const updateInputs = () => {
        clearEl(inputsContainer);
        colorInputs[selectedMode] = {};
        const currentColors = (customColors && customColors[selectedMode]) || defaultColors[selectedMode];
        
        const fields = [
          { label: 'Primary Color', key: 'primary', desc: 'Buttons, active tabs, accent elements' },
          { label: 'Primary Hover', key: 'primaryHover', desc: 'Hover state for primary elements' },
          { label: 'Border', key: 'border', desc: 'Main borders and dividers' },
          { label: 'Border Light', key: 'borderLight', desc: 'Subtle borders and separators' },
          { label: 'Text', key: 'text', desc: 'Primary text color' },
          { label: 'Text Muted', key: 'textMuted', desc: 'Secondary, less prominent text' },
          { label: 'Text Error', key: 'textError', desc: 'Error messages and warnings' },
          { label: 'Background', key: 'bg', desc: 'Main background color' },
          { label: 'Background Hover', key: 'bgHover', desc: 'Hover state for backgrounds' }
        ];
        
        fields.forEach(f => {
          const { wrapper, input, preview } = createColorInput(f.label, f.key, f.desc);
          input.value = currentColors[f.key];
          preview.style.background = currentColors[f.key];
          colorInputs[selectedMode][f.key] = input;
          inputsContainer.appendChild(wrapper);
        });
      };
      
      lightTab.onclick = () => { selectedMode = 'light'; updateTabStyles(); updateInputs(); };
      darkTab.onclick = () => { selectedMode = 'dark'; updateTabStyles(); updateInputs(); };
      
      modeTabsDiv.appendChild(lightTab);
      modeTabsDiv.appendChild(darkTab);
      content.appendChild(modeTabsDiv);
      content.appendChild(inputsContainer);
      
      // Buttons
      const buttonBar = ce('div');
      buttonBar.style.cssText = 'display:flex;gap:8px;margin-top:20px;padding-top:20px;border-top:1px solid ' + theme.border;
      
      const saveBtn = txt(ce('button'), '💾 Save');
      saveBtn.style.cssText = `padding:10px 20px;background:#0078d4;color:#fff;border:none;border-radius:4px;cursor:pointer;font-weight:600;`;
      saveBtn.onmouseenter = function(){ this.style.background = '#106ebe'; };
      saveBtn.onmouseleave = function(){ this.style.background = '#0078d4'; };
      saveBtn.onclick = () => {
        // Validate and collect values from currently displayed mode
        const newColors = { light: {}, dark: {} };
        let hasError = false;
        
        const currentMode = selectedMode;
        const currentInputs = colorInputs[currentMode];
        
        Object.keys(currentInputs).forEach(key => {
          const val = currentInputs[key].value;
          if (!/^#[0-9A-Fa-f]{6}$/.test(val)) {
            hasError = true;
            currentInputs[key].style.borderColor = '#dc3545';
          } else {
            newColors[currentMode][key] = val;
          }
        });
        
        if (hasError) {
          alert('Please fix invalid color values (must be hex format: #RRGGBB)');
          return;
        }
        
        // Get the other mode's colors (either from custom or defaults)
        const otherMode = currentMode === 'light' ? 'dark' : 'light';
        newColors[otherMode] = (customColors && customColors[otherMode]) || defaultColors[otherMode];
        
        customColors = newColors;
        try {
          localStorage.setItem('crmab-tableexplorer-colors', JSON.stringify(customColors));
          addLog('Custom colors saved successfully', 'success');
          document.body.removeChild(settingsOverlay);
          alert('Color scheme saved! Please refresh the page to apply the new colors.');
        } catch (e) {
          addLog('Failed to save custom colors: ' + e.message, 'error');
          alert('Failed to save settings: ' + e.message);
        }
      };
      
      const resetBtn = txt(ce('button'), '🔄 Reset to Defaults');
      resetBtn.style.cssText = `padding:10px 20px;background:${theme.btnBg};color:${theme.btnText};border:1px solid ${theme.border};border-radius:4px;cursor:pointer;`;
      resetBtn.onmouseenter = function(){ this.style.opacity = '0.8'; };
      resetBtn.onmouseleave = function(){ this.style.opacity = '1'; };
      resetBtn.onclick = () => {
        if (confirm('Reset all colors to defaults? This will clear your custom color scheme.')) {
          customColors = null;
          localStorage.removeItem('crmab-tableexplorer-colors');
          addLog('Colors reset to defaults', 'info');
          document.body.removeChild(settingsOverlay);
          alert('Colors reset to defaults! Please refresh the page to see changes.');
        }
      };
      
      const cancelBtn = txt(ce('button'), '✖ Cancel');
      cancelBtn.style.cssText = `padding:10px 20px;background:transparent;color:${theme.text};border:1px solid ${theme.border};border-radius:4px;cursor:pointer;`;
      cancelBtn.onmouseenter = function(){ this.style.background = theme.btnBg; };
      cancelBtn.onmouseleave = function(){ this.style.background = 'transparent'; };
      cancelBtn.onclick = () => document.body.removeChild(settingsOverlay);
      
      buttonBar.appendChild(saveBtn);
      buttonBar.appendChild(resetBtn);
      buttonBar.appendChild(cancelBtn);
      content.appendChild(buttonBar);
      
      settingsOverlay.appendChild(header);
      settingsOverlay.appendChild(content);
      document.body.appendChild(settingsOverlay);
      
      updateInputs();
      
      settingsOverlay.addEventListener('click', (e) => { 
        if (e.target === settingsOverlay) document.body.removeChild(settingsOverlay); 
      });
    };
    
    addLog('Table Explorer initialized', 'success');
    
    const entitySelectDiv = style(ce('div'), 'margin-bottom:12px;display:flex;gap:8px;align-items:center;position:relative;');
    const entitySelect = ce('select');
    entitySelect.style.cssText = 'max-width:250px;min-width:150px;padding:6px 8px;border:1px solid #ddd;border-radius:4px;';
    const envVarBtn = txt(ce('button'), '🌍 Environment Variables');
    envVarBtn.style.cssText = 'padding:6px 16px;border:1px solid #0078d4;border-radius:4px;background:#0078d4;color:#fff;cursor:pointer;font-weight:600;white-space:nowrap;';
    envVarBtn.addEventListener('mouseenter', function(){ this.style.background = '#106ebe'; });
    envVarBtn.addEventListener('mouseleave', function(){ this.style.background = '#0078d4'; });
    envVarBtn.onclick = () => switchTab('Environment Variables');
    const globalChoicesBtn = txt(ce('button'), '🌐 Global Choices');
    globalChoicesBtn.style.cssText = 'padding:6px 16px;border:1px solid #0078d4;border-radius:4px;background:#0078d4;color:#fff;cursor:pointer;font-weight:600;white-space:nowrap;';
    globalChoicesBtn.addEventListener('mouseenter', function(){ this.style.background = '#106ebe'; });
    globalChoicesBtn.addEventListener('mouseleave', function(){ this.style.background = '#0078d4'; });
    globalChoicesBtn.onclick = () => switchTab('Global Choices');
    entitySelectDiv.appendChild(txt(style(ce('span'), 'font-weight:600;'), 'Entity:'));
    entitySelectDiv.appendChild(entitySelect);
    entitySelectDiv.appendChild(envVarBtn);
    entitySelectDiv.appendChild(globalChoicesBtn);
    ui.body.appendChild(entitySelectDiv);
    
    const tabsDiv = style(ce('div'), 'display:flex;gap:4px;border-bottom:2px solid #ddd;margin-bottom:12px;flex-wrap:wrap;');
    const tabs = ['Properties', '1:N', 'N:1', 'N:N', 'Forms', 'Views', 'Fields', 'Option Sets', 'Business Rules', 'Plugins', 'JS', 'Prozesse'];
    const tabButtons = {};
    let currentTab = 'Properties';
    tabs.forEach(tab => {
      const btn = txt(ce('button'), tab);
      btn.className = 'crmab-te-tab';
      btn.dataset.tab = tab;
      btn.style.cssText = `padding:8px 16px;border:none;border-bottom:3px solid ${currentTab === tab ? '#0078d4' : 'transparent'};background:transparent;color:${currentTab === tab ? '#0078d4' : '#605e5c'};cursor:pointer;font-weight:${currentTab === tab ? '600' : '400'};`;
      btn.addEventListener('mouseenter', function(){ if (currentTab !== tab) this.style.color = '#323130'; });
      btn.addEventListener('mouseleave', function(){ if (currentTab !== tab) this.style.color = currentTab === tab ? '#0078d4' : '#605e5c'; });
      btn.addEventListener('click', () => switchTab(tab));
      tabsDiv.appendChild(btn);
      tabButtons[tab] = btn;
    });
    ui.body.appendChild(tabsDiv);
    
    const contentDiv = style(ce('div'), 'min-height:400px;');
    const searchContentWrapper = style(ce('div'), 'position:relative;margin-bottom:12px;display:none;');
    const searchContentInput = ce('input');
    searchContentInput.type = 'text';
    searchContentInput.placeholder = 'Search in current tab...';
    searchContentInput.style.cssText = 'width:100%;padding:6px 28px 6px 8px;border:1px solid #ddd;border-radius:4px;';
    const searchContentClearBtn = txt(ce('button'), '×');
    searchContentClearBtn.style.cssText = 'position:absolute;right:4px;top:50%;transform:translateY(-50%);background:transparent;border:none;color:#999;cursor:pointer;font-size:18px;line-height:1;padding:2px 6px;display:none;';
    searchContentClearBtn.onclick = () => { searchContentInput.value = ''; if (searchContentInput.oninput) searchContentInput.oninput(); searchContentInput.focus(); };
    searchContentWrapper.appendChild(searchContentInput);
    searchContentWrapper.appendChild(searchContentClearBtn);
    contentDiv.appendChild(searchContentWrapper);
    const tableContainer = ce('div');
    contentDiv.appendChild(tableContainer);
    ui.body.appendChild(contentDiv);
    
    let entities = [], currentEntity = null, entityMetadata = null, relationships = null, forms = [], views = [], fields = [], optionSets = [], businessRules = [], plugins = [], javascript = [], processes = [], envVars = [];
    let formAutomationsCache = {};
    
    const switchTab = (tab) => {
      currentTab = tab;
      Object.keys(tabButtons).forEach(t => {
        const btn = tabButtons[t];
        btn.style.borderBottomColor = t === tab ? '#0078d4' : 'transparent';
        btn.style.color = t === tab ? '#0078d4' : '#605e5c';
        btn.style.fontWeight = t === tab ? '600' : '400';
      });
      searchContentWrapper.style.display = 'block';
      searchContentInput.value = '';
      searchContentClearBtn.style.display = 'none';
      renderTab(tab);
    };
    
    const renderTab = async (tab) => {
      addLog(`Rendering tab: ${tab}`, 'info');
      clearEl(tableContainer);
      if (!currentEntity && tab !== 'Environment Variables' && tab !== 'Global Choices') { tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;text-align:center;color:#605e5c;'), 'Please select an entity')); return; }
      tableContainer.appendChild(txt(style(ce('div'), 'padding:8px;color:#605e5c;font-size:12px;'), 'Loading...'));
      
      try {
        if (tab === 'Properties') {
          await loadProperties();
        } else if (tab === '1:N') {
          await loadRelationships('OneToMany');
        } else if (tab === 'N:1') {
          await loadRelationships('ManyToOne');
        } else if (tab === 'N:N') {
          await loadRelationships('ManyToMany');
        } else if (tab === 'Forms') {
          await loadForms();
        } else if (tab === 'Views') {
          await loadViews();
        } else if (tab === 'Fields') {
          await loadFields();
        } else if (tab === 'Option Sets') {
          await loadOptionSets();
        } else if (tab === 'Business Rules') {
          await loadBusinessRules();
        } else if (tab === 'Plugins') {
          await loadPlugins();
        } else if (tab === 'JS') {
          await loadJavaScript();
        } else if (tab === 'Prozesse') {
          await loadProcesses();
        } else if (tab === 'Environment Variables') {
          await loadEnvironmentVariables();
        } else if (tab === 'Global Choices') {
          await loadGlobalChoices();
        }
      } catch (ex) {
        addLog(`Error rendering tab: ${ex.message || ex}`, 'error');
        clearEl(tableContainer);
        tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;color:#a80000;'), 'Error: ' + (ex.message || ex)));
      }
    };
    
    const loadProperties = async () => {
      if (!entityMetadata) {
        const sq = String.fromCharCode(39);
        entityMetadata = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})`);
      }
      clearEl(tableContainer);
      const skip = new Set(['Attributes', 'ManyToManyRelationships', 'ManyToOneRelationships', 'OneToManyRelationships', 'DisplayName']);
      const rows = [];
      for (const k in entityMetadata) {
        if (!entityMetadata.hasOwnProperty(k) || skip.has(k)) continue;
        rows.push({ k, v: flattenValue(entityMetadata[k]) });
      }
      rows.sort((a, b) => a.k.toLowerCase().localeCompare(b.k.toLowerCase()));
      renderTable(['Name', 'Value'], rows.map(r => [r.k, r.v]));
      if (ui.updateFooter) ui.updateFooter(`Properties: ${rows.length}`);
    };
    
    const loadRelationships = async (type) => {
      if (!relationships) {
        const sq = String.fromCharCode(39);
        try {
          const relMeta = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})?$expand=ManyToOneRelationships($select=SchemaName,ReferencingEntity,ReferencedEntity,ReferencingAttribute,ReferencedAttribute,IsCustomRelationship,IsValidForAdvancedFind,IsCustomizable),OneToManyRelationships($select=SchemaName,ReferencingEntity,ReferencedEntity,ReferencingAttribute,ReferencedAttribute,IsCustomRelationship,IsValidForAdvancedFind,IsCustomizable),ManyToManyRelationships($select=SchemaName,Entity1LogicalName,Entity2LogicalName,IntersectEntityName,IsCustomRelationship,IsValidForAdvancedFind,IsCustomizable)`);
          relationships = { ManyToOne: relMeta.ManyToOneRelationships || [], OneToMany: relMeta.OneToManyRelationships || [], ManyToMany: relMeta.ManyToManyRelationships || [] };
        } catch (e) {
          try {
            const relMeta = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})?$expand=ManyToOneRelationships($select=SchemaName,ReferencingEntity,ReferencedEntity,ReferencingAttribute,ReferencedAttribute),OneToManyRelationships($select=SchemaName,ReferencingEntity,ReferencedEntity,ReferencingAttribute,ReferencedAttribute),ManyToManyRelationships($select=SchemaName,Entity1LogicalName,Entity2LogicalName,IntersectEntityName)`);
            relationships = { ManyToOne: relMeta.ManyToOneRelationships || [], OneToMany: relMeta.OneToManyRelationships || [], ManyToMany: relMeta.ManyToManyRelationships || [] };
          } catch (e2) {
            console.warn('Error loading relationships:', e2);
            relationships = { ManyToOne: [], OneToMany: [], ManyToMany: [] };
          }
        }
      }
      clearEl(tableContainer);
      const rels = relationships[type] || [];
      let headers = [], rows = [];
      if (type === 'OneToMany') {
        headers = ['Schema', 'Referencing Entity', 'Referenced Entity', 'Referencing Attr', 'Referenced Attr', 'IsCustom', 'IsValidForAdvancedFind', 'IsCustomizable'];
        rows = rels.map(r => [
          r.SchemaName || '', 
          r.ReferencingEntity || '', 
          r.ReferencedEntity || '', 
          r.ReferencingAttribute || '', 
          r.ReferencedAttribute || '',
          r.IsCustomRelationship !== undefined ? (r.IsCustomRelationship ? 'Yes' : 'No') : '',
          r.IsValidForAdvancedFind !== undefined ? (r.IsValidForAdvancedFind ? 'Yes' : 'No') : '',
          r.IsCustomizable !== undefined ? (r.IsCustomizable ? 'Yes' : 'No') : ''
        ]);
      } else if (type === 'ManyToOne') {
        headers = ['Schema', 'Referencing Entity', 'Referenced Entity', 'Referencing Attr', 'Referenced Attr', 'IsCustom', 'IsValidForAdvancedFind', 'IsCustomizable'];
        rows = rels.map(r => [
          r.SchemaName || '', 
          r.ReferencingEntity || '', 
          r.ReferencedEntity || '', 
          r.ReferencingAttribute || '', 
          r.ReferencedAttribute || '',
          r.IsCustomRelationship !== undefined ? (r.IsCustomRelationship ? 'Yes' : 'No') : '',
          r.IsValidForAdvancedFind !== undefined ? (r.IsValidForAdvancedFind ? 'Yes' : 'No') : '',
          r.IsCustomizable !== undefined ? (r.IsCustomizable ? 'Yes' : 'No') : ''
        ]);
      } else if (type === 'ManyToMany') {
        headers = ['Schema', 'Entity 1', 'Entity 2', 'Intersect Entity', 'IsCustom', 'IsValidForAdvancedFind', 'IsCustomizable'];
        rows = rels.map(r => [
          r.SchemaName || '', 
          r.Entity1LogicalName || '', 
          r.Entity2LogicalName || '', 
          r.IntersectEntityName || '',
          r.IsCustomRelationship !== undefined ? (r.IsCustomRelationship ? 'Yes' : 'No') : '',
          r.IsValidForAdvancedFind !== undefined ? (r.IsValidForAdvancedFind ? 'Yes' : 'No') : '',
          r.IsCustomizable !== undefined ? (r.IsCustomizable ? 'Yes' : 'No') : ''
        ]);
      }
      renderTable(headers, rows, true);
      if (ui.updateFooter) ui.updateFooter(`${type}: ${rows.length}`);
    };
    
    const extractFormAutomations = async (formId) => {
      if (formAutomationsCache[formId]) return formAutomationsCache[formId];
      try {
        const formData = await fetchJson(`/systemforms(${formId})?$select=formxml,name`);
        const xml = formData.formxml || '';
        const doc = new DOMParser().parseFromString(xml, 'text/xml');
        if (doc.getElementsByTagName('parsererror').length > 0) return { libraries: [], handlers: [] };
        
        const libraries = [];
        const libNodes = doc.getElementsByTagName('Library');
        for (let i = 0; i < libNodes.length; i++) {
          const n = libNodes[i];
          const name = n.getAttribute('name') || '';
          if (name && !libraries.some(lib => lib.name === name)) {
            libraries.push({ name: name, id: n.getAttribute('libraryUniqueId') || '' });
          }
        }
        
        const handlers = [];
        const eventNodes = doc.getElementsByTagName('event');
        for (let j = 0; j < eventNodes.length; j++) {
          const ev = eventNodes[j];
          const evName = ev.getAttribute('name') || '';
          let target = 'Form';
          const targetId = ev.getAttribute('attribute') || '';
          if (targetId) target = targetId;
          
          const hs = ev.getElementsByTagName('handler');
          for (let h = 0; h < hs.length; h++) {
            const hn = hs[h];
            handlers.push({
              target: target,
              event: evName,
              library: hn.getAttribute('libraryName') || '',
              functionName: hn.getAttribute('functionName') || '',
              enabled: hn.getAttribute('enabled') !== 'false'
            });
          }
        }
        
        formAutomationsCache[formId] = { libraries, handlers };
        return formAutomationsCache[formId];
      } catch (e) {
        console.warn('Error extracting form automations:', e);
        return { libraries: [], handlers: [] };
      }
    };
    
    const loadForms = async () => {
      if (forms.length === 0) {
        try {
          const formsData = await fetchJson(`/systemforms?$select=formid,name&$filter=objecttypecode eq '${currentEntity}' and type eq 2`);
          forms = (formsData.value || []).map(f => ({ id: f.formid || f.systemformid, name: f.name || '' }));
        } catch (e) {
          try {
            const sq = String.fromCharCode(39);
            const entityMeta = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})?$select=ObjectTypeCode`);
            const otc = entityMeta && entityMeta.ObjectTypeCode != null ? entityMeta.ObjectTypeCode : null;
            if (otc != null) {
              const formsData = await fetchJson(`/systemforms?$select=formid,name&$filter=objecttypecode eq ${otc} and type eq 2`);
              forms = (formsData.value || []).map(f => ({ id: f.formid || f.systemformid, name: f.name || '' }));
            } else {
              forms = [];
            }
          } catch (e2) {
            console.warn('Error loading forms:', e2);
            forms = [];
          }
        }
      }
      clearEl(tableContainer);
      const headers = ['Name', 'ID', 'Actions'];
      const rows = forms.map(f => [f.name, f.id, f]);
      renderTable(headers, rows, true, (row, cellIdx, cell) => {
        if (cellIdx === 2) {
          const td = ce('td');
          td.style.cssText = 'padding:6px 8px;border-bottom:1px solid #eee;';
          const btnContainer = style(ce('div'), 'display:flex;gap:4px;flex-wrap:wrap;');
          const f = cell;
          const jsBtn = txt(ce('button'), '📜 JS');
          jsBtn.style.cssText = 'padding:4px 8px;background:#0078d4;color:#fff;border:none;border-radius:3px;cursor:pointer;font-size:11px;';
          jsBtn.addEventListener('mouseenter', function(){ this.style.background = '#106ebe'; });
          jsBtn.addEventListener('mouseleave', function(){ this.style.background = '#0078d4'; });
          jsBtn.addEventListener('click', async () => {
            const automations = await extractFormAutomations(f.id);
            switchTab('JS');
            javascript = automations.handlers.map(h => ({
              form: f.name || '',
              target: h.target || 'Form',
              event: h.event || '',
              library: h.library || '',
              functionName: h.functionName || '',
              enabled: h.enabled ? 'Yes' : 'No'
            }));
            await loadJavaScript();
          });
          btnContainer.appendChild(jsBtn);
          td.appendChild(btnContainer);
          return td;
        }
        return null;
      });
      if (ui.updateFooter) ui.updateFooter(`Forms: ${forms.length}`);
    };
    
    const loadViews = async () => {
      if (views.length === 0) {
        const sq = String.fromCharCode(39);
        try {
          const viewsData = await fetchJson(`/savedqueries?$filter=returnedtypecode eq ${encodeURIComponent(sq + currentEntity + sq)} and statecode eq 0&$select=name,savedqueryid,returnedtypecode,fetchxml,layoutxml`);
          views = (viewsData.value || []).map(v => ({ id: v.savedqueryid, name: v.name || '', type: 'System', fetchxml: v.fetchxml || '', layoutxml: v.layoutxml || '' }));
        } catch (e) {
          views = [];
        }
        try {
          const userViewsData = await fetchJson(`/userqueries?$filter=returnedtypecode eq ${encodeURIComponent(sq + currentEntity + sq)}&$select=name,userqueryid,returnedtypecode,fetchxml,layoutxml`);
          const userViews = (userViewsData.value || []).map(v => ({ id: v.userqueryid, name: v.name || '', type: 'Personal', fetchxml: v.fetchxml || '', layoutxml: v.layoutxml || '' }));
          views = views.concat(userViews);
        } catch (e) {}
      }
      clearEl(tableContainer);
      renderViewsTable(views);
      if (ui.updateFooter) ui.updateFooter(`Views: ${views.length}`);
    };
    
    const loadFields = async () => {
      if (fields.length === 0) {
        const sq = String.fromCharCode(39);
        try {
          const fieldsData = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})/Attributes?$select=LogicalName,SchemaName,AttributeType,DisplayName,RequiredLevel,IsCustomAttribute&$orderby=LogicalName`);
          const getBool = (val) => {
            if (val === undefined || val === null) return '';
            return val ? 'Yes' : 'No';
          };
          fields = (fieldsData.value || []).map(f => ({
            logicalName: f.LogicalName || '', 
            schemaName: f.SchemaName || '', 
            type: f.AttributeType || '', 
            displayName: dispName(f.DisplayName) || '',
            requiredLevel: f.RequiredLevel && f.RequiredLevel.Value ? f.RequiredLevel.Value : 'None',
            maxLength: '',
            minValue: '',
            maxValue: '',
            precision: '',
            isCustomAttribute: getBool(f.IsCustomAttribute)
          }));
          
          const fieldMap = new Map(fields.map(f => [f.logicalName.toLowerCase(), f]));
          
          try {
            const strData = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})/Attributes/Microsoft.Dynamics.CRM.StringAttributeMetadata?$select=LogicalName,MaxLength`);
            (strData.value || []).forEach(s => {
              const f = fieldMap.get((s.LogicalName || '').toLowerCase());
              if (f) f.maxLength = s.MaxLength != null ? String(s.MaxLength) : '';
            });
          } catch (e) {}
          
          try {
            const numData = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})/Attributes/Microsoft.Dynamics.CRM.DecimalAttributeMetadata?$select=LogicalName,MinValue,MaxValue,Precision`);
            (numData.value || []).forEach(n => {
              const f = fieldMap.get((n.LogicalName || '').toLowerCase());
              if (f) {
                f.minValue = n.MinValue != null ? String(n.MinValue) : '';
                f.maxValue = n.MaxValue != null ? String(n.MaxValue) : '';
                f.precision = n.Precision != null ? String(n.Precision) : '';
              }
            });
          } catch (e) {}
          
          try {
            const intData = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})/Attributes/Microsoft.Dynamics.CRM.IntegerAttributeMetadata?$select=LogicalName,MinValue,MaxValue`);
            (intData.value || []).forEach(i => {
              const f = fieldMap.get((i.LogicalName || '').toLowerCase());
              if (f) {
                f.minValue = i.MinValue != null ? String(i.MinValue) : '';
                f.maxValue = i.MaxValue != null ? String(i.MaxValue) : '';
              }
            });
          } catch (e) {}
          
          try {
            const dblData = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})/Attributes/Microsoft.Dynamics.CRM.DoubleAttributeMetadata?$select=LogicalName,MinValue,MaxValue,Precision`);
            (dblData.value || []).forEach(d => {
              const f = fieldMap.get((d.LogicalName || '').toLowerCase());
              if (f) {
                f.minValue = d.MinValue != null ? String(d.MinValue) : '';
                f.maxValue = d.MaxValue != null ? String(d.MaxValue) : '';
                f.precision = d.Precision != null ? String(d.Precision) : '';
              }
            });
          } catch (e) {}
          
          try {
            const monData = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})/Attributes/Microsoft.Dynamics.CRM.MoneyAttributeMetadata?$select=LogicalName,MinValue,MaxValue,Precision`);
            (monData.value || []).forEach(m => {
              const f = fieldMap.get((m.LogicalName || '').toLowerCase());
              if (f) {
                f.minValue = m.MinValue != null ? String(m.MinValue) : '';
                f.maxValue = m.MaxValue != null ? String(m.MaxValue) : '';
                f.precision = m.Precision != null ? String(m.Precision) : '';
              }
            });
          } catch (e) {}
          
          fields = Array.from(fieldMap.values());
        } catch (e) {
          console.warn('Error loading fields:', e);
          fields = [];
        }
      }
      clearEl(tableContainer);
      const requiredLevelText = (rl) => {
        if (rl === 'ApplicationRequired' || rl === 'SystemRequired') return 'Required';
        if (rl === 'Recommended') return 'Recommended';
        return 'Optional';
      };
      renderTable([
        'Logical Name', 'Schema Name', 'Type', 'Display Name', 'Required', 
        'MaxLength', 'MinValue', 'MaxValue', 'Precision',
        'IsCustomAttribute'
      ], fields.map(f => [
        f.logicalName, f.schemaName, f.type, f.displayName, requiredLevelText(f.requiredLevel),
        f.maxLength || '', f.minValue || '', f.maxValue || '', f.precision || '',
        f.isCustomAttribute
      ]), true);
      if (ui.updateFooter) ui.updateFooter(`Fields: ${fields.length}`);
    };
    
    const loadOptionSets = async () => {
      if (optionSets.length === 0) {
        const sq = String.fromCharCode(39);
        try {
          const picklistData = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName,SchemaName&$expand=OptionSet($select=Name,IsGlobal,Options)`);
          const multiSelectData = await fetchJson(`/EntityDefinitions(LogicalName=${encodeURIComponent(sq + currentEntity + sq)})/Attributes/Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata?$select=LogicalName,SchemaName&$expand=OptionSet($select=Name,IsGlobal,Options)`);
          const picklistAttrs = (picklistData.value || []).map(a => ({
            logicalName: a.LogicalName || '',
            schemaName: a.SchemaName || '',
            optionSetName: a.OptionSet && a.OptionSet.Name ? a.OptionSet.Name : '',
            isGlobal: a.OptionSet && a.OptionSet.IsGlobal ? 'Yes' : 'No',
            options: (a.OptionSet && a.OptionSet.Options || []).map(opt => ({
              value: opt.Value != null ? opt.Value : '',
              label: dispName(opt.Label) || ''
            }))
          }));
          const multiSelectAttrs = (multiSelectData.value || []).map(a => ({
            logicalName: a.LogicalName || '',
            schemaName: a.SchemaName || '',
            optionSetName: a.OptionSet && a.OptionSet.Name ? a.OptionSet.Name : '',
            isGlobal: a.OptionSet && a.OptionSet.IsGlobal ? 'Yes' : 'No',
            options: (a.OptionSet && a.OptionSet.Options || []).map(opt => ({
              value: opt.Value != null ? opt.Value : '',
              label: dispName(opt.Label) || ''
            }))
          }));
          optionSets = picklistAttrs.concat(multiSelectAttrs);
        } catch (e) {
          console.warn('Error loading option sets:', e);
          optionSets = [];
        }
      }
      clearEl(tableContainer);
      if (optionSets.length === 0) {
        tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;text-align:center;color:#605e5c;'), 'No option sets available'));
        if (ui.updateFooter) ui.updateFooter('Option Sets: 0');
        return;
      }
      const rows = [];
      optionSets.forEach(os => {
        if (os.options && os.options.length > 0) {
          os.options.forEach((opt, idx) => {
            rows.push([
              os.optionSetName || '',
              os.logicalName || '',
              os.schemaName || '',
              os.isGlobal || '',
              opt.value !== '' ? String(opt.value) : '',
              opt.label || ''
            ]);
          });
        } else {
          rows.push([
            os.optionSetName || '',
            os.logicalName || '',
            os.schemaName || '',
            os.isGlobal || '',
            '',
            '(no options)'
          ]);
        }
      });
      renderTable(['Option Set Name', 'Field Logical Name', 'Field Schema Name', 'Is Global', 'Value', 'Label'], rows, true);
      if (ui.updateFooter) ui.updateFooter(`Option Sets: ${optionSets.length} (${rows.length} values)`);
    };
    
    const createCell = (row, origIdx, cell, customCellRenderer) => {
      if (customCellRenderer) {
        const result = customCellRenderer(row, origIdx, cell);
        if (result && result.nodeType === 1) return result;
      }
      return txt(ce('td'), cell || '');
    };
    const renderRow = (row, columnOrder, customCellRenderer) => {
      const tr = ce('tr');
      columnOrder.forEach((origIdx) => {
        const cell = row[origIdx];
        const td = createCell(row, origIdx, cell, customCellRenderer);
        td.style.cssText = 'padding:6px 8px;border-bottom:1px solid #eee;';
        tr.appendChild(td);
      });
      return tr;
    };
    const downloadFile = (content, filename, mimeType) => {
      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = ce('a');
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    };
    const renderTable = (headers, rows, searchable = false, customCellRenderer = null) => {
      clearEl(tableContainer);
      if (rows.length === 0) {
        tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;text-align:center;color:#605e5c;'), 'No data available'));
        return;
      }
      const table = ce('table');
      table.style.cssText = 'width:100%;border-collapse:collapse;font-size:12px;';
      const thead = ce('thead');
      const trh = ce('tr');
      let sortColumn = -1;
      let sortDirection = 1;
      let draggedColumn = null;
      let columnOrder = headers.map((_, i) => i);
      const reorderColumns = (fromIdx, toIdx) => {
        const newOrder = [...columnOrder];
        const [removed] = newOrder.splice(fromIdx, 1);
        newOrder.splice(toIdx, 0, removed);
        columnOrder = newOrder;
        const newHeaders = columnOrder.map(i => headers[i]);
        const newRows = rows.map(row => columnOrder.map(i => row[i]));
        const originalRows = rows;
        renderTable(newHeaders, newRows, searchable, (row, cellIdx, cell) => {
          if (customCellRenderer) {
            const origIdx = columnOrder[cellIdx];
            return customCellRenderer(originalRows.find(r => r === row || JSON.stringify(r) === JSON.stringify(row)) || row, origIdx, cell);
          }
          return null;
        });
      };
      
      headers.forEach((h, colIdx) => {
        const th = ce('th');
        th.style.cssText = 'padding:8px;text-align:left;border-bottom:2px solid #ddd;font-weight:600;background:#f8fafc;cursor:pointer;user-select:none;position:relative;';
        th.appendChild(txt(ce('span'), h + ' '));
        const sortIndicator = txt(style(ce('span'), 'opacity:0.3;font-size:10px;'), '⇅');
        th.appendChild(sortIndicator);
        th.draggable = true;
        th.dataset.columnIndex = colIdx;
        th.ondragstart = (e) => {
          draggedColumn = colIdx;
          e.dataTransfer.effectAllowed = 'move';
          th.style.opacity = '0.5';
        };
        th.ondragend = () => {
          th.style.opacity = '1';
          draggedColumn = null;
        };
        th.ondragover = (e) => {
          e.preventDefault();
          e.dataTransfer.dropEffect = 'move';
          if (draggedColumn !== null && draggedColumn !== colIdx) {
            th.style.background = '#e0e7ff';
          }
        };
        th.ondragleave = () => {
          if (draggedColumn !== colIdx) {
            th.style.background = sortColumn === colIdx ? '#f0f2f5' : '#f8fafc';
          }
        };
        th.ondrop = (e) => {
          e.preventDefault();
          if (draggedColumn !== null && draggedColumn !== colIdx) {
            const fromIdx = columnOrder.indexOf(draggedColumn);
            const toIdx = columnOrder.indexOf(colIdx);
            if (fromIdx !== -1 && toIdx !== -1) {
              reorderColumns(fromIdx, toIdx);
            }
          }
          th.style.background = sortColumn === colIdx ? '#f0f2f5' : '#f8fafc';
        };
        th.onclick = () => {
          if (sortColumn === colIdx) {
            sortDirection = -sortDirection;
          } else {
            sortColumn = colIdx;
            sortDirection = 1;
          }
          const actualColIdx = columnOrder[colIdx];
          const sortedRows = [...rows].sort((a, b) => {
            const aVal = (a[actualColIdx] || '').toString().toLowerCase();
            const bVal = (b[actualColIdx] || '').toString().toLowerCase();
            const numA = parseFloat(aVal);
            const numB = parseFloat(bVal);
            if (!isNaN(numA) && !isNaN(numB)) {
              return (numA - numB) * sortDirection;
            }
            return aVal.localeCompare(bVal) * sortDirection;
          });
          Array.from(thead.querySelectorAll('th')).forEach((th, idx) => {
            const spans = th.querySelectorAll('span');
            const sortSpan = spans.length > 1 ? spans[spans.length - 1] : null;
            if (sortSpan) {
              if (idx === colIdx) {
                sortSpan.textContent = sortDirection > 0 ? ' ↑' : ' ↓';
                sortSpan.style.opacity = '1';
              } else {
                sortSpan.textContent = ' ⇅';
                sortSpan.style.opacity = '0.3';
              }
            }
          });
          clearEl(tbody);
          sortedRows.forEach((row, idx) => {
            const tr = renderRow(row, columnOrder, customCellRenderer);
            tr.dataset.rowIndex = idx;
            tbody.appendChild(tr);
          });
          filterRows();
        };
        th.onmouseenter = () => { if (sortColumn !== colIdx) th.style.background = '#f0f2f5'; };
        th.onmouseleave = () => { if (sortColumn !== colIdx) th.style.background = '#f8fafc'; };
        trh.appendChild(th);
      });
      thead.appendChild(trh);
      table.appendChild(thead);
      const tbody = ce('tbody');
      rows.forEach((row, idx) => {
        const tr = renderRow(row, columnOrder, customCellRenderer);
        tr.dataset.rowIndex = idx;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      tableContainer.appendChild(table);
      
      const filterRows = () => {
        const q = (searchContentInput.value || '').toLowerCase();
        searchContentClearBtn.style.display = q ? 'block' : 'none';
        Array.from(tbody.children).forEach(tr => {
          const text = Array.from(tr.children).map(td => td.textContent).join(' ').toLowerCase();
          tr.style.display = !q || text.includes(q) ? '' : 'none';
        });
        const visibleCount = Array.from(tbody.children).filter(tr => tr.style.display !== 'none').length;
        if (ui.updateFooter) {
          const totalCount = rows.length;
          if (q) {
            ui.updateFooter(`${visibleCount} of ${totalCount} result(s) shown`);
          } else {
            ui.updateFooter(`${totalCount} result(s)`);
          }
        }
      };
      
      searchContentWrapper.style.display = 'block';
      searchContentInput.oninput = filterRows;
      searchContentInput.placeholder = `Search in ${headers.join(', ').toLowerCase()}...`;
      
      const exportBar = style(ce('div'), 'display:flex;gap:8px;margin-bottom:12px;justify-content:flex-end;');
      const btnStyle = 'padding:4px 8px;background:#0078d4;color:#fff;border:none;border-radius:3px;cursor:pointer;font-size:11px;';
      const entityName = currentEntity || 'entity';
      const baseName = `${entityName}_${currentTab.toLowerCase().replace(/\s+/g, '_')}`;
      const createExportBtn = (label, content, ext, mime) => {
        const btn = txt(ce('button'), label);
        btn.style.cssText = btnStyle;
        btn.onclick = () => downloadFile(content, `${baseName}.${ext}`, mime);
        return btn;
      };
      const csvContent = [headers.map(h => `"${h.replace(/"/g, '""')}"`).join(',')].concat(rows.map(row => row.map(c => `"${String(c || '').replace(/"/g, '""')}"`).join(','))).join('\n');
      const jsonContent = JSON.stringify({ entity: entityName, tab: currentTab, headers, rows }, null, 2);
      const mdContent = [`# ${currentTab}: ${entityName}`, '', '| ' + headers.join(' | ') + ' |', '|' + headers.map(() => '---').join('|') + '|'].concat(rows.map(row => '| ' + row.map(c => String(c || '').replace(/\n/g, ' ')).join(' | ') + ' |')).join('\n');
      exportBar.appendChild(createExportBtn('📊 CSV', csvContent, 'csv', 'text/csv'));
      exportBar.appendChild(createExportBtn('📄 JSON', jsonContent, 'json', 'application/json'));
      exportBar.appendChild(createExportBtn('📝 MD', mdContent, 'md', 'text/markdown'));
      tableContainer.insertBefore(exportBar, table);
    };
    
    const downloadXML = (xml, filename) => {
      try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xml, 'text/xml');
        if (xmlDoc.getElementsByTagName('parsererror').length > 0) throw new Error('XML parse error');
        const formatXML = (xml, indent = 0) => {
          let formatted = '';
          const indentStr = '  '.repeat(indent);
          if (xml.nodeType === 1) {
            formatted += indentStr + '<' + xml.nodeName;
            for (let i = 0; i < xml.attributes.length; i++) {
              const attr = xml.attributes[i];
              formatted += ' ' + attr.name + '="' + attr.value + '"';
            }
            if (xml.childNodes.length === 0) {
              formatted += ' />\n';
            } else {
              formatted += '>\n';
              for (let i = 0; i < xml.childNodes.length; i++) {
                formatted += formatXML(xml.childNodes[i], indent + 1);
              }
              formatted += indentStr + '</' + xml.nodeName + '>\n';
            }
          } else if (xml.nodeType === 3) {
            const text = xml.textContent.trim();
            if (text) formatted += indentStr + text + '\n';
          }
          return formatted;
        };
        downloadFile('<?xml version="1.0" encoding="utf-8"?>\n' + formatXML(xmlDoc.documentElement), filename, 'application/xml');
      } catch (e) {
        downloadFile(xml, filename, 'application/xml');
      }
    };
    const renderViewsTable = (views) => {
      clearEl(tableContainer);
      if (views.length === 0) {
        tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;text-align:center;color:#605e5c;'), 'No views available'));
        return;
      }
      const headers = ['Name', 'ID', 'Type', 'Actions'];
      const rows = views.map(v => [v.name || '', v.id || '', v.type || '', v]);
      renderTable(headers, rows, true, (row, cellIdx, cell) => {
        if (cellIdx === 3) {
          const td = ce('td');
          td.style.cssText = 'padding:6px 8px;border-bottom:1px solid #eee;';
          const btnContainer = style(ce('div'), 'display:flex;gap:4px;flex-wrap:wrap;');
          const v = cell;
          const createXMLBtn = (label, xml, ext, color, hoverColor) => {
            if (!xml) return;
            const btn = txt(ce('button'), label);
            btn.style.cssText = `padding:4px 8px;background:${color};color:#fff;border:none;border-radius:3px;cursor:pointer;font-size:11px;`;
            btn.addEventListener('mouseenter', function(){ this.style.background = hoverColor; });
            btn.addEventListener('mouseleave', function(){ this.style.background = color; });
            btn.addEventListener('click', () => downloadXML(xml, `${(v.name || 'view').replace(/[^a-z0-9]/gi, '_')}_${v.id}.${ext}`));
            btnContainer.appendChild(btn);
          };
          createXMLBtn('📥 FetchXML', v.fetchxml, 'fetchxml', '#0078d4', '#106ebe');
          createXMLBtn('📋 LayoutXML', v.layoutxml, 'layoutxml', '#107c10', '#0e6b0d');
          if (btnContainer.children.length === 0) {
            btnContainer.appendChild(txt(ce('span'), '—'));
          }
          td.appendChild(btnContainer);
          return td;
        }
        return null;
      });
      if (ui.updateFooter) ui.updateFooter(`Views: ${views.length}`);
    };
    
    const loadEntities = async () => {
      entitySelect.appendChild(txt(ce('option'), 'Loading entities...'));
      try {
        const loadPaged = async (url) => {
          const j = await fetchJson(url);
          const arr = j && j.value ? j.value : [];
          entities = entities.concat(arr);
          const next = j['@odata.nextLink'] || j['odata.nextLink'];
          if (next) await loadPaged(next);
        };
        await loadPaged('/EntityDefinitions?$select=LogicalName,SchemaName,DisplayName');
        entities.sort((a, b) => {
          const da = (dispName(a.DisplayName) || a.LogicalName || '').toLowerCase();
          const db = (dispName(b.DisplayName) || b.LogicalName || '').toLowerCase();
          return da.localeCompare(db);
        });
        addLog(`Loaded ${entities.length} entities`, 'success');
        clearEl(entitySelect);
        entitySelect.appendChild(txt(ce('option'), 'Select entity...'));
        entities.forEach(e => {
          const opt = ce('option');
          opt.value = e.LogicalName;
          opt.textContent = (dispName(e.DisplayName) ? dispName(e.DisplayName) + ' ' : '') + '(' + e.LogicalName + ')';
          entitySelect.appendChild(opt);
        });
        let entityDropdownOverlay = null;
        const updateDropdown = () => {
          if (!entityDropdownOverlay) return;
          clearEl(entityDropdownOverlay);
          entities.forEach(e => {
            const displayName = dispName(e.DisplayName) || '';
            const logicalName = e.LogicalName || '';
            const row = style(ce('div'), 'padding:6px 8px;cursor:pointer;border-bottom:1px solid #eee;');
            row.textContent = (displayName ? displayName + ' ' : '') + '(' + logicalName + ')';
            row.onmouseenter = function(){ this.style.background = '#f3f2f1'; };
            row.onmouseleave = function(){ this.style.background = ''; };
            row.onclick = () => {
              entitySelect.value = logicalName;
              entitySelect.onchange();
              hideEntityDropdown();
            };
            entityDropdownOverlay.appendChild(row);
          });
          entityDropdownOverlay.style.display = 'block';
        };
        const showEntityDropdown = () => {
          if (entityDropdownOverlay) {
            updateDropdown();
            return;
          }
          const rect = entitySelect.getBoundingClientRect();
          entityDropdownOverlay = style(ce('div'), `position:fixed;top:${rect.bottom + 2}px;left:${rect.left}px;width:${rect.width}px;max-height:300px;overflow-y:auto;background:#fff;border:1px solid #ddd;border-radius:4px;box-shadow:0 4px 12px rgba(0,0,0,0.15);z-index:10000;display:none;`);
          document.body.appendChild(entityDropdownOverlay);
          updateDropdown();
        };
        const hideEntityDropdown = () => {
          if (entityDropdownOverlay) {
            if (entityDropdownOverlay.parentElement) entityDropdownOverlay.parentElement.removeChild(entityDropdownOverlay);
            entityDropdownOverlay = null;
          }
        };

        document.addEventListener('click', (e) => {
          if (!entitySelectDiv.contains(e.target) && (!entityDropdownOverlay || !entityDropdownOverlay.contains(e.target))) {
            hideEntityDropdown();
          }
        });
        entitySelect.onchange = async () => {
          currentEntity = entitySelect.value;
          if (!currentEntity) { clearEl(tableContainer); entityMetadata = null; relationships = null; forms = []; views = []; fields = []; businessRules = []; plugins = []; javascript = []; processes = []; envVars = []; formAutomationsCache = {}; if (ui.updateFooter) ui.updateFooter('Select entity...'); return; }
          addLog(`Entity selected: ${currentEntity}`, 'info');
          entityMetadata = null; relationships = null; forms = []; views = []; fields = []; businessRules = []; plugins = []; javascript = []; processes = []; envVars = []; formAutomationsCache = {};
          if (ui.updateFooter) ui.updateFooter(`Loading ${currentEntity}...`);
          await renderTab(currentTab);
        };
        if (ui.updateFooter) ui.updateFooter(`Loaded ${entities.length} entities`);
        addLog(`Loaded ${entities.length} entities`, 'success');
      } catch (ex) {
        addLog(`Error loading entities: ${ex.message || ex}`, 'error');
        clearEl(entitySelect);
        entitySelect.appendChild(txt(ce('option'), 'Error loading entities'));
        if (ui.updateFooter) ui.updateFooter('Error loading entities');
      }
    };
    
    const loadBusinessRules = async () => {
      if (businessRules.length === 0) {
        try {
          const rulesData = await fetchJson(`/workflows?$select=workflowid,name,statecode&$filter=primaryentity eq '${currentEntity}' and category eq 2`);
          businessRules = (rulesData.value || []).map(r => ({ id: r.workflowid, name: r.name || '', status: r.statecode === 0 ? 'Active' : 'Inactive' }));
        } catch (e) {
          console.warn('Error loading business rules:', e);
          businessRules = [];
        }
      }
      clearEl(tableContainer);
      renderTable(['Name', 'ID', 'Status'], businessRules.map(r => [r.name, r.id, r.status]), true);
      if (ui.updateFooter) ui.updateFooter(`Business Rules: ${businessRules.length}`);
    };
    
    const loadPlugins = async () => {
      if (plugins.length === 0) {
        try {
          const stepsData = await fetchJson(`/sdkmessageprocessingsteps?$select=sdkmessageprocessingstepid,name,stage,mode&$expand=sdkmessageid($select=name),sdkmessagefilterid($select=primaryobjecttypecode)&$orderby=name&$top=1000`);
          const filteredSteps = (stepsData.value || []).filter(s => {
            const filterEntity = s.sdkmessagefilterid?.primaryobjecttypecode;
            return !filterEntity || filterEntity === currentEntity;
          });
          plugins = filteredSteps.map(s => {
            const stageMap = { 10: 'Pre-Validation', 20: 'Pre-Operation', 40: 'Post-Operation' };
            const modeMap = { 0: 'Synchronous', 1: 'Asynchronous' };
            return {
              id: s.sdkmessageprocessingstepid,
              name: s.name || '',
              stage: stageMap[s.stage] || s.stage || '',
              mode: modeMap[s.mode] || s.mode || '',
              plugin: '',
              image: '',
              message: s.sdkmessageid?.name || ''
            };
          });
        } catch (e) {
          console.warn('Error loading plugins:', e);
          plugins = [];
        }
      }
      clearEl(tableContainer);
      const headers = ['Name', 'Stage', 'Mode', 'Message'];
      const rows = plugins.map(p => [p.name, p.stage, p.mode, p.message || '']);
      renderTable(headers, rows, true);
      if (ui.updateFooter) ui.updateFooter(`Plugins: ${plugins.length}`);
    };
    
    const loadJavaScript = async () => {
      if (javascript.length === 0) {
        try {
          const formsData = await fetchJson(`/systemforms?$select=formid,name,formxml&$filter=objecttypecode eq '${currentEntity}' and type eq 2`);
          const jsResources = new Set();
          const handlers = [];
          (formsData.value || []).forEach(form => {
            try {
              const doc = new DOMParser().parseFromString(form.formxml || '', 'text/xml');
              if (doc.getElementsByTagName('parsererror').length > 0) return;
              Array.from(doc.getElementsByTagName('Library')).forEach(lib => {
                const name = lib.getAttribute('name') || '';
                if (name) jsResources.add(name);
              });
              Array.from(doc.getElementsByTagName('event')).forEach(ev => {
                const evName = ev.getAttribute('name') || '';
                let target = 'Form';
                const targetId = ev.getAttribute('attribute') || '';
                if (targetId) target = targetId;
                Array.from(ev.getElementsByTagName('handler')).forEach(hn => {
                  handlers.push({
                    form: form.name || '',
                    target: target,
                    event: evName,
                    library: hn.getAttribute('libraryName') || '',
                    functionName: hn.getAttribute('functionName') || '',
                    enabled: hn.getAttribute('enabled') !== 'false' ? 'Yes' : 'No'
                  });
                });
              });
            } catch (e) {}
          });
          javascript = handlers.length > 0 ? handlers : Array.from(jsResources).map(lib => ({ form: '', target: '', event: '', library: lib, functionName: '', enabled: '' }));
        } catch (e) {
          console.warn('Error loading JavaScript:', e);
          javascript = [];
        }
      }
      clearEl(tableContainer);
      if (javascript.length === 0) {
        tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;text-align:center;color:#605e5c;'), 'No JavaScript automations found'));
        if (ui.updateFooter) ui.updateFooter('JavaScript: 0');
        return;
      }
      
      if (javascript[0].form) {
        const groupedByForm = {};
        javascript.forEach(j => {
          const formName = j.form || '(Unknown Form)';
          if (!groupedByForm[formName]) groupedByForm[formName] = {};
          const eventName = j.event || '(Unknown Event)';
          if (!groupedByForm[formName][eventName]) groupedByForm[formName][eventName] = [];
          groupedByForm[formName][eventName].push(j);
        });
        
        const rows = [];
        Object.keys(groupedByForm).sort().forEach(formName => {
          Object.keys(groupedByForm[formName]).sort().forEach(eventName => {
            groupedByForm[formName][eventName].forEach(j => {
              const targetDisplay = j.target === 'Form' ? 'Form' : `Field: ${j.target}`;
              rows.push([
                formName,
                eventName,
                targetDisplay,
                j.library || '',
                j.functionName || '',
                j.enabled || ''
              ]);
            });
          });
        });
        
        const headers = ['Form', 'Event', 'Target', 'Library', 'Function', 'Enabled'];
        renderTable(headers, rows, true);
        if (ui.updateFooter) ui.updateFooter(`JavaScript: ${javascript.length} handlers across ${Object.keys(groupedByForm).length} form(s)`);
      } else {
        const headers = ['Library Name'];
        const rows = javascript.map(j => [j.library || '']);
        renderTable(headers, rows, true);
        if (ui.updateFooter) ui.updateFooter(`JavaScript: ${javascript.length} libraries`);
      }
    };
    
    const loadProcesses = async () => {
      if (processes.length === 0) {
        try {
          const processesData = await fetchJson(`/workflows?$select=workflowid,name,statecode,category&$filter=primaryentity eq '${currentEntity}' and category eq 4`);
          processes = (processesData.value || []).map(p => ({ id: p.workflowid, name: p.name || '', status: p.statecode === 0 ? 'Active' : 'Inactive' }));
        } catch (e) {
          console.warn('Error loading processes:', e);
          processes = [];
        }
      }
      clearEl(tableContainer);
      renderTable(['Name', 'ID', 'Status'], processes.map(p => [p.name, p.id, p.status]), true);
      if (ui.updateFooter) ui.updateFooter(`Business Process Flows: ${processes.length}`);
    };
    
    const loadEnvironmentVariables = async () => {
      if (envVars.length === 0) {
        try {
          const envVarsData = await fetchJson(`/environmentvariabledefinitions?$select=environmentvariabledefinitionid,schemaname,displayname,type,defaultvalue,description&$expand=environmentvariabledefinition_environmentvariablevalue($select=environmentvariablevalueid,value)&$orderby=displayname`);
          envVars = (envVarsData.value || []).map(v => ({
            definitionId: v.environmentvariabledefinitionid,
            valueId: v.environmentvariabledefinition_environmentvariablevalue?.[0]?.environmentvariablevalueid || null,
            schemaName: v.schemaname || '',
            displayName: v.displayname || v.schemaname || '',
            type: v.type === 100000000 ? 'String' : v.type === 100000001 ? 'Number' : v.type === 100000002 ? 'Boolean' : v.type === 100000003 ? 'JSON' : 'Unknown',
            defaultValue: v.defaultvalue || '—',
            currentValue: v.environmentvariabledefinition_environmentvariablevalue?.[0]?.value || '(not set)',
            description: v.description || ''
          }));
        } catch (e) {
          console.warn('Error loading environment variables:', e);
          envVars = [];
        }
      }
      clearEl(tableContainer);
      const buttonBar = style(ce('div'), 'display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;');
      const createBtn = txt(ce('button'), '➕ Create Variable');
      createBtn.style.cssText = 'padding:8px 16px;background:#0078d4;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;';
      createBtn.addEventListener('click', () => showCreateEnvVarDialog());
      buttonBar.appendChild(createBtn);
      const refreshBtn = txt(ce('button'), '🔄 Refresh');
      refreshBtn.style.cssText = 'padding:8px 16px;background:#605e5c;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;';
      refreshBtn.addEventListener('click', async () => { envVars = []; await loadEnvironmentVariables(); });
      buttonBar.appendChild(refreshBtn);
      tableContainer.appendChild(buttonBar);
      
      if (envVars.length === 0) {
        tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;text-align:center;color:#605e5c;'), 'No environment variables found'));
        if (ui.updateFooter) ui.updateFooter('Environment Variables: 0');
        return;
      }
      
      const headers = ['Display Name', 'Schema Name', 'Type', 'Default Value', 'Current Value', 'Description', 'Actions'];
      const rows = envVars.map(v => [v.displayName, v.schemaName, v.type, v.defaultValue, v.currentValue, v.description, v]);
      renderTable(headers, rows, true, (row, cellIdx, cell) => {
        if (cellIdx === 6) {
          const td = ce('td');
          td.style.cssText = 'padding:6px 8px;border-bottom:1px solid #eee;';
          const btnContainer = style(ce('div'), 'display:flex;gap:4px;flex-wrap:wrap;');
          const v = cell;
          const editBtn = txt(ce('button'), '✏️ Edit');
          editBtn.style.cssText = 'padding:4px 8px;background:#0078d4;color:#fff;border:none;border-radius:3px;cursor:pointer;font-size:11px;';
          editBtn.addEventListener('mouseenter', function(){ this.style.background = '#106ebe'; });
          editBtn.addEventListener('mouseleave', function(){ this.style.background = '#0078d4'; });
          editBtn.addEventListener('click', () => showEditEnvVarDialog(v));
          btnContainer.appendChild(editBtn);
          td.appendChild(btnContainer);
          return td;
        }
        return null;
      });
      if (ui.updateFooter) ui.updateFooter(`Environment Variables: ${envVars.length}`);
    };
    
    const showEditEnvVarDialog = (envVar) => {
      const dialog = style(ce('div'), 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:1000000;display:flex;align-items:center;justify-content:center;');
      const content = style(ce('div'), 'background:#fff;border-radius:8px;padding:24px;max-width:500px;width:90%;max-height:90vh;overflow-y:auto;');
      const title = txt(style(ce('h3'), 'margin:0 0 16px 0;font-size:18px;color:#323130;'), 'Edit Environment Variable');
      content.appendChild(title);
      
      const form = style(ce('div'), 'display:flex;flex-direction:column;gap:12px;');
      
      const createField = (label, value, readonly = false) => {
        const wrapper = style(ce('div'), 'display:flex;flex-direction:column;gap:4px;');
        const lbl = txt(style(ce('label'), 'font-weight:600;font-size:12px;color:#323130;'), label);
        wrapper.appendChild(lbl);
        const input = ce(readonly ? 'div' : 'input');
        if (readonly) {
          input.style.cssText = 'padding:8px;border:1px solid #ddd;border-radius:4px;background:#f3f2f1;color:#605e5c;font-size:12px;';
          input.textContent = value;
        } else {
          input.type = 'text';
          input.value = value || '';
          input.style.cssText = 'padding:8px;border:1px solid #ddd;border-radius:4px;font-size:12px;';
        }
        wrapper.appendChild(input);
        return { wrapper, input };
      };
      
      const displayNameField = createField('Display Name', envVar.displayName, true);
      form.appendChild(displayNameField.wrapper);
      const schemaNameField = createField('Schema Name', envVar.schemaName, true);
      form.appendChild(schemaNameField.wrapper);
      const typeField = createField('Type', envVar.type, true);
      form.appendChild(typeField.wrapper);
      const defaultValueField = createField('Default Value', envVar.defaultValue);
      form.appendChild(defaultValueField.wrapper);
      const currentValueField = createField('Current Value', envVar.currentValue);
      form.appendChild(currentValueField.wrapper);
      const descriptionField = createField('Description', envVar.description);
      form.appendChild(descriptionField.wrapper);
      
      content.appendChild(form);
      
      const buttonBar = style(ce('div'), 'display:flex;gap:8px;justify-content:flex-end;margin-top:16px;');
      const cancelBtn = txt(ce('button'), 'Cancel');
      cancelBtn.style.cssText = 'padding:8px 16px;background:#605e5c;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;';
      cancelBtn.addEventListener('click', () => document.body.removeChild(dialog));
      buttonBar.appendChild(cancelBtn);
      
      const saveBtn = txt(ce('button'), 'Save');
      saveBtn.style.cssText = 'padding:8px 16px;background:#0078d4;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;';
      saveBtn.addEventListener('click', async () => {
        try {
          saveBtn.disabled = true;
          saveBtn.textContent = 'Saving...';
          if (defaultValueField.input.value !== envVar.defaultValue) {
            await patchJson(`/environmentvariabledefinitions(${envVar.definitionId})`, { defaultvalue: defaultValueField.input.value || null });
          }
          if (currentValueField.input.value !== envVar.currentValue) {
            if (envVar.valueId) {
              await patchJson(`/environmentvariablevalues(${envVar.valueId})`, { value: currentValueField.input.value || null });
            } else {
              await postJson('/environmentvariablevalues', {
                value: currentValueField.input.value || null,
                schemaname: envVar.schemaName,
                'EnvironmentVariableDefinitionId@odata.bind': `/environmentvariabledefinitions(${envVar.definitionId})`
              });
            }
          }
          if (descriptionField.input.value !== envVar.description) {
            await patchJson(`/environmentvariabledefinitions(${envVar.definitionId})`, { description: descriptionField.input.value || null });
          }
          document.body.removeChild(dialog);
          envVars = [];
          await loadEnvironmentVariables();
        } catch (e) {
          alert('Error saving: ' + (e.message || e));
          saveBtn.disabled = false;
          saveBtn.textContent = 'Save';
        }
      });
      buttonBar.appendChild(saveBtn);
      content.appendChild(buttonBar);
      
      dialog.appendChild(content);
      document.body.appendChild(dialog);
      dialog.addEventListener('click', (e) => { if (e.target === dialog) document.body.removeChild(dialog); });
    };
    
    const showCreateEnvVarDialog = () => {
      const dialog = style(ce('div'), 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:1000000;display:flex;align-items:center;justify-content:center;');
      const content = style(ce('div'), 'background:#fff;border-radius:8px;padding:24px;max-width:500px;width:90%;max-height:90vh;overflow-y:auto;');
      const title = txt(style(ce('h3'), 'margin:0 0 16px 0;font-size:18px;color:#323130;'), 'Create Environment Variable');
      content.appendChild(title);
      
      const form = style(ce('div'), 'display:flex;flex-direction:column;gap:12px;');
      
      const createField = (label, type = 'text') => {
        const wrapper = style(ce('div'), 'display:flex;flex-direction:column;gap:4px;');
        const lbl = txt(style(ce('label'), 'font-weight:600;font-size:12px;color:#323130;'), label);
        wrapper.appendChild(lbl);
        const input = type === 'select' ? ce('select') : ce('input');
        if (type === 'select') {
          input.style.cssText = 'padding:8px;border:1px solid #ddd;border-radius:4px;font-size:12px;';
          ['String', 'Number', 'Boolean', 'JSON'].forEach(opt => {
            const option = ce('option');
            option.value = opt;
            option.textContent = opt;
            input.appendChild(option);
          });
        } else {
          input.type = type;
          input.style.cssText = 'padding:8px;border:1px solid #ddd;border-radius:4px;font-size:12px;';
        }
        wrapper.appendChild(input);
        return { wrapper, input };
      };
      
      const displayNameField = createField('Display Name *');
      form.appendChild(displayNameField.wrapper);
      const schemaNameField = createField('Schema Name *');
      form.appendChild(schemaNameField.wrapper);
      const typeField = createField('Type *', 'select');
      form.appendChild(typeField.wrapper);
      const defaultValueField = createField('Default Value');
      form.appendChild(defaultValueField.wrapper);
      const currentValueField = createField('Current Value');
      form.appendChild(currentValueField.wrapper);
      const descriptionField = createField('Description');
      form.appendChild(descriptionField.wrapper);
      
      content.appendChild(form);
      
      const buttonBar = style(ce('div'), 'display:flex;gap:8px;justify-content:flex-end;margin-top:16px;');
      const cancelBtn = txt(ce('button'), 'Cancel');
      cancelBtn.style.cssText = 'padding:8px 16px;background:#605e5c;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;';
      cancelBtn.addEventListener('click', () => document.body.removeChild(dialog));
      buttonBar.appendChild(cancelBtn);
      
      const saveBtn = txt(ce('button'), 'Create');
      saveBtn.style.cssText = 'padding:8px 16px;background:#0078d4;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;';
      saveBtn.addEventListener('click', async () => {
        try {
          if (!displayNameField.input.value || !schemaNameField.input.value) {
            alert('Display Name and Schema Name are required');
            return;
          }
          saveBtn.disabled = true;
          saveBtn.textContent = 'Creating...';
          const typeMap = { 'String': 100000000, 'Number': 100000001, 'Boolean': 100000002, 'JSON': 100000003 };
          const defPayload = {
            displayname: displayNameField.input.value,
            schemaname: schemaNameField.input.value,
            type: typeMap[typeField.input.value] || 100000000,
            description: descriptionField.input.value || null
          };
          if (defaultValueField.input.value) defPayload.defaultvalue = defaultValueField.input.value;
          const defResult = await postJson('/environmentvariabledefinitions', defPayload);
          const definitionId = defResult.environmentvariabledefinitionid || defResult.id;
          if (currentValueField.input.value) {
            await postJson('/environmentvariablevalues', {
              value: currentValueField.input.value,
              schemaname: schemaNameField.input.value,
              'EnvironmentVariableDefinitionId@odata.bind': `/environmentvariabledefinitions(${definitionId})`
            });
          }
          document.body.removeChild(dialog);
          envVars = [];
          await loadEnvironmentVariables();
        } catch (e) {
          alert('Error creating: ' + (e.message || e));
          saveBtn.disabled = false;
          saveBtn.textContent = 'Create';
        }
      });
      buttonBar.appendChild(saveBtn);
      content.appendChild(buttonBar);
      
      dialog.appendChild(content);
      document.body.appendChild(dialog);
      dialog.addEventListener('click', (e) => { if (e.target === dialog) document.body.removeChild(dialog); });
    };
    
    const loadGlobalChoices = async () => {
      clearEl(tableContainer);
      tableContainer.appendChild(txt(style(ce('div'), 'padding:8px;color:#605e5c;font-size:12px;'), 'Loading global choices...'));
      
      try {
        const globalChoicesData = await fetchJson(`/GlobalOptionSetDefinitions`);
        const choices = (globalChoicesData.value || []).map(c => ({
          name: c.Name || '',
          displayName: c.DisplayName?.UserLocalizedLabel?.Label || c.Name || '',
          type: c.OptionSetType === 'Picklist' ? 'Picklist' : c.OptionSetType === 'State' ? 'State' : c.OptionSetType === 'Status' ? 'Status' : c.OptionSetType === 'Boolean' ? 'Boolean' : 'Unknown',
          metadataId: c.MetadataId || ''
        }));
        
        clearEl(tableContainer);
        
        if (choices.length === 0) {
          tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;text-align:center;color:#605e5c;'), 'No global choices found'));
          if (ui.updateFooter) ui.updateFooter('Global Choices: 0');
          return;
        }
        
        const headers = ['Display Name', 'Name', 'Type', 'Actions'];
        const rows = choices.map(c => [c.displayName, c.name, c.type, c]);
        renderTable(headers, rows, true, (row, cellIdx, cell) => {
          if (cellIdx === 3) {
            const td = ce('td');
            td.style.cssText = 'padding:6px 8px;border-bottom:1px solid #eee;';
            const choice = cell;
            const viewBtn = txt(ce('button'), '👁️ View Details');
            viewBtn.style.cssText = 'padding:4px 8px;background:#0078d4;color:#fff;border:none;border-radius:3px;cursor:pointer;font-size:11px;';
            viewBtn.addEventListener('mouseenter', function(){ this.style.background = '#106ebe'; });
            viewBtn.addEventListener('mouseleave', function(){ this.style.background = '#0078d4'; });
            viewBtn.addEventListener('click', async () => {
              try {
                const sq = String.fromCharCode(39);
                const detailsData = await fetchJson(`/GlobalOptionSetDefinitions(Name=${sq}${choice.name}${sq})/Microsoft.Dynamics.CRM.OptionSetMetadata`);
                const dialog = style(ce('div'), 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:1000000;display:flex;align-items:center;justify-content:center;');
                const content = style(ce('div'), 'background:#fff;border-radius:8px;padding:24px;max-width:700px;width:90%;max-height:90vh;overflow-y:auto;');
                const title = txt(style(ce('h3'), 'margin:0 0 16px 0;font-size:18px;color:#323130;'), 'Global Choice: ' + choice.displayName);
                content.appendChild(title);
                
                const info = style(ce('div'), 'display:grid;grid-template-columns:150px 1fr;gap:8px;margin-bottom:16px;font-size:12px;');
                const addRow = (label, value) => {
                  info.appendChild(txt(style(ce('div'), 'font-weight:600;color:#323130;'), label + ':'));
                  info.appendChild(txt(style(ce('div'), 'color:#605e5c;'), value || '—'));
                };
                addRow('Name', detailsData.Name);
                addRow('Display Name', detailsData.DisplayName?.UserLocalizedLabel?.Label);
                addRow('Description', detailsData.Description?.UserLocalizedLabel?.Label);
                addRow('Is Customizable', detailsData.IsCustomizable?.Value ? 'Yes' : 'No');
                content.appendChild(info);
                
                if (detailsData.Options && detailsData.Options.length > 0) {
                  const optionsTitle = txt(style(ce('h4'), 'margin:16px 0 8px 0;font-size:14px;color:#323130;'), 'Options (' + detailsData.Options.length + ')');
                  content.appendChild(optionsTitle);
                  const optionsTable = style(ce('table'), 'width:100%;border-collapse:collapse;font-size:12px;');
                  const thead = ce('thead');
                  const headerRow = ce('tr');
                  ['Value', 'Label', 'Color'].forEach(h => {
                    const th = txt(style(ce('th'), 'padding:8px;background:#f3f2f1;text-align:left;font-weight:600;border-bottom:2px solid #ddd;'), h);
                    headerRow.appendChild(th);
                  });
                  thead.appendChild(headerRow);
                  optionsTable.appendChild(thead);
                  const tbody = ce('tbody');
                  detailsData.Options.forEach(opt => {
                    const tr = ce('tr');
                    tr.style.cssText = 'border-bottom:1px solid #eee;';
                    const tdValue = txt(style(ce('td'), 'padding:8px;'), String(opt.Value));
                    const tdLabel = txt(style(ce('td'), 'padding:8px;'), opt.Label?.UserLocalizedLabel?.Label || '—');
                    const tdColor = style(ce('td'), 'padding:8px;');
                    if (opt.Color) {
                      const colorBox = style(ce('div'), `display:inline-block;width:16px;height:16px;background:${opt.Color};border:1px solid #ddd;border-radius:2px;margin-right:8px;vertical-align:middle;`);
                      tdColor.appendChild(colorBox);
                      tdColor.appendChild(txt(ce('span'), opt.Color));
                    } else {
                      tdColor.textContent = '—';
                    }
                    tr.appendChild(tdValue);
                    tr.appendChild(tdLabel);
                    tr.appendChild(tdColor);
                    tbody.appendChild(tr);
                  });
                  optionsTable.appendChild(tbody);
                  content.appendChild(optionsTable);
                }
                
                const closeBtn = txt(ce('button'), 'Close');
                closeBtn.style.cssText = 'margin-top:16px;padding:8px 16px;background:#605e5c;color:#fff;border:none;border-radius:4px;cursor:pointer;';
                closeBtn.addEventListener('click', () => document.body.removeChild(dialog));
                content.appendChild(closeBtn);
                
                dialog.appendChild(content);
                document.body.appendChild(dialog);
                dialog.addEventListener('click', (e) => { if (e.target === dialog) document.body.removeChild(dialog); });
              } catch (e) {
                alert('Error loading details: ' + (e.message || e));
              }
            });
            td.appendChild(viewBtn);
            return td;
          }
          return null;
        });
        if (ui.updateFooter) ui.updateFooter(`Global Choices: ${choices.length}`);
        addLog(`Loaded ${choices.length} global choices`, 'success');
      } catch (ex) {
        addLog(`Error loading global choices: ${ex.message || ex}`, 'error');
        clearEl(tableContainer);
        tableContainer.appendChild(txt(style(ce('div'), 'padding:20px;color:#a80000;'), 'Error: ' + (ex.message || ex)));
      }
    };
    
    // Add settings and log buttons to footer
    if (ui.footer) {
      const settingsBtn = txt(ce('button'), String.fromCharCode(9881) + ' Settings');
      settingsBtn.title = 'Customize color scheme';
      settingsBtn.style.cssText = 'margin-right:8px;padding:4px 8px;font-size:12px;background:transparent;border:1px solid currentColor;border-radius:4px;cursor:pointer;color:inherit;';
      settingsBtn.onclick = showSettings;
      ui.footer.insertBefore(settingsBtn, ui.footer.lastChild);
      
      const logBtn = txt(ce('button'), String.fromCharCode(128220) + ' Logs');
      logBtn.title = 'View activity logs';
      logBtn.style.cssText = 'margin-right:8px;padding:4px 8px;font-size:12px;background:transparent;border:1px solid currentColor;border-radius:4px;cursor:pointer;color:inherit;';
      logBtn.onclick = showLogViewer;
      ui.footer.insertBefore(logBtn, ui.footer.lastChild);
    }
    
    await loadEntities();
    await renderTab('Properties');
  } catch (ex) {
    console.error(ex);
    if (typeof addLog !== 'undefined') addLog(`Critical error: ${ex.message || ex}`, 'error');
    const err = ce('div');
    err.style.cssText = 'color:#a80000;margin-top:8px;padding:12px;';
    err.textContent = 'Error: ' + (ex.message || ex);
    ui.body.appendChild(err);
  }
})();

