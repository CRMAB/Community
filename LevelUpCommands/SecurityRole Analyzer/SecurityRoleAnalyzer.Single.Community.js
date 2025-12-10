// Security Role Analyzer Community - Single User (v1.2 - Enhanced with Settings & Logs)
(async function securityRoleAnalyzerSingle() {
  const overlayHost = window.__lvlUpOverlay || (function loadOverlayFallback() {
    alert('Overlay helpers not loaded');
    return {};
  })();
  const requiredHelpers = ['createOverlay', 'ce', 'txt', 'style', 'clearEl', 'chunk', 'sanitizeGuid', 'guidPathValue', 'guidLiteral', 'escapeODataValue', 'downloadFile'];
  const missingHelpers = requiredHelpers.filter((key) => typeof overlayHost[key] !== 'function');
  if (missingHelpers.length) {
    alert('Overlay helpers missing: ' + missingHelpers.join(', ') + '. Please load helperOverlayV2.js first.');
    return;
  }
  const createOverlay = overlayHost.createOverlay;
  const ce = overlayHost.ce;
  const txt = overlayHost.txt;
  const style = overlayHost.style;
  const clearEl = overlayHost.clearEl;
  const chunk = overlayHost.chunk;
  const sanitizeGuid = overlayHost.sanitizeGuid;
  const guidPathValue = overlayHost.guidPathValue;
  const guidLiteral = overlayHost.guidLiteral;
  const escapeODataValue = overlayHost.escapeODataValue;
  const downloadFile = overlayHost.downloadFile;

  // Color scheme system
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
      bgHover: '#f3f2f1',
      bgError: '#fde7e9',
      textOnPrimary: '#fff',
      success: '#107c10',
      successHover: '#0e6b0e',
      danger: '#d13438',
      dangerHover: '#a80000'
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
      dangerHover: '#dc2626'
    }
  };
  let customColors = null;
  try {
    const stored = localStorage.getItem('crmab-securityroleanalyzer-colors');
    if (stored) customColors = JSON.parse(stored);
  } catch (ex) {}
  
  const ui = createOverlay('Security Role Analyzer - Single User', {
    rootId: 'crmab-securityroleanalyzer-single',
    width: 1000,
    maxHeight: '95vh',
    footer: true,
    footerText: 'Ready',
  });
  if (!ui || !ui.body) return;
  
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

  ui.body.style.minHeight = '700px';
  ui.body.style.maxHeight = '90vh';
  ui.body.style.overflowY = 'auto';
  ui.body.style.paddingBottom = '20px';

  const c = getColors();

  try {
    const gc = Xrm.Utility.getGlobalContext();
    const apiBase = gc.getClientUrl() + '/api/data/v9.2';
    const orgUrl = gc.getClientUrl();
    const HTTP_HEADERS = {};
    HTTP_HEADERS['OData-MaxVersion'] = '4.0';
    HTTP_HEADERS['OData-Version'] = '4.0';
    HTTP_HEADERS['Accept'] = 'application/json';
    HTTP_HEADERS['Content-Type'] = 'application/json';
    const SQUOTE = String.fromCharCode(39);
    
    // Environment ID cache
    let cachedEnvironmentId = null;
    
    // Helper to get environment ID from organization table
    const getEnvironmentId = async () => {
      if (cachedEnvironmentId) return cachedEnvironmentId;
      
      try {
        // Method 1: Query organization table for organizationid
        const orgData = await fetchJson('/organizations?$select=organizationid');
        if (orgData.value && orgData.value.length > 0) {
          cachedEnvironmentId = orgData.value[0].organizationid;
          addLog('Environment ID retrieved from organization table', 'info');
          return cachedEnvironmentId;
        }
      } catch (e) {
        console.warn('Failed to retrieve environment ID from organization table:', e);
      }
      
      try {
        // Method 2: Parse from URL hostname
        const url = new URL(orgUrl);
        const envMatch = url.hostname.match(/([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i);
        if (envMatch) {
          cachedEnvironmentId = envMatch[1];
          addLog('Environment ID extracted from URL', 'info');
          return cachedEnvironmentId;
        }
      } catch (e) {
        console.warn('Failed to extract environment ID from URL:', e);
      }
      
      addLog('Could not determine environment ID', 'warning');
      return null;
    };
    
    // Helper to create clickable links
    const createRecordLink = (text, entityName, recordId, tooltip) => {
      const link = txt(ce('a'), text);
      link.href = `${orgUrl}/main.aspx?etn=${entityName}&id=${recordId}&newWindow=true&pagetype=entityrecord`;
      link.target = '_blank';
      link.title = tooltip || 'Open record';
      link.style.cssText = `color:${c.primary};text-decoration:none;cursor:pointer;`;
      link.addEventListener('mouseenter', function() { this.style.textDecoration = 'underline'; });
      link.addEventListener('mouseleave', function() { this.style.textDecoration = 'none'; });
      return link;
    };
    
    const createExternalLink = (text, url, tooltip) => {
      const link = txt(ce('a'), text);
      link.href = url;
      link.target = '_blank';
      link.title = tooltip || 'Open link';
      link.style.cssText = `color:${c.primary};text-decoration:none;cursor:pointer;`;
      link.addEventListener('mouseenter', function() { this.style.textDecoration = 'underline'; });
      link.addEventListener('mouseleave', function() { this.style.textDecoration = 'none'; });
      return link;
    };

    const fetchJson = async (path) => {
      const fetchOpts = {};
      fetchOpts.headers = HTTP_HEADERS;
      const response = await fetch(apiBase + path, fetchOpts);
      if (!response.ok) throw new Error('HTTP ' + response.status);
      return response.json();
    };

    const fetchWithFallback = async (primaryFactory, fallbackFactory) => {
      try {
        return await primaryFactory();
      } catch (primaryError) {
        if (!fallbackFactory) throw primaryError;
        return fallbackFactory();
      }
    };

    const showMessage = (container, message, color = null) => {
      clearEl(container);
      const msgColor = color || c.textMuted;
      container.appendChild(txt(style(ce('div'), 'padding:12px;color:' + msgColor + ';'), message));
    };

    const setFooter = (text) => { 
      if (ui.updateFooter) ui.updateFooter(text);
      addLog(text, 'info');
    };

    let allUsers = [];
    let filteredUsers = [];
    let allViews = [];
    let selectedViewFetchXml = null;

    const dropdownsContainer = style(ce('div'), 'display:flex;gap:12px;margin-bottom:12px;');
    
    const dropdownWrapper = style(ce('div'), 'position:relative;flex:2;');
    const filterInput = style(ce('input'), `width:100%;padding:8px;border:1px solid ${c.border};border-radius:4px;box-sizing:border-box;background:${c.bg};color:${c.text};`);
    filterInput.type = 'text';
    filterInput.placeholder = 'Type to filter users...';
    const userList = style(ce('div'), `display:none;position:absolute;top:100%;left:0;right:0;min-height:350px;max-height:650px;overflow-y:auto;border:1px solid ${c.border};border-radius:4px;background:${c.bg};z-index:1000;box-shadow:0 4px 8px rgba(0,0,0,0.1);`);
    dropdownWrapper.appendChild(filterInput);
    dropdownWrapper.appendChild(userList);
    
    const viewWrapper = style(ce('div'), 'position:relative;flex:1;');
    const viewSelect = style(ce('select'), `width:100%;padding:8px;border:1px solid ${c.border};border-radius:4px;background:${c.bg};color:${c.text};cursor:pointer;`);
    const allUsersOption = txt(ce('option'), 'All Users');
    allUsersOption.value = '';
    viewSelect.appendChild(allUsersOption);
    viewWrapper.appendChild(viewSelect);
    
    dropdownsContainer.appendChild(dropdownWrapper);
    dropdownsContainer.appendChild(viewWrapper);
    
    // Quick Links Section
    const quickLinksDiv = style(ce('div'), 'margin-bottom:12px;display:flex;gap:8px;flex-wrap:wrap;');
    
    const createQuickLinkBtn = (text, url, icon) => {
      const btn = txt(ce('button'), icon + ' ' + text);
      btn.style.cssText = `padding:6px 12px;background:transparent;color:${c.text};border:1px solid ${c.border};border-radius:4px;cursor:pointer;font-size:11px;font-weight:600;`;
      btn.addEventListener('mouseenter', function() { this.style.background = c.bgHover; });
      btn.addEventListener('mouseleave', function() { this.style.background = 'transparent'; });
      btn.addEventListener('click', () => {
        window.open(url, '_blank');
        addLog('Opened: ' + text, 'info');
      });
      return btn;
    };
    
    // Initialize quick links (environment-specific links will be added after env ID is retrieved)
    quickLinksDiv.appendChild(createQuickLinkBtn('All Security Roles', `${orgUrl}/main.aspx?pagetype=entitylist&etn=role`, '📋'));
    quickLinksDiv.appendChild(createQuickLinkBtn('Advanced Find: Users', `${orgUrl}/main.aspx?pagetype=advancedfind&etn=systemuser`, '🔍'));
    
    // Add admin links once environment ID is available
    getEnvironmentId().then(envId => {
      if (envId) {
        quickLinksDiv.insertBefore(createQuickLinkBtn('Admin: Users', `https://admin.powerplatform.microsoft.com/environments/${envId}/settings/users`, '👥'), quickLinksDiv.firstChild);
        quickLinksDiv.insertBefore(createQuickLinkBtn('Admin: Security Roles', `https://admin.powerplatform.microsoft.com/environments/${envId}/settings/security/roles`, '🔐'), quickLinksDiv.firstChild);
      }
    }).catch(e => console.warn('Failed to add admin links:', e));
    
    const buttonDiv = style(ce('div'), 'margin-bottom:16px;display:flex;gap:8px;');
    const refreshBtn = txt(ce('button'), 'Reload Users');
    refreshBtn.style.cssText = `padding:8px 16px;background:${c.primary};color:${c.textOnPrimary};border:none;border-radius:4px;cursor:pointer;flex:1;`;
    refreshBtn.addEventListener('mouseenter', function(){ this.style.background=c.primaryHover; });
    refreshBtn.addEventListener('mouseleave', function(){ this.style.background=c.primary; });
    const analyzeBtn = txt(ce('button'), 'Analyze Selected User');
    analyzeBtn.style.cssText = `padding:8px 16px;background:${c.success};color:${c.textOnPrimary};border:none;border-radius:4px;cursor:pointer;flex:2;`;
    analyzeBtn.addEventListener('mouseenter', function(){ if(!this.disabled) this.style.background=c.successHover; });
    analyzeBtn.addEventListener('mouseleave', function(){ if(!this.disabled) this.style.background=c.success; });
    analyzeBtn.disabled = true;
    buttonDiv.appendChild(refreshBtn);
    buttonDiv.appendChild(analyzeBtn);
    const resultDiv = style(ce('div'), 'margin-top:16px;');
    ui.body.appendChild(dropdownsContainer);
    ui.body.appendChild(buttonDiv);
    ui.body.appendChild(resultDiv);

    let selectedUser = null;

    const showDropdown = () => {
      userList.style.display = 'block';
    };

    const hideDropdown = () => {
      userList.style.display = 'none';
    };

    const populateList = (users) => {
      clearEl(userList);
      if (users.length === 0) {
        const noResult = txt(style(ce('div'), `padding:12px;color:${c.textMuted};`), 'No users found');
        userList.appendChild(noResult);
        return;
      }
      users.forEach((user) => {
        const item = style(ce('div'), `padding:8px 12px;cursor:pointer;border-bottom:1px solid ${c.borderLight};color:${c.text};`);
        const displayName = (user.fullname || '') + (user.domainname ? ' (' + user.domainname + ')' : '');
        item.textContent = displayName;
        item.addEventListener('mouseenter', () => { item.style.background = c.bgHover; });
        item.addEventListener('mouseleave', () => { item.style.background = ''; });
        item.addEventListener('mousedown', (e) => {
          e.preventDefault();
          selectedUser = user;
          filterInput.value = displayName;
          analyzeBtn.disabled = false;
          hideDropdown();
        });
        userList.appendChild(item);
      });
    };

    const filterAndDisplay = () => {
      const term = (filterInput.value || '').trim().toLowerCase();
      if (term.length === 0) {
        filteredUsers = allUsers;
      } else {
        filteredUsers = allUsers.filter((user) => {
          const fullname = (user.fullname || '').toLowerCase();
          const domainname = (user.domainname || '').toLowerCase();
          return fullname.includes(term) || domainname.includes(term);
        });
      }
      populateList(filteredUsers);
      showDropdown();
    };

    const loadAllUsers = async () => {
      try {
        refreshBtn.disabled = true;
        refreshBtn.textContent = 'Loading...';
        clearEl(resultDiv);
        showMessage(resultDiv, 'Loading users...', c.primary);
        
        let users = [];
        if (selectedViewFetchXml) {
          // Load users from selected view using FetchXML
          const records = await executeFetchXml('systemuser', selectedViewFetchXml);
          users = records.map(record => ({
            systemuserid: record.systemuserid,
            fullname: record.fullname || '',
            domainname: record.domainname || ''
          }));
        } else {
          // Load all users
          const response = await Xrm.WebApi.online.retrieveMultipleRecords('systemuser', '?$select=systemuserid,fullname,domainname&$orderby=fullname asc&$top=5000');
          users = response.entities || [];
        }
        
        allUsers = users;
        filteredUsers = allUsers;
        filterInput.value = '';
        selectedUser = null;
        analyzeBtn.disabled = true;
        clearEl(resultDiv);
        setFooter('Users loaded: ' + allUsers.length);
      } catch (error) {
        showMessage(resultDiv, 'Error loading users: ' + (error.message || error), c.textError);
      } finally {
        refreshBtn.disabled = false;
        refreshBtn.textContent = 'Reload Users';
      }
    };
    
    const executeFetchXml = async (entityName, fetchXml) => {
      const records = [];
      let query = '?fetchXml=' + encodeURIComponent(fetchXml);
      let response = await Xrm.WebApi.retrieveMultipleRecords(entityName, query);
      records.push(...(response.entities || []));
      let nextLink = response.nextLink || response['@odata.nextLink'];
      while (nextLink) {
        response = await Xrm.WebApi.retrieveMultipleRecords(entityName, nextLink);
        records.push(...(response.entities || []));
        nextLink = response.nextLink || response['@odata.nextLink'];
      }
      return records;
    };
    
    const loadSystemUserViews = async () => {
      try {
        const viewsResponse = await fetchJson('/savedqueries?$select=savedqueryid,name,fetchxml&$filter=returnedtypecode eq ' + SQUOTE + 'systemuser' + SQUOTE + ' and querytype eq 0&$orderby=name asc');
        allViews = viewsResponse.value || [];
        
        allViews.forEach(view => {
          const option = txt(ce('option'), view.name || view.savedqueryid);
          option.value = view.savedqueryid;
          option.dataset.fetchxml = view.fetchxml || '';
          viewSelect.appendChild(option);
        });
        
        addLog('Loaded ' + allViews.length + ' system user views', 'info');
      } catch (error) {
        console.error('Error loading views:', error);
        addLog('Error loading views: ' + (error.message || error), 'error');
      }
    };
    
    viewSelect.addEventListener('change', () => {
      const selectedOption = viewSelect.options[viewSelect.selectedIndex];
      if (selectedOption.value) {
        selectedViewFetchXml = selectedOption.dataset.fetchxml;
        addLog('Selected view: ' + selectedOption.textContent, 'info');
      } else {
        selectedViewFetchXml = null;
        addLog('Showing all users', 'info');
      }
      loadAllUsers();
    });

    filterInput.addEventListener('input', filterAndDisplay);
    filterInput.addEventListener('focus', () => { if (allUsers.length > 0) filterAndDisplay(); });
    filterInput.addEventListener('blur', hideDropdown);
    refreshBtn.addEventListener('click', loadAllUsers);
    analyzeBtn.addEventListener('click', async () => {
      if (selectedUser) {
        const displayName = (selectedUser.fullname || '') + (selectedUser.domainname ? ' (' + selectedUser.domainname + ')' : '');
        showMessage(resultDiv, 'Loading roles...', c.primary);
        await loadUserRoles(selectedUser.systemuserid, displayName, resultDiv);
      }
    });

    loadSystemUserViews();
    loadAllUsers();

    async function loadUserRoles(userId, userName, container) {
      try {
        const userGuid = guidPathValue(userId);
        if (!userGuid) { showMessage(container, 'Invalid systemuser identifier.', c.textError); return; }
        clearEl(container);
        
        // Create header with clickable user name
        const userHeader = style(ce('div'), 'font-weight:600;margin-bottom:8px;font-size:14px;');
        userHeader.appendChild(txt(ce('span'), 'Security Roles for: '));
        userHeader.appendChild(createRecordLink(userName, 'systemuser', userGuid, 'Open user record in Dynamics 365'));
        container.appendChild(userHeader);
        const roleDetails = await loadUserRoleDetails(userGuid);
        const directRoleIds = roleDetails.directRoleIds;
        const teamRoleIds = roleDetails.teamRoleIds;
        const allRoleIds = Array.from(new Set(directRoleIds.concat(teamRoleIds)));
        const roleMap = await fetchRoleNames(allRoleIds);
        if (allRoleIds.length === 0) {
          showMessage(container, 'No roles assigned.');
          setFooter('Roles: 0');
          return;
        }
        const table = ce('table');
        table.style.cssText = 'width:100%;border-collapse:collapse;font-size:12px;margin-top:8px;';
        const headRow = ce('tr');
        ['Role Name', 'Assignment Type'].forEach((header) => {
          const th = txt(ce('th'), header);
          th.style.cssText = `padding:8px;text-align:left;border-bottom:2px solid ${c.border};font-weight:600;background:${c.bgHover};color:${c.text};`;
          headRow.appendChild(th);
        });
        const thead = ce('thead');
        thead.appendChild(headRow);
        table.appendChild(thead);
        const tbody = ce('tbody');
        const directSet = new Set(directRoleIds);
        const teamSet = new Set(teamRoleIds);
        allRoleIds
          .sort((a, b) => (roleMap.get(a) || '').localeCompare(roleMap.get(b) || ''))
          .forEach(async (roleId) => {
            const tr = ce('tr');
            const nameCell = ce('td');
            // Create Power Platform Admin Center link for security role
            const envId = await getEnvironmentId();
            if (envId) {
              const roleLink = txt(ce('a'), roleMap.get(roleId) || roleId);
              roleLink.href = `https://admin.powerplatform.microsoft.com/settingredirect/${envId}/securityroles/${roleId}/roleeditor`;
              roleLink.target = '_blank';
              roleLink.title = 'Open security role in Power Platform Admin Center';
              roleLink.style.cssText = `color:${c.primary};text-decoration:none;cursor:pointer;`;
              roleLink.addEventListener('mouseenter', function() { this.style.textDecoration = 'underline'; });
              roleLink.addEventListener('mouseleave', function() { this.style.textDecoration = 'none'; });
              nameCell.appendChild(roleLink);
            } else {
              // Fallback to plain text if environment ID cannot be determined
              nameCell.appendChild(txt(ce('span'), roleMap.get(roleId) || roleId));
            }
            nameCell.style.cssText = `padding:6px 8px;border-bottom:1px solid ${c.borderLight};`;
            const typeParts = [];
            const isDirect = directSet.has(roleId);
            const isTeam = teamSet.has(roleId);
            if (isDirect && isTeam) {
              typeParts.push('D, T');
            } else if (isDirect) {
              typeParts.push('D');
            } else if (isTeam) {
              typeParts.push('T');
            }
            const typeCell = txt(ce('td'), typeParts.join(''));
            typeCell.style.cssText = `padding:6px 8px;border-bottom:1px solid ${c.borderLight};color:${c.text};`;
            typeCell.title = (isDirect && isTeam) ? 'Direct and Team assignment' : (isDirect ? 'Direct assignment' : 'Team assignment');
            tr.appendChild(nameCell);
            tr.appendChild(typeCell);
            tbody.appendChild(tr);
          });
        table.appendChild(tbody);
        container.appendChild(table);
        
        // Add export buttons
        addExportButtons(userName, allRoleIds, directSet, teamSet, roleMap, container);
        
        setFooter('Roles: ' + allRoleIds.length + ' (' + directRoleIds.length + ' direct, ' + teamRoleIds.length + ' via teams)');
      } catch (error) {
        showMessage(container, 'Error loading roles: ' + (error.message || error), c.textError);
      }
    }
    
    function addExportButtons(userName, roleIds, directSet, teamSet, roleMap, container) {
      const exportDiv = style(ce('div'), 'margin-top:16px;display:flex;gap:8px;justify-content:flex-end;');
      
      const createExportBtn = (text, color, onClick) => {
        const btn = txt(ce('button'), text);
        btn.style.cssText = `padding:8px 16px;background:${color};color:${c.textOnPrimary};border:none;border-radius:4px;cursor:pointer;font-weight:600;font-size:12px;`;
        btn.addEventListener('mouseenter', function() { this.style.opacity = '0.8'; });
        btn.addEventListener('mouseleave', function() { this.style.opacity = '1'; });
        btn.addEventListener('click', onClick);
        return btn;
      };
      
      const downloadFile = (content, filename, mimeType) => {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = ce('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      };
      
      const exportToCSV = () => {
        const sortedRoleIds = roleIds.slice().sort((a, b) => (roleMap.get(a) || '').localeCompare(roleMap.get(b) || ''));
        let csv = 'User,Role Name,Assignment Type\n';
        
        sortedRoleIds.forEach(roleId => {
          const roleName = roleMap.get(roleId) || roleId;
          const isDirect = directSet.has(roleId);
          const isTeam = teamSet.has(roleId);
          let assignmentType = '';
          if (isDirect && isTeam) assignmentType = 'D, T';
          else if (isDirect) assignmentType = 'D';
          else if (isTeam) assignmentType = 'T';
          
          csv += '"' + userName.replace(/"/g, '""') + '",';
          csv += '"' + roleName.replace(/"/g, '""') + '",';
          csv += assignmentType + '\n';
        });
        
        downloadFile(csv, 'security-roles-' + userName.replace(/[^a-z0-9]/gi, '-') + '.csv', 'text/csv;charset=utf-8;');
        setFooter('Exported to CSV');
        addLog('Exported roles to CSV for ' + userName, 'info');
      };
      
      const exportToJSON = () => {
        const roles = [];
        roleIds.forEach(roleId => {
          const isDirect = directSet.has(roleId);
          const isTeam = teamSet.has(roleId);
          let assignmentType = '';
          if (isDirect && isTeam) assignmentType = 'D, T';
          else if (isDirect) assignmentType = 'D';
          else if (isTeam) assignmentType = 'T';
          
          roles.push({
            roleId: roleId,
            roleName: roleMap.get(roleId) || roleId,
            assignmentType: assignmentType,
            isDirect: isDirect,
            isTeam: isTeam
          });
        });
        
        const data = {
          exportDate: new Date().toISOString(),
          user: userName,
          totalRoles: roleIds.length,
          directRoles: Array.from(directSet).length,
          teamRoles: Array.from(teamSet).length,
          roles: roles.sort((a, b) => a.roleName.localeCompare(b.roleName))
        };
        
        downloadFile(JSON.stringify(data, null, 2), 'security-roles-' + userName.replace(/[^a-z0-9]/gi, '-') + '.json', 'application/json');
        setFooter('Exported to JSON');
        addLog('Exported roles to JSON for ' + userName, 'info');
      };
      
      const exportToMarkdown = () => {
        const sortedRoleIds = roleIds.slice().sort((a, b) => (roleMap.get(a) || '').localeCompare(roleMap.get(b) || ''));
        let md = '# Security Roles Report\n\n';
        md += '**User:** ' + userName + '\n\n';
        md += '**Generated:** ' + new Date().toLocaleString() + '\n\n';
        md += '## Legend\n\n';
        md += '- **D**: Direct assignment\n';
        md += '- **T**: Team assignment\n';
        md += '- **D, T**: Both direct and team assignment\n\n';
        md += '## Roles Summary\n\n';
        md += '- Total Roles: ' + roleIds.length + '\n';
        md += '- Direct Assignments: ' + Array.from(directSet).length + '\n';
        md += '- Team Assignments: ' + Array.from(teamSet).length + '\n\n';
        md += '## Role Details\n\n';
        md += '| Role Name | Assignment Type |\n';
        md += '|-----------|----------------|\n';
        
        sortedRoleIds.forEach(roleId => {
          const roleName = roleMap.get(roleId) || roleId;
          const isDirect = directSet.has(roleId);
          const isTeam = teamSet.has(roleId);
          let assignmentType = '';
          if (isDirect && isTeam) assignmentType = 'D, T';
          else if (isDirect) assignmentType = 'D';
          else if (isTeam) assignmentType = 'T';
          
          md += '| ' + roleName + ' | ' + assignmentType + ' |\n';
        });
        
        downloadFile(md, 'security-roles-' + userName.replace(/[^a-z0-9]/gi, '-') + '.md', 'text/markdown;charset=utf-8;');
        setFooter('Exported to Markdown');
        addLog('Exported roles to Markdown for ' + userName, 'info');
      };
      
      exportDiv.appendChild(createExportBtn('📄 Export CSV', c.success, exportToCSV));
      exportDiv.appendChild(createExportBtn('📋 Export JSON', c.primary, exportToJSON));
      exportDiv.appendChild(createExportBtn('📝 Export MD', '#5c2d91', exportToMarkdown));
      
      container.appendChild(exportDiv);
    }

    async function loadUserRoleDetails(userGuid) {
      const userId = guidPathValue(userGuid);
      if (!userId) {
        const emptyResult = {};
        emptyResult.directRoleIds = [];
        emptyResult.teamRoleIds = [];
        return emptyResult;
      }
      const userLiteral = guidLiteral(userGuid);
      const directRolesData = await fetchWithFallback(
        () => fetchJson('/systemusers(' + userId + ')/systemuserroles_association?$select=roleid'),
        () => fetchJson('/systemuserroles?$select=roleid&$filter=systemuserid eq ' + userLiteral)
      );
      const directRoleIds = (directRolesData.value || []).map((role) => sanitizeGuid(role.roleid)).filter((id) => id);

      const teamMemberships = await fetchWithFallback(
        () => fetchJson('/systemusers(' + userId + ')/teammembership_association?$select=teamid'),
        () => fetchJson('/teammembership?$select=teamid&$filter=systemuserid eq ' + userLiteral)
      );
      const teamIds = (teamMemberships.value || []).map((team) => sanitizeGuid(team.teamid)).filter((id) => id);
      const teamRoleIds = [];
      if (teamIds.length > 0) {
        const teamRolesData = await fetchTeamRoles(teamIds.map(guidPathValue));
        (teamRolesData.value || []).forEach((role) => {
          const roleId = sanitizeGuid(role.roleid);
          if (roleId) teamRoleIds.push(roleId);
        });
      }
      const result = {};
      result.directRoleIds = directRoleIds;
      result.teamRoleIds = teamRoleIds;
      return result;
    }

    async function fetchTeamRoles(teamIds) {
      return fetchWithFallback(
        async () => {
          const results = await Promise.all(teamIds.map((teamId) => fetchJson('/teams(' + teamId + ')/teamroles_association?$select=roleid')));
          const combined = {};
          combined.value = results.flatMap((result) => result.value || []);
          return combined;
        },
        () => fetchJson('/teamroles?$select=roleid&$filter=' + teamIds.map((id) => 'teamid eq ' + guidLiteral(id)).join(' or '))
      );
    }

    async function fetchRoleNames(roleIds) {
      const roleMap = new Map();
      if (roleIds.length === 0) return roleMap;
      for (const batch of chunk(roleIds.map(guidPathValue), 20)) {
        const rolesData = await fetchJson('/roles?$select=roleid,name&$filter=' + batch.map((id) => 'roleid eq ' + guidLiteral(id)).join(' or '));
        (rolesData.value || []).forEach((role) => {
          roleMap.set(sanitizeGuid(role.roleid), role.name || '');
        });
      }
      return roleMap;
    }
    // Settings Dialog
    const showSettings = () => {
      const overlay = createOverlay('Color Settings', {
        rootId: 'crmab-securityroleanalyzer-settings',
        width: 700,
        maxHeight: '90vh',
        footer: false
      });
      if (!overlay || !overlay.body) return;
      const dialog = overlay.body;
      dialog.style.cssText = `padding:16px;max-height:80vh;overflow-y:auto;`;
      
      const modeTabs = ce('div');
      modeTabs.style.cssText = 'display:flex;gap:8px;margin-bottom:16px;border-bottom:2px solid ' + c.border + ';';
      let currentMode = 'light';
      
      const lightTab = txt(ce('button'), 'Light Mode');
      lightTab.style.cssText = `padding:8px 16px;background:transparent;border:none;border-bottom:2px solid ${c.primary};color:${c.primary};cursor:pointer;font-weight:600;`;
      const darkTab = txt(ce('button'), 'Dark Mode');
      darkTab.style.cssText = `padding:8px 16px;background:transparent;border:none;border-bottom:2px solid transparent;color:${c.textMuted};cursor:pointer;font-weight:600;`;
      
      const colorGrid = ce('div');
      colorGrid.style.cssText = 'display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px;';
      
      const renderColorInputs = (mode) => {
        clearEl(colorGrid);
        const colors = (customColors && customColors[mode]) || defaultColors[mode];
        Object.keys(defaultColors[mode]).forEach((key) => {
          const row = ce('div');
          row.style.cssText = 'display:flex;flex-direction:column;gap:4px;';
          const label = txt(ce('label'), key.replace(/([A-Z])/g, ' $1').trim());
          label.style.cssText = `font-size:12px;font-weight:600;color:${c.text};`;
          const inputWrapper = ce('div');
          inputWrapper.style.cssText = 'display:flex;gap:8px;align-items:center;';
          const input = ce('input');
          input.type = 'text';
          input.value = colors[key];
          input.style.cssText = `flex:1;padding:6px;border:1px solid ${c.border};border-radius:4px;font-family:monospace;font-size:12px;background:${c.bg};color:${c.text};`;
          const preview = ce('div');
          preview.style.cssText = `width:32px;height:32px;border:1px solid ${c.border};border-radius:4px;background:${colors[key]};`;
          input.addEventListener('input', () => { preview.style.background = input.value; });
          inputWrapper.appendChild(input);
          inputWrapper.appendChild(preview);
          row.appendChild(label);
          row.appendChild(inputWrapper);
          row.dataset.key = key;
          colorGrid.appendChild(row);
        });
      };
      
      const switchMode = (mode) => {
        currentMode = mode;
        if (mode === 'light') {
          lightTab.style.borderBottomColor = c.primary;
          lightTab.style.color = c.primary;
          darkTab.style.borderBottomColor = 'transparent';
          darkTab.style.color = c.textMuted;
        } else {
          darkTab.style.borderBottomColor = c.primary;
          darkTab.style.color = c.primary;
          lightTab.style.borderBottomColor = 'transparent';
          lightTab.style.color = c.textMuted;
        }
        renderColorInputs(mode);
      };
      
      lightTab.addEventListener('click', () => switchMode('light'));
      darkTab.addEventListener('click', () => switchMode('dark'));
      
      modeTabs.appendChild(lightTab);
      modeTabs.appendChild(darkTab);
      dialog.appendChild(modeTabs);
      dialog.appendChild(colorGrid);
      
      renderColorInputs('light');
      
      const buttonRow = ce('div');
      buttonRow.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;margin-top:16px;';
      
      const saveBtn = txt(ce('button'), 'Save');
      saveBtn.style.cssText = `padding:8px 16px;background:${c.success};color:${c.textOnPrimary};border:none;border-radius:4px;cursor:pointer;font-weight:600;`;
      saveBtn.addEventListener('mouseenter', function(){ this.style.background=c.successHover; });
      saveBtn.addEventListener('mouseleave', function(){ this.style.background=c.success; });
      saveBtn.addEventListener('click', () => {
        const newColors = { light: {}, dark: {} };
        ['light', 'dark'].forEach((mode) => {
          switchMode(mode);
          Array.from(colorGrid.children).forEach((row) => {
            const key = row.dataset.key;
            const input = row.querySelector('input');
            newColors[mode][key] = input.value;
          });
        });
        try {
          localStorage.setItem('crmab-securityroleanalyzer-colors', JSON.stringify(newColors));
          addLog('Color settings saved', 'info');
          if (overlay.host) overlay.host.remove();
          requestAnimationFrame(() => location.reload());
        } catch (ex) {
          alert('Error saving settings: ' + ex.message);
        }
      });
      
      const resetBtn = txt(ce('button'), 'Reset to Defaults');
      resetBtn.style.cssText = `padding:8px 16px;background:${c.danger};color:${c.textOnPrimary};border:none;border-radius:4px;cursor:pointer;font-weight:600;`;
      resetBtn.addEventListener('mouseenter', function(){ this.style.background=c.dangerHover; });
      resetBtn.addEventListener('mouseleave', function(){ this.style.background=c.danger; });
      resetBtn.addEventListener('click', () => {
        if (confirm('Reset all colors to defaults?')) {
          try {
            localStorage.removeItem('crmab-securityroleanalyzer-colors');
            addLog('Color settings reset to defaults', 'info');
            if (overlay.host) overlay.host.remove();
            requestAnimationFrame(() => location.reload());
          } catch (ex) {
            alert('Error resetting settings: ' + ex.message);
          }
        }
      });
      
      const cancelBtn = txt(ce('button'), 'Cancel');
      cancelBtn.style.cssText = `padding:8px 16px;background:transparent;color:${c.text};border:1px solid ${c.border};border-radius:4px;cursor:pointer;font-weight:600;`;
      cancelBtn.addEventListener('mouseenter', function(){ this.style.background=c.bgHover; });
      cancelBtn.addEventListener('mouseleave', function(){ this.style.background='transparent'; });
      cancelBtn.addEventListener('click', () => { if (overlay.host) overlay.host.remove(); });
      
      buttonRow.appendChild(cancelBtn);
      buttonRow.appendChild(resetBtn);
      buttonRow.appendChild(saveBtn);
      dialog.appendChild(buttonRow);
    };

    // Logs Viewer Dialog
    const showLogViewer = () => {
      const overlay = createOverlay('Activity Logs', {
        rootId: 'crmab-securityroleanalyzer-logs',
        width: 800,
        maxHeight: '90vh',
        footer: false
      });
      if (!overlay || !overlay.body) return;
      const dialog = overlay.body;
      dialog.style.cssText = `padding:16px;max-height:80vh;overflow-y:auto;`;
      
      const header = ce('div');
      header.style.cssText = 'display:flex;justify-content:flex-end;align-items:center;margin-bottom:16px;';
      const clearBtn = txt(ce('button'), 'Clear Logs');
      clearBtn.style.cssText = `padding:6px 12px;background:${c.danger};color:${c.textOnPrimary};border:none;border-radius:4px;cursor:pointer;font-weight:600;font-size:12px;`;
      clearBtn.addEventListener('mouseenter', function(){ this.style.background=c.dangerHover; });
      clearBtn.addEventListener('mouseleave', function(){ this.style.background=c.danger; });
      clearBtn.addEventListener('click', () => {
        if (confirm('Clear all activity logs?')) {
          activityLog = [];
          if (overlay.host) overlay.host.remove();
        }
      });
      header.appendChild(clearBtn);
      dialog.appendChild(header);
      
      const logContainer = ce('div');
      logContainer.style.cssText = `max-height:500px;overflow-y:auto;border:1px solid ${c.border};border-radius:4px;padding:12px;background:${c.bgHover};`;
      
      if (activityLog.length === 0) {
        const emptyMsg = txt(ce('div'), 'No activity logs yet');
        emptyMsg.style.cssText = `color:${c.textMuted};text-align:center;padding:20px;`;
        logContainer.appendChild(emptyMsg);
      } else {
        activityLog.slice().reverse().forEach((log) => {
          const entry = ce('div');
          entry.style.cssText = `padding:8px;border-bottom:1px solid ${c.borderLight};font-size:12px;`;
          const timestamp = txt(ce('span'), log.timestamp.toLocaleString() + ' - ');
          timestamp.style.cssText = `color:${c.textMuted};font-weight:600;`;
          const message = txt(ce('span'), log.message);
          message.style.cssText = log.type === 'error' ? `color:${c.textError};` : `color:${c.text};`;
          entry.appendChild(timestamp);
          entry.appendChild(message);
          logContainer.appendChild(entry);
        });
      }
      
      dialog.appendChild(logContainer);
      
      const closeBtn = txt(ce('button'), 'Close');
      closeBtn.style.cssText = `margin-top:16px;padding:8px 16px;background:${c.primary};color:${c.textOnPrimary};border:none;border-radius:4px;cursor:pointer;font-weight:600;width:100%;`;
      closeBtn.addEventListener('mouseenter', function(){ this.style.background=c.primaryHover; });
      closeBtn.addEventListener('mouseleave', function(){ this.style.background=c.primary; });
      closeBtn.addEventListener('click', () => { if (overlay.host) overlay.host.remove(); });
      dialog.appendChild(closeBtn);
    };

    // Add footer buttons (on the right side)
    const settingsBtn = txt(ce('button'), String.fromCharCode(9881) + ' Settings');
    settingsBtn.style.cssText = `background:transparent;border:1px solid ${c.border};color:${c.text};padding:6px 12px;border-radius:4px;cursor:pointer;font-weight:600;font-size:12px;margin-left:8px;`;
    settingsBtn.addEventListener('mouseenter', function(){ this.style.background=c.bgHover; });
    settingsBtn.addEventListener('mouseleave', function(){ this.style.background='transparent'; });
    settingsBtn.addEventListener('click', showSettings);
    
    const logBtn = txt(ce('button'), String.fromCharCode(128276) + ' Logs');
    logBtn.style.cssText = `background:transparent;border:1px solid ${c.border};color:${c.text};padding:6px 12px;border-radius:4px;cursor:pointer;font-weight:600;font-size:12px;margin-left:8px;`;
    logBtn.addEventListener('mouseenter', function(){ this.style.background=c.bgHover; });
    logBtn.addEventListener('mouseleave', function(){ this.style.background='transparent'; });
    logBtn.addEventListener('click', showLogViewer);
    
    if (ui.footer) {
      ui.footer.appendChild(settingsBtn);
      ui.footer.appendChild(logBtn);
    }

  } catch (error) {
    console.error(error);
    const err = document.createElement('div');
    err.style.cssText = 'color:#a80000;margin-top:8px;padding:12px;';
    err.textContent = 'Error: ' + (error.message || error);
    if (ui.body) ui.body.appendChild(err);
  }
})();