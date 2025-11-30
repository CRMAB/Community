// Field Focus v2 – Overlay navigator for fields, tabs, and sections
(function () {
  var X = window.Xrm;
  if (!X || !X.Page || !X.Page.ui) {
    alert("Form context not available.");
    return;
  }
  var overlayLib = window.__lvlUpOverlay;
  if (!overlayLib || typeof overlayLib.createOverlay !== "function") {
    alert("helperOverlayV2.js must be loaded.");
    return;
  }

  var ROOT_ID = "crmab-fieldfocus-root";
  var STYLE_ID = "crmab-fieldfocus-style";
  var overlay = overlayLib.createOverlay("Field Focus 2.0", {
    rootId: ROOT_ID,
    styleId: STYLE_ID,
    width: 540,
    maxHeight: "80vh",
    footer: true,
    footerText: "Ready",
    customStyles: buildStyles(ROOT_ID)
  });
  if (!overlay) return;
  var body = overlay.body;
  body.classList.add("crmab-ff-body");
  var ce = function (tag, cls, text) {
    var el = document.createElement(tag);
    if (cls) el.className = cls;
    if (text) el.textContent = text;
    return el;
  };
  var raf = function (fn) {
    if (window.requestAnimationFrame) return window.requestAnimationFrame(fn);
    if (typeof queueMicrotask === "function") return queueMicrotask(fn);
    if (window.Promise) return Promise.resolve().then(fn);
    return fn();
  };
  var controls = ce("div", "crmab-ff-controls");
  var modeSelect = ce("select", "crmab-ff-select");
  [{ value: "fields", label: "Fields / Columns" }, { value: "tabs", label: "Tabs" }, { value: "sections", label: "Sections" }].forEach(function (opt) {
    var option = ce("option");
    option.value = opt.value;
    option.textContent = opt.label;
    modeSelect.appendChild(option);
  });

  var searchInput = ce("input", "crmab-ff-input"); searchInput.type = "text"; searchInput.placeholder = "Search (label, logical name, tab, section...)";

  var nameToggle = ce("button", "crmab-ff-btn", "Display"); nameToggle.type = "button";
  nameToggle.title = "Toggle between display name and logical name";
  var sortToggle = ce("button", "crmab-ff-btn", "A-Z"); sortToggle.type = "button";
  sortToggle.title = "Sort order";
  var refreshBtn = ce("button", "crmab-ff-btn crmab-ff-btn-icon", "↻"); refreshBtn.type = "button";
  refreshBtn.title = "Refresh metadata";

  controls.appendChild(modeSelect); controls.appendChild(searchInput); controls.appendChild(nameToggle);
  controls.appendChild(sortToggle); controls.appendChild(refreshBtn); body.appendChild(controls);
  var list = ce("div", "crmab-ff-list"); body.appendChild(list);

  var logs = [];
  var addLog = function (message, level) {
    var timestamp = new Date().toISOString().substr(11, 8);
    level = level || "info";
    logs.push({ timestamp: timestamp, level: level, message: message });
    console.log("[" + timestamp + "] [" + level.toUpperCase() + "] " + message);
  };
  var state = {
      mode: "fields",
      query: "",
      showLogical: false,
      sort: "alpha",
      items: []
    },
    builders = {
      fields: buildFieldRows,
      tabs: buildTabRows,
      sections: buildSectionRows
    },
    modeLabels = {
      fields: "Fields / Columns",
      tabs: "Tabs",
      sections: "Sections"
    };
  addLog("Field Focus V2 initialized", "success");
  function iterate(collection, cb) {
    if (!collection) return;
    if (typeof collection.forEach === "function") {
      collection.forEach(cb);
      return;
    }
    if (typeof collection.getLength === "function" && typeof collection.get === "function") {
      for (var i = 0; i < collection.getLength(); i++) cb(collection.get(i), i);
      return;
    }
    if (typeof collection.length === "number") {
      for (var j = 0; j < collection.length; j++) cb(collection[j], j);
    }
  }
  function isRenderableControl(control) {
    if (!control || !control.getControlType) return false;
    var t = control.getControlType();
    return ["timeline", "subgrid", "quickform", "kbsearch", "webresource", "iframe"].indexOf(t) === -1;
  }
  function hostForControl(name) {
    if (!name) return null;
    return document.querySelector("[data-control-name='" + name + "']");
  }
  function getTabLabel(tab) {
    if (!tab) return "";
    if (tab.getLabel) return tab.getLabel() || "";
    if (tab.getName) return tab.getName() || "";
    return "";
  }
  function getSectionLabel(section) {
    if (!section) return "";
    if (section.getLabel) return section.getLabel() || "";
    if (section.getName) return section.getName() || "";
    return "";
  }
  function attrSelector(attr, value) {
    if (!value && value !== 0) return null;
    var safe = String(value).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
    return "[" + attr + "=\"" + safe + "\"]";
  }
  function getSectionDom(section) {
    if (!section) return null;
    var logical = section.getName ? section.getName() : "";
    var label = getSectionLabel(section);
    var logicalLower = logical ? logical.toLowerCase() : "";
    var labelLower = label ? label.toLowerCase() : "";
    var selectors = [];
    if (logical) {
      selectors.push(attrSelector("data-id", "Section_" + logical));
      selectors.push(attrSelector("data-id", logical));
      selectors.push(attrSelector("data-id", "section_" + logicalLower));
    }
    if (label) selectors.push(attrSelector("aria-label", label));
    for (var i = 0; i < selectors.length; i++) {
      var sel = selectors[i];
      if (!sel) continue;
      var el = document.querySelector(sel);
      if (el) return el;
    }
    var candidates = document.querySelectorAll("[data-id^='section_'], section[aria-label]");
    for (var j = 0; j < candidates.length; j++) {
      var node = candidates[j];
      var dataId = (node.getAttribute("data-id") || "").toLowerCase();
      if (logicalLower && dataId.indexOf(logicalLower) !== -1) return node;
      var aria = (node.getAttribute("aria-label") || "").toLowerCase();
      if (labelLower && aria === labelLower) return node;
    }
    return null;
  }
  function isVisible(entity) {
    try { return !!(entity && entity.getVisible && entity.getVisible()); }
    catch (e) { return true; }
  }
  function buildFieldRows() {
    var rows = [];
    iterate(X.Page.getAttribute(), function (attr) {
      try {
        var name = attr && attr.getName ? attr.getName() : null;
        if (!name || !attr.controls || typeof attr.controls.getLength !== "function") return;
        var chosen = null;
        for (var i = 0; i < attr.controls.getLength(); i++) {
          var ctrl = attr.controls.get(i);
          if (isRenderableControl(ctrl)) { chosen = ctrl; break; }
        }
        if (!chosen) chosen = attr.controls.getLength() ? attr.controls.get(0) : null;
        if (!chosen) return;
        var label = (chosen.getLabel && chosen.getLabel()) || name;
        var section = chosen.getParent ? chosen.getParent() : null;
        var tab = section && section.getParent ? section.getParent() : null;
        var ctrlName = chosen.getName ? chosen.getName() : null;
        rows.push({
          type: "field",
          label: label,
          logical: name,
          visible: isVisible(chosen),
          control: chosen,
          section: section,
          tab: tab,
          tabName: getTabLabel(tab),
          sectionName: getSectionLabel(section),
          host: hostForControl(ctrlName)
        });
      } catch (e) {}
    });
    return rows;
  }
  function hostsForSection(section) {
    var hosts = [];
    if (!section || !section.controls || typeof section.controls.getLength !== "function") return hosts;
    for (var i = 0; i < section.controls.getLength(); i++) {
      var ctrl = section.controls.get(i);
      var n = ctrl && ctrl.getName ? ctrl.getName() : null;
      var host = hostForControl(n);
      if (host) hosts.push(host);
    }
    return hosts;
  }

  function buildTabRows() {
    var rows = [];
    var tabs = X.Page.ui && X.Page.ui.tabs ? X.Page.ui.tabs.get() : [];
    for (var i = 0; i < tabs.length; i++) {
      var tab = tabs[i];
      var label = getTabLabel(tab);
      var logical = tab && tab.getName ? tab.getName() : label;
      var hosts = [];
      var sections = tab && tab.sections;
      if (sections && typeof sections.getLength === "function") {
        for (var s = 0; s < sections.getLength(); s++) {
          var sectionHosts = hostsForSection(sections.get(s));
          hosts = hosts.concat(sectionHosts);
        }
      }
      rows.push({
        type: "tab",
        label: label,
        logical: logical,
        visible: isVisible(tab),
        tab: tab,
        hosts: hosts,
        host: hosts.length ? hosts[0] : null,
        tabName: label
      });
    }
    return rows;
  }

  function buildSectionRows() {
    var rows = [];
    var tabs = X.Page.ui && X.Page.ui.tabs ? X.Page.ui.tabs.get() : [];
    for (var i = 0; i < tabs.length; i++) {
      var tab = tabs[i];
      var sections = tab && tab.sections;
      if (!sections || typeof sections.getLength !== "function") continue;
      for (var j = 0; j < sections.getLength(); j++) {
        var section = sections.get(j);
        var hosts = hostsForSection(section);
        var sectionDom = getSectionDom(section);
        if (sectionDom) hosts.unshift(sectionDom);
        rows.push({
          type: "section",
          label: getSectionLabel(section),
          logical: section && section.getName ? section.getName() : "",
          visible: isVisible(section),
          section: section,
          tab: tab,
          hosts: hosts,
          host: hosts.length ? hosts[0] : null,
          tabName: getTabLabel(tab),
          sectionName: getSectionLabel(section)
        });
      }
    }
    return rows;
  }

  function highlightHosts(hosts) {
    if (!hosts || !hosts.length) return;
    for (var i = 0; i < hosts.length; i++) {
      var h = hosts[i];
      try { if (h && h.classList) h.classList.add("crmab-ff-highlight"); } catch (e) {}
    }
    var start = (window.performance && performance.now) ? performance.now() : Date.now();
    var duration = 1200;
    var checker = function (now) {
      var current = now || Date.now();
      if (current - start >= duration) {
        for (var j = 0; j < hosts.length; j++) {
          var h2 = hosts[j];
          try { if (h2 && h2.classList) h2.classList.remove("crmab-ff-highlight"); } catch (e) {}
        }
      } else {
        raf(checker);
      }
    };
    raf(checker);
  }

  function scrollAndHighlight(hosts) {
    if (!hosts || !hosts.length) return;
    var first = hosts[0];
    try { first.scrollIntoView({ behavior: "smooth", block: "center" }); }
    catch (e) { first.scrollIntoView(true); }
    highlightHosts(hosts);
  }

  function expandTab(tab) {
    if (!tab) return;
    try { tab.setVisible && tab.setVisible(true); } catch (e) {}
    try { tab.setDisplayState && tab.setDisplayState("expanded"); } catch (e) {}
  }

  function safeFocus(control) {
    try { control && control.setFocus && control.setFocus(); } catch (e) {}
  }

  function handleRowClick(row) {
    var scrollPos = list.scrollTop;
    try {
      if (row.type === "field") {
        try { row.control && row.control.setVisible && row.control.setVisible(true); } catch (e) {}
        try { row.section && row.section.setVisible && row.section.setVisible(true); } catch (e) {}
        expandTab(row.tab);
        safeFocus(row.control);
        if (!row.host) row.host = hostForControl(row.control && row.control.getName ? row.control.getName() : null);
        scrollAndHighlight(row.host ? [row.host] : []);
      } else if (row.type === "tab") {
        expandTab(row.tab);
        var firstCtrl = null;
        try {
          var sections = row.tab && row.tab.sections;
          if (sections && typeof sections.getLength === "function" && sections.getLength() > 0) {
            var sec = sections.get(0);
            if (sec && sec.controls && typeof sec.controls.getLength === "function" && sec.controls.getLength() > 0) {
              firstCtrl = sec.controls.get(0);
            }
          }
        } catch (e) {}
        safeFocus(firstCtrl);
        scrollAndHighlight(row.hosts && row.hosts.length ? row.hosts : (row.host ? [row.host] : []));
      } else if (row.type === "section") {
        expandTab(row.tab);
        try { row.section && row.section.setVisible && row.section.setVisible(true); } catch (e) {}
        scrollAndHighlight(row.hosts && row.hosts.length ? row.hosts : (row.host ? [row.host] : []));
      }
    } catch (ex) {
      alert("Navigation failed: " + (ex.message || ex));
    } finally {
      state.items = builders[state.mode] ? builders[state.mode]() : [];
      render();
      list.scrollTop = scrollPos;
    }
  }

  function displayName(row) {
    if (row.type === "field") return state.showLogical ? (row.logical || "") : (row.label || row.logical || "");
    return row.label || row.logical || "";
  }

  function secondaryName(row) {
    if (row.type === "field") return state.showLogical ? (row.label || "") : (row.logical || "");
    return row.logical && row.logical !== row.label ? row.logical : "";
  }

  function metaLine(row) {
    if (row.type === "field") {
      var seg = [];
      if (row.tabName) seg.push(row.tabName);
      if (row.sectionName) seg.push(row.sectionName);
      return seg.join(" · ");
    }
    if (row.type === "tab") return "Tab";
    return row.tabName || "";
  }

  function filterRows(rows) {
    if (!state.query) return rows.slice();
    var terms = state.query.split(/\s+/).filter(function (t) { return !!t; });
    if (!terms.length) return rows.slice();
    return rows.filter(function (row) {
      var hay = [(row.label || ""), (row.logical || ""), (row.tabName || row.tabLabel || ""), (row.sectionName || "")]
        .join("|").toLowerCase();
      for (var i = 0; i < terms.length; i++) {
        if (hay.indexOf(terms[i]) === -1) return false;
      }
      return true;
    });
  }

  var collator = window.Intl && typeof Intl.Collator === "function" ? new Intl.Collator(undefined, { sensitivity: "base" }) : null;
  function compareText(a, b) {
    if (collator) return collator.compare(a || "", b || "");
    var aa = (a || "").toLowerCase();
    var bb = (b || "").toLowerCase();
    if (aa < bb) return -1;
    if (aa > bb) return 1;
    return 0;
  }

  function sortRows(rows) {
    rows.sort(function (a, b) {
      if (state.sort === "alpha" || a.type === "tab" || b.type === "tab") {
        return compareText(displayName(a), displayName(b));
      }
      var tabCompare = compareText(a.tabName || "", b.tabName || "");
      if (tabCompare) return tabCompare;
      if (a.type === "section" || b.type === "section") {
        var secCompare = compareText(a.sectionName || a.label || "", b.sectionName || b.label || "");
        if (secCompare) return secCompare;
      }
      return compareText(displayName(a), displayName(b));
    });
  }

  function buildRow(row) {
    var item = ce("button", "crmab-ff-item");
    item.type = "button";
    var textWrap = ce("div", "crmab-ff-item-text");
    var title = ce("div", "crmab-ff-item-title", displayName(row));
    var secondary = secondaryName(row);
    if (secondary) {
      var pill = ce("span", "crmab-ff-pill", secondary);
      title.appendChild(pill);
    }
    var meta = ce("div", "crmab-ff-meta", metaLine(row));
    textWrap.appendChild(title);
    textWrap.appendChild(meta);

    var badgeWrap = ce("div", "crmab-ff-badge-wrap");
    var status = ce("span", row.visible ? "crmab-ff-badge crmab-ff-badge-ok" : "crmab-ff-badge");
    status.textContent = row.visible ? "visible" : "hidden";
    badgeWrap.appendChild(status);

    item.appendChild(textWrap);
    item.appendChild(badgeWrap);
    item.addEventListener("click", function () { handleRowClick(row); });
    return item;
  }

  function render() {
    var rows = builders[state.mode] ? state.items.slice() : [];
    rows = filterRows(rows);
    sortRows(rows);
    while (list.firstChild) list.removeChild(list.firstChild);
    if (!rows.length) {
      list.appendChild(ce("div", "crmab-ff-empty", "No objects found."));
    } else {
      var frag = document.createDocumentFragment();
      for (var i = 0; i < rows.length; i++) frag.appendChild(buildRow(rows[i]));
      list.appendChild(frag);
    }
    overlay.updateFooter("Objects: " + rows.length + " • Mode: " + (modeLabels[state.mode] || ""));
    addLog("Rendered " + rows.length + " " + state.mode + (state.query ? " (filtered by: '" + state.query + "')" : ""), "info");
    updateToggles();
  }

  function updateToggles() {
    nameToggle.disabled = state.mode !== "fields";
    nameToggle.textContent = state.showLogical ? "Logical" : "Display";
    sortToggle.textContent = state.sort === "alpha" ? "A-Z" : "Tab/Section";
  }

  function refreshItems() {
    state.items = builders[state.mode] ? builders[state.mode]() : [];
    render();
  }

  searchInput.addEventListener("input", function () {
    var inputValue = (searchInput.value || "").trim();
    if (inputValue.length === 0 || inputValue.length >= 3) {
      state.query = inputValue.toLowerCase();
      render();
    }
  });

  modeSelect.addEventListener("change", function () {
    state.mode = modeSelect.value;
    if (state.mode !== "fields") state.showLogical = false;
    addLog("Mode changed to: " + state.mode, "info");
    refreshItems();
  });

  nameToggle.addEventListener("click", function () {
    if (state.mode !== "fields") return;
    state.showLogical = !state.showLogical;
    addLog("Display mode: " + (state.showLogical ? "Logical" : "Display"), "info");
    render();
  });

  sortToggle.addEventListener("click", function () {
    state.sort = state.sort === "alpha" ? "group" : "alpha";
    addLog("Sort order: " + state.sort, "info");
    render();
  });

  refreshBtn.addEventListener("click", function () {
    searchInput.value = "";
    state.query = "";
    addLog("Manual refresh triggered - clearing filter and reloading metadata", "info");
    refreshItems();
  });

  var showLogViewer = function () {
    var logOverlay = overlayLib.createOverlay("Log Viewer", {
      rootId: "crmab-ff-log-viewer",
      width: 700,
      maxHeight: "70vh",
      theme: "auto"
    });
    if (!logOverlay) return;
    var logContent = ce("pre", "crmab-ff-log-content");
    logContent.style.cssText = "background:#f5f5f5;padding:12px;overflow:auto;max-height:500px;font-size:12px;font-family:Consolas,monospace;margin:0;color:#111;";
    logContent.textContent = logs.map(function (l) {
      return "[" + l.timestamp + "] [" + l.level.toUpperCase() + "] " + l.message;
    }).join("\n");
    logOverlay.body.appendChild(logContent);
    var copyBtn = ce("button", "crmab-ff-btn", "Copy to Clipboard");
    copyBtn.style.cssText = "margin-top:12px;padding:6px 12px;background:#0078d4;color:#fff;border:none;border-radius:4px;cursor:pointer;";
    copyBtn.addEventListener("click", function () {
      navigator.clipboard.writeText(logContent.textContent).then(function () {
        alert("Logs copied to clipboard");
      });
    });
    logOverlay.body.appendChild(copyBtn);
  };

  addLog("Loaded " + state.mode + " mode", "info");
  refreshItems();
  searchInput.focus();

  if (overlay.footer) {
    var logBtn = ce("button", "crmab-ff-btn", String.fromCharCode(128220) + " Logs");
    logBtn.style.cssText = "background:transparent;border:1px solid var(--ff-border);color:var(--ff-text);padding:4px 8px;border-radius:4px;cursor:pointer;font-size:12px;margin-right:8px;";
    logBtn.addEventListener("click", showLogViewer);
    logBtn.addEventListener("mouseenter", function () {
      this.style.background = "rgba(0,0,0,0.05)";
    });
    logBtn.addEventListener("mouseleave", function () {
      this.style.background = "transparent";
    });
    var themeBtn = overlay.footer.lastChild;
    if (themeBtn && themeBtn.nodeType === 1) {
      overlay.footer.insertBefore(logBtn, themeBtn);
    } else {
      overlay.footer.appendChild(logBtn);
    }
  }

  function buildStyles(rootId) {
    var selector = "#" + rootId;
    return selector + "{" +
      "--ff-bg:#fff;--ff-bg-alt:#f8f8f8;--ff-border:#ddd;--ff-text:#111;--ff-muted:#605e5c;--ff-highlight:#0ea5e9;" +
    "}" +
    selector + "[data-theme=\"dark\"]{" +
      "--ff-bg:#0b1220;--ff-bg-alt:#111826;--ff-border:#374151;--ff-text:#eef2ff;--ff-muted:#9ca3af;--ff-highlight:#60a5fa;" +
    "}" +
    selector + " .crmab-ff-body{display:flex;flex-direction:column;gap:12px;font-family:'Segoe UI',Tahoma,Arial;font-size:13px;}" +
    selector + " .crmab-ff-controls{display:flex;flex-wrap:wrap;gap:8px;align-items:center;}" +
    selector + " .crmab-ff-select," + selector + " .crmab-ff-input{flex:1;min-width:120px;padding:6px 8px;border:1px solid var(--ff-border);border-radius:6px;background:var(--ff-bg);color:var(--ff-text);}" +
    selector + " .crmab-ff-input::placeholder{color:var(--ff-muted);}" +
    selector + " .crmab-ff-btn{border:1px solid var(--ff-border);border-radius:6px;background:var(--ff-bg-alt);color:var(--ff-text);padding:6px 10px;cursor:pointer;font-weight:600;min-width:82px;}" +
    selector + " .crmab-ff-btn-icon{min-width:40px;font-size:16px;}" +
    selector + " .crmab-ff-btn:disabled{opacity:.5;cursor:not-allowed;}" +
    selector + " .crmab-ff-list{border:1px solid var(--ff-border);border-radius:8px;padding:4px;max-height:55vh;overflow:auto;background:var(--ff-bg);}"+
    selector + " .crmab-ff-item{width:100%;border:none;background:transparent;padding:10px;border-radius:8px;display:flex;justify-content:space-between;gap:12px;text-align:left;color:var(--ff-text);cursor:pointer;}" +
    selector + " .crmab-ff-item:hover{background:var(--ff-bg-alt);}" +
    selector + " .crmab-ff-item-text{flex:1;}" +
    selector + " .crmab-ff-item-title{font-weight:600;font-size:13px;display:flex;flex-wrap:wrap;gap:6px;align-items:center;}" +
    selector + " .crmab-ff-pill{font-size:11px;padding:1px 6px;border-radius:999px;background:var(--ff-bg-alt);border:1px solid var(--ff-border);}" +
    selector + " .crmab-ff-meta{font-size:11px;color:var(--ff-muted);margin-top:2px;}" +
    selector + " .crmab-ff-badge-wrap{display:flex;align-items:center;}" +
    selector + " .crmab-ff-badge{font-size:11px;padding:2px 8px;border-radius:999px;border:1px solid var(--ff-border);color:var(--ff-muted);}" +
    selector + " .crmab-ff-badge-ok{background:#0ea5e910;color:#107c10;border-color:#3cc17c;}" +
    selector + " .crmab-ff-empty{text-align:center;padding:18px;color:var(--ff-muted);}" +
    ".crmab-ff-highlight{outline:3px solid var(--ff-highlight);outline-offset:2px;transition:outline-color .2s;}";
  }
})();
