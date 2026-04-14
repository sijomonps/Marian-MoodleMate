(() => {
  const ROOT_ID = "moodle-assistant-root";
  const MOODLE_HOST = "marianlms.in";
  const MOODLE_DASHBOARD_URL = `https://${MOODLE_HOST}/my/`;
  const LOGIN_URL = "https://marianlms.in/login/index.php";
  const ICON_PATH = "icon.png";
  const NAV_EVENT = "marianmoodlemate:navigate";
  const AUTO_SYNC_MS = 30 * 60 * 1000;
  const RENDER_TIME_BUCKET_MINUTES = 5;

  const STORAGE_KEYS = {
    buttonPosition: "marianMoodleMate.buttonPosition",
    assignments: "marianMoodleMate.assignments",
    courses: "marianMoodleMate.courses",
    lastSyncedAt: "marianMoodleMate.lastSyncedAt",
    notified: "marianMoodleMate.notified",
  };

  const state = {
    assignments: [],
    courses: [],
    notified: {},
    lastSyncedAt: 0,
    panelOpen: false,
    isSyncing: false,
    hasFreshMoodleData: false,
    observer: null,
    autoSyncTimer: null,
    refreshDebounceId: null,
    cachedRegionMain: null,
    renderCache: {
      statusText: "",
      lastSyncedText: "",
      helperKey: "",
      assignmentsKey: "",
      coursesKey: "",
    },
    drag: {
      pointerId: null,
      originX: 0,
      originY: 0,
      originLeft: 0,
      originTop: 0,
      moved: false,
      suppressClick: false,
    },
  };

  function hasStorageApi() {
    return typeof chrome !== "undefined" && !!chrome.storage?.local;
  }

  function storageGet(keys) {
    return new Promise((resolve) => {
      if (!hasStorageApi()) {
        resolve({});
        return;
      }

      chrome.storage.local.get(keys, (result) => {
        if (chrome.runtime?.lastError) {
          resolve({});
          return;
        }
        resolve(result || {});
      });
    });
  }

  function storageSet(items) {
    return new Promise((resolve) => {
      if (!hasStorageApi()) {
        resolve();
        return;
      }

      chrome.storage.local.set(items, () => {
        void chrome.runtime?.lastError;
        resolve();
      });
    });
  }

  function normalizeText(value) {
    return (value || "").replace(/\s+/g, " ").trim();
  }

  function isMoodlePage() {
    return window.location.hostname.includes(MOODLE_HOST);
  }

  function isDashboardPage(path = window.location.pathname) {
    return /^\/my\/?$/.test(path);
  }

  function isCoursesPage(path = window.location.pathname) {
    return /^\/my\/courses\.php\/?$/.test(path);
  }

  function isLikelyLoggedOut(doc = document) {
    if (!doc?.body) {
      return false;
    }

    if (doc.body.classList.contains("notloggedin")) {
      return true;
    }

    const path = doc.location?.pathname || "";
    if (path.includes("/login/index.php")) {
      return true;
    }

    return !!doc.querySelector("form[action*='/login/index.php'] input[name='username']");
  }

  function parseDueDate(dueText) {
    const text = normalizeText(dueText);
    if (!text) {
      return null;
    }

    const candidates = [];
    const keyword = text.match(/(?:due(?:\s*date)?|deadline|closes?)\s*(?:is)?\s*:?\s*(.+)$/i);
    if (keyword?.[1]) {
      candidates.push(keyword[1]);
    }

    const monthFirst = text.match(
      /([A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}(?:,?\s+\d{1,2}:\d{2}\s*(?:AM|PM)?)?)/i
    );
    if (monthFirst?.[1]) {
      candidates.push(monthFirst[1]);
    }

    const dayFirst = text.match(
      /(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}(?:,?\s+\d{1,2}:\d{2}\s*(?:AM|PM)?)?)/i
    );
    if (dayFirst?.[1]) {
      candidates.push(dayFirst[1]);
    }

    const slashDate = text.match(/(\d{1,2}\/\d{1,2}\/\d{2,4}(?:\s+\d{1,2}:\d{2}\s*(?:AM|PM)?)?)/i);
    if (slashDate?.[1]) {
      candidates.push(slashDate[1]);
    }

    candidates.push(text);

    for (const raw of candidates) {
      const cleaned = raw.replace(/\|/g, " ").replace(/\bat\b/gi, " ").replace(/\s+/g, " ").trim();
      const time = Date.parse(cleaned);
      if (!Number.isNaN(time)) {
        return new Date(time);
      }
    }

    return null;
  }

  function calculateTimeRemaining(assignment) {
    const dueDate = assignment?.dueDate ? new Date(assignment.dueDate) : parseDueDate(assignment?.dueText);
    if (!dueDate || Number.isNaN(dueDate.getTime())) {
      return assignment?.dueText || "Due date unavailable";
    }

    const diff = dueDate.getTime() - Date.now();
    const hourMs = 60 * 60 * 1000;
    const dayMs = 24 * hourMs;

    if (diff <= 0) {
      return "❌ Overdue";
    }

    if (diff < 6 * hourMs) {
      const hours = Math.max(1, Math.ceil(diff / hourMs));
      return `⚠️ Due in ${hours} hour${hours === 1 ? "" : "s"}`;
    }

    if (diff < 2 * dayMs) {
      const days = Math.max(1, Math.ceil(diff / dayMs));
      return `🔥 Due in ${days} day${days === 1 ? "" : "s"}`;
    }

    return assignment?.dueText || dueDate.toLocaleString();
  }

  // Stable visual progress mapping for reliability.
  function calculateProgress(assignment) {
    const dueDate = assignment?.dueDate ? new Date(assignment.dueDate) : parseDueDate(assignment?.dueText);
    if (!dueDate || Number.isNaN(dueDate.getTime())) {
      return { percent: 20, color: "#22c55e" };
    }

    const diff = dueDate.getTime() - Date.now();
    const hourMs = 60 * 60 * 1000;
    const dayMs = 24 * hourMs;

    if (diff <= 0) {
      return { percent: 100, color: "#ef4444" };
    }
    if (diff < 6 * hourMs) {
      return { percent: 90, color: "#ef4444" };
    }
    if (diff < dayMs) {
      return { percent: 70, color: "#f59e0b" };
    }
    if (diff < 2 * dayMs) {
      return { percent: 50, color: "#f59e0b" };
    }

    return { percent: 20, color: "#22c55e" };
  }

  function extractAssignmentsFromPage(doc) {
    const nodes = doc.querySelectorAll(".event-name-container a");
    const seen = new Set();
    const out = [];

    nodes.forEach((element) => {
      const title = normalizeText(element.textContent);
      const link = element.href;
      const dueText = normalizeText(element.getAttribute("aria-label"));

      if (!title || !link || seen.has(link)) {
        return;
      }

      const parsed = parseDueDate(dueText);
      out.push({
        title,
        link,
        dueText,
        dueDate: parsed ? parsed.toISOString() : "",
      });
      seen.add(link);
    });

    return out;
  }

  function extractCoursesFromPage(doc) {
    const selectors = [
      ".coursebox .coursename a",
      "#region-main a[href*='/course/view.php?id=']",
      "[data-region='courses-view'] a[href*='/course/view.php?id=']",
      "a[href*='/course/view.php?id=']",
    ];

    const seen = new Set();
    const out = [];

    selectors.forEach((selector) => {
      doc.querySelectorAll(selector).forEach((element) => {
        const title = normalizeText(element.textContent);
        const link = element.href;

        if (!title || !link || seen.has(link)) {
          return;
        }

        seen.add(link);
        out.push({ title, link });
      });
    });

    return out;
  }

  async function saveAssignments(assignments) {
    const createdAtById = new Map(
      state.assignments.map((item) => [`${item.link}|${item.dueDate || item.dueText}`, Number(item.createdAt) || Date.now()])
    );

    const normalized = assignments.map((item) => {
      const id = `${item.link}|${item.dueDate || item.dueText}`;
      return {
        title: item.title,
        link: item.link,
        dueText: item.dueText || "",
        dueDate: item.dueDate || "",
        createdAt: createdAtById.get(id) || Date.now(),
      };
    });

    state.assignments = normalized;
    await storageSet({ [STORAGE_KEYS.assignments]: normalized });
    return normalized;
  }

  async function loadAssignments() {
    const result = await storageGet([STORAGE_KEYS.assignments]);
    state.assignments = Array.isArray(result[STORAGE_KEYS.assignments]) ? result[STORAGE_KEYS.assignments] : [];
    return state.assignments;
  }

  async function saveCourses(courses) {
    const normalized = courses.map((item) => ({ title: item.title, link: item.link }));
    state.courses = normalized;
    await storageSet({ [STORAGE_KEYS.courses]: normalized });
    return normalized;
  }

  async function loadCourses() {
    const result = await storageGet([STORAGE_KEYS.courses]);
    state.courses = Array.isArray(result[STORAGE_KEYS.courses]) ? result[STORAGE_KEYS.courses] : [];
    return state.courses;
  }

  async function loadNotified() {
    const result = await storageGet([STORAGE_KEYS.notified]);
    const map = result[STORAGE_KEYS.notified];
    state.notified = map && typeof map === "object" ? map : {};
  }

  async function saveNotified(map) {
    state.notified = map;
    await storageSet({ [STORAGE_KEYS.notified]: map });
  }

  async function loadLastSyncedAt() {
    const result = await storageGet([STORAGE_KEYS.lastSyncedAt]);
    const value = result[STORAGE_KEYS.lastSyncedAt];
    state.lastSyncedAt = Number.isFinite(value) ? value : 0;
    return state.lastSyncedAt;
  }

  async function saveLastSyncedAt(timestamp) {
    state.lastSyncedAt = timestamp;
    await storageSet({ [STORAGE_KEYS.lastSyncedAt]: timestamp });
  }

  async function loadButtonPosition() {
    const result = await storageGet([STORAGE_KEYS.buttonPosition]);
    const pos = result[STORAGE_KEYS.buttonPosition];
    if (!pos || typeof pos.left !== "number" || typeof pos.top !== "number") {
      return null;
    }
    return pos;
  }

  async function saveButtonPosition(position) {
    await storageSet({ [STORAGE_KEYS.buttonPosition]: position });
  }

  function notify(message) {
    if (typeof chrome !== "undefined" && chrome.notifications?.create) {
      console.log("[MarianMoodleMate] Notification fired:", message);
      const id = `marianmoodlemate-${Date.now()}-${Math.random().toString(16).slice(2)}`;
      chrome.notifications.create(
        id,
        {
          type: "basic",
          title: "Assignment Reminder",
          message,
          iconUrl: chrome.runtime.getURL(ICON_PATH),
          priority: 2,
        },
        () => {
          void chrome.runtime?.lastError;
        }
      );
    }
  }

  async function sendNotifications(assignments) {
    const now = Date.now();
    const next = { ...state.notified };
    const active = new Set();

    assignments.forEach((assignment) => {
      const due = assignment.dueDate ? new Date(assignment.dueDate) : parseDueDate(assignment.dueText);
      if (!due || Number.isNaN(due.getTime())) {
        return;
      }

      const diff = due.getTime() - now;
      if (diff <= 0) {
        return;
      }

      const base = `${assignment.link}|${assignment.dueDate || assignment.dueText}`;
      const id24 = `${base}:24h`;
      const id2 = `${base}:2h`;
      active.add(id24);
      active.add(id2);

      if (diff <= 2 * 60 * 60 * 1000 && !next[id2]) {
        notify(`${assignment.title} is due soon`);
        next[id2] = true;
        next[id24] = true;
      } else if (diff <= 24 * 60 * 60 * 1000 && !next[id24]) {
        notify(`${assignment.title} is due soon`);
        next[id24] = true;
      }
    });

    Object.keys(next).forEach((key) => {
      if (!active.has(key)) {
        delete next[key];
      }
    });

    await saveNotified(next);
  }

  async function fetchAssignments() {
    if (!isMoodlePage() || !isDashboardPage() || isLikelyLoggedOut()) {
      return state.assignments;
    }

    const fetched = extractAssignmentsFromPage(document);
    if (fetched.length === 0 && state.assignments.length > 0) {
      await sendNotifications(state.assignments);
      await saveLastSyncedAt(Date.now());
      state.hasFreshMoodleData = true;
      return state.assignments;
    }

    const saved = await saveAssignments(fetched);
    await sendNotifications(saved);
    await saveLastSyncedAt(Date.now());
    state.hasFreshMoodleData = true;
    return saved;
  }

  async function fetchCourses() {
    if (!isMoodlePage() || !isCoursesPage() || isLikelyLoggedOut()) {
      return state.courses;
    }

    const fetched = extractCoursesFromPage(document);
    if (fetched.length === 0 && state.courses.length > 0) {
      await saveLastSyncedAt(Date.now());
      state.hasFreshMoodleData = true;
      return state.courses;
    }

    const saved = await saveCourses(fetched);
    await saveLastSyncedAt(Date.now());
    state.hasFreshMoodleData = true;
    return saved;
  }

  function getStatusText() {
    if (state.isSyncing) {
      return "🔄 Syncing...";
    }

    if (isMoodlePage() && state.hasFreshMoodleData) {
      return "🟢 Connected";
    }

    return "🟡 Using saved data";
  }

  function formatLastSynced(timestamp) {
    if (!timestamp) {
      return "Last synced: Never";
    }

    const diff = Date.now() - timestamp;
    if (diff < 60 * 1000) {
      return "Last synced: Just now";
    }

    const mins = Math.floor(diff / (60 * 1000));
    if (mins < 60) {
      return `Last synced: ${mins} min${mins === 1 ? "" : "s"} ago`;
    }

    const hours = Math.floor(mins / 60);
    if (hours < 24) {
      return `Last synced: ${hours} hour${hours === 1 ? "" : "s"} ago`;
    }

    return `Last synced: ${new Date(timestamp).toLocaleString()}`;
  }

  function getHelperRenderKey() {
    if (!isMoodlePage()) {
      return "outside-moodle";
    }

    return isLikelyLoggedOut() ? "logged-out" : "logged-in";
  }

  function getAssignmentsRenderKey() {
    const bucket = Math.floor(Date.now() / (RENDER_TIME_BUCKET_MINUTES * 60 * 1000));
    const payload = state.assignments
      .map((item) => `${item.link}|${item.dueDate || ""}|${item.dueText || ""}`)
      .join("~");
    return `${bucket}|${payload}`;
  }

  function getCoursesRenderKey() {
    return state.courses.map((item) => `${item.title}|${item.link}`).join("~");
  }

  function createUi() {
    if (document.getElementById(ROOT_ID)) {
      return null;
    }

    const root = document.createElement("div");
    root.id = ROOT_ID;

    const container = document.createElement("div");
    container.id = "moodle-assistant-container";

    const panel = document.createElement("div");
    panel.id = "moodle-assistant-panel";
    panel.setAttribute("aria-hidden", "true");

    const panelContent = document.createElement("div");
    panelContent.id = "moodle-assistant-content";
    panelContent.className = "ma-panel-content";

    const header = document.createElement("div");
    header.className = "ma-panel-header";

    const title = document.createElement("p");
    title.className = "ma-panel-heading";
    title.textContent = "MarianMoodleMate 🎓";

    const refreshButton = document.createElement("button");
    refreshButton.type = "button";
    refreshButton.className = "ma-refresh-btn";
    refreshButton.title = "Refresh";
    refreshButton.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M20 12a8 8 0 1 1-2.34-5.66"/><path d="M20 4v6h-6"/></svg>';

    const statusText = document.createElement("p");
    statusText.className = "ma-status-text ma-status-main";

    const lastSyncedText = document.createElement("p");
    lastSyncedText.className = "ma-status-text ma-status-sub";

    const helperArea = document.createElement("div");
    helperArea.className = "ma-helper-area";

    const assignmentsSection = document.createElement("section");
    assignmentsSection.className = "ma-section";

    const assignmentsTitle = document.createElement("p");
    assignmentsTitle.className = "ma-section-title";
    assignmentsTitle.textContent = "📌 Assignments";

    const assignmentsBody = document.createElement("div");
    assignmentsBody.className = "ma-list-scroll ma-assignments-body";

    const coursesSection = document.createElement("section");
    coursesSection.className = "ma-section";

    const coursesTitle = document.createElement("p");
    coursesTitle.className = "ma-section-title";
    coursesTitle.textContent = "📚 Courses";

    const coursesBody = document.createElement("div");
    coursesBody.className = "ma-list-scroll ma-courses-body";

    header.append(title, refreshButton);
    assignmentsSection.append(assignmentsTitle, assignmentsBody);
    coursesSection.append(coursesTitle, coursesBody);
    panelContent.append(header, statusText, lastSyncedText, helperArea, assignmentsSection, coursesSection);
    panel.appendChild(panelContent);


    const button = document.createElement("button");
    button.id = "moodle-assistant-btn";
    button.type = "button";
    button.setAttribute("aria-label", "Toggle MarianMoodleMate panel");

    // Create a square container for the icon to enforce aspect ratio
    const iconWrap = document.createElement("span");
    iconWrap.className = "ma-fab-icon-wrap";

    const img = document.createElement("img");
    img.className = "ma-fab-icon";
    img.src = chrome.runtime.getURL(ICON_PATH);
    img.alt = "MarianMoodleMate";
    img.addEventListener("error", () => {
      button.textContent = "M";
    });

    iconWrap.appendChild(img);
    button.appendChild(iconWrap);

    container.append(panel, button);
    root.appendChild(container);
    document.body.appendChild(root);

    return {
      root,
      panel,
      button,
      refreshButton,
      statusText,
      lastSyncedText,
      helperArea,
      assignmentsBody,
      coursesBody,
    };
  }

  function setPanelOpen(ui, open) {
    state.panelOpen = open;
    ui.panel.classList.toggle("is-open", open);
    ui.panel.setAttribute("aria-hidden", String(!open));
  }

  function renderHelper(ui) {
    const key = getHelperRenderKey();
    if (state.renderCache.helperKey === key) {
      return;
    }

    state.renderCache.helperKey = key;
    ui.helperArea.replaceChildren();

    if (!isMoodlePage()) {
      const info = document.createElement("p");
      info.className = "ma-status-text ma-helper-text";
      info.textContent = "Using stored data outside Moodle.";
      ui.helperArea.appendChild(info);
      return;
    }

    if (!isLikelyLoggedOut()) {
      return;
    }

    const warning = document.createElement("p");
    warning.className = "ma-warning-text";
    warning.append("⚠️ ");

    const loginLink = document.createElement("a");
    loginLink.className = "ma-login-link";
    loginLink.href = LOGIN_URL;
    loginLink.target = "_blank";
    loginLink.rel = "noopener noreferrer";
    loginLink.textContent = "Login to Moodle";

    warning.append(loginLink, " to fetch latest information");
    ui.helperArea.appendChild(warning);
  }

  function renderAssignments(ui) {
    const key = getAssignmentsRenderKey();
    if (state.renderCache.assignmentsKey === key) {
      return;
    }

    state.renderCache.assignmentsKey = key;
    ui.assignmentsBody.replaceChildren();

    if (state.assignments.length === 0) {
      const empty = document.createElement("p");
      empty.className = "ma-status-text ma-empty-state";
      empty.textContent = "No assignments found";
      ui.assignmentsBody.appendChild(empty);
      return;
    }

    const sorted = [...state.assignments].sort((a, b) => {
      const aTime = a.dueDate ? new Date(a.dueDate).getTime() : Number.MAX_SAFE_INTEGER;
      const bTime = b.dueDate ? new Date(b.dueDate).getTime() : Number.MAX_SAFE_INTEGER;
      return aTime - bTime;
    });

    const fragment = document.createDocumentFragment();
    const list = document.createElement("ul");
    list.className = "ma-assignments-list";

    sorted.forEach((assignment) => {
      const progress = calculateProgress(assignment);

      const item = document.createElement("li");
      item.className = "ma-assignment-card";

      const link = document.createElement("a");
      link.className = "ma-assignment-link";
      link.href = assignment.link;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = assignment.title;

      const timeText = document.createElement("p");
      timeText.className = "ma-time-remaining";
      timeText.textContent = calculateTimeRemaining(assignment);

      const progressRow = document.createElement("div");
      progressRow.className = "ma-progress-row";

      const track = document.createElement("div");
      track.className = "ma-progress-track";

      const fill = document.createElement("div");
      fill.className = "ma-progress-fill";
      fill.style.width = `${progress.percent}%`;
      fill.style.backgroundColor = progress.color;

      const label = document.createElement("span");
      label.className = "ma-progress-label";
      label.textContent = `${progress.percent}%`;

      track.appendChild(fill);
      progressRow.append(track, label);

      item.append(link, timeText, progressRow);
      list.appendChild(item);
    });

    fragment.appendChild(list);
    ui.assignmentsBody.appendChild(fragment);
  }

  function renderCourses(ui) {
    const key = getCoursesRenderKey();
    if (state.renderCache.coursesKey === key) {
      return;
    }

    state.renderCache.coursesKey = key;
    ui.coursesBody.replaceChildren();

    if (state.courses.length === 0) {
      const empty = document.createElement("p");
      empty.className = "ma-status-text ma-empty-state";
      empty.textContent = "No courses available";
      ui.coursesBody.appendChild(empty);
      return;
    }

    const fragment = document.createDocumentFragment();
    const list = document.createElement("ul");
    list.className = "ma-course-list";

    state.courses.forEach((course, idx) => {
      const item = document.createElement("li");
      item.className = "ma-course-item";

      const link = document.createElement("a");
      link.className = "ma-course-link";
      link.href = course.link;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = course.title;

      // Pencil (edit) icon
      const editBtn = document.createElement("button");
      editBtn.className = "ma-course-rename-btn";
      editBtn.title = "Rename subject";
      editBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z"/></svg>';
      editBtn.style.background = "none";
      editBtn.style.border = "none";
      editBtn.style.cursor = "pointer";
      editBtn.style.padding = "2px";
      editBtn.style.marginLeft = "8px";
      editBtn.style.verticalAlign = "middle";

      editBtn.addEventListener("click", (e) => {
        e.preventDefault();
        // Replace link with input
        const input = document.createElement("input");
        input.type = "text";
        input.value = course.title;
        input.className = "ma-course-rename-input";
        input.style.fontSize = "13px";
        input.style.width = "70%";
        input.style.marginRight = "8px";
        input.style.border = "1px solid #cbd5e1";
        input.style.borderRadius = "6px";
        input.style.padding = "2px 6px";
        input.style.outline = "none";

        item.replaceChild(input, link);
        input.focus();
        input.select();

        function finishEdit(save) {
          if (save) {
            const newTitle = input.value.trim();
            if (newTitle && newTitle !== course.title) {
              state.courses[idx].title = newTitle;
              storageSet({ [STORAGE_KEYS.courses]: state.courses });
            }
          }
          renderCourses(ui);
        }

        input.addEventListener("blur", () => finishEdit(true));
        input.addEventListener("keydown", (ev) => {
          if (ev.key === "Enter") finishEdit(true);
          if (ev.key === "Escape") finishEdit(false);
        });
      });

      item.appendChild(link);
      item.appendChild(editBtn);
      list.appendChild(item);
    });

    fragment.appendChild(list);
    ui.coursesBody.appendChild(fragment);
  }

  function renderPanel(ui) {
    const statusText = getStatusText();
    if (state.renderCache.statusText !== statusText) {
      state.renderCache.statusText = statusText;
      ui.statusText.textContent = statusText;
    }

    const lastSynced = formatLastSynced(state.lastSyncedAt);
    if (state.renderCache.lastSyncedText !== lastSynced) {
      state.renderCache.lastSyncedText = lastSynced;
      ui.lastSyncedText.textContent = lastSynced;
    }

    ui.refreshButton.disabled = state.isSyncing;
    renderHelper(ui);
    renderAssignments(ui);
    renderCourses(ui);
  }

  async function loadPanelData() {
    await Promise.all([loadAssignments(), loadCourses(), loadLastSyncedAt(), loadNotified()]);
  }

  async function syncCurrentPageData(ui) {
    if (!isMoodlePage() || isLikelyLoggedOut()) {
      state.hasFreshMoodleData = false;
      return false;
    }

    state.isSyncing = true;
    if (ui && state.panelOpen) {
      renderPanel(ui);
    }

    let updated = false;

    try {
      if (isDashboardPage()) {
        await fetchAssignments();
        updated = true;
      } else if (isCoursesPage()) {
        await fetchCourses();
        updated = true;
      }

      if (updated) {
        state.hasFreshMoodleData = true;
      }

      return updated;
    } finally {
      state.isSyncing = false;
      if (ui && state.panelOpen) {
        renderPanel(ui);
      }
    }
  }

  function configureObserver(ui) {
    if (state.observer) {
      state.observer.disconnect();
      state.observer = null;
    }

    if (!isMoodlePage() || (!isDashboardPage() && !isCoursesPage())) {
      return;
    }

    if (!state.cachedRegionMain || !document.contains(state.cachedRegionMain)) {
      state.cachedRegionMain = document.querySelector("#region-main");
    }

    const target = state.cachedRegionMain || document.body;
    let pending = false;

    state.observer = new MutationObserver(() => {
      if (pending) {
        return;
      }

      pending = true;
      requestAnimationFrame(async () => {
        pending = false;
        await syncCurrentPageData(ui);
      });
    });

    state.observer.observe(target, { childList: true, subtree: true });
  }

  function installNavigationHooks(onNavigate) {
    if (!window.__marianMoodleMateNavHookInstalled) {
      window.__marianMoodleMateNavHookInstalled = true;

      const patch = (method) => {
        const original = history[method];
        history[method] = function (...args) {
          const result = original.apply(this, args);
          window.dispatchEvent(new Event(NAV_EVENT));
          return result;
        };
      };

      patch("pushState");
      patch("replaceState");
    }

    window.addEventListener("popstate", onNavigate);
    window.addEventListener(NAV_EVENT, onNavigate);
  }

  function startAutoSync(ui) {
    if (state.autoSyncTimer) {
      window.clearInterval(state.autoSyncTimer);
    }

    state.autoSyncTimer = window.setInterval(async () => {
      if (!isMoodlePage()) {
        return;
      }

      await syncCurrentPageData(ui);
      if (state.panelOpen) {
        await loadPanelData();
        renderPanel(ui);
      }
    }, AUTO_SYNC_MS);
  }

  function startDrag(ui, event) {
    if (event.button !== 0) {
      return;
    }

    state.drag.pointerId = event.pointerId;
    state.drag.originX = event.clientX;
    state.drag.originY = event.clientY;

    const rect = ui.root.getBoundingClientRect();
    state.drag.originLeft = rect.left;
    state.drag.originTop = rect.top;
    state.drag.moved = false;

    ui.button.setPointerCapture(event.pointerId);
    ui.root.classList.add("is-dragging");
  }

  function moveDrag(ui, event) {
    if (event.pointerId !== state.drag.pointerId) {
      return;
    }

    const deltaX = event.clientX - state.drag.originX;
    const deltaY = event.clientY - state.drag.originY;

    if (!state.drag.moved && Math.hypot(deltaX, deltaY) < 4) {
      return;
    }

    if (!state.drag.moved) {
      state.drag.moved = true;
      ui.root.style.right = "auto";
      ui.root.style.bottom = "auto";
    }

    const maxLeft = Math.max(12, window.innerWidth - ui.root.offsetWidth - 12);
    const maxTop = Math.max(12, window.innerHeight - ui.root.offsetHeight - 12);
    const nextLeft = Math.min(Math.max(12, state.drag.originLeft + deltaX), maxLeft);
    const nextTop = Math.min(Math.max(12, state.drag.originTop + deltaY), maxTop);

    ui.root.style.left = `${nextLeft}px`;
    ui.root.style.top = `${nextTop}px`;
  }

  async function stopDrag(ui, event) {
    if (event.pointerId !== state.drag.pointerId) {
      return;
    }

    if (ui.button.hasPointerCapture(event.pointerId)) {
      ui.button.releasePointerCapture(event.pointerId);
    }

    if (state.drag.moved) {
      state.drag.suppressClick = true;
      const rect = ui.root.getBoundingClientRect();
      await saveButtonPosition({ left: Math.round(rect.left), top: Math.round(rect.top) });
      setTimeout(() => {
        state.drag.suppressClick = false;
      }, 0);
    }

    state.drag.pointerId = null;
    state.drag.moved = false;
    ui.root.classList.remove("is-dragging");
  }

  async function applySavedPosition(ui) {
    const pos = await loadButtonPosition();
    if (!pos) {
      return;
    }

    ui.root.style.left = `${pos.left}px`;
    ui.root.style.top = `${pos.top}px`;
    ui.root.style.right = "auto";
    ui.root.style.bottom = "auto";
  }

  function attachEvents(ui) {
    ui.button.addEventListener("click", async () => {
      if (state.drag.suppressClick) {
        return;
      }

      const open = !state.panelOpen;
      setPanelOpen(ui, open);

      if (!open) {
        return;
      }

      await loadPanelData();
      renderPanel(ui);
      await syncCurrentPageData(ui);
      await loadPanelData();
      renderPanel(ui);
    });

    document.addEventListener("pointerdown", (event) => {
      if (event.target instanceof Node && !ui.root.contains(event.target)) {
        setPanelOpen(ui, false);
      }
    });

    ui.button.addEventListener("pointerdown", (event) => startDrag(ui, event));
    ui.button.addEventListener("pointermove", (event) => moveDrag(ui, event));
    ui.button.addEventListener("pointerup", (event) => {
      void stopDrag(ui, event);
    });
    ui.button.addEventListener("pointercancel", (event) => {
      void stopDrag(ui, event);
    });

    ui.refreshButton.addEventListener("click", () => {
      if (state.refreshDebounceId) {
        window.clearTimeout(state.refreshDebounceId);
      }

      state.refreshDebounceId = window.setTimeout(() => {
        state.refreshDebounceId = null;
        const opened = window.open(MOODLE_DASHBOARD_URL, "_blank", "noopener,noreferrer");
        if (!opened) {
          window.location.assign(MOODLE_DASHBOARD_URL);
        }
      }, 120);
    });

    installNavigationHooks(async () => {
      await loadPanelData();
      await syncCurrentPageData(ui);
      configureObserver(ui);
      if (state.panelOpen) {
        await loadPanelData();
        renderPanel(ui);
      }
    });
  }

  async function init() {
    const ui = createUi();
    if (!ui) {
      return;
    }

    await applySavedPosition(ui);
    await loadPanelData();
    state.hasFreshMoodleData = isMoodlePage() && state.lastSyncedAt > 0;
    renderPanel(ui);

    attachEvents(ui);
    configureObserver(ui);
    startAutoSync(ui);

    await syncCurrentPageData(ui);
    await loadPanelData();
    renderPanel(ui);
  }

  if (document.body) {
    void init();
  } else {
    document.addEventListener(
      "DOMContentLoaded",
      () => {
        void init();
      },
      { once: true }
    );
  }
})();
