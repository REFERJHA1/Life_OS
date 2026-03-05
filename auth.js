(function () {
  const STORAGE_KEY = "lifeos_auth_state";
  const DATE_KEY = "lifeos_selected_date";
  const MODERATOR_PASSWORD = "Modern@456";
  const USERS = [
    { value: "viewer", label: "Viewer (View Only)" },
    { value: "moderator", label: "Moderator" },
    { value: "other", label: "Other" }
  ];

  function isAuthPage_() {
    return /\/auth\.html$/i.test(window.location.pathname);
  }

  function rootPrefix_() {
    return window.location.pathname.indexOf("/pages/") !== -1 ? "../" : "./";
  }

  function getAuthUrl_() {
    return rootPrefix_() + "auth.html";
  }

  function getIndexUrl_() {
    return rootPrefix_() + "index.html";
  }

  function redirectToAuth_() {
    const next = encodeURIComponent(window.location.pathname + window.location.search + window.location.hash);
    window.location.href = getAuthUrl_() + "?next=" + next;
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return { signedIn: false, user: "viewer", role: "viewer" };
      const parsed = JSON.parse(raw);
      if (!parsed || !parsed.user || !parsed.role) {
        return { signedIn: false, user: "viewer", role: "viewer" };
      }

      if (typeof parsed.signedIn !== "boolean") {
        parsed.signedIn = true;
      }

      return parsed;
    } catch (error) {
      return { signedIn: false, user: "viewer", role: "viewer" };
    }
  }

  function saveState(state) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function isAdmin(state) {
    return String(state.role || "").toLowerCase() === "moderator";
  }

  function clearState_() {
    localStorage.removeItem(STORAGE_KEY);
  }

  function currentDateText_() {
    const value = new Date();
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function getWorkingDate_() {
    try {
      const saved = localStorage.getItem(DATE_KEY);
      if (saved && String(saved).trim()) {
        return String(saved).trim();
      }
    } catch (error) {
    }

    return currentDateText_();
  }

  function setWorkingDate_(dateText) {
    const value = String(dateText || "").trim();
    if (!value) return;

    localStorage.setItem(DATE_KEY, value);
    window.dispatchEvent(new CustomEvent("lifeos:date-change", {
      detail: { date: value }
    }));
  }

  function renderHeaderDatePicker_() {
    const header = document.querySelector(".page-header");
    if (!header) return;
    if (document.getElementById("globalDateInput") && document.getElementById("headerSignOutBtn")) return;

    let rightControls = document.getElementById("headerRightControls");
    if (!rightControls) {
      rightControls = document.createElement("div");
      rightControls.id = "headerRightControls";
      rightControls.className = "header-right-controls";
      header.appendChild(rightControls);
    }

    const wrap = document.createElement("div");
    wrap.className = "header-date-wrap";
    wrap.innerHTML = `
      <label for="globalDateInput" class="header-date-label">Working Date</label>
      <input id="globalDateInput" class="header-date-input" type="date">
    `;

    const signOutBtn = document.createElement("button");
    signOutBtn.id = "headerSignOutBtn";
    signOutBtn.type = "button";
    signOutBtn.className = "header-signout-btn";
    signOutBtn.textContent = "Sign Out";

    rightControls.appendChild(wrap);
    rightControls.appendChild(signOutBtn);

    const input = document.getElementById("globalDateInput");
    input.value = getWorkingDate_();

    input.addEventListener("change", function () {
      const dateText = String(input.value || "").trim();
      if (!dateText) return;

      setWorkingDate_(dateText);
    });

    window.addEventListener("lifeos:date-change", function (event) {
      const nextDate = event && event.detail && event.detail.date ? String(event.detail.date) : getWorkingDate_();
      if (input.value !== nextDate) {
        input.value = nextDate;
      }
    });

    signOutBtn.addEventListener("click", function () {
      clearState_();
      window.location.href = getAuthUrl_();
    });
  }

  function authenticate_(userValue, enteredPassword) {
    if (userValue === "moderator") {
      if (enteredPassword === MODERATOR_PASSWORD) {
        return { ok: true, state: { signedIn: true, user: "moderator", role: "moderator" } };
      }

      return { ok: false, message: "Wrong password." };
    }

    if (userValue === "other") {
      return { ok: false, message: "Wrong password." };
    }

    return { ok: true, state: { signedIn: true, user: userValue, role: "viewer" } };
  }

  function getNextUrl_() {
    const params = new URLSearchParams(window.location.search);
    const nextRaw = params.get("next");
    if (!nextRaw) return getIndexUrl_();

    const next = decodeURIComponent(nextRaw);
    if (/^https?:\/\//i.test(next)) return getIndexUrl_();
    if (next.indexOf(".html") === -1 && next.indexOf("/") !== 0) return getIndexUrl_();

    return next;
  }

  function renderAuthPage_(state) {
    const mount = document.getElementById("authPageMount");
    if (!mount) return;

    mount.innerHTML = `
      <section class="card">
        <h2>Sign In</h2>
        <p class="muted">Choose user mode before opening the dashboard.</p>
        <div class="stack-form">
          <label class="auth-label" for="authUserSelect">User</label>
          <select id="authUserSelect" class="auth-control">
            ${USERS.map((user) => `<option value="${user.value}">${user.label}</option>`).join("")}
          </select>
          <label class="auth-label" for="authPassword">Password</label>
          <input id="authPassword" class="auth-control" type="password" placeholder="Password (if required)">
          <button id="authLoginBtn" type="button" class="button">Continue</button>
          <button id="authViewBtn" type="button" class="auth-btn-secondary">Continue as View Only</button>
          <p id="authStatus" class="auth-status" aria-live="polite"></p>
        </div>
      </section>
    `;

    const userSelect = document.getElementById("authUserSelect");
    const passwordInput = document.getElementById("authPassword");
    const loginBtn = document.getElementById("authLoginBtn");
    const viewBtn = document.getElementById("authViewBtn");
    const statusEl = document.getElementById("authStatus");

    if (state && state.signedIn) {
      userSelect.value = state.role === "moderator" ? "moderator" : state.user;
      statusEl.textContent = isAdmin(state)
        ? "Already signed in as Moderator."
        : `Already signed in as ${state.user}.`;
    }

    loginBtn.addEventListener("click", function () {
      const userValue = String(userSelect.value || "viewer").toLowerCase();
      const enteredPassword = String(passwordInput.value || "");
      const result = authenticate_(userValue, enteredPassword);

      if (!result.ok) {
        statusEl.textContent = result.message;
        return;
      }

      saveState(result.state);
      statusEl.textContent = "Login successful. Redirecting...";
      window.location.href = getNextUrl_();
    });

    viewBtn.addEventListener("click", function () {
      const nextState = { signedIn: true, user: "viewer", role: "viewer" };
      saveState(nextState);
      statusEl.textContent = "View mode selected. Redirecting...";
      window.location.href = getNextUrl_();
    });
  }

  function setReadOnlyControls(readOnly) {
    const controls = document.querySelectorAll(
      "main input, main select, main textarea, main button[type='submit'], main .task-done-btn"
    );

    controls.forEach(function (element) {
      element.disabled = readOnly;
    });
  }

  function applyAuthMode(state) {
    const readOnly = !isAdmin(state);
    document.body.classList.toggle("read-only-mode", readOnly);
    setReadOnlyControls(readOnly);

    if (readOnly) {
      document.body.setAttribute("data-access-mode", "view");
    } else {
      document.body.setAttribute("data-access-mode", "admin");
    }
  }

  document.addEventListener("submit", function (event) {
    const readOnly = document.body.getAttribute("data-access-mode") !== "admin";
    if (!readOnly) return;

    if (event.target && event.target.closest("main form")) {
      event.preventDefault();
      const statusEl = document.getElementById("authStatus");
      if (statusEl) statusEl.textContent = "Read-only mode: switch to Moderator to edit.";
    }
  }, true);

  document.addEventListener("click", function (event) {
    const readOnly = document.body.getAttribute("data-access-mode") !== "admin";
    if (!readOnly) return;

    const doneButton = event.target && event.target.closest(".task-done-btn");
    if (doneButton) {
      event.preventDefault();
      event.stopPropagation();
      const statusEl = document.getElementById("authStatus");
      if (statusEl) statusEl.textContent = "Read-only mode: switch to Moderator to edit.";
    }
  }, true);

  const state = loadState();

  if (isAuthPage_()) {
    renderAuthPage_(state);
    return;
  }

  if (!state.signedIn) {
    redirectToAuth_();
    return;
  }

  renderHeaderDatePicker_();
  applyAuthMode(state);
})();
