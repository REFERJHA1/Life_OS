const a_U_sK = "a_u";
const a_U_legacy_sK = "lifeos_api_url";
const a_U_b64 = "aHR0cHM6Ly9zY3JpcHQuZ29vZ2xlLmNvbS9tYWNyb3Mvcy9BS2Z5Y2J6Y21wUDJBaVNhZWMtMmlxdnFLVVZzbmxnMEY5RDZHRGJraENNeFZfUlBudXB1RFhhbEplZ2xEeFN3S1ZxSU0tekN0US9leGVj";
const d_K = "lifeos_selected_date";
const REQUEST_TIMEOUT_MS = 15000;

const a_U = r_aU_();

function r_aU_() {
  try {
    const custom = localStorage.getItem(a_U_sK);
    if (custom && String(custom).trim()) {
      return String(custom).trim();
    }

    const legacy = localStorage.getItem(a_U_legacy_sK);
    if (legacy && String(legacy).trim()) {
      const migrated = String(legacy).trim();
      localStorage.setItem(a_U_sK, migrated);
      return migrated;
    }
  } catch (error) {
  }

  return decodeBase64_(a_U_b64);
}

function decodeBase64_(value) {
  try {
    return atob(value);
  } catch (error) {
    return "";
  }
}

function currentDateText_() {
  const value = new Date();
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getWorkingDate_(overrideDate) {
  if (overrideDate && String(overrideDate).trim()) {
    return String(overrideDate).trim();
  }

  try {
    const stored = localStorage.getItem(d_K);
    if (stored && String(stored).trim()) {
      return String(stored).trim();
    }
  } catch (error) {
  }

  return currentDateText_();
}

async function request(method, payload) {
  if (!a_U) {
    throw new Error("API URL is missing. Set a_u in localStorage.");
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  let response;

  try {
    response = await fetch(a_U, {
      method,
      body: payload ? JSON.stringify(payload) : undefined,
      signal: controller.signal
    });
  } catch (error) {
    if (error && error.name === "AbortError") {
      throw new Error("Request timed out. Please try again.");
    }
    throw error;
  } finally {
    clearTimeout(timeoutId);
  }

  if (!response.ok) {
    throw new Error(`API request failed: ${response.status}`);
  }

  const text = await response.text();
  let data;

  try {
    data = JSON.parse(text);
  } catch (error) {
    throw new Error("Invalid API response format");
  }

  if (data && data.ok === false) {
    throw new Error(data.error || "API returned an error");
  }

  return data;
}

export async function getDashboardData(date) {
  return request("POST", { action: "getDashboardData", date: getWorkingDate_(date) });
}

export async function getTodayTasks(date) {
  return request("POST", { action: "getTodayTasks", date: getWorkingDate_(date) });
}

export async function getWorkTasks() {
  return request("POST", { action: "getWorkTasks" });
}

export async function getPersonalTasks() {
  return request("POST", { action: "getPersonalTasks" });
}

export async function getJournalEntries() {
  return request("POST", { action: "getJournalEntries" });
}

export async function getJournalHistory(limit = 7) {
  return request("POST", { action: "getJournalHistory", limit });
}

export async function getAnalytics(date) {
  return request("POST", { action: "getAnalytics", date: getWorkingDate_(date) });
}

export async function createWeeklyBackup(date) {
  return request("POST", { action: "createWeeklyBackup", date: getWorkingDate_(date) });
}

export async function getTimeTable(date = "") {
  return request("POST", { action: "getTimeTable", date });
}

export async function addWorkTask(task) {
  return request("POST", { action: "addWorkTask", task });
}

export async function addPersonalTask(task) {
  return request("POST", { action: "addPersonalTask", task });
}

export async function addTimeTableEntry(entry) {
  return request("POST", { action: "addTimeTableEntry", entry });
}

export async function saveHabit(payload) {
  return request("POST", { action: "saveHabit", ...payload });
}

export async function saveJournal(payload) {
  return request("POST", { action: "saveJournal", ...payload });
}

export async function saveDailyReview(payload) {
  return request("POST", { action: "saveDailyReview", ...payload });
}

export async function dailyCheckin(payload) {
  return request("POST", { action: "dailyCheckin", ...payload });
}

export async function completeWorkTask(payload) {
  return request("POST", { action: "completeWorkTask", ...payload });
}

export async function completePersonalTask(payload) {
  return request("POST", { action: "completePersonalTask", ...payload });
}

export async function completeTimeTableEntry(payload) {
  return request("POST", { action: "completeTimeTableEntry", ...payload });
}
