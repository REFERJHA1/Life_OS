import { addTimeTableEntry, completeTimeTableEntry, getTimeTable } from "../services/common.js";
import { renderStreakBadge } from "./streak.js";

const DATE_KEY = "lifeos_selected_date";

const form = document.getElementById("timetableForm");
const messageEl = document.getElementById("ttMessage");
const dateInput = document.getElementById("ttDate");
const todayPendingList = document.getElementById("ttTodayPendingList");
const otherList = document.getElementById("ttOtherList");
const doneList = document.getElementById("ttDoneList");

dateInput.value = getLocalDateOffset_(1);

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  setMessage("Saving time block...", false);

  const payload = {
    date: document.getElementById("ttDate").value,
    startTime: document.getElementById("ttStartTime").value,
    endTime: document.getElementById("ttEndTime").value,
    block: document.getElementById("ttBlock").value.trim(),
    type: document.getElementById("ttType").value,
    linkedTask: document.getElementById("ttLinkedTask").value.trim(),
    status: document.getElementById("ttStatus").value,
    remark: document.getElementById("ttRemark").value.trim()
  };

  if (payload.startTime >= payload.endTime) {
    setMessage("End time must be after start time.", true);
    return;
  }

  try {
    await addTimeTableEntry(payload);
    form.reset();
    dateInput.value = getLocalDateOffset_(1);
    setMessage("Time block added.", false);
    await loadTimeTableSections();
  } catch (error) {
    setMessage(`Failed to add time block: ${error.message}`, true);
  }
});

async function loadTimeTableSections() {
  setLoadingState_([todayPendingList, otherList, doneList], "Loading timetable...");

  try {
    const data = await getTimeTable();
    const entries = data.entries || [];
    const today = getWorkingDate_();

    const todayPending = entries.filter((entry) => {
      return entry.date === today && normalizeStatus_(entry.status) === "pending";
    });
    const otherEntries = entries.filter((entry) => normalizeStatus_(entry.status) === "other");
    const doneEntries = entries.filter((entry) => isDoneStatusValue_(entry.status));

    renderList_(todayPendingList, todayPending, "No pending time blocks for today.", (entry) => createTimeBlockRow_(entry, true));
    renderList_(otherList, otherEntries, "No other time blocks.", (entry) => createTimeBlockRow_(entry, true));
    renderList_(doneList, doneEntries, "No completed time blocks.", (entry) => {
      const li = document.createElement("li");
      li.textContent = formatTimeBlockText_(entry);
      return li;
    });
  } catch (error) {
    [todayPendingList, otherList, doneList].forEach((list) => {
      list.innerHTML = "";
      const li = document.createElement("li");
      li.textContent = "Could not load timetable.";
      list.appendChild(li);
    });
    setMessage(`Failed to load timetable: ${error.message}`, true);
  }
}

function setLoadingState_(lists, text) {
  lists.forEach((list) => {
    list.innerHTML = "";
    const li = document.createElement("li");
    li.textContent = text;
    list.appendChild(li);
  });
}

function createTimeBlockRow_(entry, includeDoneButton) {
  const li = document.createElement("li");

  if (!includeDoneButton) {
    li.textContent = formatTimeBlockText_(entry);
    return li;
  }

  const row = document.createElement("div");
  row.className = "task-row";

  const text = document.createElement("span");
  text.textContent = formatTimeBlockText_(entry);

  const doneBtn = document.createElement("button");
  doneBtn.type = "button";
  doneBtn.className = "task-done-btn";
  doneBtn.textContent = "Done";
  doneBtn.addEventListener("click", async () => {
    doneBtn.disabled = true;
    doneBtn.textContent = "...";

    try {
      await completeTimeTableEntry({
        date: entry.date,
        block: entry.block,
        startTime: entry.startTime
      });
      await loadTimeTableSections();
    } catch (error) {
      setMessage(`Failed to complete time block: ${error.message}`, true);
    } finally {
      doneBtn.disabled = false;
      doneBtn.textContent = "Done";
    }
  });

  row.appendChild(text);
  row.appendChild(doneBtn);
  li.appendChild(row);
  return li;
}

function formatTimeBlockText_(entry) {
  const linkedTask = String(entry.linkedTask || "").trim();
  const remark = String(entry.remark || "").trim();
  const extras = [
    linkedTask ? `Task: ${linkedTask}` : "",
    remark ? `Remark: ${remark}` : ""
  ].filter(Boolean).join(" | ");

  return `${entry.date} | ${entry.startTime}-${entry.endTime} | ${entry.block} | ${entry.type}${extras ? ` | ${extras}` : ""}`;
}

function renderList_(listElement, items, emptyText, renderItem) {
  listElement.innerHTML = "";

  if (!items.length) {
    const li = document.createElement("li");
    li.textContent = emptyText;
    listElement.appendChild(li);
    return;
  }

  items.forEach((item) => {
    listElement.appendChild(renderItem(item));
  });
}

function normalizeStatus_(status) {
  return String(status || "").trim().toLowerCase();
}

function isDoneStatusValue_(status) {
  const value = normalizeStatus_(status);
  return value === "done" || value === "completed" || value === "complete";
}

function getLocalDateOffset_(daysAhead) {
  const value = new Date();
  value.setDate(value.getDate() + Number(daysAhead || 0));

  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getWorkingDate_() {
  try {
    const value = localStorage.getItem(DATE_KEY);
    if (value && String(value).trim()) {
      return String(value).trim();
    }
  } catch (error) {
  }

  return getLocalDateOffset_(0);
}

function setMessage(text, isError) {
  messageEl.textContent = text;
  messageEl.classList.toggle("error", isError);
}

loadTimeTableSections();
renderStreakBadge();

window.addEventListener("lifeos:date-change", () => {
  loadTimeTableSections();
  renderStreakBadge();
});
