import { addWorkTask, completeWorkTask, getWorkTasks } from "../services/common.js";
import { renderStreakBadge } from "./streak.js";

const DATE_KEY = "lifeos_selected_date";

const form = document.getElementById("workTaskForm");
const messageEl = document.getElementById("workMessage");
const todayList = document.getElementById("workTodayList");
const feedbackList = document.getElementById("workFeedbackList");
const allList = document.getElementById("workAllList");
const doneList = document.getElementById("workDoneList");
const dateInput = document.getElementById("workDate");

dateInput.value = getLocalDateOffset_(1);

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  setMessage("Saving task...", false);

  const payload = {
    date: document.getElementById("workDate").value,
    task: document.getElementById("workTask").value.trim(),
    priority: document.getElementById("workPriority").value,
    status: document.getElementById("workStatus").value,
    notes: document.getElementById("workNotes").value.trim(),
    remark: document.getElementById("workRemark").value.trim()
  };

  try {
    await addWorkTask(payload);
    form.reset();
    dateInput.value = getLocalDateOffset_(1);
    setMessage("Task added successfully.", false);
    await loadWorkSections();
  } catch (error) {
    setMessage(`Failed to save task: ${error.message}`, true);
  }
});

async function loadWorkSections() {
  setLoadingState_([todayList, feedbackList, allList, doneList], "Loading work tasks...");

  try {
    const data = await getWorkTasks();
    const tasks = data.tasks || [];
    const today = getWorkingDate_();
    const pendingToday = tasks.filter((item) => normalizeStatus_(item.status) === "pending" && item.date === today);
    const feedbackPending = tasks.filter((item) => normalizeStatus_(item.status) === "feedback pending");
    const otherTasks = tasks.filter((item) => normalizeStatus_(item.status) === "other");
    const doneTasks = tasks.filter((item) => isDoneStatusValue_(item.status));

    renderList_(todayList, pendingToday, "No pending work tasks for today.", (item) => {
      const li = document.createElement("li");
      const row = document.createElement("div");
      row.className = "task-row";

      const text = document.createElement("span");
      text.textContent = formatWorkTaskText_(item);

      const doneBtn = document.createElement("button");
      doneBtn.type = "button";
      doneBtn.className = "task-done-btn";
      doneBtn.textContent = "Done";
      doneBtn.addEventListener("click", async () => {
        doneBtn.disabled = true;
        doneBtn.textContent = "...";
        try {
          await completeWorkTask({ date: item.date, task: item.task });
          await loadWorkSections();
        } catch (error) {
          setMessage(`Failed to complete task: ${error.message}`, true);
        } finally {
          doneBtn.disabled = false;
          doneBtn.textContent = "Done";
        }
      });

      row.appendChild(text);
      row.appendChild(doneBtn);
      li.appendChild(row);
      return li;
    });

    renderList_(feedbackList, feedbackPending, "No feedback pending work tasks.", (item) => {
      return createWorkTaskRow_(item, true);
    });

    renderList_(allList, otherTasks, "No other work tasks.", (item) => {
      return createWorkTaskRow_(item, true);
    });

    renderList_(doneList, doneTasks, "No completed work tasks.", (item) => {
      const li = document.createElement("li");
      li.textContent = formatWorkTaskText_(item);
      return li;
    });
  } catch (error) {
    [todayList, feedbackList, allList, doneList].forEach((list) => {
      list.innerHTML = "";
      const li = document.createElement("li");
      li.textContent = "Could not load work tasks.";
      list.appendChild(li);
    });
    setMessage(`Failed to load all tasks: ${error.message}`, true);
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

function setMessage(text, isError) {
  messageEl.textContent = text;
  messageEl.classList.toggle("error", isError);
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

function normalizeStatus_(status) {
  return String(status || "").trim().toLowerCase();
}

function isDoneStatusValue_(status) {
  const value = normalizeStatus_(status);
  return value === "done" || value === "completed" || value === "complete";
}

function formatWorkTaskText_(item) {
  const notes = String(item.notes || "").trim();
  const remark = String(item.remark || "").trim();
  const extras = [
    notes ? `Notes: ${notes}` : "",
    remark ? `Remark: ${remark}` : ""
  ].filter(Boolean).join(" | ");

  return `${item.date} | ${item.task} | ${item.status} | ${item.priority}${extras ? ` | ${extras}` : ""}`;
}

function createWorkTaskRow_(item, includeDoneButton) {
  const li = document.createElement("li");

  if (!includeDoneButton) {
    li.textContent = formatWorkTaskText_(item);
    return li;
  }

  const row = document.createElement("div");
  row.className = "task-row";

  const text = document.createElement("span");
  text.textContent = formatWorkTaskText_(item);

  const doneBtn = document.createElement("button");
  doneBtn.type = "button";
  doneBtn.className = "task-done-btn";
  doneBtn.textContent = "Done";
  doneBtn.addEventListener("click", async () => {
    doneBtn.disabled = true;
    doneBtn.textContent = "...";
    try {
      await completeWorkTask({ date: item.date, task: item.task });
      await loadWorkSections();
    } catch (error) {
      setMessage(`Failed to complete task: ${error.message}`, true);
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

loadWorkSections();
renderStreakBadge();

window.addEventListener("lifeos:date-change", () => {
  loadWorkSections();
  renderStreakBadge();
});
