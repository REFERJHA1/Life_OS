import { addPersonalTask, completePersonalTask, getPersonalTasks } from "../services/common.js";
import { renderStreakBadge } from "./streak.js";

const DATE_KEY = "lifeos_selected_date";

const form = document.getElementById("personalTaskForm");
const messageEl = document.getElementById("personalMessage");
const todayList = document.getElementById("personalTodayList");
const feedbackList = document.getElementById("personalFeedbackList");
const allList = document.getElementById("personalAllList");
const doneList = document.getElementById("personalDoneList");
const dateInput = document.getElementById("personalDate");

dateInput.value = getLocalDateOffset_(1);

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  setMessage("Saving personal task...", false);

  const payload = {
    date: document.getElementById("personalDate").value,
    task: document.getElementById("personalTask").value.trim(),
    category: document.getElementById("personalCategory").value,
    status: document.getElementById("personalStatus").value,
    remark: document.getElementById("personalRemark").value.trim()
  };

  try {
    await addPersonalTask(payload);
    form.reset();
    dateInput.value = getLocalDateOffset_(1);
    setMessage("Personal task added successfully.", false);
    await loadPersonalSections();
  } catch (error) {
    setMessage(`Failed to save task: ${error.message}`, true);
  }
});

async function loadPersonalSections() {
  setLoadingState_([todayList, feedbackList, allList, doneList], "Loading personal tasks...");

  try {
    const data = await getPersonalTasks();
    const tasks = data.tasks || [];
    const today = getWorkingDate_();
    const pendingToday = tasks.filter((item) => normalizeStatus_(item.status) === "pending" && item.date === today);
    const feedbackPending = tasks.filter((item) => normalizeStatus_(item.status) === "feedback pending");
    const otherTasks = tasks.filter((item) => normalizeStatus_(item.status) === "other");
    const doneTasks = tasks.filter((item) => isDoneStatusValue_(item.status));

    renderList_(todayList, pendingToday, "No pending personal tasks for today.", (item) => {
      const li = document.createElement("li");
      const row = document.createElement("div");
      row.className = "task-row";

      const text = document.createElement("span");
      text.textContent = formatPersonalTaskText_(item);

      const doneBtn = document.createElement("button");
      doneBtn.type = "button";
      doneBtn.className = "task-done-btn";
      doneBtn.textContent = "Done";
      doneBtn.addEventListener("click", async () => {
        doneBtn.disabled = true;
        doneBtn.textContent = "...";
        try {
          await completePersonalTask({ date: item.date, task: item.task });
          await loadPersonalSections();
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

    renderList_(feedbackList, feedbackPending, "No feedback pending personal tasks.", (item) => {
      return createPersonalTaskRow_(item, true);
    });

    renderList_(allList, otherTasks, "No other personal tasks.", (item) => {
      return createPersonalTaskRow_(item, true);
    });

    renderList_(doneList, doneTasks, "No completed personal tasks.", (item) => {
      const li = document.createElement("li");
      li.textContent = formatPersonalTaskText_(item);
      return li;
    });
  } catch (error) {
    [todayList, feedbackList, allList, doneList].forEach((list) => {
      list.innerHTML = "";
      const li = document.createElement("li");
      li.textContent = "Could not load personal tasks.";
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

function formatPersonalTaskText_(item) {
  const remark = String(item.remark || "").trim();
  return `${item.date} | ${item.task} | ${item.status} | ${item.category}${remark ? ` | Remark: ${remark}` : ""}`;
}

function createPersonalTaskRow_(item, includeDoneButton) {
  const li = document.createElement("li");

  if (!includeDoneButton) {
    li.textContent = formatPersonalTaskText_(item);
    return li;
  }

  const row = document.createElement("div");
  row.className = "task-row";

  const text = document.createElement("span");
  text.textContent = formatPersonalTaskText_(item);

  const doneBtn = document.createElement("button");
  doneBtn.type = "button";
  doneBtn.className = "task-done-btn";
  doneBtn.textContent = "Done";
  doneBtn.addEventListener("click", async () => {
    doneBtn.disabled = true;
    doneBtn.textContent = "...";
    try {
      await completePersonalTask({ date: item.date, task: item.task });
      await loadPersonalSections();
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

loadPersonalSections();
renderStreakBadge();

window.addEventListener("lifeos:date-change", () => {
  loadPersonalSections();
  renderStreakBadge();
});
