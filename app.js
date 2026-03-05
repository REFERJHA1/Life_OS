import { completePersonalTask, completeWorkTask, getDashboardData, saveHabit, saveJournal } from "./services/common.js";
import { renderStreakBadge } from "./pages/streak.js";

const DATE_KEY = "lifeos_selected_date";

async function loadDashboard() {
  try {
    const workingDate = getWorkingDate_();
    const data = await getDashboardData(workingDate);

    document.getElementById("dailyScore").innerText = `${data.score ?? "--"} / 10`;
    document.getElementById("moodToday").innerText = `${data.moodToday ?? "--"} / 10`;
    document.getElementById("habitScore").innerText = `${data.habitsDone ?? 0} / 4`;
    applyScoreColor(Number(data.score || 0));
    renderTopFeedback(data);

    renderTaskList("workTaskList", data.workTasks || [], async (task) => {
      await completeWorkTask({ date: data.date, task });
      await loadDashboard();
    });
    renderTaskList("personalTaskList", data.personalTasks || [], async (task) => {
      await completePersonalTask({ date: data.date, task });
      await loadDashboard();
    });

    const weekly = (data.weeklyScores || []).join(" | ");
    if (weekly) {
      document.getElementById("weeklySeries").innerText = weekly;
    }

    bindHabitToggles(data.habits || {});
  } catch (error) {
    renderTopFeedback({ motivation: `Could not load dashboard: ${error.message}` });
    renderTaskList("workTaskList", []);
    renderTaskList("personalTaskList", []);
    document.getElementById("weeklySeries").innerText = "--";
  }
}

function renderTaskList(listId, items, onDone) {
  const list = document.getElementById(listId);
  list.innerHTML = "";

  if (!items.length) {
    const li = document.createElement("li");
    li.innerText = "No tasks";
    list.appendChild(li);
    return;
  }

  items.forEach((item) => {
    const li = document.createElement("li");
    const row = document.createElement("div");
    row.className = "task-row";

    const text = document.createElement("span");
    text.textContent = item;

    row.appendChild(text);

    if (typeof onDone === "function") {
      const doneBtn = document.createElement("button");
      doneBtn.type = "button";
      doneBtn.className = "task-done-btn";
      doneBtn.textContent = "Done";
      doneBtn.addEventListener("click", async () => {
        doneBtn.disabled = true;
        doneBtn.textContent = "...";
        try {
          await onDone(item);
        } finally {
          doneBtn.disabled = false;
          doneBtn.textContent = "Done";
        }
      });
      row.appendChild(doneBtn);
    }

    li.appendChild(row);
    list.appendChild(li);
  });
}

function bindHabitToggles(habits) {
  const map = {
    habitStudy: "Study",
    habitReading: "Reading",
    habitWorkout: "Workout",
    habitPlanning: "Planning"
  };

  Object.entries(map).forEach(([id, key]) => {
    const checkbox = document.getElementById(id);
    checkbox.checked = Boolean(habits[key]);
    setHabitVisual(checkbox);

    checkbox.addEventListener("change", async () => {
      const previous = checkbox.checked ? 0 : 1;
      setHabitVisual(checkbox);
      const doneCount = getHabitDoneCount(map);
      document.getElementById("habitScore").innerText = `${doneCount} / 4`;

      try {
        await saveHabit({ date: getWorkingDate_(), habit: key, value: checkbox.checked ? 1 : 0 });
      } catch (error) {
        checkbox.checked = previous === 1;
        setHabitVisual(checkbox);
        document.getElementById("habitScore").innerText = `${getHabitDoneCount(map)} / 4`;
      }
    });
  });
}

function bindJournalForm() {
  const form = document.getElementById("journalForm");
  const button = form.querySelector("button[type='submit']");

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    button.disabled = true;
    button.textContent = "Saving...";

    const payload = {
      date: getWorkingDate_(),
      good: document.getElementById("journalGood").value,
      problem: document.getElementById("journalProblem").value,
      improvement: document.getElementById("journalImprovement").value,
      mood: Number(document.getElementById("journalMood").value),
      remark: document.getElementById("journalRemark").value
    };

    try {
      await saveJournal(payload);
      form.reset();
      button.textContent = "✓ Saved";
      setTimeout(() => {
        button.textContent = "Save Journal";
      }, 900);
    } catch (error) {
      button.textContent = "Save Journal";
    } finally {
      button.disabled = false;
    }
  });
}

function bindStartCheckin() {
  const button = document.getElementById("startCheckinBtn");
  button.addEventListener("click", () => {
    const habitSection = document.getElementById("habitTracker");
    habitSection.scrollIntoView({ behavior: "smooth", block: "start" });

    setTimeout(() => {
      document.getElementById("journalGood").focus();
    }, 450);
  });
}

function applyScoreColor(score) {
  const scoreEl = document.getElementById("dailyScore");
  scoreEl.classList.remove("score-low", "score-mid", "score-good", "score-great");

  if (score <= 3) {
    scoreEl.classList.add("score-low");
    return;
  }

  if (score <= 6) {
    scoreEl.classList.add("score-mid");
    return;
  }

  if (score <= 8) {
    scoreEl.classList.add("score-good");
    return;
  }

  scoreEl.classList.add("score-great");
}

function renderTopFeedback(data) {
  const motivationEl = document.getElementById("motivationText");
  const rewardEl = document.getElementById("rewardText");
  const reminderEl = document.getElementById("reminderText");

  motivationEl.textContent = data.motivation || "Keep going.";

  if (data.reward) {
    rewardEl.hidden = false;
    rewardEl.textContent = data.reward;
  } else {
    rewardEl.hidden = true;
    rewardEl.textContent = "";
  }

  if (data.reminder) {
    reminderEl.hidden = false;
    reminderEl.textContent = data.reminder;
  } else {
    reminderEl.hidden = true;
    reminderEl.textContent = "";
  }
}

function setHabitVisual(checkbox) {
  const label = checkbox.closest("label");
  if (!label) return;
  label.classList.toggle("active", checkbox.checked);
}

function getHabitDoneCount(map) {
  return Object.keys(map).reduce((sum, id) => {
    const checkbox = document.getElementById(id);
    return sum + (checkbox.checked ? 1 : 0);
  }, 0);
}

function getWorkingDate_() {
  try {
    const value = localStorage.getItem(DATE_KEY);
    if (value && String(value).trim()) {
      return String(value).trim();
    }
  } catch (error) {
  }

  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

window.addEventListener("lifeos:date-change", () => {
  loadDashboard();
  renderStreakBadge();
});

bindJournalForm();
bindStartCheckin();
loadDashboard();
renderStreakBadge();
