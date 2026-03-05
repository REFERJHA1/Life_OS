import { getDashboardData, getJournalHistory, saveJournal } from "../services/common.js";
import { getCurrentStreak, renderStreakBadge } from "./streak.js";

const DATE_KEY = "lifeos_selected_date";

const form = document.getElementById("journalPageForm");
const messageEl = document.getElementById("journalMessage");
const moodTodayEl = document.getElementById("journalMoodToday");
const entryStatusEl = document.getElementById("journalEntryStatus");
const streakEl = document.getElementById("journalStreak");
const dateLabelEl = document.getElementById("journalDateLabel");
const tipEl = document.getElementById("journalTip");
const historyListEl = document.getElementById("journalHistoryList");
const progressBarEl = document.getElementById("journalProgressBar");
const progressTextEl = document.getElementById("journalProgressText");
const moodInputEl = document.getElementById("journalMood");
const moodValueEl = document.getElementById("moodValue");
const saveBtnEl = document.getElementById("journalSaveBtn");
const firstFieldEl = document.getElementById("journalGood");

bindMoodSlider();
bindFormProgress();

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  setMessage("Saving journal...", false);
  saveBtnEl.textContent = "Saving...";
  saveBtnEl.disabled = true;

  const payload = {
    date: getWorkingDate_(),
    good: document.getElementById("journalGood").value.trim(),
    problem: document.getElementById("journalProblem").value.trim(),
    improvement: document.getElementById("journalImprovement").value.trim(),
    mood: Number(document.getElementById("journalMood").value),
    remark: document.getElementById("journalRemark").value.trim()
  };

  try {
    await saveJournal(payload);
    setMessage("Journal saved.", false);
    entryStatusEl.textContent = "Entry: Saved";
    moodTodayEl.textContent = `Mood Today: ${payload.mood} / 10`;
    applyMoodColor(payload.mood);
    form.reset();
    moodInputEl.value = "5";
    moodValueEl.textContent = "5";
    updateProgress(0);
    saveBtnEl.textContent = "✓ Saved";
    await loadJournalStatus();
    setTimeout(() => {
      saveBtnEl.textContent = "Save Journal";
      saveBtnEl.disabled = false;
    }, 1000);
  } catch (error) {
    setMessage(`Failed to save journal: ${error.message}`, true);
    saveBtnEl.textContent = "Save Journal";
    saveBtnEl.disabled = false;
  }
});

async function loadJournalStatus() {
  try {
    const selectedDate = getWorkingDate_();
    const [data, history] = await Promise.all([getDashboardData(selectedDate), getJournalHistory(7)]);
    const mood = Number(data.moodToday || 0);
    const today = data.date || selectedDate;
    const streak = await getCurrentStreak(today);

    dateLabelEl.textContent = `Date: ${today}`;
    streakEl.textContent = `🔥 ${streak} day${streak === 1 ? "" : "s"} streak`;
    renderHistory(history.entries || []);
    renderCompletion(data, mood);

    if (mood > 0) {
      moodTodayEl.textContent = `Mood Today: ${mood} / 10`;
      entryStatusEl.textContent = "Entry: Saved";
      tipEl.textContent = "Tip: Keep this streak alive tomorrow.";
    } else {
      moodTodayEl.textContent = "Mood Today: -- / 10";
      entryStatusEl.textContent = "Entry: Not saved";
      tipEl.textContent = "Tip: One small reflection is enough.";
    }

    applyMoodColor(mood);
  } catch (error) {
    setMessage(`Failed to load status: ${error.message}`, true);
  }
}

function bindMoodSlider() {
  moodInputEl.addEventListener("input", () => {
    moodValueEl.textContent = moodInputEl.value;
    applyMoodColor(Number(moodInputEl.value));
  });
}

function bindFormProgress() {
  const ids = ["journalGood", "journalProblem", "journalImprovement", "journalMood"];

  ids.forEach((id) => {
    const element = document.getElementById(id);
    element.addEventListener("input", () => {
      const completed = ids.reduce((sum, fieldId) => {
        const value = String(document.getElementById(fieldId).value || "").trim();
        return sum + (value ? 1 : 0);
      }, 0);
      updateProgress(Math.round((completed / ids.length) * 100));
    });
  });
}

function renderCompletion(data, mood) {
  const journalDone = mood > 0 ? 1 : 0;
  const moodDone = mood > 0 ? 1 : 0;
  const habitsDone = Number(data.habitsDone || 0) > 0 ? 1 : 0;
  const total = 3;
  const done = journalDone + moodDone + habitsDone;
  const percent = Math.round((done / total) * 100);
  updateProgress(percent);
}

function updateProgress(percent) {
  progressBarEl.style.width = `${percent}%`;
  progressTextEl.textContent = `Progress: ${percent}%`;
}

function renderHistory(entries) {
  historyListEl.innerHTML = "";

  if (!entries.length) {
    const li = document.createElement("li");
    li.textContent = "No previous entries.";
    historyListEl.appendChild(li);
    return;
  }

  entries.forEach((entry) => {
    const li = document.createElement("li");
    li.textContent = `${entry.date}  Mood ${entry.mood || "--"}`;
    historyListEl.appendChild(li);
  });
}

function applyMoodColor(mood) {
  moodTodayEl.classList.remove("mood-low", "mood-mid", "mood-good", "mood-great");

  if (mood >= 1 && mood <= 3) {
    moodTodayEl.classList.add("mood-low");
    return;
  }

  if (mood <= 6) {
    moodTodayEl.classList.add("mood-mid");
    return;
  }

  if (mood <= 8) {
    moodTodayEl.classList.add("mood-good");
    return;
  }

  moodTodayEl.classList.add("mood-great");
}

function setMessage(text, isError) {
  messageEl.textContent = text;
  messageEl.classList.toggle("error", isError);
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

loadJournalStatus();
renderStreakBadge();

window.addEventListener("lifeos:date-change", () => {
  loadJournalStatus();
  renderStreakBadge();
});

setTimeout(() => {
  firstFieldEl.focus();
}, 200);
