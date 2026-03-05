import { getDashboardData, saveHabit } from "../services/common.js";
import { renderStreakBadge } from "./streak.js";

const DATE_KEY = "lifeos_selected_date";

const habitIds = {
  Study: "habitStudy",
  Reading: "habitReading",
  Workout: "habitWorkout",
  Planning: "habitPlanning"
};

const scoreEl = document.getElementById("habitTodayScore");
const dateEl = document.getElementById("habitDateLabel");
const messageEl = document.getElementById("habitMessage");
const remarkEl = document.getElementById("habitRemark");

async function loadHabits() {
  try {
    const data = await getDashboardData(getWorkingDate_());
    const habits = data.habits || {};

    Object.entries(habitIds).forEach(([habitName, id]) => {
      const checkbox = document.getElementById(id);
      checkbox.checked = Number(habits[habitName] || 0) === 1;
      setHabitVisual(checkbox);
      checkbox.addEventListener("change", () => handleToggle(habitName, checkbox.checked));
    });

    updateScore();
    dateEl.textContent = `Date: ${data.date || getWorkingDate_()}`;
    setMessage("Habits loaded.", false);
  } catch (error) {
    setMessage(`Failed to load habits: ${error.message}`, true);
  }
}

async function handleToggle(habit, checked) {
  const checkbox = document.getElementById(habitIds[habit]);
  setHabitVisual(checkbox);

  try {
    setMessage("Saving habit...", false);

    await saveHabit({
      date: getWorkingDate_(),
      habit,
      value: checked ? 1 : 0,
      remark: remarkEl.value.trim()
    });

    updateScore();
    setMessage(`${habit} updated.`, false);
  } catch (error) {
    setMessage(`Failed to save ${habit}: ${error.message}`, true);
  }
}

function updateScore() {
  const doneCount = Object.values(habitIds).reduce((sum, id) => {
    const checkbox = document.getElementById(id);
    return sum + (checkbox.checked ? 1 : 0);
  }, 0);

  scoreEl.textContent = `${doneCount} / 4 completed`;
}

function setMessage(text, isError) {
  messageEl.textContent = text;
  messageEl.classList.toggle("error", isError);
}

function setHabitVisual(checkbox) {
  const label = checkbox.closest("label");
  if (!label) return;
  label.classList.toggle("active", checkbox.checked);
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

loadHabits();
renderStreakBadge();

window.addEventListener("lifeos:date-change", () => {
  loadHabits();
  renderStreakBadge();
});
