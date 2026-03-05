import { getJournalEntries } from "../services/common.js";

const DATE_KEY = "lifeos_selected_date";

export async function renderStreakBadge(elementId = "streakBadge") {
  const element = document.getElementById(elementId);
  if (!element) return;

  try {
    const streak = await getCurrentStreak();
    element.textContent = `Streak: ${streak} day${streak === 1 ? "" : "s"}`;
  } catch (error) {
    element.textContent = "Streak: --";
  }
}

export async function getCurrentStreak(referenceDate) {
  const data = await getJournalEntries();
  const dates = data.dates || [];
  const today = referenceDate || getWorkingDate_();
  return calculateStreak(dates, today);
}

export function calculateStreak(entryDates, today) {
  const dateSet = new Set(entryDates);
  let streak = 0;
  const cursor = new Date(`${today}T00:00:00`);

  while (true) {
    const dateKey = formatDate(cursor);
    if (!dateSet.has(dateKey)) break;
    streak += 1;
    cursor.setDate(cursor.getDate() - 1);
  }

  return streak;
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
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

  const now = new Date();
  return formatDate(now);
}
