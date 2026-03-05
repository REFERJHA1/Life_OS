import { getAnalytics } from "../services/common.js";
import { renderStreakBadge } from "./streak.js";

const DATE_KEY = "lifeos_selected_date";

const weeklyChartEl = document.getElementById("weeklyChart");
const weeklySummaryEl = document.getElementById("weeklySummary");
const monthlyHabitSummaryEl = document.getElementById("monthlyHabitSummary");
const monthlyHabitMetaEl = document.getElementById("monthlyHabitMeta");
const moodChartEl = document.getElementById("moodChart");
const moodSummaryEl = document.getElementById("moodSummary");
const yearlyValueEl = document.getElementById("yearlyConsistencyValue");
const yearlyMetaEl = document.getElementById("yearlyConsistencyMeta");

async function loadAnalytics() {
	try {
		const data = await getAnalytics(getWorkingDate_());

		renderBars(weeklyChartEl, data.weeklyProductivity || [], 10, "score");
		weeklySummaryEl.textContent = "Daily productivity score (last 7 days)";

		const month = data.monthlyHabitCompletion || {};
		monthlyHabitSummaryEl.textContent = `${month.percentage ?? 0}%`;
		monthlyHabitMetaEl.textContent = `${month.month || "--"} • ${month.completed ?? 0}/${month.possible ?? 0}`;

		renderBars(moodChartEl, data.moodTrend || [], 10, "mood");
		moodSummaryEl.textContent = "Mood trend (last 7 days)";

		const year = data.yearlyConsistency || {};
		yearlyValueEl.textContent = `${year.percentage ?? 0}%`;
		yearlyMetaEl.textContent = `${year.year || "--"} • ${year.activeDays ?? 0}/${year.daysElapsed ?? 0} active days`;
	} catch (error) {
		weeklySummaryEl.textContent = `Failed to load analytics: ${error.message}`;
		monthlyHabitSummaryEl.textContent = "--%";
		monthlyHabitMetaEl.textContent = "--";
		moodSummaryEl.textContent = "--";
		yearlyValueEl.textContent = "--%";
		yearlyMetaEl.textContent = "--";
	}
}

function renderBars(container, rows, maxValue, key) {
	container.innerHTML = "";

	if (!rows.length) {
		const empty = document.createElement("p");
		empty.className = "muted";
		empty.textContent = "No data yet.";
		container.appendChild(empty);
		return;
	}

	rows.forEach((row) => {
		const value = Number(row[key] || 0);
		const percent = Math.max(0, Math.min(100, (value / maxValue) * 100));

		const line = document.createElement("div");
		line.className = "bar-row";

		const label = document.createElement("span");
		label.className = "bar-label";
		label.textContent = row.label || "--";

		const track = document.createElement("div");
		track.className = "bar-track";

		const fill = document.createElement("div");
		fill.className = "bar-fill";
		fill.style.width = `${percent}%`;

		const valueText = document.createElement("span");
		valueText.className = "bar-value";
		valueText.textContent = String(value);

		track.appendChild(fill);
		line.appendChild(label);
		line.appendChild(track);
		line.appendChild(valueText);
		container.appendChild(line);
	});
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

loadAnalytics();
renderStreakBadge();

window.addEventListener("lifeos:date-change", () => {
	loadAnalytics();
	renderStreakBadge();
});
