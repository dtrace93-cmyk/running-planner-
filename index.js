document.addEventListener('DOMContentLoaded', () => {
const runForm = document.getElementById('run-form');
    const runLogEl = document.getElementById('run-log');
    const totalDistanceEl = document.getElementById('total-distance');
    const avgWeeklyEl = document.getElementById('avg-weekly');
    const fastestPaceEl = document.getElementById('fastest-pace');
    const longestRunEl = document.getElementById('longest-run');
    const planBodyEl = document.getElementById('plan-body');
    const addPlanBtn = document.getElementById('add-plan');
    const excelUpload = document.getElementById('excel-upload');
    const excelHistoryUpload = document.getElementById('excel-history-upload');
    const planCsvUpload = document.getElementById('plan-csv-upload');
    const planUploadMessage = document.getElementById('plan-upload-message');
    const weeklySummaryGrid = document.getElementById('weekly-summary-grid');
    const weeklySummaryRows = document.getElementById('weekly-summary-rows');
    const sessionStatusEl = document.getElementById('session-status');
    const weatherSummaryEl = document.getElementById('weather-summary');
    const hrOverallEl = document.getElementById('hr-overall');
    const hrHalvesEl = document.getElementById('hr-halves');
    const hrTrendEl = document.getElementById('hr-trend');
    const paceOverallEl = document.getElementById('pace-overall');
    const paceHalvesEl = document.getElementById('pace-halves');
    const paceFastestEl = document.getElementById('pace-fastest');
    const paceTrendEl = document.getElementById('pace-trend');
    const weeklyAvgEl = document.getElementById('weekly-avg');
    const weeklyMaxEl = document.getElementById('weekly-max');
    const weeklyJumpsEl = document.getElementById('weekly-jumps');
    const fatigueTotalEl = document.getElementById('fatigue-total');
    const fatigueTopEl = document.getElementById('fatigue-top');
    const fatigueNoteEl = document.getElementById('fatigue-note');
    const timeOfDayCountsEl = document.getElementById('tod-counts');
    const timeOfDayPacesEl = document.getElementById('tod-paces');
    const timeOfDayNoteEl = document.getElementById('tod-note');
    const acuteLoadEl = document.getElementById('acute-load');
    const chronicLoadEl = document.getElementById('chronic-load');
    const loadRatioEl = document.getElementById('load-ratio');
    const loadRiskEl = document.getElementById('load-risk');
    const insightsListEl = document.getElementById('insights-list');
    const exportJsonBtn = document.getElementById('export-json');
    const importJsonInput = document.getElementById('import-json');
    const exportJsonSettingsBtn = document.getElementById('export-json-settings');
    const importJsonSettingsInput = document.getElementById('import-json-settings');
    const userNotesEl = document.getElementById('user-notes');
    const stravaSummaryEl = document.getElementById('strava-summary');

    const navItems = document.querySelectorAll('.sidebar-nav-item');
    const appSections = document.querySelectorAll('.app-section');

    const STRAVA_ACTIVITIES_URL = 'https://probable-goggles-jjv4qj56xvqjh5wxv-4000.app.github.dev/api/strava/activities';

    function setActiveSection(sectionId) {
      appSections.forEach(sec => {
        sec.style.display = sec.id === sectionId ? 'flex' : 'none';
      });
      navItems.forEach(item => {
        const target = item.dataset.section;
        if (target === sectionId) {
          item.classList.add('active');
        } else {
          item.classList.remove('active');
        }
      });
    }

    navItems.forEach(item => {
      item.addEventListener('click', (e) => {
        e.preventDefault();
        const targetSection = item.dataset.section;
        if (targetSection) {
          setActiveSection(targetSection);
        }
      });
    });

    const STORAGE_KEY = 'runningPlannerDataV1';
    const NOTES_KEY = 'runningPlannerUserNotes';

    let runs = [];
    let runIdCounter = 1;
    let chart;

    function getActivityTimestamp(activity) {
      const candidates = [activity.startDate, activity.start_date, activity.startDateLocal, activity.start_date_local, activity.startDateUTC, activity.start_date_utc];
      for (const value of candidates) {
        const ts = value ? Date.parse(value) : NaN;
        if (!Number.isNaN(ts)) return ts;
      }
      return 0;
    }

    function milesFromActivity(activity) {
      if (typeof activity?.distanceMiles === 'number') return activity.distanceMiles;
      if (typeof activity?.distance === 'number') return activity.distance / 1609.34;
      return null;
    }

    function secondsFromActivity(activity) {
      if (typeof activity?.movingTimeSec === 'number') return activity.movingTimeSec;
      if (typeof activity?.moving_time === 'number') return activity.moving_time;
      if (typeof activity?.movingTime === 'number') return activity.movingTime;
      return 0;
    }

    function formatMinutesSeconds(totalSeconds) {
      const minutes = Math.floor(totalSeconds / 60);
      const seconds = Math.max(0, Math.round(totalSeconds - minutes * 60));
      return `${minutes}m ${seconds.toString().padStart(2, '0')}s`;
    }

    function formatSecondsToHms(totalSeconds) {
      if (!Number.isFinite(totalSeconds) || totalSeconds <= 0) return '';
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = Math.floor(totalSeconds % 60);
      return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    }

    async function fetchLatestStravaRun() {
      if (!stravaSummaryEl) return;
      stravaSummaryEl.textContent = 'Loading latest Strava run…';
      try {
        const response = await fetch(STRAVA_ACTIVITIES_URL);
        if (response.status === 401) {
          stravaSummaryEl.textContent = 'Strava not connected.';
          return;
        }
        if (!response.ok) {
          stravaSummaryEl.textContent = 'Unable to load Strava data.';
          return;
        }
        const data = await response.json();
        if (!Array.isArray(data)) {
          stravaSummaryEl.textContent = 'Unable to load Strava data.';
          return;
        }
        const added = mergeStravaRunsIntoLog(data);
        if (added > 0) {
          refreshUI();
        }

        const runsOnly = data.filter(act => (act?.type || '').toLowerCase() === 'run');
        if (runsOnly.length === 0) {
          stravaSummaryEl.textContent = 'No recent running activities found.';
          return;
        }
        let latestRun = runsOnly[0];
        let latestTs = getActivityTimestamp(latestRun);
        runsOnly.forEach((act) => {
          const ts = getActivityTimestamp(act);
          if (ts > latestTs) {
            latestTs = ts;
            latestRun = act;
          }
        });

        const miles = milesFromActivity(latestRun);
        const movingSeconds = secondsFromActivity(latestRun);
        const dateLabel = latestTs ? new Date(latestTs).toLocaleDateString() : 'Unknown date';
        const milesText = typeof miles === 'number' ? miles.toFixed(2) : 'Unknown distance';
        const timeText = formatMinutesSeconds(movingSeconds || 0);
        stravaSummaryEl.textContent = `Last run: ${milesText} miles in ${timeText} on ${dateLabel}`;
      } catch (error) {
        console.error('Failed to load Strava data', error);
        stravaSummaryEl.textContent = 'Unable to load Strava data.';
      }
    }

    function mapStravaActivityToRunEntry(activity) {
      if (!activity || (activity.type || '').toLowerCase() !== 'run') return null;
      const completedMiles = milesFromActivity(activity) ?? 0;
      const movingSeconds = secondsFromActivity(activity) || 0;
      const paceSeconds = Number.isFinite(activity.paceSecondsPerMile)
        ? activity.paceSecondsPerMile
        : (completedMiles ? movingSeconds / completedMiles : null);
      const date = normalizeDate(activity.startDate || activity.start_date || activity.startDateLocal || activity.start_date_local);
      const name = activity.name || 'Strava run';
      const sessionType = detectSessionType(name, '', completedMiles || 0);

      return addPaceToRun({
        id: createRunId(),
        source: 'strava',
        stravaId: activity.id,
        date,
        distance: completedMiles,
        plannedMiles: completedMiles,
        completedMiles,
        distanceText: `${name} (${completedMiles ? completedMiles.toFixed(2) + ' mi' : 'Run'})`,
        sessionType,
        type: sessionType,
        avgHr: Number.isFinite(activity.averageHeartrate) ? Number(activity.averageHeartrate) : null,
        effort: 5,
        temperature: null,
        duration: formatSecondsToHms(movingSeconds),
        weather: '',
        timeOfDay: 'Unknown',
        paceSeconds,
        resultText: name
      });
    }

    function mergeStravaRunsIntoLog(data) {
      if (!Array.isArray(data)) return 0;
      const existingIds = new Set(runs.filter(r => r.stravaId != null).map(r => String(r.stravaId)));
      const mapped = data
        .filter(act => (act?.type || '').toLowerCase() === 'run')
        .map(mapStravaActivityToRunEntry)
        .filter(Boolean)
        .filter(entry => !existingIds.has(String(entry.stravaId)));
      if (!mapped.length) return 0;
      runs = runs.concat(mapped);
      return mapped.length;
    }

    async function syncStravaRunsIntoLog() {
      try {
        const response = await fetch(STRAVA_ACTIVITIES_URL);
        if (response.status === 401) return; // Not connected; skip quietly
        if (!response.ok) {
          console.warn('Unable to sync Strava runs', response.status);
          return;
        }
        const data = await response.json();
        const added = mergeStravaRunsIntoLog(data);
        if (added > 0) {
          refreshUI();
        }
      } catch (error) {
        console.error('Failed to sync Strava runs', error);
      }
    }

    function parseDurationToSeconds(value) {
      if (!value && value !== 0) return 0;
      if (typeof value === 'number') {
        // Excel stores times as fraction of a day
        return Math.round(value * 24 * 60 * 60);
      }
      const parts = String(value).split(':').map(Number);
      if (parts.length === 2) {
        const [hh, mm] = parts;
        return (hh * 3600) + (mm * 60);
      }
      if (parts.length === 3) {
        const [hh, mm, ss] = parts;
        return (hh * 3600) + (mm * 60) + ss;
      }
      return 0;
    }

    function formatPace(secondsPerMile) {
      if (!isFinite(secondsPerMile) || secondsPerMile <= 0) return '–';
      const minutes = Math.floor(secondsPerMile / 60);
      const seconds = Math.round(secondsPerMile % 60).toString().padStart(2, '0');
      return `${minutes}:${seconds} /mi`;
    }

    function createRunId() {
      return `run-${Date.now()}-${runIdCounter++}`;
    }

    function loadUserNotes() {
      if (!userNotesEl) return;
      const saved = localStorage.getItem(NOTES_KEY);
      if (saved !== null) {
        userNotesEl.value = saved;
      }
    }

    function saveUserNotes() {
      if (!userNotesEl) return;
      try {
        localStorage.setItem(NOTES_KEY, userNotesEl.value || '');
      } catch (err) {
        console.warn('Failed to save notes', err);
      }
    }

    function saveRunsToStorage(data) {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
      } catch (err) {
        console.error('Failed to persist runs', err);
      }
    }

    function normalizeLoadedRun(run) {
      if (!run || typeof run !== 'object') return null;
      const base = { ...run };
      base.id = base.id || createRunId();
      base.source = base.source || 'manual';
      base.date = base.date || base.dateLabel || base.dateValue || '';
      base.distanceText = base.distanceText || '';
      const plannedGuess = base.plannedMiles ?? base.distance ?? extractMilesFromText(base.distanceText) ?? 0;
      base.plannedMiles = Number.isFinite(plannedGuess) ? plannedGuess : 0;
      const completedGuess = base.completedMiles ?? base.distance ?? null;
      base.completedMiles = Number.isFinite(completedGuess) ? completedGuess : null;
      base.effort = Number.isFinite(Number(base.effort)) ? Number(base.effort) : 5;
      base.timeOfDay = base.timeOfDay || 'Unknown';
      base.weather = base.weather || '';
      base.weekId = normalizeWeekId(base.weekId || base.week || base.weekLabel, base.date);
      base.sessionType = base.sessionType || base.type || detectSessionType(base.distanceText, base.notes || base.resultText || '', base.plannedMiles || 0);
      base.statusFlags = base.statusFlags || detectStatusFlags(base.resultText || base.notes || '', base.plannedMiles || 0, base.completedMiles || 0);
      base.riskFlags = base.riskFlags || detectRiskFlags(base.resultText || '', base.notes || '');
      base.load = calculateRunLoad(base.completedMiles || 0, base.effort);
      return base;
    }

    function loadRunsFromStorage() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return [];
        const parsed = JSON.parse(raw);
        if (!Array.isArray(parsed)) return [];
        const normalized = parsed.map(normalizeLoadedRun).filter(Boolean);
        runIdCounter = normalized.length ? normalized.length + 1 : 1;
        return normalized;
      } catch (err) {
        console.error('Failed to load runs from storage', err);
        return [];
      }
    }

    function exportRunsToJson() {
      const blob = new Blob([JSON.stringify(runs, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'running-planner-data.json';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }

    function handleJsonImport(event) {
      const file = event.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const parsed = JSON.parse(e.target.result);
          if (!Array.isArray(parsed)) throw new Error('Invalid format');
          const normalized = parsed.map(normalizeLoadedRun).filter(Boolean);
          const hasCoreFields = normalized.some(r => r.date && (r.plannedMiles || r.completedMiles || r.distance));
          if (!hasCoreFields) throw new Error('Missing core fields');
          runs = normalized;
          runIdCounter = runs.length ? runs.length + 1 : 1;
          refreshUI();
        } catch (err) {
          console.error('Failed to import JSON', err);
          alert('Invalid JSON file. Ensure it contains an array of run objects with date and distance fields.');
        }
      };
      reader.readAsText(file);
      event.target.value = '';
    }

    function addPaceToRun(run) {
      const secondsFromDuration = parseDurationToSeconds(run.duration);
      const derivedPace = secondsFromDuration && run.distance ? secondsFromDuration / run.distance : null;
      const paceSeconds = run.paceSeconds ?? derivedPace;
      return { ...run, paceSeconds, id: run.id || createRunId() };
    }

    function normalizeKey(key) {
      return String(key || '').trim().toLowerCase();
    }

    function parseDistanceValue(value) {
      if (value === undefined || value === null) return 0;
      if (typeof value === 'number') return value;
      const raw = String(value).trim();
      if (!raw) return 0;
      const match = raw.match(/([0-9]+(?:\.[0-9]+)?)/);
      if (!match) return 0;
      const num = parseFloat(match[1]);
      const isKm = /km|kilometer/i.test(raw);
      return isKm ? num * 0.621371 : num;
    }

    function parsePaceValue(value) {
      if (!value && value !== 0) return null;
      if (typeof value === 'number') return value * 60; // assume minutes per mile numeric
      const raw = String(value).trim();
      if (!raw) return null;
      const timeMatch = raw.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
      if (!timeMatch) return null;
      const [_, part1, part2, part3] = timeMatch;
      const hasHours = part3 !== undefined;
      const hours = hasHours ? parseInt(part1, 10) : 0;
      const minutes = hasHours ? parseInt(part2, 10) : parseInt(part1, 10);
      const secondsVal = hasHours ? parseInt(part3, 10) : parseInt(part2, 10);
      const seconds = (hours * 3600) + (minutes * 60) + secondsVal;
      const perKm = /km/i.test(raw);
      return perKm ? seconds * 1.60934 : seconds;
    }

    function extractMilesFromText(text) {
      if (!text) return null;
      const match = String(text).match(/([0-9]+(?:\.[0-9]+)?)\s*mi/i);
      return match ? parseFloat(match[1]) : null;
    }

    function detectSessionType(distanceText = '', notes = '', plannedMiles = 0) {
      const text = `${distanceText} ${notes}`.toLowerCase();
      if (/race|10k|half|marathon/.test(text)) return 'Race';
      if (/long/.test(text) || plannedMiles >= 8) return 'Long run';
      if (/tempo|threshold/.test(text)) return 'Tempo';
      if (/progression/.test(text)) return 'Progression';
      if (/interval|reps|\d+x\d+/.test(text)) return 'Intervals';
      if (/recovery|easy jog|easy run/.test(text)) return 'Recovery';
      if (plannedMiles && plannedMiles <= 3) return 'Easy';
      return 'General';
    }

    function calculateRunLoad(completedMiles, effort) {
      const effortValue = Number.isFinite(effort) ? effort : 5;
      if (!completedMiles || !effortValue) return 0;
      return completedMiles * effortValue;
    }

    function normalizeWeekId(weekLabel, dateValue) {
      if (weekLabel) return weekLabel;
      if (!dateValue) return 'Unknown week';
      const date = new Date(dateValue);
      if (date.toString() === 'Invalid Date') return 'Unknown week';
      const day = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
      const dayNum = day.getUTCDay() || 7;
      day.setUTCDate(day.getUTCDate() + 4 - dayNum);
      const yearStart = new Date(Date.UTC(day.getUTCFullYear(),0,1));
      const weekNo = Math.ceil((((day - yearStart) / 86400000) + 1)/7);
      return `${day.getUTCFullYear()}-W${String(weekNo).padStart(2,'0')}`;
    }

    function detectStatusFlags(resultText = '', plannedMiles = 0, completedMiles = 0) {
      const text = String(resultText || '').toLowerCase();
      const missedKeywords = ['no run', 'did not run', 'skipped', 'missed', 'dnf'];
      const illKeywords = ['ill', 'sick'];
      const injuryKeywords = ['shin splint', 'injury', 'pain', 'niggle'];
      const fatigueKeywords = ['fatigue', 'tired', 'exhausted', 'drained'];

      const missed = missedKeywords.some(k => text.includes(k)) || (completedMiles === 0 && illKeywords.some(k => text.includes(k)));
      const ill = illKeywords.some(k => text.includes(k));
      const injury = injuryKeywords.some(k => text.includes(k));
      const fatigue = fatigueKeywords.some(k => text.includes(k));
      const recovery = /recovery|rest/.test(text) && (plannedMiles <= 0 || completedMiles <= 1);

      let statusLabel = '';
      if (missed) statusLabel = 'Missed';
      else if (injury) statusLabel = 'Injury';
      else if (ill) statusLabel = 'Ill';
      else if (recovery) statusLabel = 'Recovery';

      return { missed, ill, injury, fatigue, recovery, statusLabel };
    }

    function detectRiskFlags(resultText = '', notes = '') {
      const text = `${String(resultText || '')} ${String(notes || '')}`.toLowerCase();
      const flags = [];
      if (/ill|sick|flu|cold/.test(text)) flags.push('Ill');
      if (/fatigue|tired|exhausted|drained|burnt out/.test(text)) flags.push('Fatigue');
      if (/shin splint|injury|pain|niggle/.test(text)) flags.push('Injury');
      if (/overheating|too hot|heat|humid/.test(text)) flags.push('Heat');
      if (/stopped midway|cut short|had to stop|dnf/.test(text)) flags.push('Stopped');
      if (/bad sleep|poor sleep|no sleep|woke up|insomnia/.test(text)) flags.push('Sleep');
      return flags;
    }

    function isMissedSession(resultText) {
      if (!resultText) return false;
      return /no run|ill/i.test(resultText);
    }

    function parseResultDetails(resultText) {
      const raw = String(resultText || '').trim();
      if (!raw) return { completedMiles: null, avgHr: null, paceSeconds: null };
      if (isMissedSession(raw)) return { completedMiles: 0, avgHr: null, paceSeconds: null };

      const distanceMatch = raw.match(/(\d+(\.\d+)?)\s*mi/i);
      const completedMiles = distanceMatch ? parseFloat(distanceMatch[1]) : null;

      const hrMatch = raw.match(/Avg\s+(\d{2,3})/i);
      const avgHr = hrMatch ? parseInt(hrMatch[1], 10) : null;

      let paceSeconds = null;
      const paceTimeMatch = raw.match(/(\d{1,2})[.:](\d{2})\s*\/\s*mi/i);
      if (paceTimeMatch) {
        const minutes = parseInt(paceTimeMatch[1], 10);
        const seconds = parseInt(paceTimeMatch[2], 10);
        paceSeconds = (minutes * 60) + seconds;
      } else {
        const paceDecimalMatch = raw.match(/(\d+(\.\d+)?)\s*\/\s*mi/i);
        if (paceDecimalMatch) {
          const decimalMinutes = parseFloat(paceDecimalMatch[1]);
          const minutes = Math.floor(decimalMinutes);
          const seconds = Math.round((decimalMinutes - minutes) * 60);
          paceSeconds = (minutes * 60) + seconds;
        }
      }

      return { completedMiles, avgHr, paceSeconds };
    }

    function extractFromResultsField(value) {
      const { completedMiles, avgHr, paceSeconds } = parseResultDetails(value);
      if (completedMiles === null && avgHr === null && paceSeconds === null) return {};
      return { distance: completedMiles, avgHr, paceSeconds };
    }

    function parseResultMetrics(resultText) {
      return parseResultDetails(resultText);
    }

    function parseCsv(text) {
      const rows = [];
      let current = '';
      let row = [];
      let insideQuote = false;

      const pushValue = () => {
        row.push(current);
        current = '';
      };

      for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const next = text[i + 1];

        if (char === '"') {
          if (insideQuote && next === '"') {
            current += '"';
            i++; // skip escaped quote
          } else {
            insideQuote = !insideQuote;
          }
        } else if (char === ',' && !insideQuote) {
          pushValue();
        } else if ((char === '\n' || char === '\r') && !insideQuote) {
          if (current !== '' || row.length) pushValue();
          if (row.length) rows.push(row);
          row = [];
        } else {
          current += char;
        }
      }
      if (current !== '' || row.length) {
        pushValue();
        rows.push(row);
      }
      return rows.filter(r => r.some(cell => String(cell).trim() !== ''));
    }

    function setPlanMessage(message, isError = false) {
      planUploadMessage.textContent = message;
      planUploadMessage.style.color = isError ? '#ff9f7a' : 'var(--muted)';
    }

    function mapPlanHeader(headerRow) {
      const mapping = {};
      headerRow.forEach((cell, idx) => {
        const key = normalizeKey(cell);
        if (!key) return;
        if (mapping.week === undefined && (key === 'week' || key.includes('week'))) {
          mapping.week = idx;
          return;
        }
        if (mapping.date === undefined && (key === 'date' || key.includes('date'))) {
          mapping.date = idx;
          return;
        }
        if (mapping.distance === undefined && (key === 'distance' || key.includes('distance') || key === 'planned distance')) {
          mapping.distance = idx;
          return;
        }
        if (mapping.results === undefined && (key === 'results (distance, avg hr, avg pace)' || key === 'results' || key === 'result' || key.includes('result'))) {
          mapping.results = idx;
          return;
        }
        if (mapping.notes === undefined && (key === 'notes' || key === 'note' || key.includes('note'))) {
          mapping.notes = idx;
        }
      });
      const requiredKeys = ['week', 'date', 'distance', 'results', 'notes'];
      const mappingComplete = requiredKeys.every(k => mapping[k] !== undefined);
      return { mapping, mappingComplete };
    }

    function parsePlanRowsFromCsv(text) {
      const rows = parseCsv(text);
      if (!rows.length) return { entries: [], mappingComplete: false };

      const header = rows[0];
      const { mapping, mappingComplete } = mapPlanHeader(header);
      if (!mappingComplete) return { entries: [], mappingComplete };

      const entries = rows.slice(1).map(cells => {
        const week = String(cells[mapping.week] || '').trim();
        const date = String(cells[mapping.date] || '').trim();
        const distanceText = String(cells[mapping.distance] || '').trim();
        const result = String(cells[mapping.results] || '').trim();
        const notes = String(cells[mapping.notes] || '').trim();
        if (![week, date, distanceText, result, notes].some(Boolean)) return null;
        const plannedMiles = extractMilesFromText(distanceText);
        const { completedMiles, paceSeconds, avgHr } = parseResultMetrics(result);
        const weekId = normalizeWeekId(week, date);
        const sessionType = detectSessionType(distanceText, notes, plannedMiles || 0);
        const statusFlags = detectStatusFlags(result || notes, plannedMiles || 0, completedMiles || 0);
        const riskFlags = detectRiskFlags(result || '', notes || '');
        return {
          id: createRunId(),
          source: 'plan',
          week,
          date,
          weekId,
          distanceText,
          resultText: result,
          notes,
          plannedMiles,
          completedMiles,
          paceSeconds,
          avgHr,
          sessionType,
          statusFlags,
          riskFlags,
          effort: 5,
          timeOfDay: 'Unknown',
          weather: ''
        };
      }).filter(Boolean);

      return { entries, mappingComplete };
    }

    function renderPlan() {
      planBodyEl.innerHTML = '';
      const planEntries = runs.filter(r => r.source === 'plan');
      if (!planEntries.length) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="4" style=\"text-align:center; color: var(--muted);\">Add your training sessions here</td>`;
        planBodyEl.appendChild(row);
        return;
      }
      planEntries.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
          <td>${item.week || '—'}</td>
          <td>${item.date || '—'}</td>
          <td>${item.distanceText || '—'}</td>
          <td>${item.notes || ''}</td>
        `;
        planBodyEl.appendChild(row);
      });
    }

    function buildSessions() {
      return runs.map(run => {
        const dateValue = normalizeDate(run.date || run.dateValue);
        const plannedMiles = run.plannedMiles ?? run.distance ?? 0;
        const completedMiles = run.completedMiles ?? run.distance ?? null;
        const resultText = run.resultText || run.result || run.notes || '';
        const statusFlags = run.statusFlags || detectStatusFlags(resultText, plannedMiles || 0, completedMiles || 0);
        const weekId = normalizeWeekId(run.weekId || run.week || run.weekLabel, dateValue);
        const sessionType = run.sessionType || detectSessionType(run.distanceText, run.notes || resultText, plannedMiles || 0);
        const effort = Number.isFinite(Number(run.effort)) ? Number(run.effort) : 5;
        const load = calculateRunLoad(completedMiles || 0, effort);
        const riskFlags = run.riskFlags || detectRiskFlags(resultText, run.notes || '');
        return {
          ...run,
          id: run.id || createRunId(),
          source: run.source || 'manual',
          weekId,
          weekLabel: run.week || run.weekLabel || weekId,
          dateLabel: run.date || run.dateLabel || dateValue || '—',
          dateValue,
          plannedMiles: plannedMiles || 0,
          completedMiles,
          type: sessionType || 'General',
          avgHr: run.avgHr ?? null,
          effort,
          temperature: run.temperature ?? null,
          weather: run.weather || '',
          timeOfDay: run.timeOfDay || 'Unknown',
          duration: run.duration || '',
          paceSeconds: run.paceSeconds ?? null,
          resultText,
          statusFlags,
          load,
          riskFlags
        };
      }).filter(Boolean);
    }

    function renderRuns() {
      runLogEl.innerHTML = '';
      const combined = buildSessions();

      if (!combined.length) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="15" style=\"text-align:center; color: var(--muted);\">No runs logged yet</td>`;
        runLogEl.appendChild(row);
        return;
      }

      const sorted = combined.sort((a, b) => {
        const da = new Date(a.dateValue);
        const db = new Date(b.dateValue);
        if (da.toString() === 'Invalid Date' || db.toString() === 'Invalid Date') return 0;
        return da - db;
      });

      sorted.forEach(run => {
        const typeClass = `type-${(run.type || 'general').toLowerCase().replace(/\s+/g, '-')}`;
        const statusClass = run.statusFlags?.statusLabel ? `status-${run.statusFlags.statusLabel.toLowerCase()}` : '';
        const statusText = run.statusFlags?.statusLabel || '—';
        const weatherText = run.weather || (run.temperature != null ? `${run.temperature}°C` : '—');
        const loadText = run.load ? run.load.toFixed(1) : '—';
        const flagsText = (run.riskFlags || []).map(f => `<span class="flag-chip">${f}</span>`).join('');
        const sourceLabel = run.source === 'strava' ? 'Strava' : 'Manual';
        const sourceBadge = `<span class="badge ${run.source === 'strava' ? 'info' : 'success'}">${sourceLabel}</span>`;
        const row = document.createElement('tr');
        if (run.riskFlags && run.riskFlags.length) row.classList.add('risk-row');
        row.innerHTML = `
          <td>${run.dateLabel}</td>
          <td>${run.plannedMiles ? run.plannedMiles.toFixed(1) + ' mi' : '—'}</td>
          <td>${run.completedMiles ? run.completedMiles.toFixed(1) + ' mi' : '—'}</td>
          <td><span class="table-badge ${typeClass}">${run.type || '—'}</span></td>
          <td>${sourceBadge}</td>
          <td class="${statusClass}">${statusText}${flagsText ? '<div>' + flagsText + '</div>' : ''}</td>
          <td>${run.avgHr ?? '—'}</td>
          <td>${run.effort ?? '—'}</td>
          <td>${loadText}</td>
          <td>${weatherText}</td>
          <td>${run.timeOfDay || '—'}</td>
          <td>${run.duration || '—'}</td>
          <td>${formatPace(run.paceSeconds)}</td>
          <td>${run.resultText || '—'}</td>
          <td><button class="badge danger" data-run-id="${run.id}">Delete</button></td>
        `;
        runLogEl.appendChild(row);
      });
    }

    function updateStats() {
      const sessions = buildSessions();
      const totalPlanned = sessions.reduce((sum, s) => sum + (s.plannedMiles || 0), 0);
      const totalCompleted = sessions.reduce((sum, s) => sum + (s.completedMiles || 0), 0);
      const totalRuns = sessions.length;
      const totalSeconds = sessions.reduce((sum, s) => {
        const durationSeconds = parseDurationToSeconds(s.duration);
        if (durationSeconds) return sum + durationSeconds;
        if (s.paceSeconds && s.completedMiles) return sum + (s.paceSeconds * s.completedMiles);
        return sum;
      }, 0);

      const currentWeekId = normalizeWeekId('', new Date());
      const thisWeekSessions = sessions.filter(s => normalizeWeekId(s.weekId, s.dateValue) === currentWeekId);
      const thisWeekMiles = thisWeekSessions.reduce((sum, s) => sum + (s.completedMiles || 0), 0);
      const thisWeekRuns = thisWeekSessions.length;

      const fastestCandidates = sessions.filter(s => s.paceSeconds).map(s => s.paceSeconds);
      const fastest = fastestCandidates.length ? Math.min(...fastestCandidates) : null;
      const longest = sessions.length ? Math.max(...sessions.map(s => (s.completedMiles || s.plannedMiles || 0))) : 0;

      totalDistanceEl.textContent = `${totalPlanned.toFixed(1)} mi`;
      avgWeeklyEl.textContent = `${totalCompleted.toFixed(1)} mi`;
      fastestPaceEl.textContent = fastest ? formatPace(fastest) : '–';
      longestRunEl.textContent = longest ? `${longest.toFixed(1)} mi` : '–';

      totalDistanceEl.title = `Runs: ${totalRuns} · Time: ${formatMinutesSeconds(totalSeconds)}`;
      avgWeeklyEl.title = `This week: ${thisWeekMiles.toFixed(1)} mi across ${thisWeekRuns} run${thisWeekRuns === 1 ? '' : 's'}`;

      // include completed miles in helper message when available
      if (planUploadMessage) {
        const completedNote = totalCompleted ? `Completed so far: ${totalCompleted.toFixed(1)} mi` : '';
        if (!planUploadMessage.textContent && completedNote) {
          setPlanMessage(completedNote, false);
        }
      }
    }

    function computeWeeklySummary() {
      const sessions = buildSessions();
      const weeks = new Map();

      sessions.forEach(session => {
        const weekId = normalizeWeekId(session.weekId || session.weekLabel, session.dateValue);
        const label = session.weekLabel || weekId;
        if (!weeks.has(weekId)) {
          weeks.set(weekId, { id: weekId, label, plannedMiles: 0, completedMiles: 0, sessions: [], firstDate: session.dateValue, fatigueCount: 0, illnessCount: 0, injury: false, weeklyLoad: 0, riskCounts: { illness: 0, injury: 0, heat: 0, sleep: 0, fatigue: 0, any: 0 } });
        }
        const entry = weeks.get(weekId);
        entry.plannedMiles += session.plannedMiles || 0;
        entry.completedMiles += session.completedMiles || 0;
        entry.weeklyLoad += session.load || 0;
        entry.sessions.push(session);
        if (session.statusFlags?.fatigue) entry.fatigueCount += 1;
        if (session.statusFlags?.ill) entry.illnessCount += 1;
        if (session.statusFlags?.injury) entry.injury = true;
        if (session.riskFlags?.length) {
          entry.riskCounts.any += 1;
          if (session.riskFlags.includes('Ill')) entry.riskCounts.illness += 1;
          if (session.riskFlags.includes('Injury')) entry.riskCounts.injury += 1;
          if (session.riskFlags.includes('Heat')) entry.riskCounts.heat += 1;
          if (session.riskFlags.includes('Sleep')) entry.riskCounts.sleep += 1;
          if (session.riskFlags.includes('Fatigue')) entry.riskCounts.fatigue += 1;
        }
        if (!entry.firstDate || (session.dateValue && session.dateValue < entry.firstDate)) entry.firstDate = session.dateValue;
      });

      const sortedWeeks = [...weeks.values()].sort((a, b) => new Date(a.firstDate) - new Date(b.firstDate));
      sortedWeeks.forEach((week, idx) => {
        week.percentCompleted = week.plannedMiles ? ((week.completedMiles / week.plannedMiles) * 100) : 0;
        week.difference = week.completedMiles - week.plannedMiles;
        if (idx > 0) {
          const prev = sortedWeeks[idx - 1];
          week.overtraining = prev.completedMiles > 0 && week.completedMiles > (prev.completedMiles * 1.25);
        } else {
          week.overtraining = false;
        }
      });

      const missedSessions = sessions.filter(s => s.statusFlags?.missed).length;
      const injuryWeeks = sortedWeeks.filter(w => w.injury).map(w => w.label);
      const fatigueWeeks = sortedWeeks.filter((w, idx) => {
        const prev = sortedWeeks[idx - 1];
        return prev && w.completedMiles <= prev.completedMiles * 0.7 && w.sessions.some(s => /fatigue|tired|exhausted/.test((s.resultText || '').toLowerCase()));
      }).map(w => w.label);

      const illnessHeavyWeeks = sortedWeeks.filter(w => w.illnessCount >= 2).map(w => w.id);
      const backFromIllnessWeeks = sortedWeeks.filter((w, idx) => {
        const prev = sortedWeeks[idx - 1];
        if (!prev) return false;
        const prevIll = illnessHeavyWeeks.includes(prev.id);
        return prevIll && (w.plannedMiles ? (w.completedMiles / w.plannedMiles) >= 0.7 : w.completedMiles > 0);
      }).map(w => w.label);

      const riskSummary = sortedWeeks.reduce((acc, w) => {
        acc.any += w.riskCounts?.any || 0;
        acc.illness += w.riskCounts?.illness || 0;
        acc.injury += w.riskCounts?.injury || 0;
        acc.heat += w.riskCounts?.heat || 0;
        acc.sleep += w.riskCounts?.sleep || 0;
        acc.fatigue += w.riskCounts?.fatigue || 0;
        return acc;
      }, { any: 0, illness: 0, injury: 0, heat: 0, sleep: 0, fatigue: 0 });

      return { weeks: sortedWeeks, missedSessions, injuryWeeks, fatigueWeeks, backFromIllnessWeeks, sessions, riskSummary };
    }

    function renderWeeklySummary() {
      const summary = computeWeeklySummary();
      weeklySummaryGrid.innerHTML = '';
      weeklySummaryRows.innerHTML = '';

      if (!summary.weeks.length) {
        weeklySummaryRows.innerHTML = `<tr><td colspan="6" style="text-align:center; color: var(--muted);">No weekly data yet</td></tr>`;
        sessionStatusEl.textContent = 'No sessions yet';
        weatherSummaryEl.textContent = '';
        renderTrendsStats([], []);
        acuteLoadEl.textContent = '–';
        chronicLoadEl.textContent = '–';
        loadRatioEl.textContent = '–';
        loadRiskEl.textContent = '–';
        loadRiskEl.className = 'badge';
        renderInsights([], { weeks: [] });
        return;
      }

      const todayWeekId = normalizeWeekId('', new Date());
      const currentWeek = summary.weeks.find(w => w.id === todayWeekId) || summary.weeks[summary.weeks.length - 1];
      const currentCard = document.createElement('div');
      currentCard.className = 'card';
      currentCard.innerHTML = `
        <div class="stat-label">Latest Week</div>
        <div class="stat-value">${currentWeek.label}</div>
        <p class="helper-text" style="margin:6px 0 4px;">Planned: ${currentWeek.plannedMiles.toFixed(1)} mi · Completed: ${currentWeek.completedMiles.toFixed(1)} mi (${currentWeek.percentCompleted.toFixed(1)}%)</p>
        <p class="helper-text" style="margin:0;">Training load: ${currentWeek.weeklyLoad.toFixed(0)}</p>
      `;
      weeklySummaryGrid.appendChild(currentCard);

      summary.weeks.forEach((week, idx) => {
        const jumpBadge = week.overtraining ? '<span class="badge warning">&gt;25% jump</span>' : '';
        const riskBadge = week.riskCounts?.any >= 3 ? '<span class="badge danger">High-risk week</span>' : '';
        const row = document.createElement('tr');
        row.style.background = week.overtraining ? 'rgba(255,193,71,0.08)' : 'transparent';
        row.innerHTML = `
          <td>${week.label} ${jumpBadge} ${riskBadge}</td>
          <td>${week.plannedMiles.toFixed(1)}</td>
          <td>${week.completedMiles.toFixed(1)}</td>
          <td>${week.weeklyLoad.toFixed(0)}</td>
          <td>${week.percentCompleted.toFixed(1)}%</td>
          <td>${week.difference.toFixed(1)}</td>
        `;
        weeklySummaryRows.appendChild(row);
      });

      renderSessionStatus(summary);
      renderTrainingLoadPanel(summary);
      renderTrendsStats(summary.sessions, summary.weeks);
      renderInsights(summary.sessions, summary);
    }

    function renderSessionStatus(summary) {
      const missed = summary.missedSessions;
      const injuryWeeks = summary.injuryWeeks.length;
      const recoveryWeeks = summary.fatigueWeeks.length;
      const backFromIllness = summary.backFromIllnessWeeks.length;
      const risk = summary.riskSummary || { any: 0, illness: 0, injury: 0, heat: 0, sleep: 0, fatigue: 0 };

      sessionStatusEl.innerHTML = `
        <div style="display:flex; flex-direction:column; gap:8px;">
          <div>❌ Total missed sessions: <strong>${missed}</strong></div>
          <div>🩹 Injury weeks flagged: <strong>${injuryWeeks}</strong></div>
          <div>⚠️ Recovery weeks: <strong>${recoveryWeeks}</strong></div>
          <div>❤️ Back from illness weeks: <strong>${backFromIllness}</strong></div>
          <div>🔥 Risky runs: <strong>${risk.any}</strong> (Ill: ${risk.illness}, Injury: ${risk.injury}, Heat: ${risk.heat}, Sleep: ${risk.sleep})</div>
        </div>
      `;
    }

    function computeTrainingLoadMetrics(sessions = []) {
      const dated = sessions.filter(s => s.dateValue && s.load != null);
      if (!dated.length) return { acute: null, chronic: null, ratio: null, label: 'Not enough data', latestWeekLoad: 0 };
      const byDay = new Map();
      dated.forEach(s => {
        const key = normalizeDate(s.dateValue);
        const current = byDay.get(key) || 0;
        byDay.set(key, current + (s.load || 0));
      });
      const dates = [...byDay.keys()].map(d => new Date(d)).filter(d => d.toString() !== 'Invalid Date');
      if (!dates.length) return { acute: null, chronic: null, ratio: null, label: 'Not enough data', latestWeekLoad: 0 };
      const latest = new Date(Math.max(...dates));
      const acute = [...byDay.entries()].reduce((sum, [d, load]) => {
        const diff = (latest - new Date(d)) / (1000 * 60 * 60 * 24);
        return diff >= 0 && diff <= 6 ? sum + load : sum;
      }, 0);
      const chronicSum = [...byDay.entries()].reduce((sum, [d, load]) => {
        const diff = (latest - new Date(d)) / (1000 * 60 * 60 * 24);
        return diff >= 0 && diff <= 27 ? sum + load : sum;
      }, 0);
      const totalDays = Math.max(...dates) - Math.min(...dates);
      const daySpan = totalDays / (1000 * 60 * 60 * 24);
      const chronic = daySpan < 14 ? null : chronicSum / 4;
      const ratio = chronic && chronic > 0 ? acute / chronic : null;

      let label = 'Not enough data';
      if (ratio != null) {
        if (ratio >= 1.3) label = 'High risk';
        else if (ratio >= 1.0) label = 'Increasing';
        else if (ratio >= 0.8) label = 'Safe / Maintaining';
        else label = 'Undertraining';
      }
      return { acute, chronic, ratio, label };
    }

    function renderTrainingLoadPanel(summary) {
      const metrics = computeTrainingLoadMetrics(summary.sessions);
      const latestWeek = summary.weeks[summary.weeks.length - 1];
      acuteLoadEl.textContent = metrics.acute != null ? metrics.acute.toFixed(0) : '–';
      chronicLoadEl.textContent = metrics.chronic != null ? metrics.chronic.toFixed(0) : '–';
      loadRatioEl.textContent = metrics.ratio != null ? metrics.ratio.toFixed(2) : '–';
      loadRiskEl.textContent = metrics.label;
      loadRiskEl.className = 'badge';
      if (metrics.label === 'High risk') loadRiskEl.classList.add('danger');
      else if (metrics.label === 'Increasing') loadRiskEl.classList.add('warning');
      else if (metrics.label === 'Safe / Maintaining') loadRiskEl.classList.add('success');
      else loadRiskEl.classList.add('info');
      if (latestWeek) {
        loadRiskEl.title = `This week load: ${latestWeek.weeklyLoad.toFixed(0)}`;
      }
    }

    function renderInsights(sessions, summary) {
      const insights = generateInsights(sessions, summary, computeTrainingLoadMetrics(sessions));
      insightsListEl.innerHTML = '';
      if (!insights.length) {
        insightsListEl.innerHTML = '<li class="helper-text">No insights yet — log more runs.</li>';
        return;
      }
      insights.forEach(note => {
        const li = document.createElement('li');
        li.textContent = note;
        insightsListEl.appendChild(li);
      });
    }

    function renderTrendsStats(sessions, weekly) {
      const hrStats = computeHrTrendStats(sessions);
      hrOverallEl.textContent = hrStats.overall ? `${hrStats.overall.toFixed(0)} bpm` : '–';
      hrHalvesEl.textContent = hrStats.halvesText;
      hrTrendEl.textContent = hrStats.label;

      const paceStats = computePaceTrendStats(sessions);
      paceOverallEl.textContent = paceStats.overallText;
      paceHalvesEl.textContent = paceStats.halvesText;
      paceFastestEl.textContent = paceStats.fastestText;
      paceTrendEl.textContent = paceStats.label;

      const weeklyStats = computeWeeklyDistanceStats(weekly);
      weeklyAvgEl.textContent = weeklyStats.avgText;
      weeklyMaxEl.textContent = weeklyStats.maxText;
      weeklyJumpsEl.textContent = weeklyStats.jumpsText;

      const fatigueStats = computeFatigueStats(sessions, weekly);
      fatigueTotalEl.textContent = fatigueStats.totalText;
      fatigueTopEl.textContent = fatigueStats.topWeekText;
      fatigueNoteEl.textContent = fatigueStats.note;

      const timeStats = computeTimeOfDayStats(sessions);
      timeOfDayCountsEl.textContent = timeStats.countsText;
      timeOfDayPacesEl.textContent = timeStats.paceText;
      timeOfDayNoteEl.textContent = timeStats.note;

      const weatherGroups = sessions.reduce((acc, s) => {
        if (!s.weather) return acc;
        const key = s.weather;
        if (!acc[key]) acc[key] = { count: 0, paceTotal: 0, paceCount: 0 };
        acc[key].count += 1;
        if (s.paceSeconds) { acc[key].paceTotal += s.paceSeconds; acc[key].paceCount += 1; }
        return acc;
      }, {});
      const hot = weatherGroups['Hot']?.count || 0;
      const cool = weatherGroups['Cool']?.count || 0;
      const coolPace = weatherGroups['Cool']?.paceCount ? formatPace(weatherGroups['Cool'].paceTotal / weatherGroups['Cool'].paceCount) : '–';
      const hotPace = weatherGroups['Hot']?.paceCount ? formatPace(weatherGroups['Hot'].paceTotal / weatherGroups['Hot'].paceCount) : '–';
      weatherSummaryEl.textContent = `Runs tagged Hot: ${hot} · Cool: ${cool} · Avg pace Cool: ${coolPace}, Hot: ${hotPace}`;
    }

    function splitIntoHalves(values) {
      const mid = Math.ceil(values.length / 2);
      return [values.slice(0, mid), values.slice(mid)];
    }

    function averageNumbers(values) {
      if (!values.length) return null;
      return values.reduce((sum, v) => sum + v, 0) / values.length;
    }

    function computeHrTrendStats(sessions = []) {
      const hrSessions = [...(sessions || [])].filter(s => s.avgHr != null).sort((a, b) => new Date(a.dateValue || a.dateLabel) - new Date(b.dateValue || b.dateLabel));
      if (!hrSessions.length) {
        return { overall: null, halvesText: 'Not enough data yet', label: 'Not enough data yet' };
      }
      const values = hrSessions.map(s => s.avgHr);
      const [first, second] = splitIntoHalves(values);
      const firstAvg = averageNumbers(first);
      const secondAvg = averageNumbers(second);
      const halvesText = (firstAvg != null && secondAvg != null)
        ? `First half: ${firstAvg.toFixed(0)} bpm | Second half: ${secondAvg.toFixed(0)} bpm`
        : 'Not enough data yet';
      let label = 'Not enough data yet';
      if (firstAvg != null && secondAvg != null) {
        const diff = secondAvg - firstAvg;
        if (diff <= -3) label = 'Trending down';
        else if (diff >= 3) label = 'Trending up';
        else label = 'Stable';
      }
      return { overall: averageNumbers(values), halvesText, label };
    }

    function computePaceTrendStats(sessions = []) {
      const pacedSessions = [...(sessions || [])].filter(s => s.paceSeconds != null).sort((a, b) => new Date(a.dateValue || a.dateLabel) - new Date(b.dateValue || b.dateLabel));
      if (!pacedSessions.length) {
        return { overallText: '–', halvesText: 'Not enough data yet', fastestText: 'Fastest: –', label: 'Not enough data yet' };
      }
      const values = pacedSessions.map(s => s.paceSeconds);
      const [first, second] = splitIntoHalves(values);
      const firstAvg = averageNumbers(first);
      const secondAvg = averageNumbers(second);
      const halvesText = (firstAvg != null && secondAvg != null)
        ? `First half: ${formatPace(firstAvg)} | Second half: ${formatPace(secondAvg)}`
        : 'Not enough data yet';
      const fastest = Math.min(...values);
      let label = 'Not enough data yet';
      if (firstAvg != null && secondAvg != null) {
        const diff = secondAvg - firstAvg;
        if (diff <= -10) label = 'Getting faster';
        else if (diff >= 10) label = 'Getting slower';
        else label = 'Stable';
      }
      return {
        overallText: formatPace(averageNumbers(values)),
        halvesText,
        fastestText: `Fastest: ${formatPace(fastest)}`,
        label
      };
    }

    function computeWeeklyDistanceStats(weekly = []) {
      if (!weekly || !weekly.length) {
        return { avgText: '–', maxText: 'Max week: –', jumpsText: 'Weeks with >25% jump: –' };
      }
      const completedValues = weekly.map(w => w.completedMiles || 0);
      const avg = averageNumbers(completedValues);
      const max = Math.max(...completedValues);
      const jumps = weekly.filter(w => w.overtraining).length;
      return {
        avgText: `${avg.toFixed(1)} mi`,
        maxText: `Max week: ${max.toFixed(1)} mi`,
        jumpsText: `Weeks with >25% jump: ${jumps}`
      };
    }

    function computeFatigueStats(sessions = [], weekly = []) {
      const totalFatigue = (sessions || []).filter(s => s.statusFlags?.fatigue).length;
      const topWeek = (weekly || []).reduce((best, w) => {
        if (!best || (w.fatigueCount || 0) > (best.fatigueCount || 0)) return w;
        return best;
      }, null);
      const topWeekLabel = topWeek && topWeek.fatigueCount ? `${topWeek.label} (${topWeek.fatigueCount})` : '–';
      let note = 'Not enough data yet';
      if (totalFatigue > 0) {
        note = topWeek?.fatigueCount >= 3 ? `Watch recovery in ${topWeek.label}` : 'Steady fatigue levels';
      }
      return {
        totalText: String(totalFatigue),
        topWeekText: topWeekLabel,
        note
      };
    }

    function computeTimeOfDayStats(sessions = []) {
      const normalized = (sessions || []).map(s => ({ ...s, timeOfDay: (s.timeOfDay || 'Unknown').toLowerCase() }));
      const morningRuns = normalized.filter(s => s.timeOfDay.includes('morning'));
      const eveningRuns = normalized.filter(s => s.timeOfDay.includes('evening'));
      const morningCount = morningRuns.length;
      const eveningCount = eveningRuns.length;
      const morningPace = averageNumbers(morningRuns.filter(s => s.paceSeconds).map(s => s.paceSeconds));
      const eveningPace = averageNumbers(eveningRuns.filter(s => s.paceSeconds).map(s => s.paceSeconds));
      const countsText = `${morningCount} morning · ${eveningCount} evening`;
      const paceText = `Morning pace: ${formatPace(morningPace)} · Evening pace: ${formatPace(eveningPace)}`;

      let note = 'Not enough data yet';
      if (morningPace && eveningPace) {
        const diff = morningPace - eveningPace;
        const diffSeconds = Math.abs(Math.round(diff));
        const diffText = `${diffSeconds} sec/mi`;
        if (diff <= -1) note = `You run faster in the morning by ~${diffText}`;
        else if (diff >= 1) note = `Evening runs are faster by ~${diffText}`;
        else note = 'No clear difference yet';
      } else if (morningCount || eveningCount) {
        note = 'Need more paced runs to compare times';
      }

      return { countsText, paceText, note };
    }

    function generateInsights(sessions = [], summary = { weeks: [] }, loadMetrics = {}) {
      const notes = [];
      const heatRuns = sessions.filter(s => s.riskFlags?.includes('Heat'));
      if (heatRuns.length >= 2) {
        notes.push(`You had ${heatRuns.length} heat-affected runs recently; consider earlier start times or indoor options.`);
      }

      const morningRuns = sessions.filter(s => (s.timeOfDay || '').toLowerCase().includes('morning') && s.paceSeconds);
      const eveningRuns = sessions.filter(s => (s.timeOfDay || '').toLowerCase().includes('evening') && s.paceSeconds);
      if (morningRuns.length >= 3 && eveningRuns.length >= 3) {
        const morningAvg = averageNumbers(morningRuns.map(r => r.paceSeconds));
        const eveningAvg = averageNumbers(eveningRuns.map(r => r.paceSeconds));
        const diff = morningAvg - eveningAvg;
        const diffSec = Math.abs(Math.round(diff));
        if (diff <= -10) notes.push(`You perform better in the morning by about ${diffSec} sec/mi on average.`);
        else if (diff >= 10) notes.push(`You perform better in the evening by about ${diffSec} sec/mi on average.`);
      }

      const hrComparable = sessions.filter(s => s.avgHr && s.paceSeconds);
      const overallPace = averageNumbers(hrComparable.map(s => s.paceSeconds));
      const similarPace = hrComparable.filter(s => Math.abs(s.paceSeconds - overallPace) <= 30).sort((a, b) => new Date(a.dateValue) - new Date(b.dateValue));
      if (similarPace.length >= 4) {
        const [firstHalf, secondHalf] = splitIntoHalves(similarPace.map(s => s.avgHr));
        const firstAvg = averageNumbers(firstHalf);
        const secondAvg = averageNumbers(secondHalf);
        if (firstAvg && secondAvg) {
          const diff = secondAvg - firstAvg;
          if (diff <= -5) notes.push(`Your average HR at similar pace has dropped by around ${Math.abs(Math.round(diff))} bpm — likely fitness improving.`);
          else if (diff >= 5) notes.push(`Your HR at similar pace has risen by around ${Math.round(diff)} bpm — watch fatigue, stress, or heat.`);
        }
      }

      if (loadMetrics.ratio != null) {
        const percent = Math.round((loadMetrics.ratio - 1) * 100);
        if (loadMetrics.ratio >= 1.3) notes.push(`Your 7-day load is +${percent}% vs your 4-week baseline — high injury risk. Consider a down week.`);
        else if (loadMetrics.ratio < 0.8) notes.push('Your recent load is significantly lower than your 4-week baseline — possible undertraining.');
      }

      const longRuns = sessions.filter(s => (s.type || '').toLowerCase().includes('long'));
      const easyRuns = sessions.filter(s => ['recovery', 'easy'].includes((s.type || '').toLowerCase()) && s.paceSeconds);
      const easyPace = averageNumbers(easyRuns.map(r => r.paceSeconds));
      if (longRuns.length) {
        const hardLongs = longRuns.filter(r => /tired|hard/.test((r.resultText || '').toLowerCase()) || (r.riskFlags || []).some(flag => ['Fatigue', 'Injury'].includes(flag)));
        if (hardLongs.length >= Math.ceil(longRuns.length / 2)) {
          notes.push('Your long runs are consistently feeling hard; review fueling, pacing, and sleep.');
        } else if (easyPace && averageNumbers(longRuns.filter(r => r.paceSeconds).map(r => r.paceSeconds)) <= easyPace * 1.1) {
          notes.push('Your long runs look consistent and controlled — good sign for marathon prep.');
        }
      }

      return notes;
    }

    function buildChart() {
      const ctx = document.getElementById('progressChart');
      const planEntries = runs.filter(r => r.source === 'plan');
      const hasPlan = planEntries.length > 0;
      let labels = [];
      let datasets = [];

      if (hasPlan) {
        labels = planEntries.map(p => p.date || '');
        const plannedData = planEntries.map(p => p.plannedMiles ?? null);
        const completedData = planEntries.map(p => p.completedMiles ?? null);
        datasets = [
          {
            label: 'Planned Distance (mi)',
            data: plannedData,
            borderColor: varColor('--accent'),
            backgroundColor: 'rgba(76, 194, 255, 0.15)',
            tension: 0.3,
            spanGaps: true
          },
          {
            label: 'Completed Distance (mi)',
            data: completedData,
            borderColor: varColor('--accent-2'),
            backgroundColor: 'rgba(120, 245, 197, 0.15)',
            tension: 0.3,
            spanGaps: true
          }
        ];
      } else {
        const sorted = [...runs].sort((a, b) => new Date(a.date) - new Date(b.date));
        labels = sorted.map(r => r.date);
        const paceData = sorted.map(r => r.paceSeconds ? Number((r.paceSeconds / 60).toFixed(2)) : null);
        const distanceData = sorted.map(r => Number(r.distance.toFixed(2)));
        datasets = [
          {
            label: 'Pace (min/mi)',
            data: paceData,
            borderColor: varColor('--accent'),
            backgroundColor: 'rgba(76, 194, 255, 0.15)',
            tension: 0.3,
            yAxisID: 'y'
          },
          {
            label: 'Distance (mi)',
            data: distanceData,
            borderColor: varColor('--accent-2'),
            backgroundColor: 'rgba(120, 245, 197, 0.15)',
            tension: 0.3,
            yAxisID: 'y1'
          }
        ];
      }

      if (chart) chart.destroy();
      if (!labels.length) return;

      chart = new Chart(ctx, {
        type: 'line',
        data: { labels, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { labels: { color: '#dbe3ef' } },
            tooltip: {
              callbacks: {
                label(context) {
                  if (context.dataset.label.includes('Pace')) {
                    const raw = context.raw;
                    if (raw === null || raw === undefined) return `${context.dataset.label}: –`;
                    return `${context.dataset.label}: ${formatPace(Number(raw) * 60)}`;
                  }
                  return `${context.dataset.label}: ${Number(context.raw).toFixed(1)} mi`;
                }
              }
            }
          },
          scales: (() => {
            const base = {
              x: {
                ticks: { color: '#dbe3ef' },
                grid: { color: 'rgba(255,255,255,0.05)' }
              },
              y: {
                type: 'linear',
                position: 'left',
                ticks: { color: '#dbe3ef' },
                grid: { color: 'rgba(255,255,255,0.05)' }
              }
            };
            if (!hasPlan) {
              base.y1 = {
                type: 'linear',
                position: 'right',
                ticks: { color: '#dbe3ef' },
                grid: { drawOnChartArea: false },
                suggestedMin: 0
              };
            }
            return base;
          })()
        }
      });
    }

    function varColor(name) {
      return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    }

    runForm.addEventListener('submit', event => {
      event.preventDefault();
      const date = document.getElementById('run-date').value;
      const distance = parseFloat(document.getElementById('run-distance').value);
      const type = document.getElementById('run-type').value;
      const avgHr = parseInt(document.getElementById('run-hr').value, 10);
      const effort = parseInt(document.getElementById('run-effort').value, 10) || 5;
      const temperature = parseFloat(document.getElementById('run-temp').value);
      const weather = document.getElementById('run-weather').value;
      const timeOfDay = document.getElementById('run-time-of-day').value || 'Morning';
      const duration = document.getElementById('run-duration').value;

      if (!date || !distance || !type) return;
      if (distance <= 0) {
        alert('Please enter a distance greater than 0.');
        return;
      }

      const sessionType = detectSessionType(`${distance} mi ${type}`, '', distance);
      const newRun = addPaceToRun({
        id: createRunId(),
        source: 'manual',
        date,
        distance,
        plannedMiles: distance,
        completedMiles: distance,
        distanceText: `${distance} mi ${type}`,
        sessionType,
        type: sessionType,
        avgHr,
        effort,
        temperature,
        duration,
        weather,
        timeOfDay,
        resultText: ''
      });
      runs.push(newRun);
      runForm.reset();
      refreshUI();
    });

    function normalizeDate(dateValue) {
      if (!dateValue) return '';
      if (dateValue instanceof Date) {
        const iso = dateValue.toISOString();
        return iso.split('T')[0];
      }
      const asString = String(dateValue).trim();
      const parsed = new Date(asString);
      if (parsed.toString() !== 'Invalid Date') {
        return parsed.toISOString().split('T')[0];
      }
      return asString;
    }

    function mapRowToRun(row) {
      const entries = Object.fromEntries(Object.entries(row).map(([k, v]) => [normalizeKey(k), v]));
      const resultsField = entries['results (distance, average hr, average pace)'];
      const extracted = extractFromResultsField(resultsField);

      const date = normalizeDate(entries.date || entries.day || entries['session date']);
      const distanceRaw = entries.distance || entries['distance (mi)'] || entries.miles || extracted.distance;
      const distance = parseDistanceValue(distanceRaw);
      if (!date || !distance) return null;

      const duration = entries.duration || entries.time || entries['duration (hh:mm:ss)'] || '';
      const type = entries.type || entries['type of run'] || entries['session type'] || '';
      const avgHrRaw = entries['avg hr'] ?? entries['average hr'] ?? entries.hr ?? extracted.avgHr;
      const effortRaw = entries.effort ?? entries.rpe;
      const temperatureRaw = entries.temperature ?? entries.temp ?? entries['temperature (°c)'];
      const paceSeconds = extracted.paceSeconds || parsePaceValue(entries['average pace'] || entries.pace);

      const avgHr = avgHrRaw === '' || avgHrRaw === undefined ? null : Number(avgHrRaw);
      const effort = effortRaw === '' || effortRaw === undefined ? 5 : Number(effortRaw);
      const temperature = temperatureRaw === '' || temperatureRaw === undefined ? null : Number(temperatureRaw);

      return addPaceToRun({
        id: createRunId(),
        source: 'manual',
        date,
        distance,
        plannedMiles: distance,
        completedMiles: distance,
        distanceText: entries.distance || '',
        type,
        sessionType: type,
        avgHr,
        effort,
        temperature,
        paceSeconds,
        duration,
        resultText: ''
      });
    }

    function parseExcelFile(file) {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
          const imported = rows
            .map(mapRowToRun)
            .filter(Boolean);
          if (!imported.length) {
            alert('No usable rows found. Please ensure the sheet includes Date and Distance columns (Week/Results/Notes optional).');
            return;
          }
          runs = runs.concat(imported);
          refreshUI();
        } catch (err) {
          console.error('Failed to import Excel file', err);
          alert('Unable to import this Excel file. Please upload a .xlsx, .xls, or .csv file with the expected columns.');
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function handleExcelInput(event) {
      const file = event.target.files?.[0];
      if (!file) return;
      const allowedTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv',
        'application/vnd.ms-excel.sheet.macroEnabled.12'
      ];
      const isAllowed = allowedTypes.includes(file.type) || /\.(xlsx|xls|csv)$/i.test(file.name);
      if (!isAllowed) {
        alert('Unsupported file type. Please upload an .xlsx, .xls, or .csv workbook.');
        event.target.value = '';
        return;
      }
      parseExcelFile(file);
      event.target.value = '';
    }

    function handlePlanCsvInput(event) {
      const file = event.target.files?.[0];
      if (!file) return;
      const isCsv = file.type === 'text/csv' || /\.csv$/i.test(file.name);
      if (!isCsv) {
        setPlanMessage('Please upload a CSV file exported from your spreadsheet.', true);
        event.target.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onload = e => {
        try {
          const parsed = parsePlanRowsFromCsv(e.target.result);
          if (!parsed.mappingComplete) {
            setPlanMessage('No usable rows found. Please ensure the sheet has columns: Week, Date, Distance, Results (distance, Avg hr, Avg pace), notes.', true);
            return;
          }
          if (!parsed.entries.length) {
            setPlanMessage('No plan rows found in the CSV.', true);
            return;
          }
          runs = runs.filter(r => r.source !== 'plan');
          runs = runs.concat(parsed.entries);
          setPlanMessage(`Loaded ${parsed.entries.length} plan entries.`, false);
          refreshUI();
        } catch (err) {
          console.error('Failed to parse plan CSV', err);
          setPlanMessage('No usable rows found. Please ensure the sheet has columns: Week, Date, Distance, Results (distance, Avg hr, Avg pace), notes.', true);
        }
      };
      reader.readAsText(file);
      event.target.value = '';
    }

    excelUpload.addEventListener('change', handleExcelInput);
    excelHistoryUpload.addEventListener('change', handleExcelInput);
    planCsvUpload.addEventListener('change', handlePlanCsvInput);
    exportJsonBtn.addEventListener('click', exportRunsToJson);
    importJsonInput.addEventListener('change', handleJsonImport);
    if (exportJsonSettingsBtn) exportJsonSettingsBtn.addEventListener('click', exportRunsToJson);
    if (importJsonSettingsInput) importJsonSettingsInput.addEventListener('change', handleJsonImport);
    if (userNotesEl) userNotesEl.addEventListener('input', saveUserNotes);
    document.getElementById('clear-data').addEventListener('click', () => {
      if (confirm('This will remove all loaded runs (from uploads and manual entries) from this session. Are you sure?')) {
        runs = [];
        if (chart) { chart.destroy(); chart = null; }
        try { localStorage.removeItem(STORAGE_KEY); } catch (err) { console.warn('Failed clearing storage', err); }
        refreshUI();
        setPlanMessage('', false);
      }
    });

    runLogEl.addEventListener('click', (event) => {
      const target = event.target;
      if (target.matches('button[data-run-id]')) {
        const runId = target.getAttribute('data-run-id');
        if (confirm('Delete this run?')) {
          runs = runs.filter(r => r.id !== runId);
          refreshUI();
        }
      }
    });

    addPlanBtn.addEventListener('click', () => {
      const week = document.getElementById('plan-week').value;
      const date = document.getElementById('plan-date').value;
      const distanceText = document.getElementById('plan-distance').value;
      const notes = document.getElementById('plan-notes').value;

      if (![week, date, distanceText, notes].some(Boolean)) return;
      const plannedMiles = extractMilesFromText(distanceText);
      const sessionType = detectSessionType(distanceText, notes, plannedMiles || 0);
      runs.push({
        id: createRunId(),
        source: 'plan',
        week,
        date,
        weekId: normalizeWeekId(week, date),
        distanceText,
        notes,
        plannedMiles,
        completedMiles: null,
        resultText: '',
        paceSeconds: null,
        avgHr: null,
        sessionType,
        statusFlags: detectStatusFlags(notes, plannedMiles || 0, 0),
        riskFlags: detectRiskFlags('', notes),
        timeOfDay: 'Unknown',
        weather: ''
      });
      document.getElementById('plan-week').value = '';
      document.getElementById('plan-date').value = '';
      document.getElementById('plan-distance').value = '';
      document.getElementById('plan-notes').value = '';
      refreshUI();
    });

    function refreshUI() {
      saveRunsToStorage(runs);
      renderRuns();
      updateStats();
      buildChart();
      renderPlan();
      renderWeeklySummary();
    }

    async function initPlanner() {
      setActiveSection('dashboard-section');
      loadUserNotes();
      runs = loadRunsFromStorage();
      refreshUI();
      await syncStravaRunsIntoLog();
      console.log('Unified run log:', {
        total: runs.length,
        strava: runs.filter(r => r.source === 'strava').length,
        manual: runs.filter(r => r.source !== 'strava').length
      });
      await fetchLatestStravaRun();
    }

    initPlanner().catch(console.error);
});
