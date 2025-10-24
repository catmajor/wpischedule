export default function createICSFromExcel(json) {
  const cleaned = {};
  let inEnrolled = -1;
  for (const [key, cell] of Object.entries(json.Sheets.Sheet1)) {
    const row = parseInt(key.replace(/[^0-9]/g, ''));
    if (cell.v === "My Completed Courses") {
      inEnrolled = -1;
    }
    if (inEnrolled > -1 && row >= inEnrolled) {
      cleaned[key] = cell;
    }
    if (cell.v === "Enrolled Sections") {
      inEnrolled = row + 2;
    } 
  }

  // generate-ics-from-cells.js
  // Node script: paste your JSON into the `cells` variable below, run `node generate-ics-from-cells.js`
  // Produces 'schedule.ics' in the same folder.

  /* -------------------------
     Put your JSON here (the big object you posted).
     For brevity I declare `cells` as a placeholder. Replace with actual object.
     ------------------------- */
  const cells = cleaned;
  /* REPLACE THIS COMMENT WITH YOUR JSON OBJECT */ 
  /* -------------------------
     Helpers
     ------------------------- */
  const DAYS_MAP = {
    M: 'MO', T: 'TU', W: 'WE', R: 'TH', F: 'FR', S: 'SA', U: 'SU', SU: 'SU'
  };

  function excelSerialToDate(serial) {
    // Excel serial number -> JS Date (UTC midnight)
    // Epoch: 1899-12-30 is serial 0 in this conversion which matches most Excel exports.
    // We'll treat serial as whole days (no fractional times in the provided data).
    const msPerDay = 24 * 60 * 60 * 1000;
    const epoch = Date.UTC(1899, 11, 30); // 1899-12-30
    return new Date(epoch + Math.round(serial) * msPerDay);
  }

  function pad2(n) { return n.toString().padStart(2,'0'); }

  function formatDateTimeForTZID(date) {
    // Format like: YYYYMMDDTHHMMSS (no Z) for TZID lines
    return date.getFullYear().toString()
      + pad2(date.getMonth()+1)
      + pad2(date.getDate())
      + 'T'
      + pad2(date.getHours())
      + pad2(date.getMinutes())
      + pad2(date.getSeconds());
  }

  function formatUTCDateTimeForUntil(date) {
    // format as YYYYMMDDTHHMMSSZ (UTC)
    const d = new Date(date.getTime());
    return d.getUTCFullYear().toString()
      + pad2(d.getUTCMonth()+1)
      + pad2(d.getUTCDate())
      + 'T'
      + pad2(d.getUTCHours())
      + pad2(d.getUTCMinutes())
      + pad2(d.getUTCSeconds())
      + 'Z';
  }

  function parseScheduleCell(kValue) {
    // Expected format: "<days> | <start time> - <end time> | <location>"
    // Example: "M-T-R-F | 3:00 PM - 3:50 PM | Unity Hall 520"
    if (!kValue || typeof kValue !== 'string') return null;
    const parts = kValue.split('|').map(p => p.trim());
    const [daysPart = '', timesPart = '', locPart = ''] = parts;
    // days: split by '-' or ',' or space
    const rawDays = daysPart.split(/[-, ]+/).filter(Boolean);
    // time: "3:00 PM - 3:50 PM"
    const timeMatch = timesPart.match(/(.+?)\s*-\s*(.+)/);
    const startTimeStr = timeMatch ? timeMatch[1].trim() : null;
    const endTimeStr = timeMatch ? timeMatch[2].trim() : null;
    return {
      days: rawDays,
      startTimeStr,
      endTimeStr,
      location: locPart
    };
  }

  function parseTimeStringToHM(timeStr) {
    // e.g. "3:00 PM" -> {hours:15, minutes:0}
    if (!timeStr) return null;
    const m = timeStr.match(/^(\d{1,2})(?::(\d{1,2}))?\s*(AM|PM)?$/i);
    if (!m) return null;
    let hh = parseInt(m[1],10);
    const mm = m[2] ? parseInt(m[2],10) : 0;
    const ampm = m[3] ? m[3].toUpperCase() : null;
    if (ampm) {
      if (ampm === 'PM' && hh !== 12) hh += 12;
      if (ampm === 'AM' && hh === 12) hh = 0;
    }
    return { hours: hh, minutes: mm };
  }

  function dateAddDays(date, n) {
    return new Date(date.getTime() + n * 24*60*60*1000);
  }

  function weekdayNumberToCode(n) {
    // JS getDay: 0=Sun .. 6=Sat, convert to RFC 5545 codes
    switch(n){
      case 0: return 'SU';
      case 1: return 'MO';
      case 2: return 'TU';
      case 3: return 'WE';
      case 4: return 'TH';
      case 5: return 'FR';
      case 6: return 'SA';
    }
  }

  /* -------------------------
     Main generator
     ------------------------- */
  function generateICS(cells, options = { tzid: 'America/New_York' }) {
    // group keys by row number
    const rows = {};
    for (const key of Object.keys(cells)) {
      const match = key.match(/^([A-Z]+)(\d+)$/);
      if (!match) continue;
      const col = match[1];
      const row = match[2];
      rows[row] = rows[row] || {};
      rows[row][col] = cells[key] && cells[key].v !== undefined ? cells[key].v : null;
    }

    const events = [];

    Object.keys(rows).sort((a,b)=>Number(a)-Number(b)).forEach((r, idx) => {
      const row = rows[r];
      // Need at least B (course), K (schedule), M (start), N (end)
      if (!row.B || !row.K || !row.M || !row.N) return;

      const course = String(row.B).trim();
      const section = row.G ? String(row.G).trim() : '';
      const instructor = row.L ? String(row.L).trim() : '';
      const schedule = String(row.K).trim();
      const startSerial = Number(row.M);
      const endSerial = Number(row.N);

      if (isNaN(startSerial) || isNaN(endSerial)) return;

      const startDate = excelSerialToDate(startSerial); // local midnight UTC-based
      const endDate = excelSerialToDate(endSerial);

      const parsed = parseScheduleCell(schedule);
      if (!parsed || !parsed.startTimeStr || !parsed.endTimeStr || !parsed.days) return;

      const startHM = parseTimeStringToHM(parsed.startTimeStr);
      const endHM = parseTimeStringToHM(parsed.endTimeStr);
      if (!startHM || !endHM) return;

      // Build BYDAY codes
      const byday = parsed.days.map(s => {
        const normalized = s.trim();
        // handle two-letter 'Su' etc.
        if (DAYS_MAP[normalized]) return DAYS_MAP[normalized];
        if (DAYS_MAP[normalized.toUpperCase()]) return DAYS_MAP[normalized.toUpperCase()];
        // single-letter fallback
        if (normalized.length === 1 && DAYS_MAP[normalized.toUpperCase()]) return DAYS_MAP[normalized.toUpperCase()];
        // otherwise ignore
        return null;
      }).filter(Boolean);

      if (!byday.length) return;

      // Find first occurrence (first date on/after startDate with weekday in byday)
      // map byday codes back to JS weekday number
      const bydayNums = byday.map(code => {
        switch(code) {
          case 'SU': return 0;
          case 'MO': return 1;
          case 'TU': return 2;
          case 'WE': return 3;
          case 'TH': return 4;
          case 'FR': return 5;
          case 'SA': return 6;
        }
      });
      // convert startDate to local date object at midnight in local tz
      const localStartDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      let firstDate = new Date(localStartDate.getTime());
      let attempts = 0;
      while (!bydayNums.includes(firstDate.getDay()) && attempts < 14) {
        firstDate = dateAddDays(firstDate, 1);
        attempts++;
      }

      // Set times on firstDate
      const dtStart = new Date(firstDate.getFullYear(), firstDate.getMonth(), firstDate.getDate(), startHM.hours, startHM.minutes, 0);
      const dtEnd = new Date(firstDate.getFullYear(), firstDate.getMonth(), firstDate.getDate(), endHM.hours, endHM.minutes, 0);

      // UNTIL: we use endDate at 23:59:59 UTC to be safe and inclusive
      let untilUtc = new Date(Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate(), 23, 59, 59));
      //add an extra day to UNTIL to ensure inclusion of endDate occurrences
      untilUtc.setUTCDate(untilUtc.getUTCDate() + 1); // adds 1 day
      // Build description and location
      const location = parsed.location || '';
      const descriptionParts = [
        course,
        section ? `Section: ${section}` : null,
        instructor ? `Instructor: ${instructor}` : null,
        `Schedule: ${schedule}`
      ].filter(Boolean);
      const description = descriptionParts.join(' ');

      // UID
      const uid = `${Date.now()}-${r}-${Math.random().toString(36).slice(2,9)}@generated`;

      // Build VEVENT
      const vevent = {
        uid,
        dtstamp: formatUTCDateTimeForUntil(new Date()), // current time in UTC format with Z
        dtstart: formatDateTimeForTZID(dtStart),
        dtend: formatDateTimeForTZID(dtEnd),
        tzid: options.tzid,
        summary: course,
        description,
        location,
        rrule: `FREQ=WEEKLY;BYDAY=${byday.join(',')};UNTIL=${formatUTCDateTimeForUntil(untilUtc)}`
      };

      events.push(vevent);
    });

    // Build ICS string
    const lines = [];

    lines.push('BEGIN:VCALENDAR');
    lines.push('PRODID:-//Generated by generate-ics-from-cells.js//EN');
    lines.push('VERSION:2.0');
    lines.push('CALSCALE:GREGORIAN');
    lines.push('METHOD:PUBLISH');

    // Add a VTIMEZONE block for America/New_York (DST aware)
    lines.push('BEGIN:VTIMEZONE');
    lines.push('TZID:America/New_York');
    lines.push('X-LIC-LOCATION:America/New_York');
    lines.push('BEGIN:DAYLIGHT');
    lines.push('TZOFFSETFROM:-0500');
    lines.push('TZOFFSETTO:-0400');
    lines.push('TZNAME:EDT');
    lines.push('DTSTART:19700308T020000');
    lines.push('RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU');
    lines.push('END:DAYLIGHT');
    lines.push('BEGIN:STANDARD');
    lines.push('TZOFFSETFROM:-0400');
    lines.push('TZOFFSETTO:-0500');
    lines.push('TZNAME:EST');
    lines.push('DTSTART:19701101T020000');
    lines.push('RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU');
    lines.push('END:STANDARD');
    lines.push('END:VTIMEZONE');

    // Note: For full timezone correctness you'd include a full VTIMEZONE block for America/New_York.
    // Many calendar apps accept TZID without an embedded VTIMEZONE. If you need full VTIMEZONE included,
    // I can add a standard VTIMEZONE block for America/New_York.
    for (const ev of events) {
      lines.push('BEGIN:VEVENT');
      lines.push(`UID:${ev.uid}`);
      lines.push(`DTSTAMP:${ev.dtstamp}`);
      lines.push(`DTSTART;TZID=${ev.tzid}:${ev.dtstart}`);
      lines.push(`DTEND;TZID=${ev.tzid}:${ev.dtend}`);
      lines.push(`SUMMARY:${escapeICalText(ev.summary)}`);
      lines.push(`DESCRIPTION:${escapeICalText(ev.description)}`);
      if (ev.location) lines.push(`LOCATION:${escapeICalText(ev.location)}`);
      lines.push(`RRULE:${ev.rrule};WKST=SU`);
      lines.push('END:VEVENT');
    }

    lines.push('END:VCALENDAR');

    return lines.join('\r\n');
  }

  function escapeICalText(s) {
    if (!s) return '';
    return String(s).replace(/\\/g,'\\\\').replace(/\n/g,'\\n').replace(/;/g,'\\;').replace(/,/g,'\\,');
  }

  /* -------------------------
     Run & write file
     ------------------------- */
  try {
    const ics = generateICS(cells, { tzid: 'America/New_York' });
    console.log('Wrote schedule.ics with', (ics.match(/BEGIN:VEVENT/g) || []).length, 'events.');
    return ics;
  } catch (err) {
    console.error('Error generating ICS:', err);
    return "";
  }
}
