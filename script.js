/**
 * Duke Shift Handover Generator
 * Pure frontend — parses Excel TSV data and builds formatted handover emails.
 */

/* ============================================
   Constants
   ============================================ */

const PRIORITIES = ['Critical', 'High', 'Medium', 'Low'];
const JURISDICTIONS = ['DEF', 'DEM', 'DEC', 'DEP'];

const DAILY_REPORT_LINK =
  'https://cognizantonline-my.sharepoint.com/:x:/g/personal/2309688_cognizant_com/EXvxf7MeK45Pm5hWcQKQGPkBCsemonaDPR9fW57EhzA5ZQ';

/** Email table colour palette (from original .eml template) */
const COLORS = {
  headerOrange: 'rgb(198, 89, 17)',
  headerPeach: 'rgb(247, 202, 172)',
  cellPeach: 'rgb(252, 228, 214)',
  totalGray: 'rgb(219, 219, 219)',
  actionOrange: 'rgb(244, 176, 131)',
  notesYellow: 'rgb(255, 242, 204)',
  border: 'rgb(51, 51, 51)',
};

/** Static NOTES and UPDATES content from the handover template */
const STATIC_NOTES = [
  'Kindly Go through Helpdesk scenario and ALL DCC App List for Reporting Issue File which is send by Wayne in KT Mail and follow while creating any tickets.',
  'kindly confirm with user for Maps issue is ADMS Maps and Outages maps mostly ticket will go to Modeling team but confirm first with Wayne also once',
  'If you received mail with subject "PROD: Calls not processed" which is being forwarded by Beena mail id, Please inform Beena,Kartik, Raghu or Duke_OMS_FL DL immediately.',
  'Check new SOP of Model Push [Note :- After Model Push Ask DCC control room about everything up & running or not.',
];

const STATIC_NOTES_CONTINUATION =
  '              (Customer search, Device search & Geographic Map)]';

const STATIC_NOTES_EXTRA = [
  'One minor change in UAC jobs check the changes as updated in group',
];

const STATIC_UPDATES = [
  'Strom is occur in DEF so be <b>PROACTIVE</b> also you guy\u2019s added in STORM 01/09/2024 Group so response in that Quickly. Also, PING that group in your chat.',
  'If you guys face any issue or need any kind of help regarding issues in your shift, then connect with Akansha/ Karthik or other SME\u2019S',
  'Please Notice or find out UAC job pattern and update to other shift members. (UAC job failed in 1 hr or 1hr 15 min etc like this )',
  'Please Follow same pattern (All team members) Put UAC tickets on Hold first then Resolve it after Iteration run successfully.',
  'Also For Modelling issue sent ticket to representative group like modelling DEM, DEC ETC.',
  'UAC Update (Param shared in UAC group). We are added into DL in mail where we will receive job failure alert for\u201D OMS_ADMS_STORED_PROCE_DEF\u201D we have to monitor both Mail and Resolved incident kindly check UAC group',
];

const STATIC_UPDATES_CONTINUATION =
  'Go through the latest recording of HEALTH CHECK AND MAKE SOP on it. Also check your access on all applications.';

const STATIC_UPDATES_HIGHLIGHT = [
  'From 11-03-2024 we have to create tickets for DEC/DEM ODI Load Monitoring Once we receive the alert, we need to check the LAG query which we have mentioned in the document which is shared by Pooja and we have to resolve that ticket on our names as we resolve tickets for UAC and shift member will give full handover and Update in KT call and go through that document',
  'Whenever we creating a ticket for OMS,we have to inform that on internal group same as we do for DMS IT issues.',
  'Please read internal help desk DEF group message',
];

/* ============================================
   DOM References
   ============================================ */

const els = {
  currentShift: document.getElementById('current-shift'),
  userName: document.getElementById('user-name'),
  incidentData: document.getElementById('incident-data'),
  btnGenerate: document.getElementById('btn-generate'),
  btnCopy: document.getElementById('btn-copy'),
  emailPreview: document.getElementById('email-preview'),
  statusMessage: document.getElementById('status-message'),
};

/** Holds the last generated email HTML for clipboard copy */
let generatedEmailHtml = '';

/* ============================================
   Shift Logic
   ============================================ */

/**
 * Returns the handover target shift number.
 * Shift 1 → 2, Shift 2 → 3, Shift 3 → 1
 */
function getHandoverShift(currentShift) {
  const map = { 1: 2, 2: 3, 3: 1 };
  return map[Number(currentShift)] ?? 2;
}

/* ============================================
   Incident Data Parsing
   ============================================ */

/**
 * Extracts plain text from the contenteditable paste area,
 * normalising line breaks from pasted HTML or plain text.
 */
function getPasteAreaText(element) {
  const html = element.innerHTML;
  if (html.includes('<table') || html.includes('<tr')) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const rows = doc.querySelectorAll('tr');
    if (rows.length) {
      return Array.from(rows)
        .map((row) =>
          Array.from(row.querySelectorAll('td, th'))
            .map((cell) => cell.textContent.trim())
            .join('\t')
        )
        .join('\n');
    }
  }
  return element.innerText || element.textContent || '';
}

/**
 * Splits raw pasted text into a 2-D array of cell values.
 */
function parseTsvRows(rawText) {
  return rawText
    .split(/\r?\n/)
    .map((line) => line.trimEnd())
    .filter((line) => line.trim().length > 0)
    .map((line) => line.split('\t').map((cell) => cell.trim()));
}

/**
 * Checks whether the first row looks like a header row.
 */
function isHeaderRow(cells) {
  const joined = cells.join(' ').toLowerCase();
  return (
    joined.includes('priority') ||
    joined.includes('incident') ||
    joined.includes('jurisdiction') ||
    joined.includes('application')
  );
}

/**
 * Finds a column index by matching header name fragments.
 */
function findColumnByHeader(headers, ...keywords) {
  const lower = headers.map((h) => h.toLowerCase());
  for (let i = 0; i < lower.length; i++) {
    if (keywords.some((kw) => lower[i].includes(kw))) return i;
  }
  return -1;
}

/**
 * Scores each column for how likely it holds priority values.
 */
function scorePriorityColumn(rows) {
  const scores = [];
  const maxCols = Math.max(...rows.map((r) => r.length));
  for (let col = 0; col < maxCols; col++) {
    let hits = 0;
    for (const row of rows) {
      if (PRIORITIES.includes(normalisePriority(row[col]))) hits++;
    }
    scores[col] = hits;
  }
  return scores;
}

/**
 * Scores each column for jurisdiction codes.
 */
function scoreJurisdictionColumn(rows) {
  const scores = [];
  const maxCols = Math.max(...rows.map((r) => r.length));
  for (let col = 0; col < maxCols; col++) {
    let hits = 0;
    for (const row of rows) {
      if (JURISDICTIONS.includes(extractJurisdiction(row[col]))) hits++;
    }
    scores[col] = hits;
  }
  return scores;
}

/**
 * Scores each column for application/system keywords.
 */
function scoreApplicationColumn(rows) {
  const scores = [];
  const maxCols = Math.max(...rows.map((r) => r.length));
  for (let col = 0; col < maxCols; col++) {
    let hits = 0;
    for (const row of rows) {
      const val = (row[col] || '').toUpperCase();
      if (val.includes('OUTAGE') || val.includes('DMS') || val.includes('DSCADA') || val.includes('SUPPORT')) {
        hits++;
      }
    }
    scores[col] = hits;
  }
  return scores;
}

/** Returns the index of the highest-scoring column. */
function indexOfMax(scores) {
  let maxIdx = 0;
  let maxVal = -1;
  for (let i = 0; i < scores.length; i++) {
    if (scores[i] > maxVal) {
      maxVal = scores[i];
      maxIdx = i;
    }
  }
  return maxIdx;
}

/**
 * Detects which columns hold Priority, Application, and Jurisdiction.
 */
function detectColumns(rows) {
  if (!rows.length) {
    return { priority: 3, application: 4, jurisdiction: 6 };
  }

  let dataRows = rows;
  let priorityCol = -1;
  let applicationCol = -1;
  let jurisdictionCol = -1;

  if (isHeaderRow(rows[0])) {
    const headers = rows[0];
    dataRows = rows.slice(1);
    priorityCol = findColumnByHeader(headers, 'priority');
    applicationCol = findColumnByHeader(headers, 'application', 'system', 'assignment group', 'service');
    jurisdictionCol = findColumnByHeader(headers, 'jurisdiction', 'juris');
  }

  if (dataRows.length === 0) dataRows = rows;

  if (priorityCol < 0) priorityCol = indexOfMax(scorePriorityColumn(dataRows));
  if (applicationCol < 0) applicationCol = indexOfMax(scoreApplicationColumn(dataRows));
  if (jurisdictionCol < 0) jurisdictionCol = indexOfMax(scoreJurisdictionColumn(dataRows));

  // Avoid column collisions — reassign duplicates
  const used = new Set();
  for (const [key, val] of [
    ['priority', priorityCol],
    ['application', applicationCol],
    ['jurisdiction', jurisdictionCol],
  ]) {
    if (used.has(val)) {
      if (key === 'application') applicationCol = val + 1;
      if (key === 'jurisdiction') jurisdictionCol = val + 2;
    }
    used.add(key === 'priority' ? priorityCol : key === 'application' ? applicationCol : jurisdictionCol);
  }

  return { priority: priorityCol, application: applicationCol, jurisdiction: jurisdictionCol, dataRows };
}

/** Normalises a priority string to one of the four supported values. */
function normalisePriority(value) {
  if (!value) return '';
  const cleaned = value.trim();
  const match = PRIORITIES.find((p) => cleaned.toLowerCase() === p.toLowerCase());
  return match || '';
}

/** Extracts a jurisdiction code from a cell value. */
function extractJurisdiction(value) {
  if (!value) return '';
  const upper = value.trim().toUpperCase();
  return JURISDICTIONS.find((j) => upper === j || upper.includes(j)) || '';
}

/**
 * Classifies an application/system cell as OMS, DMS, or Others.
 */
function classifyApplication(value) {
  const upper = (value || '').toUpperCase();
  if (upper.includes('OUTAGE MANAGEMENT SYSTEM')) return 'oms';
  if (upper.includes('DMS')) return 'dms';
  return 'others';
}

/**
 * Parses all incident rows and returns aggregated counts.
 */
function parseIncidents(rawText) {
  const rows = parseTsvRows(rawText);
  if (!rows.length) {
    return {
      incidents: [],
      counts: createEmptyCounts(),
      jurisdictions: createEmptyJurisdictions(),
      warnings: ['No incident data found. Tables will show zero counts.'],
    };
  }

  const { priority: priorityCol, application: applicationCol, jurisdiction: jurisdictionCol, dataRows } =
    detectColumns(rows);

  const incidents = [];
  const counts = createEmptyCounts();
  const jurisdictions = createEmptyJurisdictions();
  const warnings = [];

  for (const row of dataRows) {
    const priority = normalisePriority(row[priorityCol]);
    if (!priority) continue;

    const appClass = classifyApplication(row[applicationCol]);
    const jurisdiction = extractJurisdiction(row[jurisdictionCol]);

    incidents.push({ priority, appClass, jurisdiction, raw: row });

    counts[priority][appClass]++;
    counts[priority].transferred =
      counts[priority].oms + counts[priority].dms + counts[priority].others;

    if (jurisdiction) {
      jurisdictions[jurisdiction]++;
    }
  }

  if (incidents.length === 0) {
    warnings.push('No rows with recognised Priority values (Critical, High, Medium, Low) were found.');
  }

  // Compute column totals
  counts.totals = { oms: 0, dms: 0, others: 0, transferred: 0 };
  for (const p of PRIORITIES) {
    counts.totals.oms += counts[p].oms;
    counts.totals.dms += counts[p].dms;
    counts.totals.others += counts[p].others;
    counts.totals.transferred += counts[p].transferred;
  }

  return { incidents, counts, jurisdictions, warnings };
}

function createEmptyCounts() {
  const row = () => ({ oms: 0, dms: 0, others: 0, transferred: 0 });
  return {
    Critical: row(),
    High: row(),
    Medium: row(),
    Low: row(),
    totals: row(),
  };
}

function createEmptyJurisdictions() {
  return { DEF: 0, DEM: 0, DEC: 0, DEP: 0 };
}

/* ============================================
   Email HTML Generation
   ============================================ */

/** Inline cell style helper matching the original email template. */
function cellStyle(bg, extra = '') {
  return `border:1.3px solid ${COLORS.border};background-color:${bg};padding:4pt 5.4pt;vertical-align:top;${extra}`;
}

/** Formats a count for display — shows number or blank spaces for zero. */
function fmtCount(n) {
  return n > 0 ? String(n) : '&nbsp;';
}

/**
 * Builds the main Shift Handover Report table HTML.
 */
function buildMainTable(counts, jurisdictions) {
  const priorityRows = PRIORITIES.map((priority) => {
    const c = counts[priority];
    return `
      <tr>
        <td style="${cellStyle(COLORS.headerPeach)}">
          <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>${priority}</b></span></p>
        </td>
        <td style="${cellStyle(COLORS.cellPeach)}">
          <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black">${fmtCount(c.oms)}</span></p>
        </td>
        <td colspan="2" style="${cellStyle(COLORS.cellPeach)}">
          <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black">${fmtCount(c.dms)}</span></p>
        </td>
        <td style="${cellStyle(COLORS.cellPeach)}">
          <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black">${fmtCount(c.others)}</span></p>
        </td>
        <td style="${cellStyle(COLORS.cellPeach)}">
          <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black">&nbsp;</span></p>
        </td>
        <td style="${cellStyle(COLORS.cellPeach)}">
          <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>${fmtCount(c.transferred)}</b></span></p>
        </td>
      </tr>`;
  }).join('');

  const t = counts.totals;

  return `
    <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;max-width:730pt;margin-bottom:12pt">
      <tbody>
        <tr>
          <td style="${cellStyle(COLORS.headerOrange, 'width:67pt;height:21pt')}"></td>
          <td colspan="7" style="${cellStyle(COLORS.headerOrange, 'height:21pt;text-align:center')}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:16pt;color:black"><b>Shift Handover Report</b></span></p>
          </td>
        </tr>
        <tr>
          <td rowspan="7" style="${cellStyle(COLORS.cellPeach, 'width:67pt')}">
            <p style="margin:0"><span style="font-family:Calibri,sans-serif;font-size:10pt;color:black"><b>SNOW Incidents Created</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.headerPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Priority</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.headerPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>No of tickets created for OMS</b></span></p>
          </td>
          <td colspan="2" style="${cellStyle(COLORS.headerPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>No of tickets created for DMS</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.headerPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>No of tickets created for Others</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.headerPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Tickets resolved by L1</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.headerPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Tickets Transferred to L2</b></span></p>
          </td>
        </tr>
        ${priorityRows}
        <tr>
          <td style="${cellStyle(COLORS.totalGray)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Total Ticket Counts</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>${fmtCount(t.oms)}</b></span></p>
          </td>
          <td colspan="2" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>${fmtCount(t.dms)}</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>${fmtCount(t.others)}</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>&nbsp;</b></span></p>
          </td>
          <td style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0;text-align:center"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>${fmtCount(t.transferred)}</b></span></p>
          </td>
        </tr>
        <tr>
          <td colspan="3" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Number of UAC INC received :&nbsp;</b></span></p>
          </td>
          <td colspan="4" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Number of UAC INC Resolved :</b></span></p>
          </td>
        </tr>
        <tr>
          <td rowspan="2" style="${cellStyle(COLORS.cellPeach, 'width:67pt')}">
            <p style="margin:0"><span style="font-family:Calibri,sans-serif;font-size:10pt;color:black"><b>Jurisdiction</b></span></p>
          </td>
          <td colspan="3" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>DEF:&nbsp; ${jurisdictions.DEF || ''}</b></span></p>
          </td>
          <td colspan="4" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>&nbsp; DEP:&nbsp; ${jurisdictions.DEP || ''}</b></span></p>
          </td>
        </tr>
        <tr>
          <td colspan="3" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>DEM: ${jurisdictions.DEM || ''}</b></span></p>
          </td>
          <td colspan="4" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black">&nbsp; <b>DEC: ${jurisdictions.DEC || ''}</b></span></p>
          </td>
        </tr>
        <tr>
          <td style="${cellStyle(COLORS.cellPeach, 'width:67pt')}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Model Push Activity</b></span></p>
          </td>
          <td colspan="3" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Go/No-Go and Issue reported by DCC-&nbsp; &nbsp;No</b></span></p>
          </td>
          <td colspan="4" style="${cellStyle(COLORS.cellPeach)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>&nbsp; Go and No Issue reported</b></span></p>
          </td>
        </tr>
      </tbody>
    </table>`;
}

/** Builds list items as HTML <li> elements. */
function buildListItems(items, highlight = false) {
  return items
    .map(
      (text) =>
        `<li style="font-family:Calibri,sans-serif;font-size:11pt;color:black;margin:0"><span${highlight ? ' style="background-color:yellow"' : ''}>${text}</span></li>`
    )
    .join('');
}

/**
 * Builds the informational table (Notes, Updates, Action items).
 */
function buildInfoTable() {
  return `
    <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-top:6pt">
      <tbody>
        <tr>
          <td rowspan="2" style="${cellStyle(COLORS.actionOrange)}">
            <p style="margin:0"><span style="font-family:'Times New Roman',serif;font-size:10pt;color:black"><b>Action for next shift members</b></span></p>
          </td>
          <td style="${cellStyle('transparent', 'width:30px')}"><br></td>
          <td style="${cellStyle('transparent', 'width:30px')}"><br></td>
        </tr>
        <tr>
          <td style="${cellStyle('transparent', 'width:30px')}"><br></td>
          <td style="${cellStyle('transparent', 'width:30px')}"><br></td>
        </tr>
        <tr>
          <td colspan="3" style="${cellStyle(COLORS.notesYellow)}">
            <p style="margin:0"><span style="font-family:Calibri,sans-serif;font-size:11pt;color:black">NOTE:-</span></p>
            <ul style="margin:4pt 0;padding-left:20pt">${buildListItems(STATIC_NOTES)}</ul>
            <p style="margin:0"><span style="font-family:Calibri,sans-serif;font-size:11pt;color:black">${STATIC_NOTES_CONTINUATION}</span></p>
            <ul style="margin:4pt 0;padding-left:20pt">${buildListItems(STATIC_NOTES_EXTRA)}</ul>
            <p style="margin:8pt 0 4pt"><span style="font-family:Calibri,sans-serif;font-size:11pt;color:black">UPDATES</span></p>
            <ul style="margin:4pt 0;padding-left:20pt">${buildListItems(STATIC_UPDATES)}</ul>
            <p style="margin:4pt 0"><span style="font-family:Calibri,sans-serif;font-size:11pt;color:black">${STATIC_UPDATES_CONTINUATION}</span></p>
            <p style="margin:4pt 0"><span style="background-color:yellow">&nbsp;</span></p>
            <ul style="margin:4pt 0;padding-left:20pt">${buildListItems(STATIC_UPDATES_HIGHLIGHT, true)}</ul>
          </td>
        </tr>
      </tbody>
    </table>`;
}

/**
 * Assembles the complete handover email HTML.
 */
function generateEmailHtml(handoverShift, userName, counts, jurisdictions) {
  const linkBlock = `
    <p style="margin:0 0 12pt">
      <a href="${DAILY_REPORT_LINK}" target="_blank" style="color:blue;font-family:Calibri,sans-serif;font-size:12pt;background-color:rgb(243,242,241)">
        <b><u>&nbsp;DUKE_SHIFT HANDOVER.xlsx</u></b>
      </a>
    </p>`;

  const greeting = `
    <p style="margin:0 0 4pt;font-family:Calibri,sans-serif;color:black">Hi Team,</p>
    <p style="margin:0 0 12pt;font-family:Calibri,sans-serif;color:black">Please find below the handover for Shift ${handoverShift}</p>`;

  const mainTable = buildMainTable(counts, jurisdictions);
  const infoTable = buildInfoTable();

  const signature = `
    <p style="margin:16pt 0 4pt;font-family:Calibri,sans-serif;color:black">Thanks and Regards,</p>
    <p style="margin:0;font-family:Calibri,sans-serif;color:black"><b>${escapeHtml(userName)}</b></p>`;

  return `
    <div style="font-family:Calibri,Arial,sans-serif;font-size:12pt;color:black">
      ${linkBlock}
      ${greeting}
      ${mainTable}
      ${infoTable}
      ${signature}
    </div>`;
}

/** Escapes HTML special characters in user-provided text. */
function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* ============================================
   Clipboard
   ============================================ */

/**
 * Strips HTML tags for plain-text clipboard fallback.
 */
function htmlToPlainText(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return doc.body.textContent || '';
}

/**
 * Selects a DOM node and copies its contents via execCommand.
 * Works on file:// URLs where the async Clipboard API is blocked.
 */
function copyNodeViaExecCommand(node) {
  const selection = window.getSelection();
  const range = document.createRange();
  range.selectNodeContents(node);
  selection.removeAllRanges();
  selection.addRange(range);
  const success = document.execCommand('copy');
  selection.removeAllRanges();
  return success;
}

/**
 * Copies rich HTML to the clipboard so Gmail preserves formatting.
 */
async function copyRichHtml(html) {
  // Preferred path — works on https:// and localhost
  if (navigator.clipboard && window.ClipboardItem && window.isSecureContext) {
    try {
      const htmlBlob = new Blob([html], { type: 'text/html' });
      const textBlob = new Blob([htmlToPlainText(html)], { type: 'text/plain' });
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/html': htmlBlob,
          'text/plain': textBlob,
        }),
      ]);
      return;
    } catch {
      // Fall through to execCommand below
    }
  }

  // Fallback — reliable when opening index.html directly (file://)
  const container = document.createElement('div');
  container.innerHTML = html;
  container.style.position = 'fixed';
  container.style.left = '-9999px';
  container.setAttribute('contenteditable', 'true');
  document.body.appendChild(container);
  container.focus();

  const success = copyNodeViaExecCommand(container);
  document.body.removeChild(container);

  if (!success) throw new Error('Copy command failed');
}

/* ============================================
   UI Helpers
   ============================================ */

function showStatus(message, type = '') {
  els.statusMessage.textContent = message;
  els.statusMessage.className = `status-message${type ? ` ${type}` : ''}`;
}

function clearStatus() {
  showStatus('');
}

/* ============================================
   Event Handlers
   ============================================ */

function handleGenerate() {
  clearStatus();

  const currentShift = els.currentShift.value;
  const userName = els.userName.value.trim() || 'Adarsh';
  const handoverShift = getHandoverShift(currentShift);
  const rawText = getPasteAreaText(els.incidentData);

  const { counts, jurisdictions, incidents, warnings } = parseIncidents(rawText);
  generatedEmailHtml = generateEmailHtml(handoverShift, userName, counts, jurisdictions);

  els.emailPreview.innerHTML = generatedEmailHtml;
  els.btnCopy.disabled = false;

  const incidentSummary = incidents.length
    ? `Parsed ${incidents.length} incident(s). Handover for Shift ${handoverShift} generated.`
    : `Handover for Shift ${handoverShift} generated (no incidents parsed).`;

  if (warnings.length) {
    showStatus(`${incidentSummary} ${warnings.join(' ')}`, warnings.length && incidents.length === 0 ? 'error' : 'success');
  } else {
    showStatus(incidentSummary, 'success');
  }
}

async function handleCopy() {
  if (!generatedEmailHtml) {
    showStatus('Generate the handover first before copying.', 'error');
    return;
  }

  try {
    await copyRichHtml(generatedEmailHtml);
    showStatus('Email copied! Paste directly into Gmail to preserve formatting.', 'success');
  } catch (err) {
    showStatus(`Copy failed: ${err.message}. Try using Chrome or Edge.`, 'error');
  }
}

/* ============================================
   Initialise
   ============================================ */

function init() {
  els.btnGenerate.addEventListener('click', handleGenerate);
  els.btnCopy.addEventListener('click', handleCopy);

  // Prevent rich formatting inside the paste area — keep plain TSV text
  els.incidentData.addEventListener('paste', (e) => {
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    document.execCommand('insertText', false, text);
  });
}

init();
