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

const SIDE = {
  CTS: 'cts',
  DUKE: 'duke',
};

/* ============================================
   DOM References
   ============================================ */

const els = {
  currentShift: document.getElementById('current-shift'),
  userName: document.getElementById('user-name'),
  incidentData: document.getElementById('incident-data'),
  sideToggle: document.querySelector('.side-toggle'),
  btnGenerate: document.getElementById('btn-generate'),
  btnCopy: document.getElementById('btn-copy'),
  emailPreview: document.getElementById('email-preview'),
  statusMessage: document.getElementById('status-message'),
};

/** Holds the last generated email HTML for clipboard copy */
let generatedEmailHtml = '';
/** Currently selected email side — default CTS (Cognizant / Gmail) */
let selectedSide = SIDE.CTS;

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
 * Clean raw pasted text (e.g. stripping outer Excel quotes and unescaping double quotes).
 */
function cleanRawInput(text) {
  if (!text) return '';
  let cleaned = text.trim();
  
  // Strip outer quotes if the entire text block was wrapped in quotes by Excel
  if ((cleaned.startsWith('"') && cleaned.endsWith('"')) || (cleaned.startsWith("'") && cleaned.endsWith("'"))) {
    cleaned = cleaned.slice(1, -1).trim();
  }
  // Unescape doubled double-quotes from Excel exports
  cleaned = cleaned.replace(/""/g, '"');
  return cleaned;
}

/**
 * Segregates raw pasted text into individual incident record chunks.
 * Handles single-cell pastes containing multiple incidents, tab-separated lines,
 * or space-separated sequences.
 */
function extractRecordChunks(rawText) {
  const text = cleanRawInput(rawText);
  if (!text) return [];

  // Check for ServiceNow incident numbers (e.g. INC5305979, INC1234567, INC...)
  const incRegex = /\bINC\d+/gi;
  const incMatches = [...text.matchAll(incRegex)];

  if (incMatches.length > 1) {
    // Multiple INC numbers found in the text — split at each INC boundary
    const chunks = [];
    for (let i = 0; i < incMatches.length; i++) {
      const startIdx = i === 0 ? 0 : incMatches[i].index;
      const endIdx = i < incMatches.length - 1 ? incMatches[i + 1].index : text.length;
      const chunkText = text.substring(startIdx, endIdx).trim();
      if (chunkText) {
        chunks.push(chunkText);
      }
    }
    return chunks;
  }

  // Split by line breaks if present
  const rawLines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);

  const chunks = [];
  for (const line of rawLines) {
    const lineIncMatches = [...line.matchAll(/\bINC\d+/gi)];
    if (lineIncMatches.length > 1) {
      for (let i = 0; i < lineIncMatches.length; i++) {
        const startIdx = i === 0 ? 0 : lineIncMatches[i].index;
        const endIdx = i < lineIncMatches.length - 1 ? lineIncMatches[i + 1].index : line.length;
        const chunkText = line.substring(startIdx, endIdx).trim();
        if (chunkText) chunks.push(chunkText);
      }
    } else {
      chunks.push(line);
    }
  }

  // Fallback: If only 1 chunk exists with no INC numbers, check for multiple Priority keywords
  if (chunks.length === 1) {
    const singleChunk = chunks[0];
    const priorityRegex = /\b(Critical|High|Medium|Low)\b/gi;
    const prioMatches = [...singleChunk.matchAll(priorityRegex)];
    if (prioMatches.length > 1) {
      const prioChunks = [];
      for (let i = 0; i < prioMatches.length; i++) {
        const startIdx = i === 0 ? 0 : prioMatches[i].index;
        const endIdx = i < prioMatches.length - 1 ? prioMatches[i + 1].index : singleChunk.length;
        const chunkText = singleChunk.substring(startIdx, endIdx).trim();
        if (chunkText) prioChunks.push(chunkText);
      }
      return prioChunks;
    }
  }

  return chunks;
}

/** Normalises a priority string to one of the four supported values. */
function normalisePriority(textOrCell) {
  const str = (textOrCell || '').trim();
  if (!str) return '';

  if (/\b(1\s*-\s*)?critical\b/i.test(str) || /\bP1\b/i.test(str)) return 'Critical';
  if (/\b(2\s*-\s*)?high\b/i.test(str) || /\bP2\b/i.test(str)) return 'High';
  if (/\b(3\s*-\s*)?(medium|moderate)\b/i.test(str) || /\bP3\b/i.test(str)) return 'Medium';
  if (/\b(4\s*-\s*)?low\b/i.test(str) || /\bP4\b/i.test(str)) return 'Low';

  if (/^1$/i.test(str)) return 'Critical';
  if (/^2$/i.test(str)) return 'High';
  if (/^3$/i.test(str)) return 'Medium';
  if (/^4$/i.test(str)) return 'Low';

  for (const p of PRIORITIES) {
    if (new RegExp(`\\b${p}\\b`, 'i').test(str)) {
      return p;
    }
  }
  return '';
}

/** Extracts a jurisdiction code from a cell value or text chunk. */
function extractJurisdiction(textOrCell) {
  const str = (textOrCell || '').trim();
  if (!str) return '';

  for (const j of JURISDICTIONS) {
    if (new RegExp(`\\b${j}\\b`, 'i').test(str)) {
      return j;
    }
  }
  return '';
}

/**
 * Classifies an application/system cell or chunk as OMS, DMS, or Others.
 */
function classifyApplication(textOrCell) {
  const upper = (textOrCell || '').toUpperCase();
  if (upper.includes('OUTAGE MANAGEMENT SYSTEM') || /\bOMS\b/.test(upper) || upper.includes('OMS_') || upper.includes('OMS-')) {
    return 'oms';
  }
  if (
    upper.includes('DISTRIBUTION MANAGEMENT SYSTEM') ||
    /\bDMS\b/.test(upper) ||
    upper.includes('DMS_') ||
    upper.includes('DMS-') ||
    upper.includes('DSCADA')
  ) {
    return 'dms';
  }
  return 'others';
}

/**
 * Parses a single incident chunk to extract its Priority, Jurisdiction, and Application class.
 */
function parseIncidentChunk(chunk) {
  const cells = chunk.split('\t').map((c) => c.trim()).filter(Boolean);

  let priority = '';
  let jurisdiction = '';
  let appClass = '';
  let rawApp = '';

  // Try cell-by-cell matching if tab cells exist
  if (cells.length > 1) {
    for (const cell of cells) {
      const p = normalisePriority(cell);
      if (p) {
        priority = p;
        break;
      }
    }

    for (const cell of cells) {
      const j = extractJurisdiction(cell);
      if (j) {
        jurisdiction = j;
        break;
      }
    }

    for (const cell of cells) {
      const cls = classifyApplication(cell);
      if (cls === 'oms' || cls === 'dms') {
        appClass = cls;
        rawApp = cell;
        break;
      }
    }

    if (!appClass) {
      for (const cell of cells) {
        if (
          !normalisePriority(cell) &&
          !extractJurisdiction(cell) &&
          !/^\bINC\d+/i.test(cell) &&
          !/^(new|on hold|in progress|resolved|closed|shift\s*\d+)$/i.test(cell)
        ) {
          rawApp = cell;
          appClass = classifyApplication(cell);
          break;
        }
      }
    }
  }

  // Fallback to searching full chunk text if cell-based matching didn't resolve priority or jurisdiction
  if (!priority) {
    priority = normalisePriority(chunk);
  }

  if (!jurisdiction) {
    jurisdiction = extractJurisdiction(chunk);
  }

  if (!appClass) {
    appClass = classifyApplication(rawApp || chunk);
  }

  return { priority, appClass, jurisdiction, raw: chunk };
}

/**
 * Parses all incident rows/chunks and returns aggregated counts.
 */
function parseIncidents(rawText) {
  const chunks = extractRecordChunks(rawText);

  const incidents = [];
  const counts = createEmptyCounts();
  const jurisdictions = createEmptyJurisdictions();
  const warnings = [];

  if (!chunks.length) {
    warnings.push('No incident data found. Tables will show zero counts.');
    return { incidents, counts, jurisdictions, warnings };
  }

  for (const chunk of chunks) {
    const incident = parseIncidentChunk(chunk);
    if (!incident.priority) continue;

    incidents.push(incident);

    counts[incident.priority][incident.appClass]++;
    counts[incident.priority].transferred =
      counts[incident.priority].oms + counts[incident.priority].dms + counts[incident.priority].others;

    if (incident.jurisdiction) {
      jurisdictions[incident.jurisdiction]++;
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
 * Builds the main Shift Handover Report table HTML (CTS / Gmail side).
 */
function buildMainTable(counts, jurisdictions, forOutlook = false) {
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
  const tableMargin = forOutlook ? 'margin-bottom:0' : 'margin-bottom:12pt';

  return `
    <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;max-width:730pt;${tableMargin}">
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

/** Builds list items as HTML <li> elements (CTS side). */
function buildListItems(items, highlight = false) {
  return items
    .map(
      (text) =>
        `<li style="font-family:Calibri,sans-serif;font-size:11pt;color:black;margin:0"><span${highlight ? ' style="background-color:yellow"' : ''}>${text}</span></li>`
    )
    .join('');
}

/**
 * Builds the informational table (Notes, Updates, Action items) (CTS side).
 */
function buildInfoTable(forOutlook = false) {
  const topMargin = forOutlook ? 'margin-top:0' : 'margin-top:6pt';
  return `
    <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;${topMargin}">
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
 * Builds the official Duke (Outlook) Shift Handover Report table HTML.
 */
function buildDukeMainTable(counts, jurisdictions) {
  const priorityRows = PRIORITIES.map((p, idx) => {
    const c = counts[p];
    const rowHeight = p === 'High' ? '1.35pt' : '7.85pt';
    const irow = idx + 2;
    const omsVal = c.oms > 0 ? c.oms : '&nbsp;';
    const dmsVal = c.dms > 0 ? c.dms : '&nbsp;';
    const othersVal = c.others > 0 ? c.others : '&nbsp;';
    const transferredVal = c.transferred > 0 ? c.transferred : '&nbsp;';
    return `
 <tr style='mso-yfti-irow:${irow};height:${rowHeight}'>
  <td width=138 valign=top style='width:103.5pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#F7CAAC;padding:0in 5.4pt 0in 5.4pt;height:${rowHeight};border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>${p}</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=129 colspan=2 valign=top style='width:96.75pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:${rowHeight};border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>${omsVal}</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=221 colspan=2 valign=top style='width:165.45pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:${rowHeight};border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>${dmsVal}</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=47 colspan=2 valign=top style='width:35.15pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:${rowHeight};border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>${othersVal}</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=142 colspan=2 valign=top style='width:106.4pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:${rowHeight};border-color:currentcolor windowtext windowtext currentcolor'>
  <p align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>&nbsp;</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=125 colspan=2 valign=top style='width:94.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:${rowHeight};border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>${transferredVal}</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in;height:${rowHeight}'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>`;
  }).join('');

  const t = counts.totals;

  return `<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=976
 style='width:731.85pt;margin-left:.4pt;border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:0in 0in 0in 0in'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:20.9pt'>
  <td width=173 style='width:1.8in;border-top:solid windowtext 1.0pt;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:none;background:#C65911;padding:0in 0in 0in 0in;height:20.9pt;border-color:currentcolor'></td>
  <td width=802 colspan=11 style='width:601.25pt;background:#C65911;padding:0in 5.4pt 0in 5.4pt;height:20.9pt;border-width:initial;border-style:initial;border-color:currentcolor'>
  <p><b><span style='font-size:16.0pt;font-family:"Times New Roman",serif;color:black'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Shift Handover Report</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in;height:20.9pt'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;height:8.8pt'>
  <td width=173 rowspan=7 style='width:1.8in;border:solid windowtext 1.0pt;border-top:none;background:#FCE4D6;padding:0in 5.4pt 0in 5.4pt;height:8.8pt;border-color:currentcolor windowtext windowtext;border-image:none'>
  <p class=MsoNormal><b><span style='font-size:10.0pt;font-family:"Calibri",sans-serif;color:black'>SNOW Incidents Created</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=138 valign=top style='width:103.5pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#F7CAAC;padding:0in 5.4pt 0in 5.4pt;height:8.8pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Priority</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=129 colspan=2 valign=top style='width:96.75pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#F7CAAC;padding:0in 0in 0in 0in;height:8.8pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>No of tickets created for OMS</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=221 colspan=2 valign=top style='width:165.45pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#F7CAAC;padding:0in 0in 0in 0in;height:8.8pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>No of tickets created for DMS</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=47 colspan=2 valign=top style='width:35.15pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#F7CAAC;padding:0in 0in 0in 0in;height:8.8pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>No of tickets created for Others</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=142 colspan=2 valign=top style='width:106.4pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#F7CAAC;padding:0in 0in 0in 0in;height:8.8pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Tickets resolved by L1</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=125 colspan=2 valign=top style='width:94.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#F7CAAC;padding:0in 0in 0in 0in;height:8.8pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Tickets Transferred to L2</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in;height:8.8pt'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
${priorityRows}
 <tr style='mso-yfti-irow:6;height:7.85pt'>
  <td width=138 valign=top style='width:103.5pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#DBDBDB;padding:0in 5.4pt 0in 5.4pt;height:7.85pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Total Ticket Counts</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=129 colspan=2 valign=top style='width:96.75pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:7.85pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>${t.oms > 0 ? t.oms : '&nbsp;'}</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=221 colspan=2 valign=top style='width:165.45pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:7.85pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>${t.dms > 0 ? t.dms : '&nbsp;'}</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=47 colspan=2 valign=top style='width:35.15pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:7.85pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>${t.others > 0 ? t.others : '&nbsp;'}</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=142 colspan=2 valign=top style='width:106.4pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:7.85pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>&nbsp;</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=125 colspan=2 valign=top style='width:94.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:7.85pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>${t.transferred > 0 ? t.transferred : '&nbsp;'}</span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in;height:7.85pt'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:7;height:21.05pt'>
  <td width=336 colspan=4 style='width:252.25pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 5.4pt 0in 5.4pt;height:21.05pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Number of UAC INC received :&nbsp;</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=465 colspan=7 style='width:349.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:21.05pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Number of UAC INC Resolved :</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in;height:21.05pt'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:8;height:9.9pt'>
  <td width=173 rowspan=2 style='width:1.8in;border:solid windowtext 1.0pt;border-top:none;background:#FCE4D6;padding:0in 5.4pt 0in 5.4pt;height:9.9pt;border-color:currentcolor windowtext windowtext;border-image:none'>
  <p class=MsoNormal><b><span style='font-size:10.0pt;font-family:"Calibri",sans-serif;color:black'>Jurisdiction </span></b><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=336 colspan=4 style='width:252.25pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 5.4pt 0in 5.4pt;height:9.9pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>DEF:&nbsp; ${jurisdictions.DEF > 0 ? jurisdictions.DEF : ''}</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=465 colspan=7 style='width:349.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:9.9pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>&nbsp; DEP:&nbsp; ${jurisdictions.DEP > 0 ? jurisdictions.DEP : ''}</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in;height:9.9pt'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:9;height:17.0pt'>
  <td width=336 colspan=4 style='width:252.25pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 5.4pt 0in 5.4pt;height:17.0pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>DEM: ${jurisdictions.DEM > 0 ? jurisdictions.DEM : ''}</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=465 colspan=7 style='width:349.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:17.0pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>&nbsp; <b>DEC: &nbsp;${jurisdictions.DEC > 0 ? jurisdictions.DEC : ''}</b></span><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in;height:17.0pt'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:10;height:22.85pt'>
  <td width=173 style='width:1.8in;border-top:none;border-left:solid windowtext 1.0pt;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.5pt;background:#FBE4D5;padding:0in 5.4pt 0in 5.4pt;height:22.85pt;border-color:currentcolor windowtext windowtext'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Model Push Activity</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=336 colspan=4 style='width:252.25pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.5pt;border-right:solid windowtext 1.5pt;background:#FCE4D6;padding:0in 5.4pt 0in 5.4pt;height:22.85pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Go/No-Go and Issue reported by DCC</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=465 colspan=7 style='width:349.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.5pt;border-right:solid windowtext 1.5pt;background:#FCE4D6;padding:0in 0in 0in 0in;height:22.85pt;border-color:currentcolor windowtext windowtext currentcolor'>
  <p><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>&nbsp; Go and No issue reported</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in;height:22.85pt'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:11'>
  <td width=173 style='width:1.8in;padding:0in 0in 0in 0in'></td>
  <td width=200 colspan=2 style='width:150.05pt;padding:0in 0in 0in 0in'></td>
  <td width=136 colspan=2 style='width:102.2pt;padding:0in 0in 0in 0in'></td>
  <td width=151 style='width:113.45pt;padding:0in 0in 0in 0in'></td>
  <td width=314 colspan=6 style='width:235.55pt;padding:0in 0in 0in 0in'></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:12'>
  <td width=173 style='width:1.8in;padding:0in 0in 0in 0in'></td>
  <td width=138 style='width:103.5pt;padding:0in 0in 0in 0in'></td>
  <td width=62 style='width:46.55pt;padding:0in 0in 0in 0in'></td>
  <td width=67 style='width:50.2pt;padding:0in 0in 0in 0in'></td>
  <td width=69 style='width:52.0pt;padding:0in 0in 0in 0in'></td>
  <td width=151 style='width:113.45pt;padding:0in 0in 0in 0in'></td>
  <td width=1 style='width:1.0pt;padding:0in 0in 0in 0in'></td>
  <td width=47 colspan=2 style='width:35.15pt;padding:0in 0in 0in 0in'></td>
  <td width=142 colspan=2 style='width:106.4pt;padding:0in 0in 0in 0in'></td>
  <td width=125 colspan=2 style='width:94.0pt;padding:0in 0in 0in 0in'></td>
 </tr>
 <tr style='mso-yfti-irow:13'>
  <td width=173 style='width:129.75pt;padding:0in 0in 0in 0in'></td>
  <td width=138 style='width:103.5pt;padding:0in 0in 0in 0in'></td>
  <td width=62 style='width:46.5pt;padding:0in 0in 0in 0in'></td>
  <td width=67 style='width:50.25pt;padding:0in 0in 0in 0in'></td>
  <td width=69 style='width:51.75pt;padding:0in 0in 0in 0in'></td>
  <td width=151 style='width:113.25pt;padding:0in 0in 0in 0in'></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'></td>
  <td width=47 style='width:35.25pt;padding:0in 0in 0in 0in'></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'></td>
  <td width=140 style='width:105.0pt;padding:0in 0in 0in 0in'></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'></td>
  <td width=124 style='width:93.0pt;padding:0in 0in 0in 0in'></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'></td>
 </tr>
 <tr style='mso-yfti-irow:14'>
  <td width=173 style='width:129.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=138 style='width:103.5pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=62 style='width:46.5pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=67 style='width:50.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=69 style='width:51.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=151 style='width:113.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=47 style='width:35.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=140 style='width:105.0pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=124 style='width:93.0pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
 </tr>
 <tr style='mso-yfti-irow:15'>
  <td width=173 style='width:129.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=138 style='width:103.5pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=62 style='width:46.5pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=67 style='width:50.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=69 style='width:51.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=151 style='width:113.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=47 style='width:35.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=140 style='width:105.0pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=124 style='width:93.0pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
 </tr>
 <tr style='mso-yfti-irow:16'>
  <td width=173 style='width:129.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=138 style='width:103.5pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=62 style='width:46.5pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=67 style='width:50.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=69 style='width:51.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=151 style='width:113.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=47 style='width:35.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=140 style='width:105.0pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=124 style='width:93.0pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
 </tr>
 <tr style='mso-yfti-irow:17;mso-yfti-lastrow:yes'>
  <td width=173 style='width:129.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=138 style='width:103.5pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=62 style='width:46.5pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=67 style='width:50.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=69 style='width:51.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=151 style='width:113.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=47 style='width:35.25pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=140 style='width:105.0pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=124 style='width:93.0pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
  <td width=1 style='width:.75pt;padding:0in 0in 0in 0in'><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p></td>
 </tr>
</table>`;
}

/**
 * Builds the official Duke (Outlook) Info Table HTML.
 */
function buildDukeInfoTable() {
  return `<p class=MsoNormal><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-ansi-language:EN-IN'>&nbsp;</span><span style='font-size:11.0pt;font-family:Calibri,sans-serif'><o:p></o:p></span></p>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left
 width=1189 style='width:891.65pt;border-collapse:collapse;mso-yfti-tbllook:1184;mso-table-lspace:9.0pt;margin-left:6.75pt;mso-table-rspace:9.0pt;margin-right:6.75pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0in 0in 0in 0in'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:7.4pt'>
  <td width=965 rowspan=2 style='width:724.0pt;border:solid windowtext 1.0pt;border-bottom:solid windowtext 1.5pt;background:#F4B083;padding:0in 5.4pt 0in 5.4pt;height:7.4pt;border-image:none'>
  <p style='mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><b><span style='font-size:10.0pt;font-family:"Times New Roman",serif;color:black'>Action for next shift members</span></b><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p>
  </td>
  <td width=224 style='width:167.65pt;padding:0in 0in 0in 0in;height:7.4pt'></td>
 </tr>
 <tr style='mso-yfti-irow:1;height:4.9pt'>
  <td width=224 style='width:167.65pt;padding:0in 0in 0in 0in;height:4.9pt'></td>
 </tr>
 <tr style='mso-yfti-irow:2;mso-yfti-lastrow:yes;height:57.5pt'>
  <td width=965 style='width:724.0pt;border-top:none;border-left:solid windowtext 1.5pt;border-bottom:none;border-right:solid windowtext 1.5pt;background:#FFF2CC;padding:0in 5.4pt 0in 5.4pt;height:57.5pt;border-color:currentcolor windowtext;border-image:none'>
  <p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>NOTE:- </span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  <ul type=disc style='margin:4pt 0;padding-left:24pt'>
   <li class=MsoListParagraph style='mso-list:l2 level1 lfo1;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>Kindly Go through Helpdesk scenario and ALL DCC App List for Reporting Issue File which is send by Wayne in KT Mail and follow while creating any tickets.</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
   <li class=MsoListParagraph style='margin-top:0in;margin-bottom:0in;mso-list:l2 level1 lfo1;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span class=ui-provider><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>kindly confirm with user for Maps issue is ADMS Maps and Outages maps mostly ticket will go to Modeling team but confirm first with Wayne also once</span></span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
   <li class=MsoListParagraph style='margin-top:0in;margin-bottom:0in;mso-list:l2 level1 lfo1;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span class=ui-provider><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>If you received mail with subject  PROD: Calls not processed  which is being forwarded by Beena mail id, Please inform Beena,Kartik, Raghu or Duke_OMS_FL DL immediately</span></span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>.</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
  </ul>
  <ul style='margin-top:0in;margin-bottom:0in;padding-left:24pt' type=disc>
   <li class=MsoListParagraph style='margin-top:0in;margin-bottom:0in;mso-list:l1 level1 lfo2;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>Check new SOP of Model Push [Note :- After Model Push Ask DCC control room about everything up &amp; running or not.</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
  </ul>
  <p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (Customer search, Device search &amp; Geographic Map)]</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  <ul type=disc style='margin:4pt 0;padding-left:24pt'>
   <li class=MsoListParagraph style='mso-list:l0 level1 lfo3;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>One minor change in UAC jobs check the changes as updated in group</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
  </ul>
  <p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>&nbsp;</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  <p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>UPDATES</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  <ul type=disc style='margin:4pt 0;padding-left:24pt'>
   <li class=MsoListParagraph style='mso-list:l4 level1 lfo4;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>Strom is occur in DEF so be <b>PROACTIVE</b> also you guy s added in STORM 01/09/2024 Group so response in that Quickly. Also, PING that group in your chat.</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
   <li class=MsoListParagraph style='mso-list:l4 level1 lfo4;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>If you guys face any issue or need any kind of help regarding issues in your shift, then connect with Akansha/ Karthik or other SME S</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
   <li class=MsoListParagraph style='mso-list:l4 level1 lfo4;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>Please Notice or find out UAC job pattern and update to other shift members. (UAC job failed in 1 hr or 1hr 15 min etc like this )</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
   <li class=MsoListParagraph style='mso-list:l4 level1 lfo4;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>Please Follow same pattern (All team members) Put UAC tickets on Hold first then Resolve it after Iteration run successfully. </span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
   <li class=MsoListParagraph style='mso-list:l4 level1 lfo4;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black'>Also For Modelling issue sent ticket to representative group like modelling DEM, DEC ETC. </span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
  </ul>
  <ul type=disc style='margin:4pt 0;padding-left:24pt'>
   <li class=MsoListParagraph style='mso-list:l3 level1 lfo5;tab-stops:list .5in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span class=ui-provider><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman";color:black;background:yellow'>UAC Update (Param shared in UAC group). We are added into DL in mail where we will receive job failure alert for  OMS_ADMS_STORED_PROCE_DEF  we have to monitor both Mail and Resolved incident kindly check UAC group</span></span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-fareast-font-family:"Times New Roman"'><o:p></o:p></span></li>
  </ul>
  <p class=MsoListParagraph style='margin-left:.25in;mso-element:frame;mso-element-frame-hspace:9.0pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'><span style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'>Please read Internal help desk DEF group message which is send by Beena</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
  </td>
  <td width=224 style='width:167.65pt;padding:0in 0in 0in 0in;height:57.5pt'></td>
 </tr>
</table>`;
}

/**
 * Assembles the complete handover email HTML.
 * @param {string} side - SIDE.CTS (Gmail) or SIDE.DUKE (Outlook).
 */
function generateEmailHtml(handoverShift, userName, counts, jurisdictions, side = SIDE.CTS) {
  if (side === SIDE.DUKE) {
    const mainTable = buildDukeMainTable(counts, jurisdictions);
    const infoTable = buildDukeInfoTable();
    const signature = `<p class=MsoNormal><span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-ansi-language:EN-GB'>&nbsp;</span><span style='font-size:11.0pt;font-family:Calibri,sans-serif'><o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-ansi-language:EN-GB'>&nbsp;</span><span style='font-size:11.0pt;font-family:Calibri,sans-serif'><o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-ansi-language:EN-GB'>&nbsp;</span><span style='font-size:11.0pt;font-family:Calibri,sans-serif'><o:p></o:p></span></p>
<p class=MsoNormal><strong><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Thanks and Regards,</span></strong><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-ansi-language:EN-GB'>${escapeHtml(userName)}</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>`;

    return `<div class=WordSection1>
<p style='margin-right:5.0pt'>H<span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black;mso-ansi-language:EN-GB'>i</span><span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-ansi-language:EN-GB'> Sunil/Mohit,</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
<div style='border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in'>
<p class=MsoNormal><span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black;mso-ansi-language:EN-GB'>Please find below DCC Daily Report</span><span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;mso-ansi-language:EN-GB'>.</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
</div>
<p class=MsoNormal style='background:white'><span lang=EN-GB style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black;mso-ansi-language:EN-GB'>&nbsp;</span><span style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p></o:p></span></p>
${mainTable}
${infoTable}
${signature}
</div>`;
  }

  const linkBlock = `
    <p style="margin:0 0 12pt">
      <a href="${DAILY_REPORT_LINK}" target="_blank" style="color:blue;font-family:Calibri,sans-serif;font-size:12pt;background-color:rgb(243,242,241)">
        <b><u>&nbsp;DUKE_SHIFT HANDOVER.xlsx</u></b>
      </a>
    </p>`;

  const greeting = `
    <p style="margin:0 0 4pt;font-family:Calibri,sans-serif;color:black">Hi Team,</p>
    <p style="margin:0 0 12pt;font-family:Calibri,sans-serif;color:black">Please find below the handover for Shift ${handoverShift}</p>`;

  const mainTable = buildMainTable(counts, jurisdictions, false);
  const infoTable = buildInfoTable(false);

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
  generatedEmailHtml = generateEmailHtml(handoverShift, userName, counts, jurisdictions, selectedSide);

  els.emailPreview.innerHTML = generatedEmailHtml;
  els.btnCopy.disabled = false;

  const sideLabel = selectedSide === SIDE.DUKE ? 'Duke (Outlook)' : 'CTS (Gmail)';
  const incidentSummary = incidents.length
    ? `Parsed ${incidents.length} incident(s). ${sideLabel} handover for Shift ${handoverShift} generated.`
    : `${sideLabel} handover for Shift ${handoverShift} generated (no incidents parsed).`;

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
    const pasteTarget = selectedSide === SIDE.DUKE ? 'Outlook' : 'Gmail';
    showStatus(`Email copied! Paste directly into ${pasteTarget} to preserve formatting.`, 'success');
  } catch (err) {
    showStatus(`Copy failed: ${err.message}. Try using Chrome or Edge.`, 'error');
  }
}

function handleSideToggle(event) {
  const btn = event.target.closest('.side-toggle-btn');
  if (!btn) return;

  selectedSide = btn.dataset.side === SIDE.DUKE ? SIDE.DUKE : SIDE.CTS;

  els.sideToggle.querySelectorAll('.side-toggle-btn').forEach((b) => {
    const isActive = b === btn;
    b.classList.toggle('active', isActive);
    b.setAttribute('aria-pressed', String(isActive));
  });
}

/* ============================================
   Initialise
   ============================================ */

function init() {
  els.btnGenerate.addEventListener('click', handleGenerate);
  els.btnCopy.addEventListener('click', handleCopy);
  els.sideToggle.addEventListener('click', handleSideToggle);

  // Prevent rich formatting inside the paste area — keep plain TSV text
  els.incidentData.addEventListener('paste', (e) => {
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    document.execCommand('insertText', false, text);
  });
}

init();
