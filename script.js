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
 * Builds the main Shift Handover Report table HTML.
 * @param {boolean} forOutlook - Adds Outlook-friendly spacing when true (Duke Side).
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
 * Spacer between tables — Outlook ignores margins but respects line-height blocks.
 */
function buildOutlookTableSpacer() {
  return `
    <!--[if mso]><br style="mso-special-character:line-break" /><br style="mso-special-character:line-break" /><![endif]-->
    <div style="line-height:16pt;font-size:16pt;mso-line-height-rule:exactly;">&nbsp;</div>`;
}

/**
 * Builds the informational table (Notes, Updates, Action items).
 * @param {boolean} forOutlook - Adjusts top spacing for Outlook (Duke Side).
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
 * Assembles the complete handover email HTML.
 * @param {string} side - SIDE.CTS (Gmail) or SIDE.DUKE (Outlook).
 */
function generateEmailHtml(handoverShift, userName, counts, jurisdictions, side = SIDE.CTS) {
  const isDuke = side === SIDE.DUKE;

  const linkBlock = isDuke
    ? ''
    : `
    <p style="margin:0 0 12pt">
      <a href="${DAILY_REPORT_LINK}" target="_blank" style="color:blue;font-family:Calibri,sans-serif;font-size:12pt;background-color:rgb(243,242,241)">
        <b><u>&nbsp;DUKE_SHIFT HANDOVER.xlsx</u></b>
      </a>
    </p>`;

  const greeting = isDuke
    ? `
    <p style="margin:0 0 4pt;font-family:Calibri,sans-serif;font-size:11pt;color:black">Hi Sunil / Mohit,</p>
    <p style="margin:0 0 12pt;font-family:Calibri,sans-serif;font-size:11pt;color:black">Please find the below DCC Daily Report.</p>`
    : `
    <p style="margin:0 0 4pt;font-family:Calibri,sans-serif;color:black">Hi Team,</p>
    <p style="margin:0 0 12pt;font-family:Calibri,sans-serif;color:black">Please find below the handover for Shift ${handoverShift}</p>`;

  const mainTable = buildMainTable(counts, jurisdictions, isDuke);
  const tableSpacer = isDuke ? buildOutlookTableSpacer() : '';
  const infoTable = buildInfoTable(isDuke);

  const signature = `
    <p style="margin:16pt 0 4pt;font-family:Calibri,sans-serif;color:black">Thanks and Regards,</p>
    <p style="margin:0;font-family:Calibri,sans-serif;color:black"><b>${escapeHtml(userName)}</b></p>`;

  const wrapperFont = isDuke
    ? 'font-family:Calibri,Aptos,sans-serif;font-size:11pt;color:black'
    : 'font-family:Calibri,Arial,sans-serif;font-size:12pt;color:black';

  return `
    <div style="${wrapperFont}">
      ${linkBlock}
      ${greeting}
      ${mainTable}
      ${tableSpacer}
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
