// Paste this entire file into Google Apps Script (script.google.com)
// Then deploy as a Web App (Execute as: Me, Who has access: Anyone)
// Copy the deployment URL into index.html where indicated.

const SHEET_NAME = 'Submissions';
const SPREADSHEET_ID = '1BKjoAogVRM2ezBZWiJuyyUlgAk9vv5Ylc9T89A2azvY';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── MAIN SUBMISSIONS SHEET ──
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Student #', 'Timestamp', 'First Name', 'Last Name', 'Email', 'House',
        'Legacy Score', 'Valor Score', 'Horizon Score',
        'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10'
      ]);
      sheet.getRange(1, 1, 1, 19).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Student # = number of data rows already present + 1
    const studentNum = Math.max(sheet.getLastRow() - 1, 0) + 1;

    sheet.appendRow([
      studentNum,
      data.timestamp,
      data.firstName,
      data.lastName,
      data.email,
      data.house,
      data.scores.Legacy,
      data.scores.Valor,
      data.scores.Horizon,
      ...data.answers
    ]);

    // ── HOUSE TABS ──
    const houses = ['Legacy', 'Valor', 'Horizon'];
    const houseHeaders = ['Student #', 'Timestamp', 'First Name', 'Last Name', 'Email',
                          'Legacy Score', 'Valor Score', 'Horizon Score',
                          'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10'];
    const houseColors  = { Legacy: '#6c2fa0', Valor: '#c0392b', Horizon: '#2471a3' };
    const tabColors    = { Legacy: '#d7b8f0', Valor: '#f0b8b8', Horizon: '#b8d8f0' };

    houses.forEach(h => {
      let tab = ss.getSheetByName(h);
      if (!tab) {
        tab = ss.insertSheet(h);
        tab.appendRow(houseHeaders);
        tab.getRange(1, 1, 1, houseHeaders.length).setFontWeight('bold')
           .setBackground(houseColors[h]).setFontColor('#ffffff');
        tab.setFrozenRows(1);
        tab.setTabColor(tabColors[h]);
      }

      if (data.house === h) {
        const houseStudentNum = Math.max(tab.getLastRow() - 1, 0) + 1;
        tab.appendRow([
          houseStudentNum,
          data.timestamp,
          data.firstName,
          data.lastName,
          data.email,
          data.scores.Legacy,
          data.scores.Valor,
          data.scores.Horizon,
          ...data.answers
        ]);
      }
    });

    sendResultEmail(data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function sendResultEmail(data) {
  if (!data.email) return;

  const houseDetails = {
    Legacy: {
      emoji: '📜',
      tagline: 'Rooted in wisdom. Guided by the past.',
      desc: 'You are a keeper of history and tradition. You find strength in stories and lessons that came before you — a true guardian of what has been.',
      headerBg: 'linear-gradient(135deg, #3b1060, #1a0535)',
      badgeBg:  'linear-gradient(135deg, #6c2fa0, #3b1060)',
      accent:   '#c084f5',
      border:   '#9b59b6',
      star:     '#d4a0ff'
    },
    Valor: {
      emoji: '⚔️',
      tagline: 'Bold in action. Present in every moment.',
      desc: 'You are driven by the here and now. You throw yourself into experiences, trust your instincts, and make things happen — a true champion of the present.',
      headerBg: 'linear-gradient(135deg, #7a0000, #3d0000)',
      badgeBg:  'linear-gradient(135deg, #c0392b, #7a0000)',
      accent:   '#ff8080',
      border:   '#e74c3c',
      star:     '#ffb3b3'
    },
    Horizon: {
      emoji: '🔭',
      tagline: 'Eyes forward. Mind on what\'s possible.',
      desc: 'You are a dreamer and a planner. You see potential where others see limits, always thinking about what could be — a true architect of the future.',
      headerBg: 'linear-gradient(135deg, #003d6b, #001830)',
      badgeBg:  'linear-gradient(135deg, #2471a3, #003d6b)',
      accent:   '#5dade2',
      border:   '#2980b9',
      star:     '#aed6f1'
    }
  };

  const h = houseDetails[data.house] || {
    emoji: '🏠', tagline: '', desc: '', headerBg: '#1a1a38',
    badgeBg: '#333', accent: '#aaa', border: '#555', star: '#ccc'
  };
  const name = [data.firstName, data.lastName].filter(Boolean).join(' ') || 'Student';
  const subject = `You have been sorted into House ${data.house}!`;

  const scoreRows = [
    { house: 'Legacy',  emoji: '📜', score: data.scores.Legacy  || 0, color: '#c084f5' },
    { house: 'Valor',   emoji: '⚔️', score: data.scores.Valor   || 0, color: '#ff8080' },
    { house: 'Horizon', emoji: '🔭', score: data.scores.Horizon || 0, color: '#5dade2' }
  ].sort((a, b) => b.score - a.score);

  const maxScore = Math.max(...scoreRows.map(r => r.score), 1);

  const scoreBarRows = scoreRows.map(r => {
    const pct = Math.round((r.score / maxScore) * 100);
    const isBold = r.house === data.house;
    return `
      <tr>
        <td style="padding:6px 10px 6px 0;font-size:0.9rem;color:${isBold ? '#f0ead6' : '#9090b8'};white-space:nowrap;font-weight:${isBold ? 'bold' : 'normal'};">
          ${r.emoji} ${r.house}
        </td>
        <td style="padding:6px 0;width:100%;">
          <div style="background:#1a1a38;border-radius:4px;height:12px;overflow:hidden;">
            <div style="background:${r.color};height:12px;width:${pct}%;border-radius:4px;"></div>
          </div>
        </td>
        <td style="padding:6px 0 6px 10px;font-size:0.9rem;color:${r.color};font-weight:bold;text-align:right;">${r.score}</td>
      </tr>`;
  }).join('');

  const stars = Array.from({length: 18}, (_, i) => {
    const top  = Math.floor(Math.random() * 90) + 2;
    const left = Math.floor(Math.random() * 96) + 2;
    const size = Math.random() < 0.3 ? 4 : 2;
    const op   = (Math.random() * 0.5 + 0.3).toFixed(2);
    return `<div style="position:absolute;top:${top}%;left:${left}%;width:${size}px;height:${size}px;background:${h.star};border-radius:50%;opacity:${op};"></div>`;
  }).join('');

  const body = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:20px 0;background:#080818;font-family:Georgia,'Times New Roman',serif;">

  <div style="max-width:560px;margin:0 auto;">

    <!-- HEADER BANNER -->
    <div style="background:${h.headerBg};border-radius:16px 16px 0 0;padding:48px 40px 40px;text-align:center;position:relative;overflow:hidden;">
      ${stars}
      <div style="position:relative;z-index:1;">
        <p style="margin:0 0 6px;font-size:0.75rem;letter-spacing:0.2em;text-transform:uppercase;color:${h.accent};opacity:0.8;">The Sorting Has Spoken</p>
        <h1 style="margin:0 0 20px;font-size:2.4rem;letter-spacing:0.1em;color:#f0ead6;text-shadow:0 0 30px ${h.accent};">The Sorting</h1>
        <div style="display:inline-block;background:${h.badgeBg};border:2px solid ${h.border};border-radius:120px 120px 80px 80px;padding:24px 40px 28px;box-shadow:0 0 40px ${h.border}66;">
          <div style="font-size:3.8rem;line-height:1;margin-bottom:10px;">${h.emoji}</div>
          <div style="font-size:1.9rem;font-weight:bold;color:#ffffff;letter-spacing:0.08em;">House ${data.house}</div>
          <div style="font-size:0.85rem;color:${h.accent};font-style:italic;margin-top:6px;">${h.tagline}</div>
        </div>
      </div>
    </div>

    <!-- BODY -->
    <div style="background:#0f0f24;border-left:1px solid ${h.border}44;border-right:1px solid ${h.border}44;padding:36px 40px;">

      <p style="margin:0 0 8px;font-size:1.1rem;color:#f0ead6;">Welcome to your house, <strong style="color:${h.accent};">${name}</strong>.</p>
      <p style="margin:0 0 28px;font-size:0.95rem;color:#9090b8;line-height:1.7;">${h.desc}</p>

      <!-- SCORE BREAKDOWN -->
      <div style="background:#0a0a1e;border:1px solid #222244;border-radius:12px;padding:20px 24px;margin-bottom:28px;">
        <p style="margin:0 0 14px;font-size:0.72rem;letter-spacing:0.14em;text-transform:uppercase;color:#555588;">Your Score Breakdown</p>
        <table style="width:100%;border-collapse:collapse;">${scoreBarRows}</table>
      </div>

      <!-- SPAM NOTE -->
      <div style="background:#0a0a1e;border:1px solid #222244;border-radius:8px;padding:14px 18px;margin-bottom:24px;text-align:center;">
        <p style="margin:0;font-size:0.78rem;color:#555588;line-height:1.6;">
          Found this in spam? Mark it <strong style="color:#9090b8;">Not Spam</strong> and add
          <strong style="color:#9090b8;">ssa.data.management@gmail.com</strong> to your contacts
          so future messages arrive in your inbox.
        </p>
      </div>

      <!-- DIVIDER -->
      <div style="border-top:1px solid #1e1e40;margin:0 0 24px;"></div>

      <p style="margin:0;font-size:0.78rem;color:#44446a;text-align:center;">
        Sorted on ${new Date(data.timestamp).toLocaleDateString('en-US', {weekday:'long',year:'numeric',month:'long',day:'numeric'})}<br>
        ${data.email}
      </p>
    </div>

    <!-- FOOTER -->
    <div style="background:#080818;border:1px solid #111130;border-radius:0 0 16px 16px;padding:20px 40px;text-align:center;">
      <p style="margin:0;font-size:0.75rem;color:#333358;letter-spacing:0.08em;">LEGACY &nbsp;·&nbsp; VALOR &nbsp;·&nbsp; HORIZON</p>
    </div>

  </div>
</body>
</html>`;

  MailApp.sendEmail({ to: data.email, from: 'ssa.data.management@gmail.com', name: 'SSA House Sorting', subject: subject, htmlBody: body });
}

// ── SYNC: rebuilds house tabs from Submissions ──
function syncHouseTabs() {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const subSheet = ss.getSheetByName(SHEET_NAME);
  if (!subSheet) return;

  const rows = subSheet.getDataRange().getValues();
  if (rows.length < 2) {
    // No data rows — clear house tabs
    ['Legacy','Valor','Horizon'].forEach(h => {
      const tab = ss.getSheetByName(h);
      if (tab && tab.getLastRow() > 1) tab.deleteRows(2, tab.getLastRow() - 1);
    });
    return;
  }

  const header = rows[0]; // Submissions header row
  const houseColIndex = header.indexOf('House'); // column F (index 5)

  const houseHeaders = ['Student #', 'Timestamp', 'First Name', 'Last Name', 'Email',
                        'Legacy Score', 'Valor Score', 'Horizon Score',
                        'Q1','Q2','Q3','Q4','Q5','Q6','Q7','Q8','Q9','Q10'];
  const houseColors  = { Legacy: '#6c2fa0', Valor: '#c0392b', Horizon: '#2471a3' };
  const tabColors    = { Legacy: '#d7b8f0', Valor: '#f0b8b8', Horizon: '#b8d8f0' };

  ['Legacy','Valor','Horizon'].forEach(h => {
    let tab = ss.getSheetByName(h);
    if (!tab) {
      tab = ss.insertSheet(h);
      tab.appendRow(houseHeaders);
      tab.getRange(1,1,1,houseHeaders.length).setFontWeight('bold')
         .setBackground(houseColors[h]).setFontColor('#ffffff');
      tab.setFrozenRows(1);
      tab.setTabColor(tabColors[h]);
    } else {
      // Clear existing data rows, keep header
      if (tab.getLastRow() > 1) tab.deleteRows(2, tab.getLastRow() - 1);
    }

    // Re-populate from Submissions rows that match this house
    let houseNum = 1;
    rows.slice(1).forEach(row => {
      if (row[houseColIndex] === h) {
        tab.appendRow([
          houseNum++,
          row[1],  // Timestamp
          row[2],  // First Name
          row[3],  // Last Name
          row[4],  // Email
          row[6],  // Legacy Score
          row[7],  // Valor Score
          row[8],  // Horizon Score
          row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18]
        ]);
      }
    });
  });
}

// Run this ONCE to set up the auto-sync trigger
function createSyncTrigger() {
  // Remove any existing onChange triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onSheetChange') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(SpreadsheetApp.openById(SPREADSHEET_ID))
    .onChange()
    .create();
}

function onSheetChange(e) {
  syncHouseTabs();
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'counts') {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const counts = {};
    ['Legacy', 'Valor', 'Horizon'].forEach(h => {
      const tab = ss.getSheetByName(h);
      counts[h] = tab ? Math.max(tab.getLastRow() - 1, 0) : 0;
    });
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', counts }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
