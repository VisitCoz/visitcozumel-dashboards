/**
 * Celebrity Email Parser — Google Apps Script
 * Auto-parses all Celebrity Cruises emails into a structured Google Sheet.
 *
 * SETUP:
 * 1. Create a new Google Sheet called "Celebrity Email Tracker"
 * 2. Open Extensions > Apps Script
 * 3. Paste this entire script
 * 4. Run setupSheet() once to create headers and formatting
 * 5. Run createTimeTrigger() once to set up auto-processing every 30 minutes
 * 6. Done — emails will be parsed automatically
 *
 * Created: 2026-03-30 | Visit Cozumel Command Center
 */

// === CONFIGURATION ===
const CONFIG = {
  SHEET_NAME: 'Email Log',
  CONTACTS_SHEET: 'Contacts',
  STATS_SHEET: 'Stats',
  PROCESSED_LABEL: 'Processed/Celebrity',
  SEARCH_QUERY: 'from:(@celebrity.com OR @rccl.com OR @intercruises.com) -label:Processed/Celebrity',
  MAX_EMAILS_PER_RUN: 50,
  DRIVE_FOLDER_NAME: 'Celebrity Email Attachments',
};

// Ship code to name mapping
const SHIP_MAP = {
  'XC': 'Celebrity Xcel',
  'EC': 'Celebrity Eclipse',
  'SI': 'Celebrity Silhouette',
  'CS': 'Celebrity Constellation',
  'AX': 'Celebrity Apex',
  'SM': 'Celebrity Summit',
  'BT': 'Celebrity Beyond',
  'ED': 'Celebrity Edge',
  'FL': 'Celebrity Flora',
  'RF': 'Celebrity Reflection',
  'SL': 'Celebrity Solstice',
  'EQ': 'Celebrity Equinox',
};

// Email type classification patterns
const EMAIL_TYPES = [
  { pattern: /invoice/i, type: 'Invoice' },
  { pattern: /updated?\s*counts?/i, type: 'Tour Counts (Updated)' },
  { pattern: /preliminary.*counts?/i, type: 'Tour Counts (Preliminary)' },
  { pattern: /tour\s*counts?/i, type: 'Tour Counts' },
  { pattern: /private.*(?:tour|journey|request)/i, type: 'Private Tour Request' },
  { pattern: /inventory\s*request/i, type: 'Inventory Request' },
  { pattern: /no\s*booking/i, type: 'No Booking Notice' },
  { pattern: /tcw|tour\s*content/i, type: 'Tour Content Worksheet' },
  { pattern: /training|celebrity\s*way/i, type: 'Training' },
  { pattern: /fam\s*trip/i, type: 'Fam Trip' },
  { pattern: /GI\s*Gastrointestinal|illness|sanitation/i, type: 'Health Advisory' },
  { pattern: /cancel/i, type: 'Cancellation' },
];

// === MAIN PROCESSING ===

function processNewEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  const label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);

  const threads = GmailApp.search(CONFIG.SEARCH_QUERY, 0, CONFIG.MAX_EMAILS_PER_RUN);

  if (threads.length === 0) {
    Logger.log('No new Celebrity emails to process.');
    return;
  }

  Logger.log(`Processing ${threads.length} threads...`);
  let rowsAdded = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      const from = message.getFrom();

      // Only process emails FROM Celebrity/RCCL/Intercruises (skip our own replies)
      if (!isTargetSender(from)) continue;

      const parsed = parseEmail(message);
      appendRow(sheet, parsed);
      rowsAdded++;
    }

    // Label the thread as processed
    thread.addLabel(label);
  }

  Logger.log(`Done. Added ${rowsAdded} rows.`);

  // Update stats
  updateStats(ss);
}

// === EMAIL PARSING ===

function parseEmail(message) {
  const from = message.getFrom();
  const subject = message.getSubject();
  const date = message.getDate();
  const messageId = message.getId();

  // Extract sender name and email
  const senderInfo = extractSender(from);

  // Detect ship
  const shipInfo = detectShip(from, subject);

  // Classify email type
  const emailType = classifyEmail(subject);

  // Extract operation date from subject (the tour date, not email date)
  const operationDate = extractOperationDate(subject);

  // Check for attachments
  const attachments = message.getAttachments();
  const hasAttachments = attachments.length > 0;
  let driveLink = '';

  if (hasAttachments) {
    driveLink = saveAttachments(message, attachments, shipInfo.code, operationDate);
  }

  // Check if email needs a response (heuristic)
  const needsResponse = checkNeedsResponse(subject, emailType);

  // Extract snippet (first 200 chars of body, no HTML)
  const bodySnippet = message.getPlainBody().substring(0, 200).replace(/\n/g, ' ').trim();

  return {
    date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    shipCode: shipInfo.code,
    shipName: shipInfo.name,
    emailType: emailType,
    operationDate: operationDate,
    senderName: senderInfo.name,
    senderEmail: senderInfo.email,
    subject: subject,
    hasAttachments: hasAttachments ? 'Yes' : 'No',
    driveLink: driveLink,
    needsResponse: needsResponse,
    bodySnippet: bodySnippet,
    messageId: messageId,
    status: 'New',
  };
}

function extractSender(fromField) {
  // "Karina Torres <xc_seniordestinationsmanager@celebrity.com>"
  const match = fromField.match(/^"?([^"<]+)"?\s*<?([^>]+)>?$/);
  if (match) {
    return { name: match[1].trim(), email: match[2].trim() };
  }
  return { name: fromField, email: fromField };
}

function detectShip(from, subject) {
  // Method 1: Email prefix (most reliable)
  const emailMatch = from.match(/<?(\w{2})_\w+@celebrity\.com/i);
  if (emailMatch) {
    const code = emailMatch[1].toUpperCase();
    if (SHIP_MAP[code]) {
      return { code: code, name: SHIP_MAP[code] };
    }
  }

  // Method 2: Subject line
  for (const [code, name] of Object.entries(SHIP_MAP)) {
    const shipNameShort = name.replace('Celebrity ', '');
    if (subject.includes(shipNameShort) || subject.includes(`CEL ${code}`) || subject.includes(`Celebrity ${code}`)) {
      return { code: code, name: name };
    }
  }

  // Method 3: Check for RCCL/Intercruises (not ship-specific)
  if (from.includes('rccl.com')) return { code: 'RCCL', name: 'Royal Caribbean Group' };
  if (from.includes('intercruises.com')) return { code: 'DCL', name: 'Disney Cruise Line (Intercruises)' };

  return { code: '??', name: 'Unknown' };
}

function classifyEmail(subject) {
  for (const { pattern, type } of EMAIL_TYPES) {
    if (pattern.test(subject)) return type;
  }
  return 'Other';
}

function extractOperationDate(subject) {
  // Try patterns like "March 18th", "Mar 5", "03.25.2026", "February 4th"
  const patterns = [
    // "March 18th" or "March 18"
    /(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})(?:st|nd|rd|th)?/i,
    // "Mar 5" abbreviated
    /(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})(?:st|nd|rd|th)?/i,
    // "03.25.2026" or "03-25-2026"
    /(\d{2})[.\-](\d{2})[.\-](\d{4})/,
    // ISO date "2026-03-25"
    /(\d{4})-(\d{2})-(\d{2})/,
  ];

  for (const p of patterns) {
    const match = subject.match(p);
    if (match) {
      return match[0].replace(/(?:st|nd|rd|th)/i, '');
    }
  }
  return '';
}

function checkNeedsResponse(subject, emailType) {
  // Invoices always need approval
  if (emailType === 'Invoice') return 'Approve Invoice';
  if (emailType === 'Private Tour Request') return 'Send Quote';
  if (emailType === 'Tour Counts (Preliminary)') return 'Confirm Min/Max/Times';
  if (emailType === 'Inventory Request') return 'Add Allotment';
  if (emailType === 'Tour Content Worksheet') return 'Review & Confirm';
  if (emailType === 'Health Advisory') return 'Acknowledge';
  return '';
}

// === ATTACHMENTS ===

function saveAttachments(message, attachments, shipCode, opDate) {
  const folder = getOrCreateDriveFolder();
  const links = [];

  for (const att of attachments) {
    const name = att.getName();
    // Skip signature images and tiny files
    if (att.getSize() < 5000 && /\.(png|jpg|gif)$/i.test(name)) continue;

    const prefix = `${shipCode}_${opDate || 'nodate'}_`;
    const file = folder.createFile(att.copyBlob().setName(prefix + name));
    links.push(file.getUrl());
  }

  return links.join('\n');
}

function getOrCreateDriveFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
}

// === SHEET HELPERS ===

function appendRow(sheet, data) {
  sheet.appendRow([
    data.date,
    data.shipCode,
    data.shipName,
    data.emailType,
    data.operationDate,
    data.senderName,
    data.senderEmail,
    data.subject,
    data.hasAttachments,
    data.driveLink,
    data.needsResponse,
    data.status,
    data.bodySnippet,
    data.messageId,
  ]);
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

function isTargetSender(from) {
  return /celebrity\.com|rccl\.com|intercruises\.com/i.test(from);
}

// === STATS SHEET ===

function updateStats(ss) {
  const logSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!logSheet || logSheet.getLastRow() < 2) return;

  let statsSheet = ss.getSheetByName(CONFIG.STATS_SHEET);
  if (!statsSheet) {
    statsSheet = ss.insertSheet(CONFIG.STATS_SHEET);
  }
  statsSheet.clear();

  // Pull all data
  const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 14).getValues();

  // Count by ship
  const shipCounts = {};
  const shipInvoices = {};
  const senderCounts = {};
  const typeCounts = {};
  const pendingResponses = [];

  for (const row of data) {
    const shipCode = row[1];
    const shipName = row[2];
    const emailType = row[3];
    const senderName = row[5];
    const needsResponse = row[10];
    const status = row[11];

    // Ship totals
    shipCounts[shipCode] = (shipCounts[shipCode] || 0) + 1;

    // Ship invoices
    if (emailType === 'Invoice') {
      shipInvoices[shipCode] = (shipInvoices[shipCode] || 0) + 1;
    }

    // Sender totals
    if (senderName) {
      senderCounts[senderName] = (senderCounts[senderName] || 0) + 1;
    }

    // Type totals
    typeCounts[emailType] = (typeCounts[emailType] || 0) + 1;

    // Pending
    if (needsResponse && status === 'New') {
      pendingResponses.push([row[0], shipName, emailType, needsResponse, row[7]]);
    }
  }

  // Write stats
  let r = 1;
  statsSheet.getRange(r, 1).setValue('CELEBRITY EMAIL STATS').setFontWeight('bold').setFontSize(14);
  statsSheet.getRange(r, 2).setValue(`Updated: ${new Date().toLocaleString()}`);
  r += 2;

  // Emails by ship
  statsSheet.getRange(r, 1).setValue('EMAILS BY SHIP').setFontWeight('bold');
  r++;
  statsSheet.getRange(r, 1, 1, 3).setValues([['Ship', 'Total Emails', 'Invoices']]).setFontWeight('bold');
  r++;
  const sorted = Object.entries(shipCounts).sort((a, b) => b[1] - a[1]);
  for (const [code, count] of sorted) {
    statsSheet.getRange(r, 1, 1, 3).setValues([[code, count, shipInvoices[code] || 0]]);
    r++;
  }
  r++;

  // Top senders
  statsSheet.getRange(r, 1).setValue('TOP SENDERS').setFontWeight('bold');
  r++;
  const sortedSenders = Object.entries(senderCounts).sort((a, b) => b[1] - a[1]).slice(0, 10);
  for (const [name, count] of sortedSenders) {
    statsSheet.getRange(r, 1, 1, 2).setValues([[name, count]]);
    r++;
  }
  r++;

  // By type
  statsSheet.getRange(r, 1).setValue('EMAILS BY TYPE').setFontWeight('bold');
  r++;
  const sortedTypes = Object.entries(typeCounts).sort((a, b) => b[1] - a[1]);
  for (const [type, count] of sortedTypes) {
    statsSheet.getRange(r, 1, 1, 2).setValues([[type, count]]);
    r++;
  }
  r++;

  // Pending responses
  statsSheet.getRange(r, 1).setValue(`PENDING RESPONSES (${pendingResponses.length})`).setFontWeight('bold').setFontColor('#cc0000');
  r++;
  if (pendingResponses.length > 0) {
    statsSheet.getRange(r, 1, 1, 5).setValues([['Date', 'Ship', 'Type', 'Action Needed', 'Subject']]).setFontWeight('bold');
    r++;
    for (const row of pendingResponses) {
      statsSheet.getRange(r, 1, 1, 5).setValues([row]);
      r++;
    }
  }

  // Auto-size columns
  for (let c = 1; c <= 5; c++) {
    statsSheet.autoResizeColumn(c);
  }
}

// === SETUP (RUN ONCE) ===

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);

  // Set headers
  const headers = [
    'Email Date',      // A
    'Ship Code',       // B
    'Ship Name',       // C
    'Email Type',      // D
    'Operation Date',  // E
    'Sender Name',     // F
    'Sender Email',    // G
    'Subject',         // H
    'Attachments?',    // I
    'Drive Link',      // J
    'Needs Response',  // K
    'Status',          // L
    'Body Snippet',    // M
    'Message ID',      // N
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 140);  // Date
  sheet.setColumnWidth(2, 60);   // Ship Code
  sheet.setColumnWidth(3, 160);  // Ship Name
  sheet.setColumnWidth(4, 180);  // Type
  sheet.setColumnWidth(5, 100);  // Op Date
  sheet.setColumnWidth(6, 160);  // Sender
  sheet.setColumnWidth(7, 250);  // Email
  sheet.setColumnWidth(8, 400);  // Subject
  sheet.setColumnWidth(9, 80);   // Attachments
  sheet.setColumnWidth(10, 200); // Drive Link
  sheet.setColumnWidth(11, 160); // Needs Response
  sheet.setColumnWidth(12, 80);  // Status
  sheet.setColumnWidth(13, 300); // Snippet
  sheet.setColumnWidth(14, 180); // Message ID

  // Add data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Responded', 'Approved', 'Flagged', 'Ignored'], true)
    .build();
  sheet.getRange('L2:L1000').setDataValidation(statusRule);

  // Conditional formatting — highlight rows needing response
  const range = sheet.getRange('K2:K1000');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Approve')
    .setBackground('#fce4ec')
    .setRanges([range])
    .build();
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Quote')
    .setBackground('#fff3e0')
    .setRanges([range])
    .build();
  sheet.setConditionalFormatRules([rule, rule2]);

  // Create Contacts sheet
  const contactSheet = getOrCreateSheet(ss, CONFIG.CONTACTS_SHEET);
  contactSheet.getRange(1, 1, 1, 7).setValues([[
    'Name', 'Email', 'Role', 'Ship Code', 'Ship Name', 'WhatsApp', 'Notes'
  ]]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  contactSheet.setFrozenRows(1);

  Logger.log('Sheet setup complete. Now run createTimeTrigger() to enable auto-processing.');
}

function createTimeTrigger() {
  // Remove existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'processNewEmails') {
      ScriptApp.deleteTrigger(t);
    }
  }

  // Create new trigger — runs every 30 minutes
  ScriptApp.newTrigger('processNewEmails')
    .timeBased()
    .everyMinutes(30)
    .create();

  Logger.log('Trigger created. processNewEmails() will run every 30 minutes.');
}

// === MANUAL BACKFILL ===

function backfillExistingEmails() {
  /**
   * Run this ONCE to process all existing Celebrity emails.
   * After this, the auto-trigger handles new ones.
   *
   * WARNING: This may process hundreds of emails.
   * Google Apps Script has a 6-minute execution limit.
   * If it times out, just run it again — already-labeled emails are skipped.
   */
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  const label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);

  // Search ALL Celebrity emails (not just unprocessed)
  const query = 'from:(@celebrity.com OR @rccl.com OR @intercruises.com) -label:Processed/Celebrity';
  let start = 0;
  let total = 0;

  while (true) {
    const threads = GmailApp.search(query, start, 50);
    if (threads.length === 0) break;

    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        if (!isTargetSender(message.getFrom())) continue;
        const parsed = parseEmail(message);
        appendRow(sheet, parsed);
        total++;
      }
      thread.addLabel(label);
    }

    start += 50;
    Logger.log(`Processed ${total} emails so far...`);

    // Safety: check remaining execution time
    if (total > 400) {
      Logger.log('Hit safety limit. Run again to continue backfill.');
      break;
    }
  }

  Logger.log(`Backfill complete. ${total} emails processed.`);
  updateStats(ss);
}

// === UTILITY ===

function testWithOneEmail() {
  /** Test the parser on the most recent Celebrity email */
  const threads = GmailApp.search('from:@celebrity.com', 0, 1);
  if (threads.length === 0) {
    Logger.log('No Celebrity emails found.');
    return;
  }
  const msg = threads[0].getMessages()[0];
  const parsed = parseEmail(msg);
  Logger.log(JSON.stringify(parsed, null, 2));
}
