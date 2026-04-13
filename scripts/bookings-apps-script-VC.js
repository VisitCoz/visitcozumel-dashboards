// ============================================================
// Visit Cozumel — Booking System (VC-only)
// Handles: Stripe PaymentIntent + Sheets + Email + Calendar + Telegram
// + Telegram bot inbound + Daily manifest + Revenue dashboard
// + Stripe Payment Links webhook (legacy)
// ============================================================
// DEPLOYMENT INSTRUCTIONS:
// 1. Open the Visit Cozumel bookings Google Sheet
// 2. Extensions → Apps Script → replace ALL code with this file
// 3. Save (Cmd+S)
// 4. Deploy → Manage deployments → pencil icon on existing deployment
//    → Version: New version → Deploy
//    The deployment URL stays the same — no Squarespace change needed.
// ============================================================
// SECURITY NOTE:
// Stripe and Telegram secrets are read from Script Properties first
// (File → Project Properties → Script Properties), with the constants
// below as fallback. To migrate: set VC_STRIPE_SK and TELEGRAM_BOT_TOKEN
// in Script Properties, then blank the constants below.
// ============================================================

// ---------- Config ----------
var EMAILS = [
  'contabilidad@visitcozumel.com.mx',
  'hello@visitcozumel.com.mx',
  'operations@visitcozumel.com.mx',
  'admin@visitcozumel.com.mx',
  'e.magnusson@visitcozumel.com.mx'
];

var VC_CALENDAR_ID = 'd96d4681c1925423cffbe11b74dea43fba4eb25df036ff46368fc76e70abc32a@group.calendar.google.com';

// Fallback constants — prefer Script Properties for production
var VC_STRIPE_SK_FALLBACK = ''; // Set in Script Properties → VC_STRIPE_SK
var TELEGRAM_BOT_TOKEN_FALLBACK = ''; // Set in Script Properties → TELEGRAM_BOT_TOKEN

var VC_DISCOUNT_CODES = {
  'SUNNY15': { percent: 15 }
};

var USD_TO_MXN_RATE = 18.5;

var MIN_AMOUNT_CENTS = 5000;
var MAX_AMOUNT_CENTS = 18500000;

// ---------- Helpers ----------
function getSecret(key, fallback) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    return v || fallback;
  } catch (e) {
    return fallback;
  }
}

function formatDate(dateStr) {
  if (!dateStr) return 'N/A';
  try {
    var d;
    if (dateStr instanceof Date) {
      d = dateStr;
    } else {
      d = new Date(String(dateStr).substring(0, 10) + 'T12:00:00');
    }
    var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    return d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
  } catch (e) {
    return String(dateStr);
  }
}

function parseTime(input) {
  if (!input) return null;
  var s = String(input).trim();

  var m12 = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (m12) {
    var h = parseInt(m12[1], 10);
    var min = parseInt(m12[2], 10);
    var ampm = m12[3].toUpperCase();
    if (ampm === 'PM' && h < 12) h += 12;
    if (ampm === 'AM' && h === 12) h = 0;
    return { hours: h, minutes: min };
  }

  var m24 = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m24) {
    var h2 = parseInt(m24[1], 10);
    var min2 = parseInt(m24[2], 10);
    if (h2 >= 0 && h2 <= 23 && min2 >= 0 && min2 <= 59) {
      return { hours: h2, minutes: min2 };
    }
  }

  return null;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// CREATE VC PAYMENT INTENT
// ============================================================
function createVcPaymentIntent(data) {
  try {
    var usdCents = Math.round(Number(data.amount));
    var amount = Math.round(usdCents * USD_TO_MXN_RATE);

    if (!isFinite(amount) || amount < MIN_AMOUNT_CENTS || amount > MAX_AMOUNT_CENTS) {
      return jsonResponse({ error: 'Invalid amount' });
    }

    var couponCode = (data.coupon_code || '').toUpperCase();

    var stripeKey = getSecret('VC_STRIPE_SK', VC_STRIPE_SK_FALLBACK);
    var payload = 'amount=' + amount
      + '&currency=mxn'
      + '&automatic_payment_methods[enabled]=true'
      + '&metadata[booking_id]=' + encodeURIComponent(data.booking_id || '')
      + '&metadata[coupon_code]=' + encodeURIComponent(couponCode)
      + '&metadata[usd_display]=' + encodeURIComponent((usdCents / 100).toFixed(2))
      + '&metadata[fx_rate]=' + encodeURIComponent(USD_TO_MXN_RATE);

    if (data.email) {
      payload += '&receipt_email=' + encodeURIComponent(data.email);
    }

    var response = UrlFetchApp.fetch('https://api.stripe.com/v1/payment_intents', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + stripeKey,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: payload,
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode === 200) {
      var intent = JSON.parse(responseText);
      return jsonResponse({
        clientSecret: intent.client_secret,
        booking_id: data.booking_id
      });
    }

    var error = JSON.parse(responseText);
    return jsonResponse({
      error: error.error ? error.error.message : 'Failed to create payment intent'
    });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ============================================================
// MAIN WEBHOOK
// ============================================================
function doPost(e) {
  var raw = e.postData.contents;
  var data = JSON.parse(raw);

  if (data.type === 'checkout.session.completed') {
    return handleStripeWebhook(data);
  }

  if (data.message || data.callback_query) {
    try {
      handleTelegram(data);
    } catch (err) {}
    return ContentService.createTextOutput('ok');
  }

  if (data.action === 'create_vc_intent') {
    return createVcPaymentIntent(data);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VC Bookings');
  if (!sheet) {
    sheet = ss.insertSheet('VC Bookings');
    sheet.appendRow([
      'Booking ID', 'Timestamp', 'Name', 'Email', 'WhatsApp',
      'Arrival Date', 'Cruise Ship', 'Pickup Time', 'Pickup Destination',
      'Guests', 'Product Type', 'Vehicle', 'Price USD', 'Charged MXN',
      'Discount Code', 'Discount %', 'Special Requests', 'Payment Status',
      'Stripe Payment ID'
    ]);
  }

  var mxnCharged = data.price ? Math.round(Number(data.price) * USD_TO_MXN_RATE * 100) / 100 : '';

  sheet.appendRow([
    data.booking_id || '',
    data.timestamp || new Date().toISOString(),
    data.name || '',
    data.email || '',
    data.whatsapp || '',
    data.arrival_date || '',
    data.cruise_ship || '',
    data.pickup_time || '',
    data.destination || '',
    data.guests || '',
    data.product_type || 'excursion',
    data.tier || '',
    data.price || '',
    mxnCharged,
    data.discount_code || '',
    data.discount_percent || '',
    data.special_requests || '',
    'Pending',
    data.stripe_payment_id || ''
  ]);

  var dateFormatted = formatDate(data.arrival_date);
  var hasDiscount = data.discount_code && data.discount_code !== 'None';
  var discountNote = hasDiscount
    ? ' — ' + data.discount_code + ' (' + data.discount_percent + ' off)'
    : '';
  var subject = 'New Booking: ' + data.name + ' — ' + data.tier
    + ' ($' + data.price + ')' + discountNote + ' — ' + dateFormatted;

  var body = 'NEW BOOKING FROM VISITCOZUMEL.COM.MX\n\n'
    + 'Booking ID: ' + (data.booking_id || 'N/A') + '\n'
    + 'Name: ' + data.name + '\n'
    + 'Email: ' + data.email + '\n'
    + 'WhatsApp: ' + data.whatsapp + '\n'
    + 'Arrival Date: ' + dateFormatted + '\n'
    + 'Cruise Ship: ' + data.cruise_ship + '\n'
    + 'Pickup Time: ' + data.pickup_time + '\n'
    + 'Destination: ' + data.destination + '\n'
    + 'Guests: ' + data.guests + '\n'
    + 'Product: ' + (data.product_type || 'excursion') + '\n'
    + 'Vehicle: ' + data.tier + '\n'
    + 'Price: $' + data.price + ' USD\n'
    + 'Charged: $' + mxnCharged + ' MXN (rate ' + USD_TO_MXN_RATE + ')\n'
    + (hasDiscount ? 'Discount: ' + data.discount_code + ' (' + data.discount_percent + ' off)\n' : '')
    + 'Special Requests: ' + (data.special_requests || 'None') + '\n\n'
    + 'Payment Status: PENDING\n';

  EMAILS.forEach(function (email) { MailApp.sendEmail(email, subject, body); });

  var telegramMsg = 'New Visit Cozumel Booking\n'
    + data.name + ' — ' + data.tier + '\n'
    + dateFormatted + ' at ' + data.pickup_time + '\n'
    + data.guests + ' guests — $' + data.price + ' USD (charged $' + mxnCharged + ' MXN) — PENDING'
    + (hasDiscount ? '\n' + data.discount_code + ' (' + data.discount_percent + ' off)' : '');
  sendTelegramNotification(telegramMsg);

  try {
    var cal = CalendarApp.getCalendarById(VC_CALENDAR_ID);
    if (cal && data.arrival_date) {
      var startDate = new Date(data.arrival_date + 'T12:00:00');
      var parsed = parseTime(data.pickup_time);
      if (parsed) {
        startDate.setHours(parsed.hours, parsed.minutes, 0);
      } else if (data.pickup_time) {
        Logger.log('VC: could not parse pickup_time "' + data.pickup_time + '" — defaulting to noon');
      }
      var endDate = new Date(startDate.getTime() + 2 * 60 * 60 * 1000);
      var eventTitle = data.name + ' — ' + data.tier + ' (' + data.guests + ' guests)';
      var eventDesc = 'Booking ID: ' + (data.booking_id || 'N/A') + '\n'
        + 'Email: ' + data.email + '\nWhatsApp: ' + data.whatsapp + '\n'
        + 'Ship: ' + (data.cruise_ship || 'N/A') + '\n'
        + 'Destination: ' + (data.destination || 'N/A') + '\n'
        + 'Total: $' + data.price + ' USD\n'
        + 'Charged: $' + mxnCharged + ' MXN (rate ' + USD_TO_MXN_RATE + ')\n'
        + (hasDiscount ? 'Discount: ' + data.discount_code + ' (' + data.discount_percent + ' off)\n' : '')
        + 'Special: ' + (data.special_requests || 'None') + '\nPayment: PENDING';
      cal.createEvent(eventTitle, startDate, endDate, {
        description: eventDesc,
        location: 'Cozumel, Mexico'
      });
    }
  } catch (calErr) {
    Logger.log('VC Calendar error: ' + calErr);
  }

  return jsonResponse({ status: 'success' });
}

// ============================================================
// STRIPE WEBHOOK (legacy Payment Links flow)
// ============================================================
function handleStripeWebhook(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var session = data.data && data.data.object ? data.data.object : {};
  var bookingId = session.client_reference_id || '';
  var paymentStatus = session.payment_status || 'unknown';

  if (!bookingId) {
    return jsonResponse({ status: 'error', message: 'No booking ID' });
  }

  var sheet = ss.getSheetByName('VC Bookings');
  if (!sheet) {
    return jsonResponse({ status: 'error', message: 'VC Bookings sheet not found' });
  }

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === bookingId) {
      sheet.getRange(i + 1, 18).setValue('PAID');
      var name = values[i][2];
      var subject = 'PAYMENT RECEIVED: ' + name + ' — ' + bookingId;
      var body = 'PAYMENT CONFIRMED\n\nBooking ID: ' + bookingId
        + '\nName: ' + name + '\nStripe: ' + paymentStatus;
      EMAILS.forEach(function (email) { MailApp.sendEmail(email, subject, body); });
      sendTelegramNotification('Payment Received\n' + name + '\nBooking: ' + bookingId);
      return jsonResponse({ status: 'success', booking: bookingId });
    }
  }

  return jsonResponse({ status: 'not_found', booking: bookingId });
}

// ============================================================
// DAILY MANIFEST
// ============================================================
function sendDailyManifest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, 'America/Cancun', 'yyyy-MM-dd');
  var tomorrowFormatted = formatDate(tomorrowStr);

  var vcSheet = ss.getSheetByName('VC Bookings');
  if (!vcSheet || vcSheet.getLastRow() <= 1) return;

  var vcData = vcSheet.getDataRange().getValues();
  var bookings = [];
  for (var j = 1; j < vcData.length; j++) {
    var dv = vcData[j][5];
    var bd = (dv instanceof Date)
      ? Utilities.formatDate(dv, 'America/Cancun', 'yyyy-MM-dd')
      : String(dv).substring(0, 10);
    if (bd === tomorrowStr) {
      bookings.push({
        name: vcData[j][2],
        time: vcData[j][7],
        tier: vcData[j][11],
        guests: vcData[j][9],
        price: vcData[j][12],
        ship: vcData[j][6],
        dest: vcData[j][8],
        whatsapp: vcData[j][4],
        requests: vcData[j][16],
        payment: vcData[j][17]
      });
    }
  }

  if (bookings.length === 0) return;

  var manifest = 'VISIT COZUMEL — ' + bookings.length + ' booking'
    + (bookings.length > 1 ? 's' : '') + '\n━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  for (var v = 0; v < bookings.length; v++) {
    var c = bookings[v];
    manifest += '\n' + (v + 1) + '. ' + c.name + '\n'
      + '   Pickup: ' + c.time + '\n'
      + '   Vehicle: ' + c.tier + '\n'
      + '   Guests: ' + c.guests + '\n'
      + '   Dest: ' + c.dest + '\n'
      + '   Total: $' + c.price + ' USD\n'
      + '   Ship: ' + (c.ship || 'N/A') + '\n'
      + '   WhatsApp: ' + c.whatsapp + '\n'
      + '   Payment: ' + c.payment + '\n';
    if (c.requests && c.requests !== 'None' && c.requests !== 'N/A' && c.requests !== '') {
      manifest += '   Note: ' + c.requests + '\n';
    }
  }

  var subject = 'Tomorrow\'s Manifest — ' + tomorrowFormatted + ' — '
    + bookings.length + ' booking' + (bookings.length > 1 ? 's' : '');
  var body = 'BOOKINGS FOR TOMORROW: ' + tomorrowFormatted + '\n'
    + bookings.length + ' total\n\n' + manifest
    + '\n━━━━━━━━━━━━━━━━━━━━━━━━━\nAutomated daily reminder at 6 PM';
  EMAILS.forEach(function (email) { MailApp.sendEmail(email, subject, body); });
  sendTelegramNotification('*Tomorrow\'s Manifest — ' + tomorrowFormatted + '*\n'
    + bookings.length + ' booking' + (bookings.length > 1 ? 's' : '') + '\n\n' + manifest);
}

// ============================================================
// REVENUE DASHBOARD
// ============================================================
function updateRevenueDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Revenue');
  if (!dashboard) {
    dashboard = ss.insertSheet('Revenue');
    dashboard.getRange('A1:C1').setValues([['Month', 'VC Bookings', 'VC Revenue']]);
    dashboard.getRange('A1:C1').setFontWeight('bold').setBackground('#1a7f4e').setFontColor('#fff');
  }

  var months = {};

  var vcSheet = ss.getSheetByName('VC Bookings');
  if (vcSheet && vcSheet.getLastRow() > 1) {
    var vcData = vcSheet.getDataRange().getValues();
    for (var j = 1; j < vcData.length; j++) {
      var dv = vcData[j][5];
      var monthKey = '';
      if (dv instanceof Date) {
        monthKey = Utilities.formatDate(dv, 'America/Cancun', 'yyyy-MM');
      } else {
        monthKey = String(dv).substring(0, 7);
      }
      if (!monthKey || monthKey.length < 7) continue;
      if (!months[monthKey]) months[monthKey] = { count: 0, rev: 0 };
      months[monthKey].count++;
      months[monthKey].rev += parseFloat(vcData[j][12]) || 0;
    }
  }

  var sortedMonths = Object.keys(months).sort();
  if (sortedMonths.length > 0) {
    if (dashboard.getLastRow() > 1) {
      dashboard.getRange(2, 1, dashboard.getLastRow() - 1, 3).clearContent();
    }
    for (var m = 0; m < sortedMonths.length; m++) {
      var mk = sortedMonths[m];
      var d = months[mk];
      dashboard.getRange(m + 2, 1, 1, 3).setValues([[mk, d.count, d.rev]]);
    }
    dashboard.getRange(2, 3, sortedMonths.length, 1).setNumberFormat('$#,##0');
  }
}

// ============================================================
// TELEGRAM
// ============================================================
function sendTelegramNotification(text) {
  var token = getSecret('TELEGRAM_BOT_TOKEN', TELEGRAM_BOT_TOKEN_FALLBACK);
  if (!token) return;
  try {
    var chatIds = getChatIds();
    for (var i = 0; i < chatIds.length; i++) {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: chatIds[i],
          text: text,
          parse_mode: 'Markdown'
        }),
        muteHttpExceptions: true
      });
    }
  } catch (e) {
    Logger.log('Telegram send error: ' + e);
  }
}

function getChatIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName('Config');
  if (!config) return [];
  var data = config.getDataRange().getValues();
  var ids = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 'telegram_chat_id') ids.push(String(data[i][1]));
  }
  return ids;
}

function handleTelegram(update) {
  var token = getSecret('TELEGRAM_BOT_TOKEN', TELEGRAM_BOT_TOKEN_FALLBACK);
  if (!token) return;
  var msg = update.message;
  if (!msg || !msg.text) return;

  var chatId = msg.chat.id;
  var text = msg.text.trim().toLowerCase();

  if (text === '/start') {
    saveChatId(chatId);
    replyTelegram(chatId, 'Connected! You\'ll receive booking notifications here.\n\nCommands:\n/tomorrow — tomorrow\'s bookings\n/today — today\'s bookings\n/search Name — find a booking\n/revenue — monthly revenue\n/pending — unpaid bookings');
    return;
  }

  if (text === '/tomorrow') {
    var result = getBookingsForDate(1);
    replyTelegram(chatId, result || 'No bookings for tomorrow.');
    return;
  }

  if (text === '/today') {
    var result = getBookingsForDate(0);
    replyTelegram(chatId, result || 'No bookings for today.');
    return;
  }

  if (text.indexOf('/search ') === 0) {
    var query = msg.text.trim().substring(8);
    var result = searchBookings(query);
    replyTelegram(chatId, result || 'No bookings found for "' + query + '"');
    return;
  }

  if (text === '/revenue') {
    updateRevenueDashboard();
    var result = getRevenueReport();
    replyTelegram(chatId, result);
    return;
  }

  if (text === '/pending') {
    var result = getPendingBookings();
    replyTelegram(chatId, result || 'No pending payments.');
    return;
  }

  replyTelegram(chatId, 'Commands:\n/tomorrow — tomorrow\'s bookings\n/today — today\'s bookings\n/search Name — find a booking\n/revenue — monthly revenue\n/pending — unpaid bookings');
}

function saveChatId(chatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName('Config');
  if (!config) {
    config = ss.insertSheet('Config');
    config.getRange('A1:B1').setValues([['Key', 'Value']]);
  }
  var data = config.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 'telegram_chat_id' && String(data[i][1]) === String(chatId)) return;
  }
  config.appendRow(['telegram_chat_id', chatId]);
}

function replyTelegram(chatId, text) {
  var token = getSecret('TELEGRAM_BOT_TOKEN', TELEGRAM_BOT_TOKEN_FALLBACK);
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: String(chatId), text: text, parse_mode: 'Markdown' })
  });
}

function getBookingsForDate(daysFromNow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target = new Date();
  target.setDate(target.getDate() + daysFromNow);
  var targetStr = Utilities.formatDate(target, 'America/Cancun', 'yyyy-MM-dd');
  var targetFormatted = formatDate(targetStr);
  var label = daysFromNow === 0 ? 'Today' : 'Tomorrow';

  var vcSheet = ss.getSheetByName('VC Bookings');
  if (!vcSheet || vcSheet.getLastRow() <= 1) return '';

  var vcData = vcSheet.getDataRange().getValues();
  var lines = [];
  for (var j = 1; j < vcData.length; j++) {
    var dv = vcData[j][5];
    var bd = (dv instanceof Date) ? Utilities.formatDate(dv, 'America/Cancun', 'yyyy-MM-dd') : String(dv).substring(0, 10);
    if (bd === targetStr) {
      var line = '*' + vcData[j][2] + '* — ' + vcData[j][11] + '\n'
        + '   ' + vcData[j][7] + ' · ' + vcData[j][9] + ' guests\n'
        + '   $' + vcData[j][12] + ' · ' + vcData[j][17] + '\n'
        + '   Ship: ' + (vcData[j][6] || 'N/A') + '\n';
      lines.push(line);
    }
  }

  if (lines.length === 0) return '';
  return '*' + label + ': ' + targetFormatted + '* — ' + lines.length + ' booking'
    + (lines.length > 1 ? 's' : '') + '\n\n' + lines.join('\n');
}

function searchBookings(query) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var q = query.toLowerCase();
  var vcSheet = ss.getSheetByName('VC Bookings');
  if (!vcSheet || vcSheet.getLastRow() <= 1) return '';

  var vcData = vcSheet.getDataRange().getValues();
  var lines = [];
  for (var j = 1; j < vcData.length; j++) {
    if (String(vcData[j][2]).toLowerCase().indexOf(q) > -1
        || String(vcData[j][3]).toLowerCase().indexOf(q) > -1) {
      lines.push('*' + vcData[j][2] + '*\n'
        + '   ' + formatDate(vcData[j][5]) + ' · ' + vcData[j][11]
        + ' · $' + vcData[j][12] + ' · ' + vcData[j][17]);
    }
  }

  if (lines.length === 0) return '';
  return '*Search: "' + query + '"*\n\n' + lines.join('\n\n');
}

function getPendingBookings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vcSheet = ss.getSheetByName('VC Bookings');
  if (!vcSheet || vcSheet.getLastRow() <= 1) return '';

  var vcData = vcSheet.getDataRange().getValues();
  var lines = [];
  for (var j = 1; j < vcData.length; j++) {
    if (String(vcData[j][17]).toLowerCase() === 'pending') {
      lines.push(vcData[j][2] + ' — $' + vcData[j][12] + '\n'
        + '   ' + formatDate(vcData[j][5]) + ' · ' + vcData[j][11]);
    }
  }

  if (lines.length === 0) return '';
  return '*Pending Payments*\n\n' + lines.join('\n\n');
}

function getRevenueReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Revenue');
  if (!dashboard || dashboard.getLastRow() < 2) return 'No revenue data yet.';

  var data = dashboard.getDataRange().getValues();
  var result = '*Revenue Report*\n\n';
  var grandTotal = 0;

  for (var i = 1; i < data.length; i++) {
    var month = data[i][0];
    var count = data[i][1];
    var rev = data[i][2] || 0;
    grandTotal += rev;
    result += '*' + month + '* — ' + count + ' bookings — $' + rev + '\n';
  }

  result += '\n*Total: $' + grandTotal + ' USD*';
  return result;
}

// ============================================================
// TELEGRAM WEBHOOK SETUP
// ============================================================
function clearTelegramQueue() {
  var token = getSecret('TELEGRAM_BOT_TOKEN', TELEGRAM_BOT_TOKEN_FALLBACK);
  var res1 = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/deleteWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ drop_pending_updates: true })
  });
  Logger.log('Delete webhook: ' + res1.getContentText());

  var WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbzoXj9Dt6uoqgWEmmb6K9540-wRmKEHUJ6_NY9koDiQyk_e8aVtysHRWXIJ8YO8AxUz/exec';
  var res2 = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ url: WEBAPP_URL })
  });
  Logger.log('Set webhook: ' + res2.getContentText());
}

function setTelegramWebhook() {
  var token = getSecret('TELEGRAM_BOT_TOKEN', TELEGRAM_BOT_TOKEN_FALLBACK);
  var WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbzoXj9Dt6uoqgWEmmb6K9540-wRmKEHUJ6_NY9koDiQyk_e8aVtysHRWXIJ8YO8AxUz/exec';
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setWebhook', {
    method: 'post',
    payload: { url: WEBAPP_URL }
  });
  Logger.log(res.getContentText());
}

function checkTelegramWebhook() {
  var token = getSecret('TELEGRAM_BOT_TOKEN', TELEGRAM_BOT_TOKEN_FALLBACK);
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getWebhookInfo');
  Logger.log(res.getContentText());
}

function doGet() {
  return jsonResponse({ status: 'ok', message: 'Visit Cozumel booking endpoint is live' });
}
