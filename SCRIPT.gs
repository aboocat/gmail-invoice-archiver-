/**
 * Gmail Invoice Collector → Drive Archive
 * Built by Daniel Schnitterbaum (with ChatGPT as co-pilot)
 * Works. No SaaS. No bullshit.
 *
 * Features:
 * - PDF attachments → Drive
 * - HTML invoices → PDF
 * - Vendor folder auto-creation
 * - Dedupe via label + filename check
 * - Batched runs (Apps Script quota-safe)
 * - Weekly digest
 * - Reprocess function
 *
 * License: MIT
 */


/**** CONFIG ****/
const ROOT_FOLDER_ID   = '1SJnBUuUGyX9i16latsggxIhC4Md5hFCd'; // <- dein Root-Folder (Hauptaccount)
const PROCESSED_LABEL  = 'invoices_exported';
const DIGEST_LABEL     = 'invoices_exported'; // gleiche Basis
const SEARCH_QUERY =
  'newer_than:365d has:nouserlabels (' +
    // PDFs im Anhang
    '(has:attachment filename:pdf subject:(invoice OR rechnung OR receipt)) ' +
    'OR ' +
    // mails ohne Anhang (Apple, Stripe etc.)
    '(subject:(invoice OR rechnung OR receipt OR "Ihre Rechnung" OR "Your receipt"))' +
  ')';

/**** MAIN: export ****/
function exportInvoices() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const processed = GmailApp.getUserLabelByName(PROCESSED_LABEL) || GmailApp.createLabel(PROCESSED_LABEL);

  const PAGE = 150;                // Threads pro Seite
  const SOFT_LIMIT_MS = 5 * 60e3;  // ~5 Minuten, unter 6-Min Apps Script Limit
  const startTime = Date.now();

  let start = 0;
  while (true) {
    // Hole eine Seite
    const threads = GmailApp.search(SEARCH_QUERY, start, PAGE);
    if (!threads.length) break;

    // Verarbeite Seite
    for (const t of threads) {
      let touched = false;
      for (const m of t.getMessages()) {
        const vendor = detectVendor(m);
        const vendorFolder = ensureVendorFolder(root, vendor);

        // 1) PDF-Anhänge sichern
        const atts = m.getAttachments({ includeInlineImages: false });
        for (const att of atts) {
          if (isPdf(att)) {
            const name = buildFilename(m, vendor, att.getName());
            if (!fileExists(vendorFolder, name)) {
              vendorFolder.createFile(att.copyBlob()).setName(name);
              touched = true;
            }
          }
        }

        // 2) Wenn keine PDFs gesichert: Mail → PDF (inkl. Inline-Images)
        if (!touched) {
          const html = inlineRemoteImages(m.getBody());
          const name = buildFilename(m, vendor, sanitize(m.getSubject()) + '.pdf');
          if (!fileExists(vendorFolder, name)) {
            const pdf = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(name);
            vendorFolder.createFile(pdf);
            touched = true;
          }
        }
      }
      if (touched) t.addLabel(processed);
    }

    // Nächste Seite anfragen
    start += PAGE;

    // Soft time-out, damit lange Postfächer in mehreren Läufen abgearbeitet werden
    if (Date.now() - startTime > SOFT_LIMIT_MS) {
      Logger.log('Zeitbudget erreicht – weiterer Durchlauf übernimmt den Rest.');
      break;
    }
  }
}

/**** WEEKLY DIGEST ****/
function weeklyDigest() {
  // Threads, die in der letzten Woche verarbeitet wurden
  const since = new Date();
  since.setDate(since.getDate() - 7);
  const label = GmailApp.getUserLabelByName(DIGEST_LABEL);
  if (!label) return;

  const threads = label.getThreads().filter(t => t.getLastMessageDate() >= since);
  const tally = {}; // { vendor: count }
  const samples = []; // kleine Liste als Beispiel

  threads.forEach(t => {
    t.getMessages().forEach(m => {
      const v = detectVendor(m);
      tally[v] = (tally[v] || 0) + 1;
      if (samples.length < 10) samples.push(`${formatDate(m.getDate())} · ${v} · ${m.getSubject()}`);
    });
  });

  const lines = Object.keys(tally)
    .sort((a,b) => tally[b]-tally[a])
    .map(v => `• ${v}: ${tally[v]}`);

  const body =
    `Invoices archived last 7 days:\n\n` +
    (lines.length ? lines.join('\n') : '– none –') +
    `\n\nExamples:\n${samples.length ? samples.join('\n') : '–'}`;

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: `Weekly Invoice Digest (${new Date().toDateString()})`,
    body
  });
}

/**** ONE-CLICK REPROCESS ****/
// Entfernt das PROCESSED_LABEL, damit exportInvoices() erneut greift.
// Optional: query übergeben (z.B. 'from:@apple.com newer_than:90d'), sonst alles.
function unprocessAll(query) {
  const label = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!label) return;
  const threads = query ? GmailApp.search(query) : label.getThreads();
  threads.forEach(t => t.removeLabel(label));
}

/**** HELPERS ****/
function detectVendor(message) {
  // Primär: Absender-Domain
  const from = message.getFrom(); // e.g. "Apple <no_reply@email.apple.com>"
  const email = from.match(/<([^>]+)>/) ? RegExp.$1 : from;
  const domain = (email.split('@')[1] || '').toLowerCase();
  let vendor = domain.replace(/^www\./,'').split('.')[0]; // email.apple.com -> email -> 'email'
  // Korrigieren für gängige Subdomains
  const map = { 'email': 'apple', 'invoices': 'amazon', 'billing': 'billing', 'mailer': 'mailer' };
  if (map[vendor]) vendor = map[vendor];

  // Fallback: Subject-Marken
  const subj = message.getSubject().toLowerCase();
  if (/telekom|t-mobile|magenta/.test(subj)) vendor = 'telekom';
  if (/vodafone/.test(subj)) vendor = 'vodafone';
  if (/o2/.test(subj)) vendor = 'o2';
  if (/paypal/.test(subj)) vendor = 'paypal';
  if (/amazon/.test(subj)) vendor = 'amazon';
  if (/stripe/.test(subj)) vendor = 'stripe';
  if (/apple/.test(subj)) vendor = 'apple';

  return sanitize(vendor || 'misc');
}

function ensureVendorFolder(root, vendor) {
  const iter = root.getFoldersByName(vendor);
  return iter.hasNext() ? iter.next() : root.createFolder(vendor);
}

function isPdf(att) {
  const n = att.getName().toLowerCase();
  const ct = (att.getContentType() || '').toLowerCase();
  return n.endsWith('.pdf') || ct === 'application/pdf';
}

function buildFilename(message, vendor, baseName) {
  const date = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyyMMdd');
  return `${date}_${vendor}_${sanitize(baseName)}`;
}

function sanitize(s) {
  return String(s)
    .replace(/\r?\n/g, ' ')
    .replace(/[^\p{L}\p{N}\-_. ]+/gu, '')
    .trim()
    .replace(/\s+/g, '_')
    .slice(0, 120);
}

function fileExists(folder, name) {
  return folder.getFilesByName(name).hasNext();
}

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Versucht, <img src="https://..."> in data: URLs zu konvertieren,
 * damit sie im PDF erscheinen. (Best effort; cid: Bilder lassen wir i.d.R. in Ruhe.)
 */
function inlineRemoteImages(html) {
  const maxImages = 10;
  let count = 0;
  return html.replace(/<img\s+[^>]*src=["'](http[^"']+)["'][^>]*>/gi, (match, url) => {
    if (count >= maxImages) return match;
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
        const blob = res.getBlob();
        const b64 = Utilities.base64Encode(blob.getBytes());
        const mime = blob.getContentType() || 'image/png';
        count++;
        return match.replace(url, `data:${mime};base64,${b64}`);
      }
    } catch (e) { /* ignore */ }
    return match; // Fallback original
  });
}
