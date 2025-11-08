# Gmail Invoice Archiver → Google Drive (Vendor-Sorted)

Automated invoice extraction from Gmail into structured Google Drive folders — **no SaaS, no subscriptions, private, fast.**

## What it does

- 📨 Detects invoice emails in Gmail  
- 📎 Saves PDF attachments (Telekom, Vodafone, PayPal, etc.)
- 🧾 Converts invoice emails **without attachments** to PDF (Apple, Stripe, etc.)
- 🗂️ Auto-creates vendor folders (amazon, apple, telekom, etc.)
- 🔁 Duplicate-safe (idempotent)
- ⏱️ Processes in batches (resumes automatically)
- 📅 Weekly digest email with vendor summary
- 🔖 Uses Gmail labels for processing & reprocessing

> Built to replace third-party “invoice extraction” SaaS tools.
> Runs entirely in your Google Workspace account.  
> Privacy by design.

## Why

I got tired of:
- digging invoices out of Gmail
- forwarding to accounting mailboxes
- SaaS tools taking mailbox access + monthly fees

So I built a clean, maintainable automation that behaves like a private assistant.

## Features

| Feature | Status |
|---|---|
PDF attachment extraction | ✅  
HTML-mail → PDF conversion | ✅  
Vendor auto-foldering | ✅  
Label-based dedupe | ✅  
Pagination (no quota crashes) | ✅  
Weekly digest | ✅  
Reprocess function | ✅  
Google Sheets log | 🔜 (optional)  
UI button panel | 🔜  

## Setup (5 minutes)

### 1) Create a Drive folder  
Example ID: `1SJnBUuUGyX9i16latsggxIhC4Md5hFCd`

### 2) Go to  
https://script.google.com → **New project**

### 3) Paste the script from `script.gs`, edit:

```js
const ROOT_FOLDER_ID = 'YOUR_FOLDER_ID';
