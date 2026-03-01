# TaxAct-Acceptance-Draft-Reply-Script
A Google Apps Script for Gmail that helps accountants speed up client communication by automatically creating draft replies for TaxAct “Return Accepted” notifications.

# TaxAct Acceptance Draft Reply Script

A Google Apps Script for Gmail that helps accountants speed up client communication by automatically creating draft replies for TaxAct “Return Accepted” notifications.

## What this script does

This script searches your Gmail inbox for TaxAct e-file acceptance emails and creates ready-to-send draft replies. It is designed for accountants, tax preparers, and bookkeeping assistants who want to save time when forwarding acceptance notifications to clients.

Instead of manually opening each TaxAct message, writing a reply, and copying the original content, the script does the repetitive part for you.

## Main features

- Scans Gmail for TaxAct “Return Accepted” emails
- Creates a draft reply directly in the email thread
- Detects whether the return is:
  - Federal
  - State
  - Extension
- Uses short, standardized wording such as:
  - `IRS accepted`
  - `CA accepted`
  - `IRS extension accepted`
- Adds payment instructions for federal non-extension acceptances
- Skips threads that were already processed
- Skips threads where a later reply from your email already exists
- Labels processed threads to prevent duplicates

## Who this is for

This script is especially helpful for:

- Accountants
- Tax preparers
- Bookkeepers
- Admin assistants handling tax workflow
- Anyone managing a large number of TaxAct acceptance notifications in Gmail

If you regularly forward return acceptance notices to clients, this script can significantly reduce manual work.

## How it works

The script:

1. Searches Gmail for TaxAct acceptance emails in your inbox
2. Looks for subjects matching the TaxAct acceptance pattern
3. Extracts:
   - tax year
   - jurisdiction
   - whether it is an extension
4. Builds a draft reply with a short intro message
5. Includes the original message below
6. Applies a Gmail label to mark the thread as processed

## Example output

For a subject like:

`Your Requested 2025 TaxAct E-file Notice: Federal Return Accepted`

the draft reply will begin with:

```text
IRS accepted. The fee is $ and you can pay via:
[payment methods]
```
For a state return like:

`Your Requested 2025 TaxAct E-file Notice: California Return Accepted`

the draft reply will begin with:

```text
CA accepted
```

## Requirements

- A Google account
- Gmail
- Google Apps Script access

## Installation

1. Open [Google Apps Script](https://script.google.com/)
2. Create a new project
3. Paste the script into the editor
4. Save the project
5. Run it
6. Authorize the required Gmail permissions when prompted

## Configuration

You can adjust the following values in the script:

- `MAX_THREADS_PER_RUN`  
  Controls how many Gmail threads are processed in one run

- `PROCESSED_LABEL`  
  The Gmail label used to mark processed threads

You can also customize:

- payment instructions
- reply wording
- subject parsing rules
- state abbreviations
- label name

## Notes

- The script creates **draft replies**, not sent emails
- It processes only threads that are not already labeled as processed
- It creates only one draft per thread
- It checks whether you have already replied later in the thread and skips those cases
- Original email content is included in plain text
