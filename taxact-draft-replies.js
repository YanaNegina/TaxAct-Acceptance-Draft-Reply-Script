function createTaxActAcceptanceDraftReplies() {
  const MAX_THREADS_PER_RUN = 100;

  const query = 'in:inbox "Your Requested" "TaxAct E-file Notice:" "Return Accepted"';
  const threads = GmailApp.search(query, 0, MAX_THREADS_PER_RUN);

  const myEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();

  let createdCount = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();

    // Skip this thread if it already has a later reply from you anywhere
    if (threadHasReplyFromMe_(messages, myEmail)) {
      continue;
    }

    // Find the first matching message in the thread
    for (let i = 0; i < messages.length; i++) {
      const msg = messages[i];
      const subject = msg.getSubject() || '';

      const parsed = parseTaxActSubject_(subject);
      if (!parsed) continue;

      const introText = buildDraftIntroText_(parsed);
      if (!introText) continue;

      const plainBody = buildPlainReplyWithOriginalMessage_(msg, introText);
      const htmlBody = buildHtmlReplyWithOriginalMessage_(msg, introText);

      msg.createDraftReply(plainBody, {
        htmlBody: htmlBody
      });

      createdCount++;
      break; // only one draft per thread
    }
  }

  Logger.log('Draft replies created: ' + createdCount);
}

function parseTaxActSubject_(subject) {
  const regex = /^Your Requested\s+(\d{4})\s+TaxAct E-file Notice:\s+(.*?)\s+(Extension\s+)?Return Accepted\b/i;
  const match = subject.match(regex);

  if (!match) return null;

  return {
    year: match[1],
    jurisdictionRaw: normalizeSpaces_(match[2]),
    isExtension: !!match[3]
  };
}

function buildDraftIntroText_(parsed) {
  const year = parsed.year;
  const jurisdictionRaw = parsed.jurisdictionRaw;
  const isExtension = parsed.isExtension;
  const isCurrentYear = year === '2025';

  // Federal
  if (/^federal$/i.test(jurisdictionRaw)) {
    const baseText = isExtension ? 'IRS extension accepted' : 'IRS accepted';
    const mainText = isCurrentYear ? baseText : `${baseText} (${year})`;

    // Payment text only for federal non-extension
    if (!isExtension) {
      return (
        `${mainText}. The fee is $ and you can pay via:\n` +
        `[payment methods]`
      );
    }

    return mainText;
  }

  // U.S. state
  const stateAbbr = getStateAbbreviation_(jurisdictionRaw);

  let baseText;
  if (stateAbbr) {
    baseText = isExtension
      ? `${stateAbbr} extension accepted`
      : `${stateAbbr} accepted`;
  } else {
    // Not a U.S. state and not Federal -> use exactly what is in subject
    baseText = isExtension
      ? `${jurisdictionRaw} extension accepted`
      : `${jurisdictionRaw} accepted`;
  }

  return isCurrentYear ? baseText : `${baseText} (${year})`;
}

function buildPlainReplyWithOriginalMessage_(msg, introText) {
  const from = msg.getFrom() || '';
  const date = msg.getDate();
  const subject = msg.getSubject() || '';
  const originalPlainBody = msg.getPlainBody() || '';

  return (
    introText +
    '\n\n' +
    '---------- Forwarded message ----------\n' +
    'From: ' + from + '\n' +
    'Date: ' + date + '\n' +
    'Subject: ' + subject + '\n\n' +
    originalPlainBody
  );
}

function buildHtmlReplyWithOriginalMessage_(msg, introText) {
  const from = escapeHtml_(msg.getFrom() || '');
  const date = escapeHtml_(String(msg.getDate() || ''));
  const subject = escapeHtml_(msg.getSubject() || '');
  const originalHtml = msg.getBody() || '';

  const introHtml = textToHtml_(introText);

  return (
    '<div>' + introHtml + '</div>' +
    '<br>' +
    '<div>---------- Forwarded message ----------</div>' +
    '<div><b>From:</b> ' + from + '</div>' +
    '<div><b>Date:</b> ' + date + '</div>' +
    '<div><b>Subject:</b> ' + subject + '</div>' +
    '<br>' +
    '<blockquote style="margin:0 0 0 0.8em;border-left:1px solid #ccc;padding-left:1em;">' +
    originalHtml +
    '</blockquote>'
  );
}

function textToHtml_(text) {
  return escapeHtml_(text).replace(/\n/g, '<br>');
}

function escapeHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getStateAbbreviation_(stateName) {
  const states = {
    'Alabama': 'AL',
    'Alaska': 'AK',
    'Arizona': 'AZ',
    'Arkansas': 'AR',
    'California': 'CA',
    'Colorado': 'CO',
    'Connecticut': 'CT',
    'Delaware': 'DE',
    'Florida': 'FL',
    'Georgia': 'GA',
    'Hawaii': 'HI',
    'Idaho': 'ID',
    'Illinois': 'IL',
    'Indiana': 'IN',
    'Iowa': 'IA',
    'Kansas': 'KS',
    'Kentucky': 'KY',
    'Louisiana': 'LA',
    'Maine': 'ME',
    'Maryland': 'MD',
    'Massachusetts': 'MA',
    'Michigan': 'MI',
    'Minnesota': 'MN',
    'Mississippi': 'MS',
    'Missouri': 'MO',
    'Montana': 'MT',
    'Nebraska': 'NE',
    'Nevada': 'NV',
    'New Hampshire': 'NH',
    'New Jersey': 'NJ',
    'New Mexico': 'NM',
    'New York': 'NY',
    'North Carolina': 'NC',
    'North Dakota': 'ND',
    'Ohio': 'OH',
    'Oklahoma': 'OK',
    'Oregon': 'OR',
    'Pennsylvania': 'PA',
    'Rhode Island': 'RI',
    'South Carolina': 'SC',
    'South Dakota': 'SD',
    'Tennessee': 'TN',
    'Texas': 'TX',
    'Utah': 'UT',
    'Vermont': 'VT',
    'Virginia': 'VA',
    'Washington': 'WA',
    'West Virginia': 'WV',
    'Wisconsin': 'WI',
    'Wyoming': 'WY',
    'District of Columbia': 'DC'
  };

  return states[normalizeSpaces_(stateName)] || null;
}

function normalizeSpaces_(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

function threadHasReplyFromMe_(messages, myEmail) {
  for (let i = 1; i < messages.length; i++) {
    const from = (messages[i].getFrom() || '').toLowerCase();
    if (myEmail && from.indexOf(myEmail) !== -1) {
      return true;
    }
  }
  return false;
}
