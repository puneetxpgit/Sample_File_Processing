// n8n Code node: Run once for each item.
// Outlook forwards use the inner To line against towerRecipients and groundRecipients.
// Direct Gmail messages use the sender From address against towerSenders and groundSenders.

const config = $('Classification Config').first().json;
const email = $input.item.json;

function asList(value) {
  if (Array.isArray(value)) return value.map((item) => String(item).trim()).filter(Boolean);
  return String(value || '').split(/[,;\n]/).map((item) => item.trim()).filter(Boolean);
}

function normalize(address) {
  return String(address || '').trim().toLowerCase();
}

function extractEmails(value) {
  return [...String(value || '').matchAll(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi)].map((match) =>
    normalize(match[0])
  );
}

function htmlToText(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|h\d|li|table)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&quot;/gi, '"');
}

function bodyText(source) {
  const body = source.body || source.sampleBody;
  if (typeof body === 'string') return body;
  if (body && typeof body === 'object' && typeof body.content === 'string') {
    const type = String(body.contentType || '').toLowerCase();
    return type.includes('html') ? htmlToText(body.content) : body.content;
  }
  if (typeof source.text === 'string') return source.text;
  if (typeof source.bodyPreview === 'string') return source.bodyPreview;
  return '';
}

function parseBlocks(text) {
  const marker = text.search(/^-{2,}\s*Original Message\s*-{2,}|^Begin forwarded message/im);
  const region = marker >= 0 ? text.slice(marker) : text;
  const lines = region.split(/\r?\n/);
  const blocks = [];
  let current = null;
  const push = () => {
    if (current && (current.to.length || (current.from.length && current.dateRaw))) blocks.push(current);
  };
  for (let i = 0; i < lines.length; i += 1) {
    let line = lines[i];
    while (i + 1 < lines.length && /^[ \t]/.test(lines[i + 1])) {
      const next = lines[i + 1].trim();
      if (/^(from|sent|date|to|cc|bcc|subject)\s*:/i.test(next)) break;
      i += 1;
      line += ` ${next}`;
    }
    const match = line.match(/^(from|sent|date|to)\s*:\s*(.*)$/i);
    if (!match) continue;
    const key = match[1].toLowerCase();
    const value = match[2].trim();
    if (key === 'from') {
      push();
      current = { from: extractEmails(value), to: [], dateRaw: '' };
    } else if (!current) {
      continue;
    } else if (key === 'to') {
      current.to.push(...extractEmails(value));
    } else {
      current.dateRaw = value;
    }
  }
  push();
  return blocks;
}

function addressesFromValue(value) {
  if (value == null || value === '') return [];
  if (typeof value === 'string') return extractEmails(value);
  if (Array.isArray(value)) {
    const found = [];
    value.forEach((item) => found.push(...addressesFromValue(item)));
    return [...new Set(found)];
  }
  if (typeof value === 'object') {
    if (value.emailAddress) {
      return addressesFromValue(value.emailAddress.address || value.emailAddress.email || '');
    }
    for (const key of ['address', 'email', 'value']) {
      if (value[key]) return addressesFromValue(value[key]);
    }
  }
  return [];
}

function objectAddresses(source) {
  for (const key of ['from', 'sender', 'fromAddress']) {
    const found = addressesFromValue(source[key]);
    if (found.length) return found;
  }
  return [];
}

function scanHeaders(text) {
  const lines = String(text || '').split(/\r?\n/);
  const headers = [];
  for (let i = 0; i < lines.length; i += 1) {
    let line = lines[i];
    while (i + 1 < lines.length && /^[ \t]/.test(lines[i + 1])) {
      const next = lines[i + 1].trim();
      if (/^(from|sent|date|to|cc|bcc|subject)\s*:/i.test(next)) break;
      i += 1;
      line += ` ${next}`;
    }
    const match = line.match(/^(from|sent|date|to)\s*:\s*(.*)$/i);
    if (match) headers.push([match[1].toLowerCase(), match[2].trim()]);
  }
  return headers;
}

function isForward(text) {
  return /^-{2,}\s*Original Message\s*-{2,}|^Begin forwarded message/im.test(text || '');
}

function senderAddresses(source, text) {
  const found = objectAddresses(source);
  if (found.length) return found;
  const header = scanHeaders(text).find(([key]) => key === 'from');
  return header ? extractEmails(header[1]) : [];
}

function directTo(source, text) {
  for (const key of ['to', 'toAddress']) {
    const found = addressesFromValue(source[key]);
    if (found.length) return found.join('; ');
  }
  const header = scanHeaders(text).find(([key]) => key === 'to');
  return header ? extractEmails(header[1]).join('; ') : '';
}

function directDate(source, text) {
  for (const key of ['date', 'originalDate']) {
    if (typeof source[key] === 'string' && source[key].trim()) return source[key].trim();
  }
  const header = scanHeaders(text).find(([key, value]) => (key === 'sent' || key === 'date') && value);
  return header ? header[1] : '';
}

function review(reason) {
  return {
    json: {
      emailId: email.internetMessageId || email.id || email.emailId || '',
      category: 'Review',
      originalTo: '',
      originalFrom: '',
      originalDate: '',
      reason,
      status: 'review_required',
      body: '',
      bodyBase64: '',
      attachments: email.attachments || [],
      spreadsheetPath: email.spreadsheetPath || '',
      mopPaths: email.mopPaths || '',
    },
  };
}

function classifyForward(text, tower, ground) {
  const blocks = parseBlocks(text);
  const extracted = [...new Set(blocks.flatMap((block) => block.to))];
  const innermost = [...blocks].reverse().find((block) => block.from.length) || blocks[blocks.length - 1];
  const base = {
    emailId: email.internetMessageId || email.id || email.emailId || 'manual-sample',
    originalFrom: innermost?.from?.[0] || '',
    originalDate: innermost?.dateRaw || '',
    body: text,
    bodyBase64: Buffer.from(text, 'utf8').toString('base64'),
    attachments: email.attachments || [],
    subject: email.subject || '',
    spreadsheetPath: email.spreadsheetPath || '',
    mopPaths: email.mopPaths || '',
  };

  if (!extracted.length) {
    return {
      json: {
        ...base,
        category: 'Review',
        originalTo: '',
        reason: 'Original To line could not be parsed',
        status: 'review_required',
        needsAiToExtraction: true,
      },
    };
  }

  const towerHits = extracted.filter((address) => tower.has(address));
  const groundHits = extracted.filter((address) => ground.has(address));
  if (towerHits.length && groundHits.length) {
    return {
      json: {
        ...base,
        category: 'Review',
        originalTo: [...towerHits, ...groundHits].join('; '),
        reason: 'Forwarded To addresses match both Tower and Ground lists',
        status: 'review_required',
      },
    };
  }
  if (towerHits.length || groundHits.length) {
    const hits = towerHits.length ? towerHits : groundHits;
    const category = towerHits.length ? 'Tower' : 'Ground';
    return {
      json: {
        ...base,
        category,
        originalTo: hits.join('; '),
        reason: `Matched ${category} recipient ${hits.join(', ')}`,
        status: 'classified',
      },
    };
  }
  return {
    json: {
      ...base,
      category: 'Review',
      originalTo: extracted.join('; '),
      reason: `Unrecognized forwarded To address: ${extracted.join(', ')}`,
      status: 'review_required',
    },
  };
}

function classifySender(text, towerSenders, groundSenders) {
  const senders = senderAddresses(email, text);
  const base = {
    emailId: email.internetMessageId || email.id || email.emailId || 'manual-sample',
    originalFrom: senders.join('; '),
    originalDate: directDate(email, text),
    body: text,
    bodyBase64: Buffer.from(text, 'utf8').toString('base64'),
    attachments: email.attachments || [],
    subject: email.subject || '',
    spreadsheetPath: email.spreadsheetPath || '',
    mopPaths: email.mopPaths || '',
    toSource: 'from',
    originalTo: directTo(email, text),
  };
  if (!senders.length) {
    return {
      json: {
        ...base,
        category: 'Review',
        reason: 'Sender From address could not be read',
        status: 'review_required',
      },
    };
  }
  const towerHits = senders.filter((address) => towerSenders.has(address));
  const groundHits = senders.filter((address) => groundSenders.has(address));
  if (towerHits.length && groundHits.length) {
    return {
      json: {
        ...base,
        category: 'Review',
        reason: 'Sender addresses match both Tower and Ground lists',
        status: 'review_required',
      },
    };
  }
  if (towerHits.length || groundHits.length) {
    const hits = towerHits.length ? towerHits : groundHits;
    const category = towerHits.length ? 'Tower' : 'Ground';
    return {
      json: {
        ...base,
        originalFrom: hits.join('; '),
        category,
        reason: `Matched ${category} sender ${hits.join(', ')}`,
        status: 'classified',
      },
    };
  }
  return {
    json: {
      ...base,
      category: 'Review',
      reason: `Unrecognized sender: ${senders.join(', ')}`,
      status: 'review_required',
    },
  };
}

try {
  const tower = new Set(asList(config.towerRecipients).map(normalize));
  const ground = new Set(asList(config.groundRecipients).map(normalize));
  const towerSenders = new Set(asList(config.towerSenders).map(normalize));
  const groundSenders = new Set(asList(config.groundSenders).map(normalize));
  const configuredOnBoth = [
    ...new Set([
      ...[...tower].filter((address) => ground.has(address)),
      ...[...towerSenders].filter((address) => groundSenders.has(address)),
    ]),
  ];
  if (configuredOnBoth.length) {
    return review(`Address is configured on both lists: ${configuredOnBoth.join(', ')}`);
  }
  if (!tower.size && !ground.size && !towerSenders.size && !groundSenders.size) {
    return review('Tower and Ground recipient and sender lists are empty');
  }

  const text = bodyText(email);
  if (isForward(text)) {
    if (!text.trim()) return review('Email body is missing or empty');
    return classifyForward(text, tower, ground);
  }
  if (!text.trim() && !objectAddresses(email).length) return review('Email body is missing or empty');
  return classifySender(text, towerSenders, groundSenders);
} catch (error) {
  return review(`Malformed email body: ${error.message}`);
}
