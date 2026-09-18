// Match the server's canonicalization, while preserving users' original text.
export function urlKey(value) {
  try {
    const u = new URL(value.trim());
    return `${u.protocol}//${u.host}${u.pathname.replace(/\/+$/, '')}`;
  } catch { return value.trim(); }
}

export function existingUrls(text) {
  return new Set(text.split(/\r?\n/).map(s => s.trim())
    .filter(s => s && !s.startsWith('#')).map(urlKey));
}

export function appendIssues(text, urls) {
  const seen = existingUrls(text);
  const added = [];
  for (const url of urls) {
    const key = urlKey(url);
    if (!seen.has(key)) { seen.add(key); added.push(key); }
  }
  return {
    text: added.length ? text + (text && !text.endsWith('\n') ? '\n' : '') + added.join('\n') + '\n' : text,
    added: added.length,
    skipped: urls.length - added.length,
  };
}
