// Pure logic for turning an email into a Stuff task row. No Office.js or DOM
// here, so tests/task.test.js can exercise it under plain Node.
(function (root) {
  const TITLE_MAX = 500; // tasks_title_len (0003_constraints.sql)
  const NOTES_MAX = 20000; // tasks_notes_len
  const BODY_MAX = 4000; // keep pasted email text readable, well under NOTES_MAX

  /** Local calendar date as YYYY-MM-DD (do_date / deadline are date-only). */
  function isoDate(d) {
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  }

  function addDays(d, n) {
    const out = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    out.setDate(out.getDate() + n);
    return out;
  }

  /** Classic OWA deep link; takes the EWS item id Office.js hands us. */
  function outlookLink(ewsItemId) {
    if (!ewsItemId) return null;
    return (
      'https://outlook.office.com/owa/?ItemID=' +
      encodeURIComponent(ewsItemId) +
      '&exvsurl=1&viewmodel=ReadMessageItem'
    );
  }

  function formatReceived(d) {
    if (!d) return null;
    return d.toLocaleString('en-US', {
      month: 'short', day: 'numeric', year: 'numeric',
      hour: 'numeric', minute: '2-digit',
    });
  }

  /**
   * email: { subject, fromName, fromEmail, received: Date, itemId, bodyText }
   * opts:  { title, doDate, deadline, includeSender, includeLink, includeBody }
   * Returns the insert payload for public.tasks. owner_id, sort_key, etc. are
   * filled by DB defaults, same as the Siri and MCP write paths.
   */
  function buildTask(email, opts) {
    const title = (opts.title || email.subject || '').trim() || '(no subject)';
    const lines = [];
    if (opts.includeSender && (email.fromName || email.fromEmail)) {
      const who = email.fromName && email.fromEmail
        ? `${email.fromName} <${email.fromEmail}>`
        : email.fromName || email.fromEmail;
      lines.push(`From: ${who}`);
    }
    if (opts.includeSender && email.received) {
      lines.push(`Received: ${formatReceived(email.received)}`);
    }
    const link = opts.includeLink ? outlookLink(email.itemId) : null;
    if (link) lines.push(`Open in Outlook: ${link}`);
    if (opts.includeBody && email.bodyText) {
      let body = email.bodyText.replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
      if (body.length > BODY_MAX) body = body.slice(0, BODY_MAX).trimEnd() + '…';
      if (body) lines.push(...(lines.length ? ['', body] : [body]));
    }
    const notes = lines.join('\n').slice(0, NOTES_MAX);

    const row = {
      title: title.slice(0, TITLE_MAX),
      // Mirrors the app's insert invariant: a dated task is 'scheduled' so it
      // leaves the Inbox instead of showing in both places.
      status: opts.doDate ? 'scheduled' : 'inbox',
    };
    if (notes) row.notes = notes;
    if (opts.doDate) row.do_date = opts.doDate;
    if (opts.deadline) row.deadline = opts.deadline;
    return row;
  }

  /** Human label for where the task landed. */
  function destination(row, today) {
    if (!row.do_date) return 'Inbox';
    const t = isoDate(today);
    if (row.do_date <= t) return 'Today';
    if (row.do_date === isoDate(addDays(today, 1))) return 'Tomorrow';
    return 'Upcoming';
  }

  const api = { isoDate, addDays, outlookLink, buildTask, destination, TITLE_MAX, BODY_MAX };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.StuffTask = api;
})(this);
