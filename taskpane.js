// Outlook task pane: email -> Stuff task. Writes straight to Supabase as the
// signed-in user (RLS fills owner_id); PowerSync fans it out to every device.
(function () {
  const $ = (id) => document.getElementById(id);
  const cfg = window.STUFF_CONFIG;
  const T = window.StuffTask;

  const sb = window.supabase.createClient(cfg.supabaseUrl, cfg.supabaseKey, {
    auth: { persistSession: true, autoRefreshToken: true, storageKey: 'stuff-outlook-auth' },
  });

  let email = null; // the email currently shown in Outlook
  let when = 'none';
  let lastTaskId = null;

  function show(view) {
    for (const id of ['boot', 'signin', 'compose', 'done']) $(id).hidden = id !== view;
  }

  function setError(el, msg) {
    el.textContent = msg || '';
    el.hidden = !msg;
  }

  // ---- Reading the email -------------------------------------------------

  function readItem() {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) return null;
    return {
      subject: item.subject || '',
      fromName: item.from ? item.from.displayName : '',
      fromEmail: item.from ? item.from.emailAddress : '',
      received: item.dateTimeCreated || null,
      itemId: item.itemId || null,
      bodyText: null, // fetched lazily, only if "Email text" is ticked
      _item: item,
    };
  }

  function readBody(e) {
    return new Promise((resolve) => {
      if (!e._item || !e._item.body) return resolve(e.bodyText || '');
      e._item.body.getAsync(Office.CoercionType.Text, (r) =>
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ''));
    });
  }

  // Opened in a plain browser (local preview), not inside Outlook.
  function previewItem() {
    return {
      subject: 'Re: Budget draft for Q1 — need your edits by Friday',
      fromName: 'Pat Example', fromEmail: 'pat@example.com',
      received: new Date(), itemId: 'AAMkPREVIEW=', bodyText: 'Preview body text.',
    };
  }

  // ---- Compose -------------------------------------------------------------

  function setWhen(value) {
    when = value;
    for (const b of $('when').querySelectorAll('.chip')) b.classList.toggle('on', b.dataset.when === value);
    $('do-date').hidden = value !== 'pick';
    if (value === 'pick' && !$('do-date').value) $('do-date').value = T.isoDate(T.addDays(new Date(), 2));
  }

  function doDate() {
    const now = new Date();
    if (when === 'today') return T.isoDate(now);
    if (when === 'tomorrow') return T.isoDate(T.addDays(now, 1));
    if (when === 'pick') return $('do-date').value || null;
    return null;
  }

  function resetCompose() {
    $('title').value = (email && email.subject) || '';
    $('deadline').value = '';
    setWhen('none');
    setError($('save-error'), '');
    $('save').disabled = !email;
    show('compose');
    $('title').focus();
    $('title').select();
  }

  async function save(ev) {
    ev.preventDefault();
    if (!email) return;
    $('save').disabled = true;
    setError($('save-error'), '');
    try {
      const includeBody = $('inc-body').checked;
      if (includeBody && email.bodyText == null) email.bodyText = await readBody(email);
      const row = T.buildTask(email, {
        title: $('title').value,
        doDate: doDate(),
        deadline: $('deadline').value || null,
        includeLink: $('inc-link').checked,
        includeSender: $('inc-sender').checked,
        includeBody,
      });
      const { data, error } = await sb.from('tasks').insert(row).select('id').single();
      if (error) throw error;
      lastTaskId = data.id;
      $('done-where').textContent = T.destination(row, new Date());
      show('done');
    } catch (e) {
      setError($('save-error'), friendly(e));
      $('save').disabled = false;
    }
  }

  async function undo() {
    if (!lastTaskId) return;
    $('undo').disabled = true;
    const { error } = await sb.from('tasks').delete().eq('id', lastTaskId);
    $('undo').disabled = false;
    if (error) return alertInline(friendly(error));
    lastTaskId = null;
    resetCompose();
  }

  function alertInline(msg) {
    resetCompose();
    setError($('save-error'), msg);
  }

  function friendly(e) {
    const msg = (e && (e.message || e.error_description)) || String(e);
    if (/fetch|network/i.test(msg)) return "Couldn't reach Stuff. Check your connection and try again.";
    if (/JWT|token|auth/i.test(msg)) return 'Your sign-in expired. Sign out and back in.';
    return msg;
  }

  // ---- Auth ------------------------------------------------------------------

  async function signIn(ev) {
    ev.preventDefault();
    const btn = $('signin').querySelector('button');
    btn.disabled = true;
    setError($('signin-error'), '');
    const { error } = await sb.auth.signInWithPassword({
      email: $('email').value.trim(),
      password: $('password').value,
    });
    btn.disabled = false;
    $('password').value = '';
    if (error) return setError($('signin-error'), error.message);
    await render();
  }

  async function signOut() {
    await sb.auth.signOut();
    await render();
  }

  async function render() {
    const { data } = await sb.auth.getSession();
    const user = data.session && data.session.user;
    // Never run on an anonymous session (same rule as the app).
    if (!user || user.is_anonymous) {
      $('account').hidden = true;
      show('signin');
      return;
    }
    $('who').textContent = user.email || '';
    $('account').hidden = false;
    resetCompose();
  }

  // ---- Boot --------------------------------------------------------------------

  function wire() {
    $('signin').addEventListener('submit', signIn);
    $('compose').addEventListener('submit', save);
    $('signout').addEventListener('click', signOut);
    $('undo').addEventListener('click', undo);
    $('again').addEventListener('click', resetCompose);
    $('when').addEventListener('click', (e) => {
      const b = e.target.closest('.chip');
      if (b) setWhen(b.dataset.when);
    });
    // ⌘/Ctrl+Enter saves from the title box.
    $('title').addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) $('compose').requestSubmit();
    });
  }

  Office.onReady(async (info) => {
    wire();
    if (info.host === Office.HostType.Outlook) {
      email = readItem();
      // Pinned pane: follow the selection as you move through the inbox.
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, async () => {
        email = readItem();
        lastTaskId = null;
        if (!$('compose').hidden || !$('done').hidden) resetCompose();
      });
    } else {
      email = previewItem();
    }
    await render();
  });
})();
