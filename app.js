(() => {
  'use strict';

  const cfg = window.APP_CONFIG || {};
  const $ = (id) => document.getElementById(id);
  const TAG = '[brain]';
  const SCOPES = ['Files.ReadWrite', 'Mail.Send'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';

  // ---------- opslag op het toestel ----------
  const store = {
    get(key, fallback) {
      try { const v = localStorage.getItem('qc.' + key); return v === null ? fallback : v; }
      catch { return fallback; }
    },
    set(key, value) { try { localStorage.setItem('qc.' + key, value); } catch { /* geen opslag */ } }
  };

  const settings = {
    vaultPath: store.get('vaultPath', 'Desktop/Obsidian Vault/Vault/Inbox'),
    workEmail: store.get('workEmail', ''),
    lang: store.get('lang', 'nl-NL')
  };

  // ---------- status ----------
  function status(message, kind) {
    const el = $('status');
    el.textContent = message || '';
    el.className = kind || '';
  }

  // ---------- bestemming ----------
  let dest = 'prive';
  function setDest(value) {
    dest = value === 'werk' ? 'werk' : 'prive';
    document.body.dataset.dest = dest;
    $('destPrive').setAttribute('aria-pressed', String(dest === 'prive'));
    $('destWerk').setAttribute('aria-pressed', String(dest === 'werk'));
    $('send').textContent = dest === 'werk' ? 'Verzend naar Werk' : 'Verzend naar Privé';
    store.set('dest', dest);
    renderAccount();
  }

  // ---------- inloggen bij Microsoft ----------
  const redirectUri = location.origin + location.pathname.replace(/[^/]*$/, '');
  let msalApp = null;
  let account = null;

  async function initAuth() {
    if (!cfg.clientId) {
      $('setupBanner').hidden = false;
      $('setupBanner').textContent = 'De app is nog niet gekoppeld aan Microsoft. Vul eerst de client-ID in config.js in.';
      return;
    }
    if (!window.msal) {
      $('setupBanner').hidden = false;
      $('setupBanner').textContent = 'De inlogbibliotheek kon niet worden geladen. Controleer je internetverbinding en herlaad.';
      return;
    }
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: cfg.clientId,
        authority: 'https://login.microsoftonline.com/consumers',
        redirectUri
      },
      cache: { cacheLocation: 'localStorage' }
    });
    if (typeof msalApp.initialize === 'function') await msalApp.initialize();
    let cameFromLogin = false;
    try {
      const result = await msalApp.handleRedirectPromise();
      cameFromLogin = !!result;
      account = (result && result.account) || msalApp.getAllAccounts()[0] || null;
    } catch (e) {
      status('Inloggen mislukt: ' + (e.errorMessage || e.message), 'err');
    }
    renderAccount();
    // Was je halverwege het verzenden naar de inlogpagina gestuurd? Dan maken we het nu af.
    const pending = store.get('pendingSend', '');
    store.set('pendingSend', '');
    if (pending && cameFromLogin && account) send();
  }

  function renderAccount() {
    // Inloggen is alleen verplicht voor Privé; Werk kan zonder (via je mail-app).
    $('loginBanner').hidden = !msalApp || !!account || dest !== 'prive';
    const el = $('account');
    el.textContent = '';
    if (!msalApp) return;
    if (account) {
      el.append('Ingelogd als ' + account.username + '. ');
      const out = document.createElement('button');
      out.textContent = 'Uitloggen';
      out.onclick = () => msalApp.logoutRedirect({ account });
      el.append(out);
    } else {
      el.textContent = 'Niet ingelogd.';
    }
  }

  function login() {
    saveDraft();
    return msalApp.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
  }

  async function getToken() {
    try {
      const result = await msalApp.acquireTokenSilent({ scopes: SCOPES, account });
      return result.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        saveDraft();
        store.set('pendingSend', '1');
        status('Je moet opnieuw inloggen. Je notitie blijft bewaard.');
        await msalApp.acquireTokenRedirect({ scopes: SCOPES, account });
      }
      throw e;
    }
  }

  // ---------- Microsoft Graph ----------
  async function graph(method, path, token, body, contentType) {
    const headers = { Authorization: 'Bearer ' + token };
    if (contentType) headers['Content-Type'] = contentType;
    const response = await fetch(GRAPH + path, { method, headers, body });
    if (!response.ok) {
      let message = 'HTTP ' + response.status;
      try { const j = await response.json(); if (j.error && j.error.message) message = j.error.message; } catch { /* geen json */ }
      const err = new Error(message);
      err.status = response.status;
      throw err;
    }
    return response.status === 202 || response.status === 204 ? null : response.json();
  }

  const encodePath = (p) => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

  // ---------- notitie opbouwen ----------
  const pad = (n) => String(n).padStart(2, '0');

  // Maakt een tekst geschikt als bestandsnaam en mailonderwerp.
  function cleanTitle(value, max) {
    return value.replace(/[\\/:*?"<>|#^[\]%&{}~]/g, ' ').replace(/\s+/g, ' ').trim()
      .slice(0, max).replace(/[.,;\s]+$/, '').trim();
  }

  function suggestTitle(text) {
    return cleanTitle(text.replace(/\s+/g, ' ').trim().split(' ').slice(0, 6).join(' '), 50);
  }

  // Het onderwerp volgt je tekst, totdat je het zelf aanpast. Maak je het leeg, dan volgt het weer.
  let titleEdited = false;
  function updateTitleSuggestion() {
    if (!titleEdited) $('title').value = suggestTitle($('text').value);
  }

  function buildNote(text) {
    const d = new Date();
    const date = d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    const title = cleanTitle($('title').value, 80) || suggestTitle(text) || 'notitie';
    const name = date + ' ' + pad(d.getHours()) + pad(d.getMinutes()) + ' ' + title;
    const created = date + 'T' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
    const markdown = '---\ncreated: ' + created + '\nsource: snelle-notitie\ntags:\n  - inbox\n---\n\n' + text + '\n';
    return { name, markdown };
  }

  async function sendPrive(text, token) {
    const folder = encodePath(settings.vaultPath);
    try {
      await graph('GET', '/me/drive/root:/' + folder + '?$select=id,folder', token);
    } catch (e) {
      if (e.status === 404) throw new Error('Map "' + settings.vaultPath + '" bestaat niet in je OneDrive. Pas het pad aan bij Instellingen.');
      throw e;
    }
    const note = buildNote(text);
    const file = encodeURIComponent(note.name + '.md');
    await graph('PUT', '/me/drive/root:/' + folder + '/' + file + ':/content?@microsoft.graph.conflictBehavior=rename',
      token, note.markdown, 'text/plain');
    return 'Opgeslagen in je privé-vault ✓';
  }

  async function sendWerk(text, token) {
    if (!settings.workEmail) throw new Error('Vul eerst je werk-mailadres in bij Instellingen.');
    const note = buildNote(text);
    const mail = {
      message: {
        subject: TAG + ' ' + note.name,
        body: { contentType: 'Text', content: text },
        toRecipients: [{ emailAddress: { address: settings.workEmail } }]
      },
      saveToSentItems: true
    };
    await graph('POST', '/me/sendMail', token, JSON.stringify(mail), 'application/json');
    return 'Gemaild naar je werkadres ✓';
  }

  // Zonder Microsoft-login: open de mail-app met een vooraf ingevuld bericht.
  function sendWerkViaMailApp(text) {
    if (!settings.workEmail) { status('Vul eerst je werk-mailadres in bij Instellingen.', 'err'); return; }
    const note = buildNote(text);
    location.href = 'mailto:' + encodeURIComponent(settings.workEmail) +
      '?subject=' + encodeURIComponent(TAG + ' ' + note.name) +
      '&body=' + encodeURIComponent(text);
    status('Je mail-app is geopend. Tik daar op verzenden; daarna kun je de tekst hier wissen.');
  }

  let sending = false;
  async function send() {
    if (sending) return;
    const text = $('text').value.trim();
    if (!text) { status('Er staat nog geen tekst.', 'err'); return; }
    stopListening();
    if (dest === 'werk' && !account) { sendWerkViaMailApp(text); return; }
    if (!msalApp) { status('De app is nog niet gekoppeld aan Microsoft.', 'err'); return; }
    if (!account) {
      // Privé kan alleen met een Microsoft-login: stuur direct door en verzend daarna vanzelf.
      store.set('pendingSend', '1');
      status('Je moet eerst inloggen bij Microsoft. Je notitie blijft bewaard.');
      login();
      return;
    }
    sending = true;
    $('send').disabled = true;
    status('Bezig met verzenden…');
    try {
      const token = await getToken();
      const message = dest === 'werk' ? await sendWerk(text, token) : await sendPrive(text, token);
      $('text').value = '';
      $('title').value = '';
      titleEdited = false;
      saveDraft();
      status(message, 'ok');
      if (navigator.vibrate) navigator.vibrate(40);
    } catch (e) {
      const offline = !navigator.onLine || e instanceof TypeError;
      status(offline ? 'Geen verbinding. Je notitie blijft bewaard; probeer het zo opnieuw.' : 'Niet gelukt: ' + (e.errorMessage || e.message), 'err');
    } finally {
      sending = false;
      $('send').disabled = false;
    }
  }

  // ---------- concept bewaren ----------
  function saveDraft() {
    store.set('draft', $('text').value);
    store.set('draftTitle', titleEdited ? $('title').value : '');
  }

  // ---------- spraak naar tekst ----------
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  let recognition = null;
  let listening = false;

  function appendSpoken(fragment) {
    fragment = fragment.trim();
    if (!fragment) return;
    const area = $('text');
    const current = area.value;
    if (!current.trim() || /[.!?]\s*$/.test(current)) fragment = fragment[0].toUpperCase() + fragment.slice(1);
    area.value = current + (current && !/\s$/.test(current) ? ' ' : '') + fragment;
    area.scrollTop = area.scrollHeight;
    updateTitleSuggestion();
    saveDraft();
  }

  function startRecognition() {
    recognition = new SR();
    recognition.lang = settings.lang;
    recognition.continuous = false; // op Android betrouwbaarder; we starten zelf opnieuw na elke zin
    recognition.interimResults = true;
    recognition.onresult = (event) => {
      let interim = '';
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const result = event.results[i];
        if (result.isFinal) appendSpoken(result[0].transcript);
        else interim += result[0].transcript;
      }
      $('interim').textContent = interim;
    };
    recognition.onerror = (event) => {
      if (event.error === 'no-speech' || event.error === 'aborted') return;
      listening = false;
      const messages = {
        'not-allowed': 'Geen toegang tot de microfoon. Sta die toe in de site-instellingen van Chrome.',
        'service-not-allowed': 'Spraakherkenning is op dit toestel niet beschikbaar.',
        'network': 'Spraakherkenning heeft een internetverbinding nodig.'
      };
      status(messages[event.error] || 'Spraakherkenning gestopt: ' + event.error, 'err');
    };
    recognition.onend = () => {
      $('interim').textContent = '';
      if (listening) setTimeout(() => { if (listening) startRecognition(); }, 150);
      else renderMic();
    };
    try { recognition.start(); } catch { /* al gestart */ }
  }

  function renderMic() {
    $('mic').setAttribute('aria-pressed', String(listening));
    $('mic').textContent = listening ? '■' : '🎤';
  }

  function startListening() {
    if (!SR) { status('Deze browser ondersteunt geen spraakherkenning. Gebruik Chrome.', 'err'); return; }
    listening = true;
    status('Luistert…');
    renderMic();
    startRecognition();
  }

  function stopListening() {
    if (!listening) return;
    listening = false;
    try { recognition.stop(); } catch { /* al gestopt */ }
    status('');
    renderMic();
  }

  // ---------- instellingen ----------
  function openSettings() {
    $('vaultPath').value = settings.vaultPath;
    $('workEmail').value = settings.workEmail;
    $('lang').value = settings.lang;
    renderAccount();
    $('settings').showModal();
  }

  function saveSettings() {
    settings.vaultPath = $('vaultPath').value.trim().replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
    settings.workEmail = $('workEmail').value.trim();
    settings.lang = $('lang').value;
    store.set('vaultPath', settings.vaultPath);
    store.set('workEmail', settings.workEmail);
    store.set('lang', settings.lang);
    $('settings').close();
    status('Instellingen opgeslagen.', 'ok');
  }

  // ---------- start ----------
  $('destPrive').onclick = () => setDest('prive');
  $('destWerk').onclick = () => setDest('werk');
  $('mic').onclick = () => (listening ? stopListening() : startListening());
  $('send').onclick = send;
  $('login').onclick = login;
  $('openSettings').onclick = openSettings;
  $('cancelSettings').onclick = () => $('settings').close();
  $('saveSettings').onclick = saveSettings;
  $('text').addEventListener('input', () => { updateTitleSuggestion(); saveDraft(); });
  $('title').addEventListener('input', () => {
    titleEdited = $('title').value.trim() !== '';
    saveDraft();
  });
  $('title').addEventListener('blur', updateTitleSuggestion);
  document.addEventListener('visibilitychange', () => { if (document.hidden) stopListening(); });

  $('text').value = store.get('draft', '');
  const savedTitle = store.get('draftTitle', '');
  if (savedTitle) { $('title').value = savedTitle; titleEdited = true; }
  updateTitleSuggestion();
  setDest(new URLSearchParams(location.search).get('dest') || store.get('dest', 'prive'));
  initAuth();
})();
