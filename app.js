(() => {
  'use strict';

  const cfg = window.APP_CONFIG || {};
  const $ = (id) => document.getElementById(id);
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
  const session = {
    get(key) { try { return sessionStorage.getItem('qc.' + key); } catch { return null; } },
    set(key, value) { try { sessionStorage.setItem('qc.' + key, value); } catch { /* geen opslag */ } },
    remove(key) { try { sessionStorage.removeItem('qc.' + key); } catch { /* geen opslag */ } }
  };

  // Map, werkadres en label horen bij het ingelogde account; elke gebruiker heeft zijn eigen instellingen.
  const settings = { vaultId: '', vaultPath: '', workEmail: '', tag: 'brain', dateInSubject: false, lang: 'nl-NL' };
  const userKey = (key) => key + '.' + (account ? account.homeAccountId : 'anon');
  function loadUserSettings() {
    settings.vaultId = store.get(userKey('vaultId'), '');       // id van de gekozen OneDrive-map
    settings.vaultPath = store.get(userKey('vaultPath'), '');   // leesbaar pad, alleen voor weergave
    settings.workEmail = store.get(userKey('workEmail'), '');
    settings.tag = store.get(userKey('tag'), 'brain');          // label vooraan in het mailonderwerp
    settings.dateInSubject = store.get(userKey('dateInSubject'), '') === '1';
  }

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
  }

  // ---------- inloggen bij Microsoft ----------
  const redirectUri = location.origin + location.pathname.replace(/[^/]*$/, '');
  let msalApp = null;
  let account = null;

  // Laat het inlogscherm of de app zien, nooit allebei.
  function renderAuth() {
    $('gate').hidden = !!account;
    $('app').hidden = !account;
    $('accountName').textContent = account ? account.username : '';
  }

  // Uitloggen uit de app, zodat je met een ander account kunt inloggen.
  // Je blijft wel aangemeld bij Microsoft zelf in de browser.
  async function logout() {
    stopListening();
    try { await msalApp.logoutRedirect({ account, onRedirectNavigate: () => false }); } catch { /* cache al leeg */ }
    account = null;
    msalApp.setActiveAccount(null);
    loadUserSettings();
    if ($('settings').open) $('settings').close();
    status('');
    $('gateLogin').disabled = false;
    gateMessage('');
    renderAuth();
  }

  function gateMessage(message) { $('gateMsg').textContent = message || ''; }

  async function initAuth() {
    if (!cfg.clientId || !window.msal) {
      renderAuth();
      $('gateLogin').disabled = true;
      gateMessage(!cfg.clientId
        ? 'De app is nog niet gekoppeld aan Microsoft (client-ID ontbreekt in config.js).'
        : 'De inlogbibliotheek kon niet worden geladen. Controleer je internetverbinding en herlaad.');
      return;
    }
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: cfg.clientId,
        authority: 'https://login.microsoftonline.com/consumers',
        redirectUri
      },
      // localStorage: je blijft ingelogd, ook nadat je de app sluit.
      cache: { cacheLocation: 'localStorage' }
    });
    if (typeof msalApp.initialize === 'function') await msalApp.initialize();

    let cameFromLogin = false;
    try {
      const result = await msalApp.handleRedirectPromise();
      cameFromLogin = !!result;
      account = (result && result.account) || msalApp.getActiveAccount() || msalApp.getAllAccounts()[0] || null;
    } catch (e) {
      gateMessage('Inloggen is niet gelukt: ' + (e.errorMessage || e.message));
      account = msalApp.getAllAccounts()[0] || null;
    }
    if (account) msalApp.setActiveAccount(account);
    loadUserSettings();
    renderAuth();
    if (!account) return;

    if (cameFromLogin) {
      session.remove('reauth');
      // Was je halverwege het verzenden naar de inlogpagina gestuurd? Dan maken we het nu af.
      const pending = store.get('pendingSend', '');
      store.set('pendingSend', '');
      if (pending) send();
    } else {
      store.set('pendingSend', '');
      refreshSession();
    }
  }

  function login() {
    gateMessage('');
    $('gateLogin').disabled = true;
    msalApp.loginRedirect({ scopes: SCOPES, prompt: 'select_account' })
      .catch((e) => { $('gateLogin').disabled = false; gateMessage('Inloggen is niet gelukt: ' + (e.errorMessage || e.message)); });
  }

  const NETWORK_ERRORS = ['no_network_connectivity', 'post_request_failed', 'endpoints_resolution_error', 'network_error'];
  const isNetworkError = (e) => !navigator.onLine || e instanceof TypeError || NETWORK_ERRORS.includes(e && e.errorCode);

  // Bij het openen vernieuwen we de sessie alvast op de achtergrond. Lukt dat niet stil, dan gaan we
  // één keer langs Microsoft; met een actieve Microsoft-sessie hoef je daar niets in te vullen.
  async function refreshSession() {
    if (!navigator.onLine) return;
    try {
      await msalApp.acquireTokenSilent({ scopes: SCOPES, account });
      session.remove('reauth');
    } catch (e) {
      if (isNetworkError(e) || session.get('reauth')) return;
      session.set('reauth', '1');
      saveDraft();
      msalApp.acquireTokenRedirect({ scopes: SCOPES, account }).catch(() => { /* later opnieuw bij verzenden */ });
    }
  }

  async function getToken(resumeSend) {
    try {
      const result = await msalApp.acquireTokenSilent({ scopes: SCOPES, account });
      return result.accessToken;
    } catch (e) {
      if (!isNetworkError(e)) {
        saveDraft();
        if (resumeSend) store.set('pendingSend', '1');
        status('Je sessie wordt vernieuwd. Je notitie blijft bewaard.');
        await msalApp.acquireTokenRedirect({ scopes: SCOPES, account });
      }
      throw e;
    }
  }

  // ---------- Microsoft Graph ----------
  async function graph(method, path, token, body, contentType) {
    const headers = { Authorization: 'Bearer ' + token };
    if (contentType) headers['Content-Type'] = contentType;
    const response = await fetch(path.startsWith('http') ? path : GRAPH + path, { method, headers, body });
    if (!response.ok) {
      let message = 'HTTP ' + response.status;
      let code = '';
      try { const j = await response.json(); if (j.error) { message = j.error.message || message; code = j.error.code || ''; } } catch { /* geen json */ }
      const err = new Error(message);
      err.code = code;
      err.status = response.status;
      throw err;
    }
    return response.status === 202 || response.status === 204 ? null : response.json();
  }

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

  function buildNote(text) {
    const d = new Date();
    const date = d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    // Onderwerp is optioneel: leeg gelaten, dan de eerste woorden van de notitie.
    const title = cleanTitle($('title').value, 80) || suggestTitle(text) || 'notitie';
    const name = date + ' ' + pad(d.getHours()) + pad(d.getMinutes()) + ' ' + title;
    const created = date + 'T' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
    const markdown = '---\ncreated: ' + created + '\nsource: snelle-notitie\ntags:\n  - inbox\n---\n\n' + text + '\n';
    return { name, title, markdown };
  }

  async function sendPrive(text, token) {
    const note = buildNote(text);
    const file = encodeURIComponent(note.name + '.md');
    try {
      await graph('PUT', '/me/drive/items/' + encodeURIComponent(settings.vaultId) + ':/' + file +
        ':/content?@microsoft.graph.conflictBehavior=rename', token, note.markdown, 'text/plain');
    } catch (e) {
      if (e.status === 404) throw new Error('De gekozen map bestaat niet meer in je OneDrive. Kies hem opnieuw bij Instellingen.');
      throw e;
    }
    return 'Opgeslagen in je privé-vault ✓';
  }

  const cleanTag = (value) => value.trim().toLowerCase().replace(/[^a-z0-9._-]/g, '');

  // De mail gaat vanaf het standaardadres van je Microsoft-account. Het label staat vooraan in het
  // onderwerp, zodat je er op je werk op kunt filteren.
  async function sendWerk(text, token) {
    const note = buildNote(text);
    const mail = {
      message: {
        subject: '[' + (settings.tag || 'brain') + '] ' + (settings.dateInSubject ? note.name : note.title),
        body: { contentType: 'Text', content: text },
        toRecipients: [{ emailAddress: { address: settings.workEmail } }]
      },
      saveToSentItems: true
    };
    await graph('POST', '/me/sendMail', token, JSON.stringify(mail), 'application/json');
    return 'Gemaild naar je werkadres ✓';
  }

  let sending = false;
  async function send() {
    if (sending || !account) return;
    const text = $('text').value.trim();
    if (!text) { status('Er staat nog geen tekst.', 'err'); return; }
    stopListening();
    if (dest === 'prive' && !settings.vaultId) {
      status('Kies eerst de map waar je privé-notities heen moeten.');
      openPicker();
      return;
    }
    if (dest === 'werk' && !settings.workEmail) {
      status('Vul eerst je werk-mailadres in.');
      openSettings('werk');
      $('workEmail').focus();
      return;
    }
    sending = true;
    $('send').disabled = true;
    status('Bezig met verzenden…');
    try {
      const token = await getToken(true);
      const message = dest === 'werk' ? await sendWerk(text, token) : await sendPrive(text, token);
      $('text').value = '';
      $('title').value = '';
      saveDraft();
      status(message, 'ok');
      if (navigator.vibrate) navigator.vibrate(40);
    } catch (e) {
      status(isNetworkError(e) ? 'Geen verbinding. Je notitie blijft bewaard; probeer het zo opnieuw.'
        : 'Niet gelukt: ' + (e.errorMessage || e.message), 'err');
    } finally {
      sending = false;
      $('send').disabled = false;
    }
  }

  // ---------- concept bewaren ----------
  function saveDraft() {
    store.set('draft', $('text').value);
    store.set('draftTitle', $('title').value);
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

  // ---------- mapkiezer (bladert door je OneDrive) ----------
  let pickStack = []; // [{ id, name }] vanaf de hoofdmap
  let pickLoad = 0;

  function renderVaultPath() {
    $('vaultPathView').textContent = settings.vaultPath || 'Nog geen map gekozen';
  }

  async function loadFolder() {
    const myLoad = ++pickLoad;
    const list = $('folderList');
    const current = pickStack[pickStack.length - 1];
    $('crumbs').textContent = ['OneDrive'].concat(pickStack.map((f) => f.name)).join(' › ');
    $('pickChoose').disabled = !current;
    list.innerHTML = '<div class="empty">Laden…</div>';
    try {
      const token = await getToken(false);
      let url = (current ? '/me/drive/items/' + encodeURIComponent(current.id) : '/me/drive/root') +
        '/children?$select=id,name,folder&$top=200';
      const folders = [];
      while (url) {
        const page = await graph('GET', url, token);
        folders.push(...page.value.filter((item) => item.folder));
        url = page['@odata.nextLink'] || null;
      }
      if (myLoad !== pickLoad) return;
      folders.sort((a, b) => a.name.localeCompare(b.name, 'nl', { sensitivity: 'base' }));
      list.textContent = '';
      if (current) {
        const up = document.createElement('button');
        up.textContent = '↑  Eén map omhoog';
        up.onclick = () => { pickStack.pop(); loadFolder(); };
        list.append(up);
      }
      for (const folder of folders) {
        const row = document.createElement('button');
        row.textContent = '📁  ' + folder.name;
        row.onclick = () => { pickStack.push({ id: folder.id, name: folder.name }); loadFolder(); };
        list.append(row);
      }
      if (!folders.length) {
        const empty = document.createElement('div');
        empty.className = 'empty';
        empty.textContent = 'Geen submappen. Je kunt deze map kiezen.';
        list.append(empty);
      }
    } catch (e) {
      if (myLoad !== pickLoad) return;
      list.innerHTML = '';
      const err = document.createElement('div');
      err.className = 'empty';
      err.textContent = isNetworkError(e) ? 'Geen verbinding met OneDrive.' : 'Mappen laden is niet gelukt: ' + (e.errorMessage || e.message);
      list.append(err);
    }
  }

  function openPicker() {
    pickStack = [];
    $('picker').showModal();
    loadFolder();
  }

  function choosePicked() {
    const current = pickStack[pickStack.length - 1];
    if (!current) return;
    settings.vaultId = current.id;
    settings.vaultPath = pickStack.map((f) => f.name).join('/');
    store.set(userKey('vaultId'), settings.vaultId);
    store.set(userKey('vaultPath'), settings.vaultPath);
    renderVaultPath();
    $('picker').close();
    status('Map gekozen: ' + settings.vaultPath, 'ok');
  }

  // ---------- instellingen ----------
  // Wijzigingen worden direct bewaard; er is geen aparte opslaan-knop.
  function showTab(tab) {
    tab = tab === 'werk' ? 'werk' : 'prive';
    $('settings').dataset.tab = tab;
    $('tabPrive').setAttribute('aria-selected', String(tab === 'prive'));
    $('tabWerk').setAttribute('aria-selected', String(tab === 'werk'));
    $('panelPrive').hidden = tab !== 'prive';
    $('panelWerk').hidden = tab !== 'werk';
    askLogout(false);
  }

  // Uitloggen gaat in twee stappen. De bevestiging staat links en "Nee" rechts, op de plek van de
  // eerste knop, zodat twee keer snel tikken je nooit uitlogt.
  function askLogout(asking) {
    $('accountRow').hidden = asking;
    $('logoutConfirm').hidden = !asking;
  }

  function openSettings(tab) {
    renderVaultPath();
    $('workEmail').value = settings.workEmail;
    $('tag').value = settings.tag;
    $('dateInSubject').checked = settings.dateInSubject;
    showTab(tab || dest);
    $('settings').showModal();
  }

  // ---------- start ----------
  $('gateLogin').onclick = login;
  $('destPrive').onclick = () => setDest('prive');
  $('destWerk').onclick = () => setDest('werk');
  $('mic').onclick = () => (listening ? stopListening() : startListening());
  $('send').onclick = send;
  $('openSettings').onclick = () => openSettings();
  $('closeSettings').onclick = () => $('settings').close();
  $('doneSettings').onclick = () => $('settings').close();
  $('tabPrive').onclick = () => showTab('prive');
  $('tabWerk').onclick = () => showTab('werk');
  $('workEmail').addEventListener('input', () => {
    settings.workEmail = $('workEmail').value.trim();
    store.set(userKey('workEmail'), settings.workEmail);
  });
  $('tag').addEventListener('input', () => {
    settings.tag = cleanTag($('tag').value);
    store.set(userKey('tag'), settings.tag);
  });
  $('tag').addEventListener('blur', () => { $('tag').value = settings.tag; });
  $('dateInSubject').addEventListener('change', () => {
    settings.dateInSubject = $('dateInSubject').checked;
    store.set(userKey('dateInSubject'), settings.dateInSubject ? '1' : '');
  });
  $('logoutAsk').onclick = () => askLogout(true);
  $('logoutNo').onclick = () => askLogout(false);
  $('logoutYes').onclick = logout;
  $('browse').onclick = openPicker;
  $('pickCancel').onclick = () => $('picker').close();
  $('pickChoose').onclick = choosePicked;
  $('text').addEventListener('input', saveDraft);
  $('title').addEventListener('input', saveDraft);
  document.addEventListener('visibilitychange', () => { if (document.hidden) stopListening(); });

  $('text').value = store.get('draft', '');
  $('title').value = store.get('draftTitle', '');
  setDest(new URLSearchParams(location.search).get('dest') || store.get('dest', 'prive'));
  initAuth();
})();
