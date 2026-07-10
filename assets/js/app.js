(() => {
  'use strict';

  const cfg = window.APP_CONFIG;
  let msalApp;
  let api;
  let account = null;
  let me = null;
  let options = { teams: [], players: [] };

  const $ = (id) => document.getElementById(id);
  const qsa = (selector) => [...document.querySelectorAll(selector)];

  document.addEventListener('DOMContentLoaded', () => {
    boot().catch(error => {
      console.error('Application boot failed', error);
      const toast = document.getElementById('toast');
      if (toast) {
        toast.textContent = `Startup error: ${error.message}`;
        toast.classList.add('show', 'error');
      }
    });
  });

  async function boot() {
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: cfg.entraClientId, authority: cfg.entraAuthority, redirectUri: cfg.redirectUri },
      cache: { cacheLocation: cfg.cacheLocation, storeAuthStateInCookie: false },
      system: { allowNativeBroker: false }
    });
    await msalApp.initialize();
    const redirect = await msalApp.handleRedirectPromise();
    account = redirect?.account || msalApp.getAllAccounts()[0] || null;
    if (account) msalApp.setActiveAccount(account);
    api = new Api(cfg.apiBaseUrl, getToken);
    bindUi();
    await loadPublicConfig();
    updateAuthUi();
    if (account) await loadAuthenticatedState();
    const params = new URLSearchParams(location.search);
    if (params.get('payment') === 'success') toast('Payment completed. Stripe is confirming it now.');
    if (params.has('payment')) history.replaceState({}, '', cfg.redirectUri);
  }

  function bindUi() {
    $('signInButton').addEventListener('click', signIn);
    $('heroSignIn').addEventListener('click', () => account ? showView('predictions') : signIn());
    $('signOutButton').addEventListener('click', () => msalApp.logoutRedirect({ account, postLogoutRedirectUri: cfg.redirectUri }));
    qsa('[data-view]').forEach(el => el.addEventListener('click', () => showView(el.dataset.view)));
    $('payButton').addEventListener('click', startPayment);
    $('predictionForm').addEventListener('submit', savePredictions);
    $('refreshLeaderboard').addEventListener('click', loadLeaderboard);
    $('scheduleForm').addEventListener('submit', saveSchedule);
    $('syncFootballButton').addEventListener('click', syncFootballData);
    $('goldenBootOptionsForm').addEventListener('submit', e => saveAwardOptions(e, 'golden_boot_player'));
    $('goldenGloveOptionsForm').addEventListener('submit', e => saveAwardOptions(e, 'golden_glove_player'));
    $('resultsForm').addEventListener('submit', saveResults);
  }

  async function signIn() {
    await msalApp.loginRedirect({ scopes: ['openid', 'profile', 'email', cfg.apiScope], prompt: 'select_account' });
  }

  async function getToken() {
    account = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
    if (!account) throw new Error('Sign in required.');
    try {
      const result = await msalApp.acquireTokenSilent({ account, scopes: [cfg.apiScope] });
      return result.accessToken;
    } catch {
      const result = await msalApp.acquireTokenPopup({ account, scopes: [cfg.apiScope] });
      return result.accessToken;
    }
  }

  async function loadPublicConfig() {
    try {
      const res = await fetch(`${cfg.apiBaseUrl}/api/public/config`);
      const data = await res.json();
      $('competitionName').textContent = data.name || 'Sofaking Predictions';
      $('competitionSeason').textContent = data.season || '2026/2027';
      $('competitionStatus').textContent = String(data.status || 'setup').toUpperCase();
      $('deadlineText').textContent = data.submissionDeadlineAt ? formatDate(data.submissionDeadlineAt) : 'Waiting for fixture release';
    } catch { $('deadlineText').textContent = 'API unavailable'; }
  }

  function updateAuthUi() {
    const signedIn = Boolean(account);
    $('signInButton').hidden = signedIn;
    $('signOutButton').hidden = !signedIn;
    $('accountName').textContent = signedIn ? account.name || account.username : '';
    qsa('.auth-only').forEach(el => el.hidden = !signedIn);
  }

  async function loadAuthenticatedState() {
    try {
      // Authentication is established by MSAL before any football-data request.
      // A provider/database sync failure must never make the user appear signed out.
      me = await api.get('/api/me');
      qsa('.admin-only').forEach(el => el.hidden = !me.user.isAdmin);
      renderEntryState();

      try {
        await loadOptions();
        await loadMyPredictions();
      } catch (error) {
        console.error('Prediction options failed to load', error);
        toast(`Signed in, but football options failed: ${error.message}`, true);
      }

      await Promise.allSettled([loadLeaderboard(), loadLive()]);
      if (me.user.isAdmin) {
        try {
          await loadAdmin();
        } catch (error) {
          console.error('Admin data failed to load', error);
          toast(`Signed in, but admin data failed: ${error.message}`, true);
        }
      }
    } catch (error) {
      console.error('Authenticated profile failed to load', error);
      toast(error.message, true);
    }
  }

  function showView(name) {
    if (!account && name !== 'home') { signIn(); return; }
    qsa('.view').forEach(v => v.classList.remove('active'));
    qsa('.nav-button').forEach(v => v.classList.toggle('active', v.dataset.view === name));
    const view = $(`view-${name}`) || $('view-home');
    view.classList.add('active');
    if (name === 'leaderboard') loadLeaderboard();
    if (name === 'live') loadLive();
    scrollTo({ top: 0, behavior: 'smooth' });
  }

  async function loadOptions() {
    options = await api.get('/api/options');
    renderPredictionFields();
  }

  function renderPredictionFields() {
    const top4 = ['1st', '2nd', '3rd', '4th'];
    const bottom3 = ['18th', '19th', '20th'];
    $('top4Fields').innerHTML = top4.map((label, i) => selectHtml(`top4_${i + 1}`, label, options.teams)).join('');
    $('bottom3Fields').innerHTML = bottom3.map((label, i) => selectHtml(`bottom3_${i + 1}`, label, options.teams)).join('');
    fillSelect(document.querySelector('[name="fa_cup"]'), options.englishCupTeams || options.teams);
    fillSelect(document.querySelector('[name="league_cup"]'), options.englishCupTeams || options.teams);
    fillSelect(document.querySelector('[name="champions_league"]'), options.championsLeagueTeams || options.teams);
    fillSelect(
      document.querySelector('[name="golden_boot"]'),
      options.goldenBootPlayers || options.players || []
    );
    fillSelect(
      document.querySelector('[name="golden_glove"]'),
      options.goldenGlovePlayers || options.goalkeepers || []
    );
    renderAdminResultFields();
  }

  function renderAdminResultFields() {
    const host = $('adminResultsFields');
    if (!host) return;
    const fields = [
      ...[1,2,3,4].map(i => ({ name:`result_top4_${i}`, label:`Top four #${i}`, list:options.teams })),
      ...[1,2,3].map(i => ({ name:`result_bottom3_${i}`, label:`Bottom three #${i}`, list:options.teams })),
      { name:'result_golden_boot', label:'Golden Boot', list:options.goldenBootPlayers || options.players || [] },
      { name:'result_golden_glove', label:'Golden Glove', list:options.goldenGlovePlayers || options.goalkeepers || [] },
      { name:'result_fa_cup', label:'FA Cup', list:options.englishCupTeams || options.teams },
      { name:'result_league_cup', label:'League Cup', list:options.englishCupTeams || options.teams },
      { name:'result_champions_league', label:'Champions League', list:options.championsLeagueTeams || options.teams }
    ];
    host.innerHTML = fields.map(f => selectHtml(f.name, f.label, f.list)).join('');
  }

  function selectHtml(name, label, list) {
    return `<label>${escapeHtml(label)}<select name="${name}" required>${optionHtml(list)}</select></label>`;
  }
  function optionHtml(list) { return `<option value="">Choose...</option>${list.map(o => `<option value="${escapeHtml(o.key)}">${escapeHtml(o.label)}</option>`).join('')}`; }
  function fillSelect(el, list) { if (el) el.innerHTML = optionHtml(list); }

  function renderEntryState() {
    const status = me.entry.paymentStatus;
    $('entryBadge').textContent = status.toUpperCase();
    $('entryBadge').className = `status-pill ${status}`;
    const paid = ['paid', 'waived'].includes(status);
    $('paymentPanel').hidden = paid;
    $('predictionForm').hidden = !paid;
    if (me.competition.status === 'locked' || me.competition.status === 'live' || me.competition.status === 'finished') {
      qsa('#predictionForm select, #predictionForm button[type="submit"]').forEach(el => el.disabled = true);
      $('predictionMessage').textContent = 'Predictions are locked.';
    }
  }

  async function startPayment() {
    try {
      $('payButton').disabled = true;
      const data = await api.post('/api/payments/checkout', {});
      location.assign(data.checkoutUrl);
    } catch (error) { toast(error.message, true); $('payButton').disabled = false; }
  }

  async function loadMyPredictions() {
    const data = await api.get('/api/predictions');
    for (const p of data.predictions) {
      const name = p.category === 'top4' || p.category === 'bottom3' ? `${p.category}_${p.position}` : p.category;
      const el = document.querySelector(`[name="${name}"]`);
      if (el) el.value = p.selection_key;
    }
  }

  async function savePredictions(event) {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const payload = {
      top4: [1,2,3,4].map(i => form.get(`top4_${i}`)),
      bottom3: [1,2,3].map(i => form.get(`bottom3_${i}`)),
      golden_boot: form.get('golden_boot'), golden_glove: form.get('golden_glove'),
      fa_cup: form.get('fa_cup'), league_cup: form.get('league_cup'), champions_league: form.get('champions_league')
    };
    try {
      await api.put('/api/predictions', payload);
      $('predictionMessage').textContent = 'Saved securely.';
      toast('Predictions saved.');
    } catch (error) { $('predictionMessage').textContent = error.message; toast(error.message, true); }
  }

  async function loadLeaderboard() {
    if (!account) return;
    try {
      const data = await api.get('/api/leaderboard');
      $('leaderboardBody').innerHTML = data.leaderboard.length ? data.leaderboard.map(row => `<tr><td>${row.rank}</td><td>${escapeHtml(row.name)}</td><td><strong>${row.score}</strong></td><td>${row.exactCount}</td><td>${row.misplacedCount}</td><td class="${row.paymentStatus}">${escapeHtml(row.paymentStatus)}</td></tr>`).join('') : '<tr><td colspan="6">No submitted entries yet.</td></tr>';
      try {
        const revealed = await api.get('/api/predictions/all');
        renderRevealed(revealed.predictions);
      } catch (error) {
        $('revealedPredictions').innerHTML = `<h2>Everyone's predictions</h2><p>${escapeHtml(error.message)}</p>`;
      }
    } catch (error) { toast(error.message, true); }
  }

  function renderRevealed(rows) {
    const grouped = new Map();
    rows.forEach(r => { if (!grouped.has(r.display_name)) grouped.set(r.display_name, []); grouped.get(r.display_name).push(r); });
    $('revealedPredictions').innerHTML = `<h2>Everyone's predictions</h2>${[...grouped.entries()].map(([name,picks]) => `<h3>${escapeHtml(name)}</h3><div class="prediction-list">${picks.map(p => `<span>${escapeHtml(labelCategory(p.category, p.position))}</span><b>${escapeHtml(p.selection_label)}</b>`).join('')}</div>`).join('')}`;
  }

  async function loadLive() {
    if (!account) return;
    try {
      const data = await api.get('/api/live');
      $('liveUpdated').textContent = data.updatedAt ? `Updated ${formatDate(data.updatedAt)}` : data.message || 'Not connected';
      $('fixturesList').innerHTML = data.fixtures?.length ? data.fixtures.map(f => `<div class="fixture"><div><strong>${escapeHtml(f.home)} vs ${escapeHtml(f.away)}</strong><small>Matchweek ${f.matchday || '-'}</small></div><div>${f.homeScore ?? ''}${f.homeScore != null ? '–' : ''}${f.awayScore ?? ''}<small>${escapeHtml(f.status)}</small></div></div>`).join('') : `<p>${escapeHtml(data.message || 'No fixtures available.')}</p>`;
      $('leagueTableBody').innerHTML = data.table?.length ? data.table.map(r => `<tr><td>${r.position}</td><td>${escapeHtml(r.team)}</td><td>${r.played}</td><td>${r.goalDifference}</td><td><strong>${r.points}</strong></td></tr>`).join('') : '<tr><td colspan="5">No table data.</td></tr>';
    } catch (error) { toast(error.message, true); }
  }

  async function loadAdmin() {
    const data = await api.get('/api/admin/overview');
    if (data.competition.first_kickoff_at) $('firstKickoffInput').value = toLocalInput(data.competition.first_kickoff_at);
    $('adminEntriesBody').innerHTML = data.entries.map(e => `<tr><td>${escapeHtml(e.display_name)}</td><td>${escapeHtml(e.email || '')}</td><td>${escapeHtml(e.payment_status)}</td><td>${e.submitted_at ? formatDate(e.submitted_at) : 'No'}</td><td>${['paid','waived'].includes(e.payment_status) ? '' : `<button class="button button-ghost waive-button" data-entry="${e.id}">Waive</button>`}</td></tr>`).join('');
    qsa('.waive-button').forEach(b => b.addEventListener('click', () => waivePayment(Number(b.dataset.entry))));
    renderFootballSync(data.footballSync);
    $('goldenBootOptions').value = (options.goldenBootPlayers || options.players || [])
      .map(option => option.label)
      .join('\n');
    $('goldenGloveOptions').value = (options.goldenGlovePlayers || options.goalkeepers || [])
      .map(option => option.label)
      .join('\n');
    data.results.forEach(r => {
      const name = (r.category === 'top4' || r.category === 'bottom3') ? `result_${r.category}_${r.position}` : `result_${r.category}`;
      const el = document.querySelector(`[name="${name}"]`);
      if (el) el.value = r.selection_key;
      if (r.is_final) $('resultsFinal').checked = true;
    });
  }

  async function saveSchedule(event) {
    event.preventDefault();
    try {
      const iso = new Date($('firstKickoffInput').value).toISOString();
      const data = await api.put('/api/admin/competition', { firstKickoffAt: iso });
      $('adminMessage').textContent = `Opened. Deadline: ${formatDate(data.submissionDeadlineAt)}`;
      await loadPublicConfig();
    } catch (error) { toast(error.message, true); }
  }



  function renderFootballSync(sync) {
    if (!sync) {
      $('footballSyncStatus').textContent = 'Not synced';
      $('footballSyncTime').textContent = 'Never';
      $('footballSyncCounts').textContent = '0 teams · 0 players';
      return;
    }
    $('footballSyncStatus').textContent = String(sync.status || 'unknown').toUpperCase();
    $('footballSyncTime').textContent = sync.completed_at ? formatDate(sync.completed_at) : 'In progress';
    $('footballSyncCounts').textContent =
      `${sync.premier_league_team_count || 0} PL teams · ` +
      `${sync.champions_league_team_count || 0} UCL teams · ` +
      `${sync.player_count || 0} players · ${sync.goalkeeper_count || 0} keepers`;
    $('footballSyncMessage').textContent = sync.error_message || '';
  }

  async function syncFootballData() {
    const button = $('syncFootballButton');
    button.disabled = true;
    button.textContent = 'Syncing…';

    try {
      let pass = 0;
      let complete = false;

      while (!complete && pass < 6) {
        pass++;
        $('footballSyncMessage').textContent =
          pass === 1
            ? 'Importing teams and the first squad batch…'
            : `Importing the next squad batch (pass ${pass})…`;

        const result = await api.post('/api/admin/sync-football-data', {});
        complete = Boolean(result.synced);

        $('footballSyncStatus').textContent = String(result.status || 'running').toUpperCase();
        $('footballSyncCounts').textContent =
          `${result.premierLeagueTeams || 0} PL teams · ` +
          `${result.championsLeagueTeams || 0} UCL teams · ` +
          `${result.players || 0} players · ${result.goalkeepers || 0} keepers`;

        if (complete) {
          $('footballSyncMessage').textContent =
            `Sync complete: ${result.players} players and ${result.goalkeepers} goalkeepers imported.` +
            (result.warnings?.length ? ` ${result.warnings.join(' ')}` : '');
          toast('Football squads synced.');
          break;
        }

        $('footballSyncMessage').textContent =
          `${result.players || 0} players imported. ` +
          `${result.remainingTeams || 0} teams remain. ` +
          `Waiting ${result.retryAfterSeconds || 65} seconds for the API rate limit…`;

        await new Promise(resolve =>
          setTimeout(resolve, (result.retryAfterSeconds || 65) * 1000)
        );
      }

      await loadOptions();
      await loadAdmin();

      if (!complete) {
        $('footballSyncMessage').textContent =
          'The sync paused before completion. Click “Sync teams and players” again to continue.';
      }
    } catch (error) {
      console.error('Football squad sync failed', error);
      $('footballSyncMessage').textContent = error.message;
      toast(error.message, true);
    } finally {
      button.disabled = false;
      button.textContent = 'Sync teams and players';
    }
  }


  async function saveAwardOptions(event, type) {
    event.preventDefault();

    const isBoot = type === 'golden_boot_player';
    const textarea = isBoot
      ? $('goldenBootOptions')
      : $('goldenGloveOptions');

    const message = isBoot
      ? $('goldenBootOptionsMessage')
      : $('goldenGloveOptionsMessage');

    const label = isBoot
      ? 'Golden Boot'
      : 'Golden Glove';

    const values = textarea.value
      .split('\n')
      .map(value => value.trim())
      .filter(Boolean);

    if (!values.length) {
      message.textContent = `Add at least one ${label} option.`;
      return;
    }

    try {
      await api.put('/api/admin/options', {
        type,
        options: values
      });

      message.textContent =
        `${values.length} ${label} options saved.`;

      await loadOptions();
      await loadAdmin();

      toast(`${label} list saved.`);
    } catch (error) {
      message.textContent = error.message;
      toast(error.message, true);
    }
  }

  async function saveResults(event) {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const payload = {
      top4: [1,2,3,4].map(i => form.get(`result_top4_${i}`)),
      bottom3: [1,2,3].map(i => form.get(`result_bottom3_${i}`)),
      golden_boot: form.get('result_golden_boot'), golden_glove: form.get('result_golden_glove'),
      fa_cup: form.get('result_fa_cup'), league_cup: form.get('result_league_cup'), champions_league: form.get('result_champions_league'),
      isFinal: $('resultsFinal').checked
    };
    try { await api.put('/api/admin/results', payload); toast('Results saved and leaderboard recalculated.'); await loadLeaderboard(); }
    catch (error) { toast(error.message, true); }
  }

  async function waivePayment(entryId) {
    if (!confirm('Mark this entry as paid/waived?')) return;
    try { await api.post('/api/admin/payments/waive', { entryId }); await loadAdmin(); }
    catch (error) { toast(error.message, true); }
  }

  function labelCategory(category, position) {
    const labels = { top4: `Top four #${position}`, bottom3: `Bottom three #${position}`, golden_boot: 'Golden Boot', golden_glove: 'Golden Glove', fa_cup: 'FA Cup', league_cup: 'League Cup', champions_league: 'Champions League' };
    return labels[category] || category;
  }
  function formatDate(value) { return new Intl.DateTimeFormat('en-GB',{dateStyle:'medium',timeStyle:'short',timeZone:'Europe/London'}).format(new Date(value)); }
  function toLocalInput(value) { const d = new Date(value); const offset = d.getTimezoneOffset(); return new Date(d.getTime()-offset*60000).toISOString().slice(0,16); }
  function escapeHtml(value) { return String(value ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
  function toast(message, isError=false) { const el=$('toast'); el.textContent=message; el.className=`toast show${isError?' error':''}`; clearTimeout(toast.timer); toast.timer=setTimeout(()=>el.className='toast',4200); }
})();
