const APP_VERSION = 1;
const STORAGE_KEY = 'arian-outreach-dashboard-v1';
const THEME_KEY = 'arian-outreach-theme';

const state = {
  data: null,
  activeView: 'dashboard',
  actionUrgency: 'ALL',
  actionWorkflow: 'ALL',
  actionSearch: '',
  professorSearch: '',
  professorCountry: 'ALL',
  professorStatus: 'ALL',
  professorTrack: 'ALL',
  researchSearch: '',
  careerSearch: ''
};

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

function text(v) { return String(v ?? '').trim(); }
function lower(v) { return text(v).toLowerCase(); }
function escapeHtml(v) {
  return text(v).replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));
}
function unique(arr) { return [...new Set(arr.filter(Boolean))].sort((a,b)=>a.localeCompare(b)); }
function nowStamp() { return new Date().toLocaleString('en-GB', {timeZone:'Asia/Dhaka', dateStyle:'medium', timeStyle:'short'}); }
function isBlankRow(row) { return !row || Object.values(row).every(v => text(v) === ''); }

function showToast(message) {
  const el = $('#toast');
  el.textContent = message;
  el.classList.add('show');
  clearTimeout(showToast.t);
  showToast.t = setTimeout(() => el.classList.remove('show'), 2600);
}

function sheetRows(workbook, name) {
  const sheet = workbook.Sheets[name];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, {defval:'', raw:false}).filter(r => !isBlankRow(r));
}

function compactMaster(rows) {
  return rows.map(r => ({
    country:text(r['Country']), region:text(r['Region']), university:text(r['University']), department:text(r['Department / School']),
    professor:text(r['Professor / Group']), email:text(r['Email / Contact']), degree:text(r['Degree Target']), track:text(r['Research Track']),
    cv:text(r['CV Version']), priority:text(r['Priority']), fit:text(r['Fit Score']), recommended:text(r['Recommended Action']),
    recruiting:text(r['Recruiting Status']), funding:text(r['Funding Signal']), match:text(r['Match Summary']), hook:text(r['Personalized Hook']),
    profile:text(r['Profile URL']), opening:text(r['Opening / Funding URL']), publication:text(r['Publication / Project URL']), verified:text(r['Last Verified']),
    status:text(r['Outreach Status']), dateSent:text(r['Date Sent']), followDue:text(r['Follow-up Due']), reply:text(r['Reply Status']), replyDate:text(r['Reply Date']),
    meeting:text(r['Interview / Meeting']), application:text(r['Application Status']), next:text(r['Next Action']), subject:text(r['Gmail Thread Subject']),
    notes:text(r['Notes']), key:text(r['Unique Key']), gmailState:text(r['Gmail State']), followStage:text(r['Follow-up Stage']),
    responseDue:text(r['Response Action Due']), suppressed:text(r['Follow-up Suppressed?']), suppressionReason:text(r['Suppression Reason']),
    conversationPriority:text(r['Conversation Priority']), interactionUpdated:text(r['Interaction Updated At'])
  })).filter(r => r.professor || r.university);
}

function compactActions(rows) {
  return rows.map(r => ({
    urgency:text(r['Urgency']), workflow:text(r['Workflow']), action:text(r['Action']), due:text(r['Due / Trigger']),
    country:text(r['Country / Region']), university:text(r['University / Organization']), contact:text(r['Professor / Contact']),
    track:text(r['Track / Role']), state:text(r['Status / Reply']), next:text(r['Next Action']), reference:text(r['Reference'])
  })).filter(r => r.action || r.contact);
}

function compactResearch(rows) {
  return rows.map(r => ({
    institution:text(r['Institution']), contact:text(r['Professor / Contact']), country:text(r['Country / Region']), track:text(r['Track']),
    email:text(r['Email / Contact']), status:text(r['Outreach Status']), reply:text(r['Reply Status']), dateSent:text(r['Date Sent']),
    followDue:text(r['Follow-up Due']), meeting:text(r['Interview / Meeting']), next:text(r['Next Action']), subject:text(r['Gmail Subject']),
    lastSync:text(r['Last Gmail Sync']), notes:text(r['Notes'])
  })).filter(r => r.contact || r.institution);
}

function compactCareer(rows) {
  return rows.map(r => ({
    organization:text(r['Organization']), role:text(r['Role / Program']), category:text(r['Category']), status:text(r['Status']),
    last:text(r['Last Contact / Event']), next:text(r['Next Action']), due:text(r['Deadline / Date']), subject:text(r['Gmail Subject']), notes:text(r['Notes'])
  })).filter(r => r.organization || r.role);
}

function compactInteractions(rows) {
  return rows.map(r => ({
    date:text(r['Date / Time']), workflow:text(r['Workflow']), direction:text(r['Direction']), country:text(r['Country / Region']),
    institution:text(r['University / Institution']), contact:text(r['Professor / Contact']), type:text(r['Interaction Type']),
    summary:text(r['Summary']), outcome:text(r['Outcome / State Change']), due:text(r['Action Due']), subject:text(r['Gmail Subject / Thread']),
    calendar:text(r['Calendar / Reference']), source:text(r['Source']), key:text(r['Unique Key'])
  })).filter(r => r.date || r.contact || r.summary);
}

function compactPlaybook(rows) {
  return rows.map(r => ({
    type:text(r['Response Type']), signal:text(r['Typical Signal']), responseTime:text(r['Target Response Time']), trackerState:text(r['Tracker State / Reply Status']),
    followRule:text(r['Follow-up Rule']), calendar:text(r['Calendar?']), action:text(r['Default Action']), notes:text(r['Exceptions / Notes'])
  })).filter(r => r.type);
}

function parseWorkbook(workbook, fileName) {
  const data = {
    version: APP_VERSION,
    importedAt: new Date().toISOString(),
    fileName,
    master: compactMaster(sheetRows(workbook, 'Master Tracker')),
    actions: compactActions(sheetRows(workbook, 'Action Queue')),
    research: compactResearch(sheetRows(workbook, 'Research Outreach')),
    career: compactCareer(sheetRows(workbook, 'Career Tracker')),
    interactions: compactInteractions(sheetRows(workbook, 'Interaction Log')),
    playbook: compactPlaybook(sheetRows(workbook, 'Response Playbook'))
  };
  if (!data.master.length && !data.actions.length) throw new Error('This does not look like the Fall 2027 outreach workbook.');
  return data;
}

async function loadFile(file) {
  if (!file) return;
  try {
    if (typeof XLSX === 'undefined') throw new Error('Spreadsheet parser did not load. Check your internet connection and refresh.');
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, {type:'array'});
    const parsed = parseWorkbook(wb, file.name);
    state.data = parsed;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(parsed));
    hydrateFilters();
    renderAll();
    showToast(`Loaded ${file.name}`);
  } catch (err) {
    console.error(err);
    showToast(err.message || 'Could not read workbook.');
  }
}

function loadCached() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    if (parsed.version !== APP_VERSION) return false;
    state.data = parsed;
    return true;
  } catch { return false; }
}

function clearLocalData() {
  localStorage.removeItem(STORAGE_KEY);
  state.data = null;
  $('#appContent').classList.add('hidden');
  $('#emptyState').classList.remove('hidden');
  $('#syncText').textContent = 'No workbook loaded';
  $('#navActionCount').textContent = '0';
  showToast('Local tracker data cleared');
}

function statusClass(status) {
  const s = lower(status);
  if (s.includes('closed')) return 'closed';
  if (s.includes('replied') || s.includes('interview')) return 'replied';
  if (s.includes('scheduled')) return 'scheduled';
  if (s.includes('sent') || s.includes('waiting')) return 'sent';
  return '';
}

function metric(label, value, cls='') {
  return `<div class="metric ${cls}"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div>`;
}

function computeMetrics() {
  const d = state.data;
  const a = d.actions;
  const m = d.master;
  const replyMeeting = a.filter(x => ['PREPARE MEETING','HANDLE REPLY','HANDLE REFERRAL'].includes(x.action)).length;
  const followups = a.filter(x => ['FOLLOW UP','REVIEW / CLOSE'].includes(x.action)).length;
  const applications = a.filter(x => x.action === 'APPLICATION ACTION').length;
  const drafts = m.filter(x => x.status === 'Draft Prepared / Ready for Arian Review').length;
  const scheduled = m.filter(x => x.status === 'Scheduled by Arian').length + d.research.filter(x => x.status === 'Scheduled').length;
  const sent = m.filter(x => x.status === 'Sent / Awaiting Reply').length;
  const replied = m.filter(x => x.status === 'Replied').length;
  const closed = m.filter(x => x.status === 'Closed').length;
  const activeCareer = d.career.filter(x => !/closed|rejected|completed.*no/i.test(x.status)).length;
  return {replyMeeting, followups, applications, drafts, scheduled, sent, replied, closed, activeCareer, total:m.length};
}

function renderDashboard() {
  const metrics = computeMetrics();
  $('#metricGrid').innerHTML = [
    metric('Replies / meetings', metrics.replyMeeting, metrics.replyMeeting ? 'urgent':''),
    metric('Follow-ups due', metrics.followups, metrics.followups ? 'warn':''),
    metric('Application actions', metrics.applications),
    metric('Drafts ready', metrics.drafts),
    metric('Scheduled', metrics.scheduled),
    metric('Awaiting reply', metrics.sent)
  ].join('');

  const top = state.data.actions.slice(0, 7);
  $('#priorityList').innerHTML = top.length ? top.map(item => `
    <article class="priority-item">
      <span class="priority-badge ${escapeHtml(item.urgency)}">${escapeHtml(item.urgency || 'INFO')}</span>
      <div><strong>${escapeHtml(item.action)} · ${escapeHtml(item.contact || item.university)}</strong>
      <p>${escapeHtml(item.university)}${item.country ? ` · ${escapeHtml(item.country)}`:''}<br>${escapeHtml(item.next)}</p></div>
      <time>${escapeHtml(item.due || '')}</time>
    </article>`).join('') : '<p class="muted">No active actions in the imported queue.</p>';

  const pipeline = [
    ['Not started', state.data.master.filter(x=>['Not Started','Researching'].includes(x.status)).length],
    ['Draft ready', metrics.drafts], ['Scheduled', state.data.master.filter(x=>x.status==='Scheduled by Arian').length],
    ['Awaiting', metrics.sent], ['Replied', metrics.replied], ['Closed', metrics.closed]
  ];
  const max = Math.max(1, ...pipeline.map(x=>x[1]));
  $('#pipelineBars').innerHTML = pipeline.map(([label,count])=>`<div class="bar-row"><span>${escapeHtml(label)}</span><div class="bar"><i style="width:${Math.max(3, count/max*100)}%"></i></div><b>${count}</b></div>`).join('');

  const meeting = state.data.research.find(x => x.status === 'Interview' || lower(x.reply).includes('meeting confirmed')) || state.data.master.find(x => x.meeting);
  $('#nextEvent').innerHTML = meeting ? `<strong>${escapeHtml(meeting.contact || meeting.professor)}</strong><p>${escapeHtml(meeting.institution || meeting.university)}<br>${escapeHtml(meeting.meeting || meeting.next || '')}</p>` : '<strong>No confirmed meeting found</strong><p>Confirmed meetings from the workbook will appear here.</p>';

  const recent = [...state.data.interactions].reverse().slice(0,6);
  $('#recentInteractions').innerHTML = recent.length ? recent.map(i => `<div class="timeline-item"><small>${escapeHtml(i.date)}</small><strong>${escapeHtml(i.contact || i.institution || i.workflow)}</strong><p>${escapeHtml(i.type || i.outcome)}${i.summary ? ` · ${escapeHtml(i.summary)}`:''}</p></div>`).join('') : '<p class="muted">No Interaction Log entries in the imported workbook.</p>';

  $('#navActionCount').textContent = state.data.actions.filter(x=>['NOW','TODAY'].includes(x.urgency)).length;
}

function renderActions() {
  let rows = [...state.data.actions];
  if (state.actionUrgency !== 'ALL') rows = rows.filter(x => x.urgency === state.actionUrgency);
  if (state.actionWorkflow !== 'ALL') rows = rows.filter(x => x.workflow === state.actionWorkflow);
  const q = lower(state.actionSearch);
  if (q) rows = rows.filter(x => lower(Object.values(x).join(' ')).includes(q));
  $('#actionCards').innerHTML = rows.length ? rows.map(x => `<article class="task-card">
    <div class="urgency ${escapeHtml(x.urgency)}">${escapeHtml(x.urgency || 'INFO')}</div>
    <div><h3>${escapeHtml(x.action)} · ${escapeHtml(x.contact || x.university)}</h3><div class="meta">${escapeHtml(x.workflow)} · ${escapeHtml(x.university)}${x.country?` · ${escapeHtml(x.country)}`:''}</div><div class="pill-row"><span class="pill">${escapeHtml(x.state)}</span>${x.track?`<span class="pill">${escapeHtml(x.track)}</span>`:''}</div><p class="next">${escapeHtml(x.next)}</p></div>
    <div class="due">${escapeHtml(x.due || '')}</div>
  </article>`).join('') : '<div class="panel">No matching action items.</div>';
}

function renderProfessors() {
  let rows = [...state.data.master];
  const q = lower(state.professorSearch);
  if (q) rows = rows.filter(x => lower([x.professor,x.university,x.country,x.track,x.next,x.reply].join(' ')).includes(q));
  if (state.professorCountry !== 'ALL') rows = rows.filter(x => x.country === state.professorCountry);
  if (state.professorStatus !== 'ALL') rows = rows.filter(x => x.status === state.professorStatus);
  if (state.professorTrack !== 'ALL') rows = rows.filter(x => x.cv === state.professorTrack || x.track === state.professorTrack);
  $('#professorCount').textContent = rows.length;
  $('#professorTable').innerHTML = rows.map(x => `<tr>
    <td><strong>${escapeHtml(x.professor)}</strong>${x.email?`<br><small>${escapeHtml(x.email)}</small>`:''}</td>
    <td>${escapeHtml(x.university)}</td><td>${escapeHtml(x.country)}</td><td>${escapeHtml(x.track)}</td><td>${escapeHtml(x.priority)}</td>
    <td><span class="status-chip ${statusClass(x.status)}">${escapeHtml(x.status || '—')}</span></td><td>${escapeHtml(x.reply || '—')}</td><td>${escapeHtml(x.next || '—')}</td>
  </tr>`).join('');
}

function genericCards(rows, type) {
  if (!rows.length) return '<div class="panel">No matching records.</div>';
  return rows.map(x => {
    const name = type === 'research' ? x.contact : x.role;
    const org = type === 'research' ? x.institution : x.organization;
    const status = x.status || '';
    const reply = type === 'research' ? x.reply : x.category;
    const due = type === 'research' ? (x.meeting || x.followDue) : x.due;
    return `<article class="task-card"><div class="urgency ${statusClass(status).toUpperCase()}">•</div><div><h3>${escapeHtml(name)}</h3><div class="meta">${escapeHtml(org)}${type==='research'&&x.country?` · ${escapeHtml(x.country)}`:''}</div><div class="pill-row"><span class="pill">${escapeHtml(status)}</span>${reply?`<span class="pill">${escapeHtml(reply)}</span>`:''}</div><p class="next">${escapeHtml(x.next || '')}</p></div><div class="due">${escapeHtml(due || '')}</div></article>`;
  }).join('');
}

function renderResearch() {
  let rows = [...state.data.research];
  const q = lower(state.researchSearch);
  if (q) rows = rows.filter(x => lower(Object.values(x).join(' ')).includes(q));
  $('#researchCards').innerHTML = genericCards(rows, 'research');
}

function renderCareer() {
  let rows = [...state.data.career];
  const q = lower(state.careerSearch);
  if (q) rows = rows.filter(x => lower(Object.values(x).join(' ')).includes(q));
  $('#careerCards').innerHTML = genericCards(rows, 'career');
}

function renderPlaybook() {
  $('#playbookGrid').innerHTML = state.data.playbook.length ? state.data.playbook.map(x => `<article class="play-card"><h3>${escapeHtml(x.type)}</h3><dl>
    <dt>Signal</dt><dd>${escapeHtml(x.signal)}</dd><dt>Respond</dt><dd>${escapeHtml(x.responseTime)}</dd><dt>Tracker state</dt><dd>${escapeHtml(x.trackerState)}</dd><dt>Follow-up</dt><dd>${escapeHtml(x.followRule)}</dd><dt>Calendar</dt><dd>${escapeHtml(x.calendar)}</dd><dt>Action</dt><dd>${escapeHtml(x.action)}</dd>${x.notes?`<dt>Notes</dt><dd>${escapeHtml(x.notes)}</dd>`:''}
  </dl></article>`).join('') : '<div class="panel">Response Playbook sheet not found in this workbook.</div>';
}

function hydrateFilters() {
  const d = state.data;
  const workflows = unique(d.actions.map(x=>x.workflow));
  $('#actionWorkflowFilter').innerHTML = '<option value="ALL">All workflows</option>' + workflows.map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  $('#countryFilter').innerHTML = '<option value="ALL">All countries</option>' + unique(d.master.map(x=>x.country)).map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  $('#statusFilter').innerHTML = '<option value="ALL">All statuses</option>' + unique(d.master.map(x=>x.status)).map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  const tracks = unique(d.master.map(x=>x.cv).filter(Boolean));
  $('#trackFilter').innerHTML = '<option value="ALL">All tracks</option>' + tracks.map(x=>`<option>${escapeHtml(x)}</option>`).join('');
}

function renderAll() {
  if (!state.data) return;
  $('#emptyState').classList.add('hidden');
  $('#appContent').classList.remove('hidden');
  $('#syncText').textContent = `Local snapshot · imported ${new Date(state.data.importedAt).toLocaleString()} · ${state.data.fileName}`;
  renderDashboard(); renderActions(); renderProfessors(); renderResearch(); renderCareer(); renderPlaybook();
}

function switchView(view) {
  state.activeView = view;
  $$('.nav-item').forEach(x => x.classList.toggle('active', x.dataset.view === view));
  $$('.view').forEach(x => x.classList.toggle('active', x.id === `view-${view}`));
  const titles = {dashboard:'Control Center', actions:'Action Queue', professors:'Professor Tracker', research:'Research / RA', career:'Career Tracker', playbook:'Response Playbook'};
  $('#pageTitle').textContent = titles[view] || 'Outreach OS';
  $('#sidebar').classList.remove('open');
  window.scrollTo({top:0,behavior:'smooth'});
}

function bindUI() {
  $$('.nav-item').forEach(btn => btn.addEventListener('click', () => switchView(btn.dataset.view)));
  $$('[data-go]').forEach(btn => btn.addEventListener('click', () => switchView(btn.dataset.go)));
  $('#workbookInput').addEventListener('change', e => loadFile(e.target.files[0]));
  $('#clearLocalData').addEventListener('click', clearLocalData);
  $('#menuBtn').addEventListener('click', () => $('#sidebar').classList.toggle('open'));
  $('#themeBtn').addEventListener('click', () => {
    document.body.classList.toggle('dark');
    localStorage.setItem(THEME_KEY, document.body.classList.contains('dark') ? 'dark' : 'light');
  });

  $$('#actionUrgencyFilter button').forEach(btn => btn.addEventListener('click', () => {
    $$('#actionUrgencyFilter button').forEach(x=>x.classList.remove('active')); btn.classList.add('active'); state.actionUrgency = btn.dataset.value; renderActions();
  }));
  $('#actionWorkflowFilter').addEventListener('change', e => {state.actionWorkflow=e.target.value;renderActions();});
  $('#actionSearch').addEventListener('input', e => {state.actionSearch=e.target.value;renderActions();});
  $('#professorSearch').addEventListener('input', e => {state.professorSearch=e.target.value;renderProfessors();});
  $('#countryFilter').addEventListener('change', e => {state.professorCountry=e.target.value;renderProfessors();});
  $('#statusFilter').addEventListener('change', e => {state.professorStatus=e.target.value;renderProfessors();});
  $('#trackFilter').addEventListener('change', e => {state.professorTrack=e.target.value;renderProfessors();});
  $('#researchSearch').addEventListener('input', e => {state.researchSearch=e.target.value;renderResearch();});
  $('#careerSearch').addEventListener('input', e => {state.careerSearch=e.target.value;renderCareer();});

  const drop = $('#dropZone');
  ['dragenter','dragover'].forEach(evt => drop.addEventListener(evt, e => {e.preventDefault(); drop.classList.add('drag');}));
  ['dragleave','drop'].forEach(evt => drop.addEventListener(evt, e => {e.preventDefault(); drop.classList.remove('drag');}));
  drop.addEventListener('drop', e => loadFile(e.dataTransfer.files[0]));
}

function startClock() {
  const tick = () => {
    const now = new Date();
    $('#todayLabel').textContent = now.toLocaleDateString('en-GB',{timeZone:'Asia/Dhaka',weekday:'long',day:'numeric',month:'long'});
    $('#localClock').textContent = now.toLocaleTimeString('en-GB',{timeZone:'Asia/Dhaka',hour:'2-digit',minute:'2-digit'});
  };
  tick(); setInterval(tick, 30000);
}

function init() {
  if (localStorage.getItem(THEME_KEY) === 'dark' || (!localStorage.getItem(THEME_KEY) && matchMedia('(prefers-color-scheme: dark)').matches)) document.body.classList.add('dark');
  bindUI(); startClock();
  if (loadCached()) { hydrateFilters(); renderAll(); }
}

document.addEventListener('DOMContentLoaded', init);
