const APP_VERSION = 2;
const STORAGE_KEY = 'arian-outreach-dashboard-v1';
const THEME_KEY = 'arian-outreach-theme';

const state = {
  data: null,
  activeView: 'dashboard',
  actionUrgency: 'ALL',
  actionWorkflow: 'ALL',
  actionSearch: '',
  actionFocus: 'ALL',
  professorSearch: '',
  professorCountry: 'ALL',
  professorStatus: 'ALL',
  professorTrack: 'ALL',
  professorSort: 'priority',
  professorView: window.matchMedia('(max-width: 640px)').matches ? 'cards' : 'table',
  researchSearch: '',
  careerSearch: '',
  playbookSearch: ''
};

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
const icon = id => `<svg aria-hidden="true"><use href="#${id}"/></svg>`;

function text(v) { return String(v ?? '').trim(); }
function lower(v) { return text(v).toLowerCase(); }
function escapeHtml(v) { return text(v).replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch])); }
function escapeAttr(v) { return escapeHtml(v).replace(/`/g, '&#96;'); }
function unique(arr) { return [...new Set(arr.filter(Boolean))].sort((a,b)=>a.localeCompare(b)); }
function isBlankRow(row) { return !row || Object.values(row).every(v => text(v) === ''); }
function ageHours(iso) { const t = new Date(iso).getTime(); return Number.isFinite(t) ? (Date.now()-t)/36e5 : Infinity; }
function dateValue(v) { const d = new Date(v); return Number.isNaN(d.getTime()) ? 0 : d.getTime(); }
function safeUrl(v) { const s = text(v); return /^https?:\/\//i.test(s) ? s : ''; }
function gmailUrl(email) { return email ? `https://mail.google.com/mail/u/0/#search/${encodeURIComponent(email)}` : ''; }
function mailtoUrl(email) { return email ? `mailto:${encodeURIComponent(email)}` : ''; }

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
    lastOutbound:text(r['Last Outbound Date']), followSent:text(r['Follow-up Sent Date']), responseDue:text(r['Response Action Due']),
    suppressed:text(r['Follow-up Suppressed?']), suppressionReason:text(r['Suppression Reason']), conversationPriority:text(r['Conversation Priority']),
    interactionUpdated:text(r['Interaction Updated At']), lastSync:text(r['Last Gmail Sync'])
  })).filter(r => r.professor || r.university);
}
function compactActions(rows) {
  return rows.map(r => ({urgency:text(r['Urgency']), workflow:text(r['Workflow']), action:text(r['Action']), due:text(r['Due / Trigger']), country:text(r['Country / Region']), university:text(r['University / Organization']), contact:text(r['Professor / Contact']), track:text(r['Track / Role']), state:text(r['Status / Reply']), next:text(r['Next Action']), reference:text(r['Reference'])})).filter(r => r.action || r.contact);
}
function compactResearch(rows) {
  return rows.map(r => ({institution:text(r['Institution']), contact:text(r['Professor / Contact']), country:text(r['Country / Region']), track:text(r['Track']), email:text(r['Email / Contact']), status:text(r['Outreach Status']), reply:text(r['Reply Status']), dateSent:text(r['Date Sent']), followDue:text(r['Follow-up Due']), meeting:text(r['Interview / Meeting']), next:text(r['Next Action']), subject:text(r['Gmail Subject']), lastSync:text(r['Last Gmail Sync']), notes:text(r['Notes'])})).filter(r => r.contact || r.institution);
}
function compactCareer(rows) {
  return rows.map(r => ({organization:text(r['Organization']), role:text(r['Role / Program']), category:text(r['Category']), status:text(r['Status']), last:text(r['Last Contact / Event']), next:text(r['Next Action']), due:text(r['Deadline / Date']), subject:text(r['Gmail Subject']), notes:text(r['Notes'])})).filter(r => r.organization || r.role);
}
function compactInteractions(rows) {
  return rows.map(r => ({date:text(r['Date / Time']), workflow:text(r['Workflow']), direction:text(r['Direction']), country:text(r['Country / Region']), institution:text(r['University / Institution']), contact:text(r['Professor / Contact']), type:text(r['Interaction Type']), summary:text(r['Summary']), outcome:text(r['Outcome / State Change']), due:text(r['Action Due']), subject:text(r['Gmail Subject / Thread']), calendar:text(r['Calendar / Reference']), source:text(r['Source']), key:text(r['Unique Key'])})).filter(r => r.date || r.contact || r.summary);
}
function compactPlaybook(rows) {
  return rows.map(r => ({type:text(r['Response Type']), signal:text(r['Typical Signal']), responseTime:text(r['Target Response Time']), trackerState:text(r['Tracker State / Reply Status']), followRule:text(r['Follow-up Rule']), calendar:text(r['Calendar?']), action:text(r['Default Action']), notes:text(r['Exceptions / Notes'])})).filter(r => r.type);
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

function normalizeCached(parsed) {
  if (!parsed || ![1,2].includes(parsed.version)) return null;
  for (const key of ['master','actions','research','career','interactions','playbook']) if (!Array.isArray(parsed[key])) parsed[key] = [];
  return parsed;
}

async function loadFile(file) {
  if (!file) return;
  try {
    if (typeof XLSX === 'undefined') throw new Error('Spreadsheet parser did not load. Check your internet connection and refresh.');
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, {type:'array'});
    const parsed = parseWorkbook(wb, file.name);
    state.data = parsed;
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(parsed)); }
    catch { showToast('Loaded, but this browser could not cache the full snapshot.'); }
    resetFilters(); hydrateFilters(); renderAll(); closeModal('updateModal');
    showToast(`Loaded ${file.name}`);
  } catch (err) { console.error(err); showToast(err.message || 'Could not read workbook.'); }
  $('#workbookInput').value = '';
}

function loadCached() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY); if (!raw) return false;
    const parsed = normalizeCached(JSON.parse(raw)); if (!parsed) return false;
    state.data = parsed; return true;
  } catch { return false; }
}
function clearLocalData() {
  localStorage.removeItem(STORAGE_KEY); state.data = null;
  $('#appContent').classList.add('hidden'); $('#emptyState').classList.remove('hidden'); $('#staleBanner').classList.add('hidden');
  updateFreshness(); closeModal('updateModal'); $('#navActionCount').textContent='0'; $('#mobileActionCount').textContent='0';
  showToast('Local tracker data cleared');
}
function resetFilters(){
  Object.assign(state,{actionUrgency:'ALL',actionWorkflow:'ALL',actionSearch:'',actionFocus:'ALL',professorSearch:'',professorCountry:'ALL',professorStatus:'ALL',professorTrack:'ALL',professorSort:'priority',researchSearch:'',careerSearch:'',playbookSearch:''});
}

function statusClass(status) {
  const s=lower(status); if(s.includes('closed'))return'closed'; if(s.includes('replied')||s.includes('interview'))return'replied'; if(s.includes('scheduled'))return'scheduled'; if(s.includes('sent')||s.includes('waiting'))return'sent'; return'';
}
function priorityRank(v){ const m=text(v).match(/P\s*([0-9]+)/i); return m?Number(m[1]):99; }
function actionCategory(a){
  if(['PREPARE MEETING','HANDLE REPLY','HANDLE REFERRAL'].includes(a.action))return'reply';
  if(['FOLLOW UP','REVIEW / CLOSE'].includes(a.action))return'followup';
  if(a.action==='APPLICATION ACTION')return'application';
  if(a.action==='REVIEW / SCHEDULE')return'draft';
  if(a.action==='WAIT FOR DELIVERY')return'scheduled';
  return'other';
}

function computeMetrics(){
  const d=state.data,a=d.actions,m=d.master;
  return {
    replyMeeting:a.filter(x=>actionCategory(x)==='reply').length,
    followups:a.filter(x=>actionCategory(x)==='followup').length,
    applications:a.filter(x=>actionCategory(x)==='application').length,
    drafts:m.filter(x=>x.status==='Draft Prepared / Ready for Arian Review').length,
    scheduled:m.filter(x=>x.status==='Scheduled by Arian').length+d.research.filter(x=>x.status==='Scheduled').length,
    sent:m.filter(x=>x.status==='Sent / Awaiting Reply').length,
    replied:m.filter(x=>x.status==='Replied').length,
    closed:m.filter(x=>x.status==='Closed').length,
    total:m.length
  };
}
function metricCard(label,value,detail,cls,focus){ return `<button class="metric ${cls||''}" data-metric-focus="${focus}"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span><small>${escapeHtml(detail)}</small></button>`; }

function updateFreshness(){
  const chip=$('#freshnessChip'), txt=$('#freshnessText');
  if(!state.data){chip.dataset.state='empty';txt.textContent='No data';return;}
  const hours=ageHours(state.data.importedAt); const stale=hours>24; chip.dataset.state=stale?'stale':'fresh';
  txt.textContent=stale?`${Math.floor(hours/24)}d old`:(hours<1?'Just updated':`${Math.floor(hours)}h old`);
  $('#staleBanner').classList.toggle('hidden',!stale);
}
function updateSummary(){
  const box=$('#updateSummary'); if(!state.data){box.innerHTML='<strong>No local snapshot loaded.</strong> Import the workbook to start.';return;}
  box.innerHTML=`<strong>Current snapshot:</strong> ${escapeHtml(state.data.fileName||'Workbook')}<br><span>Imported ${escapeHtml(new Date(state.data.importedAt).toLocaleString())} · ${state.data.master.length} professor records · ${state.data.actions.length} active queue items</span>`;
}

function renderDashboard(){
  const m=computeMetrics();
  $('#metricGrid').innerHTML=[
    metricCard('Replies / meetings',m.replyMeeting,'Handle active conversations first',m.replyMeeting?'urgent':'','reply'),
    metricCard('Follow-ups due',m.followups,'One concise routine follow-up',m.followups?'warn':'','followup'),
    metricCard('Application actions',m.applications,'Formal applications requested or due','','application'),
    metricCard('Drafts ready',m.drafts,'Review and schedule when ready','','draft'),
    metricCard('Scheduled',m.scheduled,'Await actual Gmail delivery','','scheduled'),
    metricCard('Awaiting reply',m.sent,'Passive waiting, no action needed',m.sent?'ok':'','awaiting')
  ].join('');
  const top=state.data.actions.slice(0,7);
  $('#priorityList').innerHTML=top.length?top.map((x,i)=>`<article class="priority-item" data-open-action="${i}"><span class="priority-badge ${escapeAttr(x.urgency)}">${escapeHtml(x.urgency||'INFO')}</span><div><strong>${escapeHtml(x.action)} · ${escapeHtml(x.contact||x.university)}</strong><p>${escapeHtml(x.university)}${x.country?` · ${escapeHtml(x.country)}`:''}<br>${escapeHtml(x.next)}</p></div><time>${escapeHtml(x.due||'')}</time></article>`).join(''):'<p class="muted">No active actions in the imported queue.</p>';
  const pipeline=[['Not started',state.data.master.filter(x=>['Not Started','Researching'].includes(x.status)).length],['Draft ready',m.drafts],['Scheduled',state.data.master.filter(x=>x.status==='Scheduled by Arian').length],['Awaiting',m.sent],['Replied',m.replied],['Closed',m.closed]];
  $('#pipelineTotal').textContent=`${m.total} records`; const max=Math.max(1,...pipeline.map(x=>x[1]));
  $('#pipelineBars').innerHTML=pipeline.map(([label,count])=>`<div class="bar-row"><span>${escapeHtml(label)}</span><div class="bar"><i style="width:${Math.max(2,count/max*100)}%"></i></div><b>${count}</b></div>`).join('');
  const meeting=state.data.research.find(x=>x.status==='Interview'||lower(x.reply).includes('meeting confirmed'))||state.data.master.find(x=>x.meeting);
  $('#nextEvent').innerHTML=meeting?`<div class="event-icon">${icon('i-calendar')}</div><strong>${escapeHtml(meeting.contact||meeting.professor)}</strong><p>${escapeHtml(meeting.institution||meeting.university)}<br>${escapeHtml(meeting.meeting||meeting.next||'')}</p>`:`<div class="event-icon">${icon('i-calendar')}</div><strong>No confirmed meeting found</strong><p>Meetings and interviews will surface here automatically from your workbook.</p>`;
  const recent=[...state.data.interactions].sort((a,b)=>dateValue(b.date)-dateValue(a.date)).slice(0,6);
  $('#recentInteractions').innerHTML=recent.length?recent.map(i=>`<div class="timeline-item"><small>${escapeHtml(i.date)}</small><strong>${escapeHtml(i.contact||i.institution||i.workflow)}</strong><p>${escapeHtml(i.type||i.outcome)}${i.summary?` · ${escapeHtml(i.summary)}`:''}</p></div>`).join(''):'<p class="muted">No Interaction Log entries in this snapshot.</p>';
  const urgent=state.data.actions.filter(x=>['NOW','TODAY'].includes(x.urgency)).length; $('#navActionCount').textContent=urgent; $('#mobileActionCount').textContent=urgent;
  $('#heroHeading').textContent=m.replyMeeting?`${m.replyMeeting} active conversation${m.replyMeeting===1?'':'s'} need your attention.`:(m.followups?`${m.followups} follow-up${m.followups===1?'':'s'} are due.`:'Your active queue is under control.');
  $('#heroSummary').textContent=m.replyMeeting?'Handle replies or meeting preparation before sending new outreach.':m.followups?'Work through due follow-ups, then review scheduled and draft outreach.':'Use the Action Queue to decide what to prepare next.';
}

function filteredActions(){
  let rows=[...state.data.actions];
  if(state.actionUrgency!=='ALL')rows=rows.filter(x=>x.urgency===state.actionUrgency);
  if(state.actionWorkflow!=='ALL')rows=rows.filter(x=>x.workflow===state.actionWorkflow);
  if(state.actionFocus!=='ALL')rows=rows.filter(x=>actionCategory(x)===state.actionFocus || (state.actionFocus==='awaiting'&&false));
  const q=lower(state.actionSearch); if(q)rows=rows.filter(x=>lower(Object.values(x).join(' ')).includes(q));
  return rows;
}
function renderActions(){
  const rows=filteredActions(); $('#actionVisibleCount').textContent=rows.length;
  $('#actionCards').innerHTML=rows.length?rows.map((x,i)=>`<article class="task-card" data-action-index="${state.data.actions.indexOf(x)}"><span class="urgency-badge ${escapeAttr(x.urgency)}">${escapeHtml(x.urgency||'INFO')}</span><div><h3>${escapeHtml(x.action)} · ${escapeHtml(x.contact||x.university)}</h3><div class="meta">${escapeHtml(x.workflow)} · ${escapeHtml(x.university)}${x.country?` · ${escapeHtml(x.country)}`:''}</div><div class="pill-row"><span class="pill">${escapeHtml(x.state)}</span>${x.track?`<span class="pill">${escapeHtml(x.track)}</span>`:''}</div><p class="next">${escapeHtml(x.next)}</p></div><div class="due">${escapeHtml(x.due||'')}</div></article>`).join(''):'<div class="panel"><strong>No matching action items.</strong><p class="muted">Try clearing a filter or search term.</p></div>';
}

function sortedProfessors(rows){
  const copy=[...rows];
  if(state.professorSort==='university')return copy.sort((a,b)=>a.university.localeCompare(b.university)||a.professor.localeCompare(b.professor));
  if(state.professorSort==='status')return copy.sort((a,b)=>a.status.localeCompare(b.status)||priorityRank(a.priority)-priorityRank(b.priority));
  if(state.professorSort==='recent')return copy.sort((a,b)=>Math.max(dateValue(b.interactionUpdated),dateValue(b.replyDate),dateValue(b.dateSent))-Math.max(dateValue(a.interactionUpdated),dateValue(a.replyDate),dateValue(a.dateSent)));
  return copy.sort((a,b)=>priorityRank(a.priority)-priorityRank(b.priority)||a.university.localeCompare(b.university));
}
function filteredProfessors(){
  let rows=[...state.data.master]; const q=lower(state.professorSearch);
  if(q)rows=rows.filter(x=>lower([x.professor,x.university,x.country,x.track,x.next,x.reply,x.notes,x.match].join(' ')).includes(q));
  if(state.professorCountry!=='ALL')rows=rows.filter(x=>x.country===state.professorCountry);
  if(state.professorStatus!=='ALL')rows=rows.filter(x=>x.status===state.professorStatus);
  if(state.professorTrack!=='ALL')rows=rows.filter(x=>x.cv===state.professorTrack||x.track===state.professorTrack);
  return sortedProfessors(rows);
}
function renderActiveFilters(){
  const items=[];
  if(state.professorCountry!=='ALL')items.push(['country',state.professorCountry]);
  if(state.professorStatus!=='ALL')items.push(['status',state.professorStatus]);
  if(state.professorTrack!=='ALL')items.push(['track',state.professorTrack]);
  if(state.professorSearch)items.push(['search',`Search: ${state.professorSearch}`]);
  $('#activeFilterRow').innerHTML=items.map(([k,v])=>`<button class="filter-chip" data-clear-filter="${k}"><b>${escapeHtml(v)}</b> ×</button>`).join('');
}
function professorCard(x,index){return `<article class="record-card" data-professor-index="${index}"><div class="record-card-top"><div><h3>${escapeHtml(x.professor)}</h3><div class="org">${escapeHtml(x.university)}${x.country?` · ${escapeHtml(x.country)}`:''}</div></div><span class="priority-token">${escapeHtml(x.priority||'—')}</span></div><div class="pill-row"><span class="status-chip ${statusClass(x.status)}">${escapeHtml(x.status||'—')}</span>${x.reply?`<span class="pill">${escapeHtml(x.reply)}</span>`:''}</div><p class="record-next">${escapeHtml(x.next||x.track||'No next action recorded.')}</p><div class="record-card-foot"><span>${escapeHtml(x.track||x.cv||'')}</span>${icon('i-chevron')}</div></article>`;}
function renderProfessors(){
  const rows=filteredProfessors(); $('#professorCount').textContent=rows.length; renderActiveFilters();
  $('#professorTable').innerHTML=rows.map(x=>{const idx=state.data.master.indexOf(x);return `<tr data-professor-index="${idx}"><td><strong>${escapeHtml(x.professor)}</strong>${x.email?`<br><small>${escapeHtml(x.email)}</small>`:''}</td><td>${escapeHtml(x.university)}</td><td>${escapeHtml(x.country)}</td><td>${escapeHtml(x.track)}</td><td>${escapeHtml(x.priority||'—')}</td><td><span class="status-chip ${statusClass(x.status)}">${escapeHtml(x.status||'—')}</span></td><td>${escapeHtml(x.reply||'—')}</td><td>${escapeHtml(x.next||'—')}</td><td><button class="row-arrow" tabindex="-1">${icon('i-chevron')}</button></td></tr>`}).join('');
  $('#professorCardGrid').innerHTML=rows.map(x=>professorCard(x,state.data.master.indexOf(x))).join('');
  $('#professorTableWrap').classList.toggle('hidden',state.professorView!=='table'); $('#professorCardGrid').classList.toggle('hidden',state.professorView!=='cards');
  $$('[data-prof-view]').forEach(b=>b.classList.toggle('active',b.dataset.profView===state.professorView));
}

function genericRecordCard(x,type,index){
  const isResearch=type==='research', name=isResearch?x.contact:x.role, org=isResearch?x.institution:x.organization, status=x.status||'', secondary=isResearch?(x.reply||x.track):x.category, due=isResearch?(x.meeting||x.followDue):x.due;
  return `<article class="record-card" data-record-type="${type}" data-record-index="${index}"><div class="record-card-top"><div><h3>${escapeHtml(name)}</h3><div class="org">${escapeHtml(org)}${isResearch&&x.country?` · ${escapeHtml(x.country)}`:''}</div></div><span class="status-chip ${statusClass(status)}">${escapeHtml(status||'—')}</span></div><div class="pill-row">${secondary?`<span class="pill">${escapeHtml(secondary)}</span>`:''}${due?`<span class="pill">${escapeHtml(due)}</span>`:''}</div><p class="record-next">${escapeHtml(x.next||'No next action recorded.')}</p><div class="record-card-foot"><span>${escapeHtml(isResearch?x.track:(x.last||''))}</span>${icon('i-chevron')}</div></article>`;
}
function renderResearch(){
  let rows=[...state.data.research],q=lower(state.researchSearch); if(q)rows=rows.filter(x=>lower(Object.values(x).join(' ')).includes(q)); $('#researchCount').textContent=rows.length;
  $('#researchCards').innerHTML=rows.length?rows.map(x=>genericRecordCard(x,'research',state.data.research.indexOf(x))).join(''):'<div class="panel">No matching research records.</div>';
}
function renderCareer(){
  let rows=[...state.data.career],q=lower(state.careerSearch); if(q)rows=rows.filter(x=>lower(Object.values(x).join(' ')).includes(q)); $('#careerCount').textContent=rows.length;
  $('#careerCards').innerHTML=rows.length?rows.map(x=>genericRecordCard(x,'career',state.data.career.indexOf(x))).join(''):'<div class="panel">No matching career records.</div>';
}
function renderPlaybook(){
  let rows=[...state.data.playbook],q=lower(state.playbookSearch); if(q)rows=rows.filter(x=>lower(Object.values(x).join(' ')).includes(q));
  $('#playbookGrid').innerHTML=rows.length?rows.map(x=>`<article class="play-card"><div class="play-card-head"><h3>${escapeHtml(x.type)}</h3><span class="response-time">${escapeHtml(x.responseTime||'Review')}</span></div><p class="signal">${escapeHtml(x.signal)}</p><dl class="play-rule"><dt>Tracker state</dt><dd>${escapeHtml(x.trackerState)}</dd><dt>Follow-up</dt><dd>${escapeHtml(x.followRule)}</dd><dt>Calendar</dt><dd>${escapeHtml(x.calendar)}</dd><dt>Default action</dt><dd>${escapeHtml(x.action)}</dd>${x.notes?`<dt>Exceptions</dt><dd>${escapeHtml(x.notes)}</dd>`:''}</dl></article>`).join(''):'<div class="panel">No matching response rules.</div>';
}

function hydrateFilters(){
  const d=state.data; $('#actionWorkflowFilter').innerHTML='<option value="ALL">All workflows</option>'+unique(d.actions.map(x=>x.workflow)).map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  $('#countryFilter').innerHTML='<option value="ALL">All countries</option>'+unique(d.master.map(x=>x.country)).map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  $('#statusFilter').innerHTML='<option value="ALL">All statuses</option>'+unique(d.master.map(x=>x.status)).map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  $('#trackFilter').innerHTML='<option value="ALL">All CV tracks</option>'+unique(d.master.map(x=>x.cv).filter(Boolean)).map(x=>`<option>${escapeHtml(x)}</option>`).join('');
}
function renderAll(){
  if(!state.data)return; $('#emptyState').classList.add('hidden'); $('#appContent').classList.remove('hidden'); updateFreshness(); updateSummary();
  renderDashboard();renderActions();renderProfessors();renderResearch();renderCareer();renderPlaybook();
}

function switchView(view){
  state.activeView=view; $$('.nav-item').forEach(x=>x.classList.toggle('active',x.dataset.view===view)); $$('.mobile-bottom-nav [data-view]').forEach(x=>x.classList.toggle('active',x.dataset.view===view)); $$('.view').forEach(x=>x.classList.toggle('active',x.id===`view-${view}`));
  const meta={dashboard:['Dashboard','Your Fall 2027 outreach control center'],actions:['Action Queue','Only items that require an action'],professors:['Professors','Your canonical graduate-outreach database'],research:['Research / RA','Research positions, internships and RA outreach'],career:['Career','Jobs, internships and fellowships'],playbook:['Response Playbook','Rules for replies, follow-ups and state changes']};
  $('#pageTitle').textContent=meta[view]?.[0]||'Outreach OS'; $('#pageSubtitle').textContent=meta[view]?.[1]||''; closeSidebar(); window.scrollTo({top:0,behavior:'smooth'});
}
function applyMetricFocus(focus){
  switchView('actions'); state.actionFocus=focus; state.actionUrgency='ALL'; state.actionSearch='';
  $$('#actionUrgencyFilter button').forEach(b=>b.classList.toggle('active',b.dataset.value==='ALL')); $('#actionSearch').value=''; renderActions();
  if(focus==='awaiting'){switchView('professors');state.professorStatus='Sent / Awaiting Reply';$('#statusFilter').value=state.professorStatus;renderProfessors();}
}

function openDrawer(type,index){
  if(!state.data)return; let item,eyebrow,title,subtitle,actions='',body='';
  if(type==='professor'){
    item=state.data.master[index]; if(!item)return; eyebrow='PROFESSOR RECORD';title=item.professor||'Professor';subtitle=[item.university,item.country].filter(Boolean).join(' · ');
    const links=[]; if(item.email){links.push(`<a class="btn btn-primary" href="${mailtoUrl(item.email)}">${icon('i-mail')} Email</a>`);links.push(`<a class="btn btn-secondary" href="${gmailUrl(item.email)}" target="_blank" rel="noopener">${icon('i-search')} Gmail</a>`);}
    if(safeUrl(item.profile))links.push(`<a class="btn btn-secondary" href="${escapeAttr(item.profile)}" target="_blank" rel="noopener">${icon('i-external')} Profile</a>`); if(item.email)links.push(`<button class="btn btn-secondary" data-copy="${escapeAttr(item.email)}">${icon('i-copy')} Copy email</button>`); actions=links.join('');
    body=detailSection('Current workflow',detailGrid([['Outreach status',item.status],['Reply status',item.reply],['Follow-up stage',item.followStage],['Conversation priority',item.conversationPriority],['Date sent',item.dateSent],['Follow-up due',item.followDue],['Response action due',item.responseDue],['Meeting / interview',item.meeting]]))+detailSection('Research fit',detailGrid([['Degree target',item.degree],['Research track',item.track],['CV version',item.cv],['Priority',item.priority],['Fit score',item.fit],['Recruiting status',item.recruiting],['Funding signal',item.funding],['Recommended action',item.recommended]]))+detailSection('Next action',detailCopy(item.next||'No next action recorded.'))+(item.match?detailSection('Match summary',detailCopy(item.match)):'')+(item.hook?detailSection('Personalized hook',detailCopy(item.hook)):'')+(item.notes?detailSection('Notes',detailCopy(item.notes)):'')+detailSection('Communication',detailGrid([['Email',item.email],['Gmail state',item.gmailState],['Last outbound',item.lastOutbound],['Follow-up sent',item.followSent],['Last Gmail sync',item.lastSync],['Subject',item.subject]]))+linkSection([['Professor / group profile',item.profile],['Opening / funding evidence',item.opening],['Publication / project',item.publication]]);
  } else if(type==='action'){
    item=state.data.actions[index]; if(!item)return;eyebrow=`${item.urgency||'INFO'} · ${item.workflow}`;title=`${item.action} · ${item.contact||item.university}`;subtitle=[item.university,item.country].filter(Boolean).join(' · ');actions=`<button class="btn btn-primary" data-go="professors">${icon('i-users')} Open professor tracker</button>`;body=detailSection('Action',detailCopy(item.next))+detailSection('Context',detailGrid([['Due / trigger',item.due],['Current state',item.state],['Track / role',item.track],['Reference',item.reference]]));
  } else if(type==='research'){
    item=state.data.research[index];if(!item)return;eyebrow='RESEARCH / RA';title=item.contact||'Research outreach';subtitle=[item.institution,item.country].filter(Boolean).join(' · ');actions=item.email?`<a class="btn btn-primary" href="${mailtoUrl(item.email)}">${icon('i-mail')} Email</a><a class="btn btn-secondary" href="${gmailUrl(item.email)}" target="_blank" rel="noopener">${icon('i-search')} Gmail</a>`:'';body=detailSection('Workflow',detailGrid([['Status',item.status],['Reply',item.reply],['Date sent',item.dateSent],['Follow-up due',item.followDue],['Meeting',item.meeting],['Track',item.track],['Last Gmail sync',item.lastSync],['Subject',item.subject]]))+detailSection('Next action',detailCopy(item.next||'No next action recorded.'))+(item.notes?detailSection('Notes',detailCopy(item.notes)):'');
  } else if(type==='career'){
    item=state.data.career[index];if(!item)return;eyebrow='CAREER';title=item.role||'Career item';subtitle=item.organization;body=detailSection('Workflow',detailGrid([['Category',item.category],['Status',item.status],['Last contact / event',item.last],['Deadline / date',item.due],['Subject',item.subject]]))+detailSection('Next action',detailCopy(item.next||'No next action recorded.'))+(item.notes?detailSection('Notes',detailCopy(item.notes)):'');
  }
  $('#drawerEyebrow').textContent=eyebrow;$('#drawerTitle').textContent=title;$('#drawerSubtitle').textContent=subtitle;$('#drawerActions').innerHTML=actions;$('#drawerBody').innerHTML=body;$('#drawerBackdrop').classList.remove('hidden');$('#detailDrawer').classList.add('open');$('#detailDrawer').setAttribute('aria-hidden','false');
}
function detailSection(title,content){return content?`<section class="detail-section"><h3>${escapeHtml(title)}</h3>${content}</section>`:'';}
function detailGrid(items){const usable=items.filter(([,v])=>text(v));return usable.length?`<div class="detail-grid">${usable.map(([k,v])=>`<div class="detail-item"><span>${escapeHtml(k)}</span><strong>${escapeHtml(v)}</strong></div>`).join('')}</div>`:'';}
function detailCopy(v){return text(v)?`<div class="detail-copy">${escapeHtml(v)}</div>`:'';}
function linkSection(items){const usable=items.filter(([,url])=>safeUrl(url));return usable.length?detailSection('Evidence & links',`<div class="detail-links">${usable.map(([label,url])=>`<a class="detail-link" href="${escapeAttr(url)}" target="_blank" rel="noopener">${icon('i-link')}<span>${escapeHtml(label)}</span><span class="external">${icon('i-external')}</span></a>`).join('')}</div>`):'';}
function closeDrawer(){ $('#drawerBackdrop').classList.add('hidden');$('#detailDrawer').classList.remove('open');$('#detailDrawer').setAttribute('aria-hidden','true'); }

function openModal(id){const el=$(`#${id}`);if(!el)return;if(id==='updateModal')updateSummary();el.classList.remove('hidden');if(id==='searchModal')setTimeout(()=>$('#globalSearchInput').focus(),30);}
function closeModal(id){const el=$(`#${id}`);if(el)el.classList.add('hidden');}
function openSidebar(){$('#sidebar').classList.add('open');} function closeSidebar(){$('#sidebar').classList.remove('open');}

function searchEverything(q){
  q=lower(q);if(!q){$('#globalSearchResults').innerHTML=`<div class="search-hint">${icon('i-command')}<p>Type anything to search across your imported workspace.</p></div>`;return;}
  const groups=[];
  const prof=state.data.master.filter(x=>lower([x.professor,x.university,x.country,x.track,x.next,x.reply].join(' ')).includes(q)).slice(0,6).map(x=>({type:'professor',index:state.data.master.indexOf(x),title:x.professor,sub:`${x.university} · ${x.track}`,meta:x.status,icon:'i-users'}));
  const acts=state.data.actions.filter(x=>lower(Object.values(x).join(' ')).includes(q)).slice(0,5).map(x=>({type:'action',index:state.data.actions.indexOf(x),title:`${x.action} · ${x.contact||x.university}`,sub:x.university,meta:x.urgency,icon:'i-check'}));
  const research=state.data.research.filter(x=>lower(Object.values(x).join(' ')).includes(q)).slice(0,4).map(x=>({type:'research',index:state.data.research.indexOf(x),title:x.contact,sub:x.institution,meta:x.status,icon:'i-flask'}));
  const career=state.data.career.filter(x=>lower(Object.values(x).join(' ')).includes(q)).slice(0,4).map(x=>({type:'career',index:state.data.career.indexOf(x),title:x.role,sub:x.organization,meta:x.status,icon:'i-briefcase'}));
  if(prof.length)groups.push(['Professors',prof]);if(acts.length)groups.push(['Actions',acts]);if(research.length)groups.push(['Research / RA',research]);if(career.length)groups.push(['Career',career]);
  $('#globalSearchResults').innerHTML=groups.length?groups.map(([name,items])=>`<div class="search-section-title">${escapeHtml(name)}</div>${items.map(r=>`<div class="search-result" data-search-type="${r.type}" data-search-index="${r.index}"><div class="search-result-icon">${icon(r.icon)}</div><div><strong>${escapeHtml(r.title)}</strong><span>${escapeHtml(r.sub)}</span></div><small>${escapeHtml(r.meta||'')}</small></div>`).join('')}`).join(''):'<div class="search-hint"><p>No matching records.</p></div>';
}

function bindUI(){
  $$('.nav-item').forEach(btn=>btn.addEventListener('click',()=>switchView(btn.dataset.view))); $$('.mobile-bottom-nav [data-view]').forEach(btn=>btn.addEventListener('click',()=>switchView(btn.dataset.view))); $$('[data-go]').forEach(btn=>btn.addEventListener('click',()=>switchView(btn.dataset.go)));
  $('#menuBtn').addEventListener('click',openSidebar);$('#sidebarClose').addEventListener('click',closeSidebar);$('#sidebarScrim').addEventListener('click',closeSidebar);$('#mobileMore').addEventListener('click',openSidebar);
  $('#updateDataBtn').addEventListener('click',()=>openModal('updateModal'));$('#freshnessChip').addEventListener('click',()=>openModal('updateModal'));$$('[data-update-data]').forEach(b=>b.addEventListener('click',()=>openModal('updateModal')));
  $('#emptyImportBtn').addEventListener('click',()=>$('#workbookInput').click());$('#modalImportBtn').addEventListener('click',()=>$('#workbookInput').click());$('#dropZone').addEventListener('click',()=>$('#workbookInput').click());$('#workbookInput').addEventListener('change',e=>loadFile(e.target.files[0]));$('#clearLocalData').addEventListener('click',clearLocalData);
  $$('[data-close-modal]').forEach(b=>b.addEventListener('click',()=>closeModal(b.dataset.closeModal))); $$('.modal-backdrop').forEach(m=>m.addEventListener('mousedown',e=>{if(e.target===m)closeModal(m.id);}));
  $('#themeBtn').addEventListener('click',toggleTheme);$('#globalSearchBtn').addEventListener('click',()=>state.data?openModal('searchModal'):showToast('Import the workbook first.'));$('#globalSearchInput').addEventListener('input',e=>searchEverything(e.target.value));
  $('#drawerClose').addEventListener('click',closeDrawer);$('#drawerBackdrop').addEventListener('click',closeDrawer);
  $$('#actionUrgencyFilter button').forEach(btn=>btn.addEventListener('click',()=>{$$('#actionUrgencyFilter button').forEach(x=>x.classList.remove('active'));btn.classList.add('active');state.actionUrgency=btn.dataset.value;state.actionFocus='ALL';renderActions();}));
  $('#actionWorkflowFilter').addEventListener('change',e=>{state.actionWorkflow=e.target.value;renderActions();});$('#actionSearch').addEventListener('input',e=>{state.actionSearch=e.target.value;renderActions();});
  $('#professorSearch').addEventListener('input',e=>{state.professorSearch=e.target.value;renderProfessors();});$('#countryFilter').addEventListener('change',e=>{state.professorCountry=e.target.value;renderProfessors();});$('#statusFilter').addEventListener('change',e=>{state.professorStatus=e.target.value;renderProfessors();});$('#trackFilter').addEventListener('change',e=>{state.professorTrack=e.target.value;renderProfessors();});$('#professorSort').addEventListener('change',e=>{state.professorSort=e.target.value;renderProfessors();});
  $$('[data-prof-view]').forEach(b=>b.addEventListener('click',()=>{state.professorView=b.dataset.profView;renderProfessors();}));
  $('#researchSearch').addEventListener('input',e=>{state.researchSearch=e.target.value;renderResearch();});$('#careerSearch').addEventListener('input',e=>{state.careerSearch=e.target.value;renderCareer();});$('#playbookSearch').addEventListener('input',e=>{state.playbookSearch=e.target.value;renderPlaybook();});
  const drop=$('#dropZone');['dragenter','dragover'].forEach(evt=>drop.addEventListener(evt,e=>{e.preventDefault();drop.classList.add('drag');}));['dragleave','drop'].forEach(evt=>drop.addEventListener(evt,e=>{e.preventDefault();drop.classList.remove('drag');}));drop.addEventListener('drop',e=>loadFile(e.dataTransfer.files[0]));
  document.addEventListener('click',delegatedClick);document.addEventListener('keydown',keyboardShortcuts);
}
function delegatedClick(e){
  const metric=e.target.closest('[data-metric-focus]');if(metric){applyMetricFocus(metric.dataset.metricFocus);return;}
  const act=e.target.closest('[data-action-index]');if(act){openDrawer('action',Number(act.dataset.actionIndex));return;}
  const pri=e.target.closest('[data-open-action]');if(pri){openDrawer('action',Number(pri.dataset.openAction));return;}
  const prof=e.target.closest('[data-professor-index]');if(prof){openDrawer('professor',Number(prof.dataset.professorIndex));return;}
  const rec=e.target.closest('[data-record-type]');if(rec){openDrawer(rec.dataset.recordType,Number(rec.dataset.recordIndex));return;}
  const sr=e.target.closest('[data-search-type]');if(sr){closeModal('searchModal');$('#globalSearchInput').value='';searchEverything('');openDrawer(sr.dataset.searchType,Number(sr.dataset.searchIndex));return;}
  const clear=e.target.closest('[data-clear-filter]');if(clear){const k=clear.dataset.clearFilter;if(k==='country'){state.professorCountry='ALL';$('#countryFilter').value='ALL';}if(k==='status'){state.professorStatus='ALL';$('#statusFilter').value='ALL';}if(k==='track'){state.professorTrack='ALL';$('#trackFilter').value='ALL';}if(k==='search'){state.professorSearch='';$('#professorSearch').value='';}renderProfessors();return;}
  const copy=e.target.closest('[data-copy]');if(copy){navigator.clipboard?.writeText(copy.dataset.copy).then(()=>showToast('Copied to clipboard')).catch(()=>showToast('Could not copy'));return;}
  const go=e.target.closest('[data-go]');if(go){closeDrawer();switchView(go.dataset.go);}
}
function keyboardShortcuts(e){
  if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();if(state.data)openModal('searchModal');return;}
  if(e.key==='Escape'){closeModal('searchModal');closeModal('updateModal');closeDrawer();closeSidebar();return;}
  if(['INPUT','TEXTAREA','SELECT'].includes(document.activeElement?.tagName))return;
  const map={'1':'dashboard','2':'actions','3':'professors','4':'research','5':'career','6':'playbook'};if(map[e.key])switchView(map[e.key]);
}
function toggleTheme(){document.body.classList.toggle('dark');localStorage.setItem(THEME_KEY,document.body.classList.contains('dark')?'dark':'light');syncThemeIcon();}
function syncThemeIcon(){$('#themeIcon use').setAttribute('href',document.body.classList.contains('dark')?'#i-sun':'#i-moon');}
function startClock(){const tick=()=>{const now=new Date();$('#todayLabel').textContent=now.toLocaleDateString('en-GB',{timeZone:'Asia/Dhaka',weekday:'long',day:'numeric',month:'long'});$('#localClock').textContent=now.toLocaleTimeString('en-GB',{timeZone:'Asia/Dhaka',hour:'2-digit',minute:'2-digit'});updateFreshness();};tick();setInterval(tick,60000);}

function init(){
  if(localStorage.getItem(THEME_KEY)==='dark'||(!localStorage.getItem(THEME_KEY)&&matchMedia('(prefers-color-scheme: dark)').matches))document.body.classList.add('dark');syncThemeIcon();bindUI();startClock();
  if(loadCached()){hydrateFilters();renderAll();}
}
document.addEventListener('DOMContentLoaded',init);
