// ================== STATE ==================
const AppState = { clients: [] };
let USERS = [
  { id: 1, username: 'manager', password: '1234', role: 'manager', clientIds: [] }
];
let uploadedFiles = [];
let audienceDraft = {};
let selectedAudienceColor = 'all';
let filterAudienceColor = 'all';
let filterAudienceProcedure = 'all';
let filterAudienceTribunal = 'all';
let filterSuiviProcedure = 'all';
let filterSuiviTribunal = 'all';
let filterDiligenceProcedure = 'all';
let editingDossier = null;
let editingOriginalProcedures = [];
let customProcedures = [];
let suppressProcedureChange = false;
let currentUser = null;
let editingTeamUserId = null;
let dashboardCalendarCursor = new Date(new Date().getFullYear(), new Date().getMonth(), 1);

// ================== HELPERS ==================
const $ = (id) => document.getElementById(id);

function isManager(){
  return currentUser?.role === 'manager';
}

function isAdmin(){
  return currentUser?.role === 'admin';
}

function isViewer(){
  return currentUser?.role === 'viewer';
}

function canEditData(){
  return isManager() || isAdmin();
}

function canDeleteData(){
  return isManager();
}

function canManageTeam(){
  return isManager();
}

function canViewClient(client){
  if(!client) return false;
  if(isManager()) return true;
  if(isAdmin()) return true;
  const ids = Array.isArray(currentUser?.clientIds) ? currentUser.clientIds : [];
  return ids.includes(client.id);
}

function canEditClient(client){
  if(!canEditData()) return false;
  return canViewClient(client);
}

function getVisibleClients(){
  return AppState.clients.filter(c=>canViewClient(c));
}

function collectDeepValues(value, out = []){
  if(value === null || value === undefined) return out;
  if(Array.isArray(value)){
    value.forEach(v => collectDeepValues(v, out));
    return out;
  }
  if(typeof value === 'object'){
    Object.values(value).forEach(v => collectDeepValues(v, out));
    return out;
  }
  out.push(String(value));
  return out;
}

function escapeAttr(value){
  return escapeHtml(String(value ?? ''));
}

function makeAudienceDraftKey(ci, di, procKey){
  return `${ci}::${di}::${encodeURIComponent(String(procKey))}`;
}

function parseAudienceDraftKey(key){
  const value = String(key);
  if(value.includes('::')){
    const [ci = '', di = '', ...rest] = value.split('::');
    return { ci, di, procKey: decodeURIComponent(rest.join('::')) };
  }
  const [ci = '', di = '', ...rest] = value.split('_');
  return { ci, di, procKey: rest.join('_') };
}

function fileToDataUrl(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onload = ()=>resolve(String(reader.result || ''));
    reader.onerror = ()=>reject(new Error('Lecture fichier impossible'));
    reader.readAsDataURL(file);
  });
}

async function serializeUploadedFiles(files){
  const result = [];
  for(const file of files){
    if(!file) continue;
    if(file.dataUrl && file.name){
      result.push({
        name: String(file.name || ''),
        size: Number(file.size || 0),
        type: String(file.type || ''),
        dataUrl: String(file.dataUrl || '')
      });
      continue;
    }
    if(typeof File !== 'undefined' && file instanceof File){
      const dataUrl = await fileToDataUrl(file);
      result.push({
        name: file.name,
        size: file.size,
        type: file.type,
        dataUrl
      });
    }
  }
  return result;
}

function getStoredFileSource(file){
  if(!file) return '';
  if(file.dataUrl) return String(file.dataUrl);
  if(typeof File !== 'undefined' && file instanceof File){
    return URL.createObjectURL(file);
  }
  return '';
}

function openStoredFile(file){
  const src = getStoredFileSource(file);
  if(!src){
    alert('Fichier non disponible pour aperçu');
    return;
  }
  window.open(src, '_blank');
  if(typeof File !== 'undefined' && file instanceof File){
    setTimeout(()=>URL.revokeObjectURL(src), 10000);
  }
}

function downloadStoredFile(file){
  const src = getStoredFileSource(file);
  if(!src){
    alert('Fichier non disponible pour téléchargement');
    return;
  }
  const a = document.createElement('a');
  a.href = src;
  a.download = String(file?.name || 'document');
  document.body.appendChild(a);
  a.click();
  a.remove();
  if(typeof File !== 'undefined' && file instanceof File){
    setTimeout(()=>URL.revokeObjectURL(src), 10000);
  }
}

// ================== INIT ==================
function initApplication(){
  setupEvents();
  renderClients();
  renderDashboard();
  updateClientDropdown();
  renderSuivi();
  renderAudience();
  renderDiligence();
  renderEquipe();
}

// ================== EVENTS ==================
function setupEvents(){
  $('dashboardLink').onclick = ()=>showView('dashboard');
  $('clientsLink').onclick = ()=>showView('clients');
  $('creationLink').onclick = ()=>showView('creation');
  $('suiviLink').onclick = ()=>showView('suivi');
  $('audienceLink').onclick = ()=>showView('audience');
  $('diligenceLink').onclick = ()=>showView('diligence');
  $('equipeLink')?.addEventListener('click', ()=>showView('equipe'));

  $('loginBtn').onclick = login;
  $('logoutBtn').onclick = logout;
  $('closeDossierModalBtn')?.addEventListener('click', closeDossierModal);
  $('dossierModal')?.addEventListener('click', (e)=>{
    if(e.target?.id === 'dossierModal') closeDossierModal();
  });
  document.addEventListener('keydown', (e)=>{
    if(e.key === 'Escape') closeDossierModal();
  });
  $('addClientBtn').onclick = ()=>addClient($('clientName').value);
  $('addDossierBtn').onclick = addDossier;
  $('uploadBtn')?.addEventListener('click', ()=> $('fileInput')?.click());
  $('fileInput')?.addEventListener('change', handleFiles);

  const dz = $('dropzone');
  if(dz){
    dz.addEventListener('dragover', (e)=>{
      e.preventDefault();
      dz.classList.add('dragover');
    });
    dz.addEventListener('dragleave', ()=> dz.classList.remove('dragover'));
    dz.addEventListener('drop', (e)=>{
      e.preventDefault();
      dz.classList.remove('dragover');
      handleFiles(e);
    });
  }

  document.querySelectorAll('.proc-check').forEach(cb=>{
    cb.addEventListener('change', e=>{
      if(suppressProcedureChange) return;
      const label = e.target.closest('label');
      if(label) label.classList.toggle('active', e.target.checked);
      renderProcedureDetails();
    });
  });
  $('procedureCustom')?.addEventListener('keydown', (e)=>{
    if(e.key === 'Enter'){
      e.preventDefault();
      addCustomProcedure();
    }
  });
  $('addProcedureBtn')?.addEventListener('click', addCustomProcedure);

  $('searchClientInput')?.addEventListener('input', renderClients);

  $('filterGlobal')?.addEventListener('input', renderSuivi);
  $('filterSuiviProcedure')?.addEventListener('change', (e)=>{
    filterSuiviProcedure = e.target.value;
    renderSuivi();
  });
  $('filterSuiviTribunal')?.addEventListener('change', (e)=>{
    filterSuiviTribunal = e.target.value;
    renderSuivi();
  });
  $('filterAudience')?.addEventListener('input', renderAudience);
  $('diligenceProcedureFilter')?.addEventListener('change', (e)=>{
    filterDiligenceProcedure = e.target.value;
    renderDiligence();
  });
  $('diligenceSearchInput')?.addEventListener('input', renderDiligence);
  $('teamRole')?.addEventListener('change', updateTeamClientSelectorState);
  $('teamSaveBtn')?.addEventListener('click', saveTeamUser);
  $('teamResetBtn')?.addEventListener('click', resetTeamForm);
  $('teamAddClientBtn')?.addEventListener('click', addClientFromTeam);
  $('teamNewClientName')?.addEventListener('keydown', (e)=>{
    if(e.key === 'Enter'){
      e.preventDefault();
      addClientFromTeam();
    }
  });
  $('filterAudienceColor')?.addEventListener('change', (e)=>{
    filterAudienceColor = e.target.value;
    renderAudience();
  });
  $('filterAudienceProcedure')?.addEventListener('change', (e)=>{
    filterAudienceProcedure = e.target.value;
    renderAudience();
  });
  $('filterAudienceTribunal')?.addEventListener('change', (e)=>{
    filterAudienceTribunal = e.target.value;
    renderAudience();
  });

  $('saveAudienceBtn')?.addEventListener('click', saveAllAudience);
  $('exportAudienceBtn')?.addEventListener('click', exportAudienceXLS);
  $('calendarPrevBtn')?.addEventListener('click', ()=>{
    dashboardCalendarCursor = new Date(dashboardCalendarCursor.getFullYear(), dashboardCalendarCursor.getMonth() - 1, 1);
    renderDashboardCalendar();
  });
  $('calendarNextBtn')?.addEventListener('click', ()=>{
    dashboardCalendarCursor = new Date(dashboardCalendarCursor.getFullYear(), dashboardCalendarCursor.getMonth() + 1, 1);
    renderDashboardCalendar();
  });

  // ===== Audience color filters =====
  document.querySelectorAll('.color-btn').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const color = btn.dataset.color;
      setSelectedAudienceColor(color, false);
    });
  });
}

function setSelectedAudienceColor(color, syncFilter){
  selectedAudienceColor = color;
  document.querySelectorAll('.color-btn').forEach(b=>b.classList.remove('active'));
  const btn = document.querySelector(`.color-btn[data-color="${color}"]`);
  if(btn) btn.classList.add('active');
  if(syncFilter){
    const sel = $('filterAudienceColor');
    if(sel) sel.value = color;
  }
}

// ================== NAV ==================
function showView(v){
  $('dashboardSection').style.display = v==='dashboard'?'block':'none';
  $('clientSection').style.display = v==='clients'?'block':'none';
  $('creationSection').style.display = v==='creation'?'block':'none';
  $('suiviSection').style.display = v==='suivi'?'block':'none';
  $('audienceSection').style.display = v==='audience'?'block':'none';
  $('diligenceSection').style.display = v==='diligence'?'block':'none';
  if($('equipeSection')) $('equipeSection').style.display = v==='equipe'?'block':'none';

  document.querySelectorAll('.nav-link').forEach(n=>n.classList.remove('active'));
  const target = $(v+'Link');
  if(target) target.classList.add('active');
  if(v === 'equipe') renderEquipe();
}

// ================== LOGIN ==================
function login(){
  const u = USERS.find(x=>x.username===$('username').value.trim() && x.password===$('password').value.trim());
  if(!u){ $('errorMsg').style.display='block'; return; }
  currentUser = u;
  $('errorMsg').style.display='none';
  $('loginScreen').style.display='none';
  $('appContent').style.display='flex';
  applyRoleUI();
  renderClients();
  renderDashboard();
  updateClientDropdown();
  renderSuivi();
  renderAudience();
  renderDiligence();
  renderEquipe();
}

function logout(){
  closeDossierModal();
  currentUser = null;
  $('appContent').style.display='none';
  $('loginScreen').style.display='flex';
  $('username').value='';
  $('password').value='';
}

function applyRoleUI(){
  const viewer = isViewer();
  const manager = isManager();

  if($('creationLink')) $('creationLink').style.display = viewer ? 'none' : '';
  if($('clientsLink')) $('clientsLink').style.display = viewer ? 'none' : '';
  if($('equipeLink')) $('equipeLink').style.display = manager ? '' : 'none';
  if($('addClientBtn')) $('addClientBtn').style.display = manager ? '' : 'none';

  const audienceEditable = canEditData();
  document.querySelectorAll('.color-btn').forEach(btn=> btn.disabled = !audienceEditable);
  if($('saveAudienceBtn')) $('saveAudienceBtn').style.display = audienceEditable ? '' : 'none';

  if(viewer && $('creationSection')?.style.display !== 'none'){
    showView('suivi');
  }
}

// ================== CLIENTS ==================
function addClient(name){
  if(!isManager()) return alert('Accès refusé');
  name = name.trim();
  if(!name) return alert('Nom obligatoire');
  const existing = AppState.clients.find(c=>c.name.trim().toLowerCase() === name.toLowerCase());
  if(existing){
    alert('Client déjà موجود, on ouvre le client existant');
    goToCreation(existing.id);
    return;
  }
  AppState.clients.push({ id: Date.now(), name, dossiers: [] });
  $('clientName').value='';
  renderClients();
  updateClientDropdown();
  renderDashboard();
  renderAudience();
  renderEquipe();
}

function updateClientDropdown(){
  const selectClient = $('selectClient');
  if(!selectClient) return;
  selectClient.innerHTML = '<option value="">Client</option>';
  getVisibleClients().filter(c=>canEditClient(c)).forEach(c=>{
    selectClient.innerHTML += `<option value="${c.id}">${escapeHtml(c.name)}</option>`;
  });
}

function renderClients(){
  const q = $('searchClientInput')?.value?.toLowerCase() || '';
  const clientsBody = $('clientsBody');
  clientsBody.innerHTML='';

  getVisibleClients()
    .filter(c => c.name.toLowerCase().includes(q))
    .forEach(c=>{
      const canEdit = canEditClient(c);
      clientsBody.innerHTML += `
        <tr>
          <td>${escapeHtml(c.name)}</td>
          <td>${c.dossiers.length}</td>
          <td>
            <button class="btn-primary" onclick="goToCreation(${c.id})" ${canEdit ? '' : 'disabled'}>
              <i class="fa-solid fa-folder-plus"></i>
            </button>
          </td>
        </tr>
      `;
    });
}

function goToCreation(clientId){
  const client = AppState.clients.find(c=>c.id == clientId);
  if(!client || !canEditClient(client)) return alert('Accès refusé');
  resetCreationForm(clientId);
  showView('creation');
}

// ================== DOSSIERS ==================
async function addDossier(){
  try{
    if(!canEditData()) return alert('Accès refusé');
    const wasEditing = !!editingDossier;
    const client = AppState.clients.find(c=>c.id == $('selectClient').value);
    if(!client) return alert('Choisir client');
    if(!canEditClient(client)) return alert('Accès refusé');

    let selected = [...document.querySelectorAll('.proc-check:checked')].map(cb=>cb.value);
    const activeLabels = [...document.querySelectorAll('.checkbox-group label.active')]
      .map(l=>l.dataset.proc)
      .filter(Boolean);
    const customList = customProcedures.slice();
    const cardProcs = [...document.querySelectorAll('#procedureDetails .proc-card h4')]
      .map(h=>h.innerText.trim())
      .filter(Boolean);
    selected.push(...activeLabels, ...customList, ...cardProcs);
    if(editingDossier){
      const prev = AppState.clients
        .find(c=>c.id == editingDossier.clientId)
        ?.dossiers?.[editingDossier.index];
      const prevList = normalizeProcedures(prev || {});
      selected.push(...prevList);
    }
    selected = [...new Set(selected.map(v=>String(v).trim()).filter(Boolean))].filter(v=>v !== 'Autre');
    if(selected.length === 0){
      if(editingDossier){
        const prev = AppState.clients
          .find(c=>c.id == editingDossier.clientId)
          ?.dossiers?.[editingDossier.index];
        const fallback = normalizeProcedures(prev || {});
        selected.push(...fallback);
      }
    }
    if(selected.length === 0) return alert('Choisir au moins une procédure');

    const details = {};
    const cards = [...document.querySelectorAll('#procedureDetails .proc-card')];
    cards.forEach(card=>{
      const procName = card.querySelector('h4')?.innerText.trim();
      if(!procName) return;
      details[procName] = {};
      card.querySelectorAll('input, select').forEach(fieldEl=>{
        details[procName][fieldEl.dataset.field] = fieldEl.value.trim();
      });
    });

    const dossier = {
      debiteur: $('debiteurInput').value.trim(),
      referenceClient: $('referenceClientInput').value.trim(),
      dateAffectation: $('dateAffectation').value || new Date().toISOString().slice(0,10),
      procedure: selected.join(', '),
      procedureList: selected.slice(),
      procedureDetails: details,
      ville: $('villeInput').value.trim(),
      adresse: $('adresseInput').value.trim(),
      montant: $('montantInput').value.trim(),
      ww: $('wwInput').value.trim(),
      marque: $('marqueInput').value.trim(),
      type: $('typeInput').value.trim(),
      note: $('noteInput')?.value.trim() || '',
      avancement: $('avancementInput')?.value.trim() || '',
      statut: $('statutInput')?.value || 'En cours',
      files: await serializeUploadedFiles(uploadedFiles)
    };
    console.log('[ADD DOSSIER]', JSON.stringify(dossier, null, 2));

    if(editingDossier){
      const prevClient = AppState.clients.find(c=>c.id == editingDossier.clientId);
      if(prevClient && prevClient.id === client.id){
        prevClient.dossiers[editingDossier.index] = dossier;
      }else{
        if(prevClient) prevClient.dossiers.splice(editingDossier.index, 1);
        client.dossiers.push(dossier);
      }
    }else{
      client.dossiers.push(dossier);
    }

    renderSuivi();
    renderDashboard();
    renderClients();
    renderAudience();
    renderDiligence();
    if(wasEditing){
      alert('Dossier mis à jour ✅');
    }else{
      alert('Dossier créé ✅');
    }
    resetCreationForm();
    showView('suivi');
  }catch(err){
    console.error(err);
    alert('Erreur pendant la sauvegarde du dossier');
  }
}

function handleFiles(e){
  let files = [];
  if(e.dataTransfer?.files) files = [...e.dataTransfer.files];
  if(e.target?.files) files = [...e.target.files];
  if(!files.length) return;
  files.forEach(f=>{
    uploadedFiles.push(f);
  });
  renderFileList();
}

function renderFileList(){
  const list = $('fileList');
  if(!list) return;
  list.innerHTML = '';
  uploadedFiles.forEach((f, idx)=>{
    const li = document.createElement('li');
    const name = document.createElement('span');
    const fileName = String(f?.name || `Fichier ${idx + 1}`);
    const fileSize = Number(f?.size || 0);
    const sizeText = fileSize ? ` (${Math.round(fileSize / 1024)} KB)` : '';
    name.textContent = `${fileName}${sizeText}`;

    const actions = document.createElement('span');
    actions.className = 'file-actions';

    const viewBtn = document.createElement('button');
    viewBtn.type = 'button';
    viewBtn.className = 'btn-view';
    viewBtn.textContent = 'Voir';
    viewBtn.addEventListener('click', ()=>viewFile(idx));

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn-remove';
    removeBtn.textContent = 'Supprimer';
    removeBtn.addEventListener('click', ()=>removeFile(idx));

    const downloadBtn = document.createElement('button');
    downloadBtn.type = 'button';
    downloadBtn.className = 'btn-primary';
    downloadBtn.textContent = 'Télécharger';
    downloadBtn.addEventListener('click', ()=>downloadFile(idx));

    actions.appendChild(viewBtn);
    actions.appendChild(downloadBtn);
    actions.appendChild(removeBtn);
    li.appendChild(name);
    li.appendChild(actions);
    list.appendChild(li);
  });
}

function viewFile(index){
  const f = uploadedFiles[index];
  if(!f) return;
  openStoredFile(f);
}

function downloadFile(index){
  const f = uploadedFiles[index];
  if(!f) return;
  downloadStoredFile(f);
}

function removeFile(index){
  uploadedFiles.splice(index, 1);
  renderFileList();
}

// ================== SUIVI ==================
function renderSuivi(){
  const q = $('filterGlobal')?.value?.toLowerCase() || '';
  const suiviBody = $('suiviBody');
  suiviBody.innerHTML='';
  const allRowsMeta = [];

  getVisibleClients().forEach(c=>{
    c.dossiers.forEach((d, index)=>{
      const procSource = normalizeProcedures(d);
      const procSet = new Set(procSource);
      const tribunalList = Object.values(d.procedureDetails || {})
        .map(p=>String(p?.tribunal || '').trim())
        .filter(Boolean);

      allRowsMeta.push({ procSet, tribunalList });
      if(filterSuiviProcedure !== 'all' && !procSet.has(filterSuiviProcedure)) return;
      if(filterSuiviTribunal !== 'all' && !tribunalList.includes(filterSuiviTribunal)) return;

      if(q){
        const haystack = buildSuiviSearchHaystack(c.name, d, procSource, tribunalList);
        if(!haystack.includes(q)) return;
      }

      suiviBody.innerHTML += `
        <tr>
          <td>${escapeHtml(c.name)}</td>
          <td>${escapeHtml(d.dateAffectation || '-')}</td>
          <td class="procedure-cell">${renderProcedureBadges(procSource)}</td>
          <td>${escapeHtml(d.referenceClient || '-')}</td>
          <td>${escapeHtml(d.debiteur || '-')}</td>
          <td>${escapeHtml(d.montant || '-')}</td>
          <td>${escapeHtml(d.ville || '-')}</td>
          <td>${renderStatusBadge(d.statut || 'En cours')}</td>
          <td>${escapeHtml(d.avancement || '-')}</td>
          <td>
            <button class="btn-primary" onclick="openDossierDetails(${c.id}, ${index})">
              <i class="fa-solid fa-eye"></i>
            </button>
            <button class="btn-primary" onclick="editDossier(${c.id}, ${index})" ${canEditClient(c) ? '' : 'disabled'}>
              <i class="fa-solid fa-pen"></i>
            </button>
            <button class="btn-danger" onclick="deleteDossier(${c.id}, ${index})" ${canDeleteData() ? '' : 'disabled'}>
              <i class="fa-solid fa-trash"></i>
            </button>
          </td>
        </tr>
      `;
    });
  });
  syncSuiviFilterOptions(allRowsMeta);
}

function syncSuiviFilterOptions(rowsMeta){
  const procedureSelect = $('filterSuiviProcedure');
  const tribunalSelect = $('filterSuiviTribunal');
  if(!procedureSelect || !tribunalSelect) return;

  const procedures = new Set();
  const tribunaux = new Set();
  rowsMeta.forEach(row=>{
    row.procSet.forEach(v=>procedures.add(String(v)));
    row.tribunalList.forEach(v=>tribunaux.add(String(v)));
  });

  const sortedProcedures = [...procedures].sort((a,b)=>a.localeCompare(b, 'fr'));
  const sortedTribunaux = [...tribunaux].sort((a,b)=>a.localeCompare(b, 'fr'));

  procedureSelect.innerHTML = `<option value="all">Toutes</option>${sortedProcedures.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('')}`;
  tribunalSelect.innerHTML = `<option value="all">Tous</option>${sortedTribunaux.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('')}`;

  if(filterSuiviProcedure !== 'all' && !procedures.has(filterSuiviProcedure)){
    filterSuiviProcedure = 'all';
  }
  if(filterSuiviTribunal !== 'all' && !tribunaux.has(filterSuiviTribunal)){
    filterSuiviTribunal = 'all';
  }

  procedureSelect.value = filterSuiviProcedure;
  tribunalSelect.value = filterSuiviTribunal;
}

function renderProcedureBadges(procedureText){
  if(!procedureText) return '-';
  const items = Array.isArray(procedureText)
    ? procedureText.map(p=>String(p).trim()).filter(Boolean)
    : procedureText.split(',').map(p=>p.trim()).filter(Boolean);
  if(!items.length) return '-';
  const pills = items.map(name=>{
    let cls = 'proc-autre';
    if(name === 'ASS') cls = 'proc-ass';
    if(name === 'Restitution') cls = 'proc-restitution';
    if(name === 'Nantissement') cls = 'proc-nantissement';
    if(name === 'SFDC') cls = 'proc-sfdc';
    if(name === 'Injonction') cls = 'proc-injonction';
    return `<span class="proc-pill ${cls}">${escapeHtml(name)}</span>`;
  }).join('');
  return `<div class="proc-pill-list">${pills}</div>`;
}

function renderStatusBadge(status){
  const value = String(status || 'En cours');
  let cls = 'status-encours';
  if(value === 'Soldé') cls = 'status-solde';
  if(value === 'Arrêt définitif') cls = 'status-arret';
  if(value === 'Clôture') cls = 'status-cloture';
  if(value === 'Suspension') cls = 'status-suspension';
  return `<span class="status-badge ${cls}">${escapeHtml(value)}</span>`;
}

function buildSuiviSearchHaystack(clientName, dossier, procedures, tribunaux){
  const fileNames = Array.isArray(dossier?.files)
    ? dossier.files.map(f=>String(f?.name || '').trim()).filter(Boolean)
    : [];
  const procedureDetailsValues = collectDeepValues(dossier?.procedureDetails || {});
  const diligenceValues = collectDeepValues(dossier?.diligence || {});
  const staticFields = [
    clientName || '',
    dossier?.debiteur || '',
    dossier?.referenceClient || '',
    dossier?.dateAffectation || '',
    dossier?.procedure || '',
    ...(Array.isArray(dossier?.procedureList) ? dossier.procedureList : []),
    dossier?.ville || '',
    dossier?.adresse || '',
    dossier?.montant || '',
    dossier?.ww || '',
    dossier?.marque || '',
    dossier?.type || '',
    dossier?.note || '',
    dossier?.avancement || '',
    dossier?.statut || '',
    ...(Array.isArray(procedures) ? procedures : []),
    ...(Array.isArray(tribunaux) ? tribunaux : []),
    ...fileNames
  ];
  return [...staticFields, ...procedureDetailsValues, ...diligenceValues]
    .map(v=>String(v).toLowerCase())
    .join(' ');
}

function buildAudienceSearchHaystack(clientName, dossier, procKey, procedureData, draftData){
  const fileNames = Array.isArray(dossier?.files)
    ? dossier.files.map(f=>String(f?.name || '').trim()).filter(Boolean)
    : [];
  const dossierValues = [
    clientName || '',
    dossier?.debiteur || '',
    dossier?.referenceClient || '',
    dossier?.dateAffectation || '',
    dossier?.ville || '',
    dossier?.adresse || '',
    dossier?.montant || '',
    dossier?.ww || '',
    dossier?.marque || '',
    dossier?.type || '',
    dossier?.note || '',
    dossier?.avancement || '',
    dossier?.statut || '',
    dossier?.procedure || '',
    ...(Array.isArray(dossier?.procedureList) ? dossier.procedureList : []),
    ...fileNames
  ];
  const detailsValues = collectDeepValues(dossier?.procedureDetails || {});
  const diligenceValues = collectDeepValues(dossier?.diligence || {});
  const procValues = [
    procKey || '',
    ...collectDeepValues(procedureData || {}),
    ...collectDeepValues(draftData || {})
  ];
  return [...dossierValues, ...detailsValues, ...diligenceValues, ...procValues]
    .map(v=>String(v).toLowerCase())
    .join(' ');
}

function editDossier(clientId, index){
  if(!canEditData()) return alert('Accès refusé');
  const client = AppState.clients.find(c=>c.id == clientId);
  if(!client) return;
  if(!canEditClient(client)) return alert('Accès refusé');
  const d = client.dossiers[index];
  if(!d) return;
  console.log('[EDIT DOSSIER]', JSON.stringify(d, null, 2));

  editingDossier = { clientId, index };
  editingOriginalProcedures = normalizeProcedures(d);
  showView('creation');

  $('selectClient').value = clientId;
  $('debiteurInput').value = d.debiteur || '';
  $('referenceClientInput').value = d.referenceClient || '';
  $('dateAffectation').value = d.dateAffectation || '';
  $('villeInput').value = d.ville || '';
  $('adresseInput').value = d.adresse || '';
  $('montantInput').value = d.montant || '';
  $('wwInput').value = d.ww || '';
  $('marqueInput').value = d.marque || '';
  $('typeInput').value = d.type || '';
  if($('noteInput')) $('noteInput').value = d.note || '';
  if($('avancementInput')) $('avancementInput').value = d.avancement || '';
  if($('statutInput')) $('statutInput').value = d.statut || 'En cours';
  uploadedFiles = Array.isArray(d.files) ? d.files.map(f=>({ ...f })) : [];
  renderFileList();

  document.querySelectorAll('.proc-check').forEach(cb=>cb.checked=false);
  document.querySelectorAll('.checkbox-group label').forEach(l=>l.classList.remove('active'));

  const standard = new Set(['ASS','Restitution','Nantissement','SFDC','Injonction']);
  const procs = normalizeProcedures(d);

  d.procedureList = procs.slice();
  d.procedure = procs.join(', ');
  customProcedures = procs.filter(p=>!standard.has(p));
  $('procedureCustom').value = '';
  renderCustomProcedures();

  procs.forEach(p=>{
    const cb = document.querySelector(`.proc-check[value="${p}"]`);
    if(cb){
      cb.checked = true;
      const label = cb.closest('label');
      if(label) label.classList.add('active');
    }
  });

  const details = d.procedureDetails || {};
  const detailsKeys = Object.keys(details);
  renderProcedureDetails(procs);
  if(!document.querySelectorAll('#procedureDetails .proc-card').length && detailsKeys.length){
    renderProcedureDetails(detailsKeys);
  }

  const detailsMap = {};
  Object.entries(details).forEach(([k, v])=>{
    const nk = String(k || '').trim().toLowerCase();
    if(nk) detailsMap[nk] = v;
  });

  [...document.querySelectorAll('#procedureDetails .proc-card')].forEach(card=>{
    const name = card.querySelector('h4')?.innerText.trim() || '';
    const fields = details[name] || detailsMap[name.toLowerCase()] || {};
    card.querySelectorAll('input, select').forEach(fieldEl=>{
      const key = fieldEl.dataset.field;
      if(key && fields[key] !== undefined) fieldEl.value = fields[key];
    });
  });

  const addBtn = $('addDossierBtn');
  if(addBtn) addBtn.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Mettre à jour';
}

function deleteDossier(clientId, index){
  if(!canDeleteData()) return alert('Seul le gestionnaire peut supprimer un dossier');
  const client = AppState.clients.find(c=>c.id == clientId);
  if(!client) return;
  const dossier = client.dossiers[index];
  if(!dossier) return;
  const ref = dossier.referenceClient || dossier.debiteur || `#${index + 1}`;
  if(!window.confirm(`Supprimer définitivement le dossier "${ref}" ?`)) return;
  client.dossiers.splice(index, 1);
  closeDossierModal();
  renderClients();
  renderDashboard();
  renderSuivi();
  renderAudience();
  renderDiligence();
}

function resetCreationForm(clientId = ''){
  editingDossier = null;
  editingOriginalProcedures = [];
  customProcedures = [];
  uploadedFiles = [];

  document.querySelectorAll('#creationSection input').forEach(i=> i.value='');
  document.querySelectorAll('.proc-check').forEach(cb=>cb.checked=false);
  document.querySelectorAll('.checkbox-group label').forEach(l=>l.classList.remove('active'));
  $('procedureDetails').innerHTML = '';
  $('procedureCustom').value = '';
  renderCustomProcedures();
  if($('noteInput')) $('noteInput').value = '';
  if($('avancementInput')) $('avancementInput').value = '';
  if($('statutInput')) $('statutInput').value = 'En cours';
  if($('fileInput')) $('fileInput').value = '';
  if($('fileList')) $('fileList').innerHTML = '';

  const selectClient = $('selectClient');
  if(selectClient){
    selectClient.value = clientId ? String(clientId) : '';
  }
  const addBtn = $('addDossierBtn');
  if(addBtn) addBtn.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Créer le Dossier';
}

function closeDossierModal(){
  const modal = $('dossierModal');
  const body = $('dossierModalBody');
  if(body) body.innerHTML = '';
  if(modal) modal.style.display = 'none';
}

function openDossierDetails(clientId, index){
  const client = AppState.clients.find(c=>c.id == clientId);
  if(!client) return;
  const dossier = client.dossiers[index];
  if(!dossier) return;

  const body = $('dossierModalBody');
  const modal = $('dossierModal');
  if(!body || !modal) return;

  const files = Array.isArray(dossier.files) ? dossier.files : [];
  const detailsRows = [
    ['Client', client.name || '-'],
    ['Débiteur', dossier.debiteur || '-'],
    ['Référence Client', dossier.referenceClient || '-'],
    ['Date d’affectation', dossier.dateAffectation || '-'],
    ['Procédure', dossier.procedure || '-'],
    ['Montant', dossier.montant || '-'],
    ['Ville', dossier.ville || '-'],
    ['Adresse', dossier.adresse || '-'],
    ['WW', dossier.ww || '-'],
    ['Marque', dossier.marque || '-'],
    ['Type', dossier.type || '-'],
    ['Statut', dossier.statut || 'En cours'],
    ['Avancement', dossier.avancement || '-'],
    ['Note', dossier.note || '-']
  ];

  const detailsHtml = detailsRows.map(([label, value])=>`
    <div class="details-row">
      <div class="details-label">${escapeHtml(label)}</div>
      <div class="details-value">${label === 'Statut' ? renderStatusBadge(value) : escapeHtml(value)}</div>
    </div>
  `).join('');

  const filesHtml = files.length
    ? files.map((f, i)=>`
      <div class="details-file-item">
        <span>${escapeHtml(String(f?.name || `Fichier ${i + 1}`))}</span>
        <span class="details-file-actions">
          <button type="button" class="btn-primary" onclick="viewDossierFile(${client.id}, ${index}, ${i})">Voir</button>
          <button type="button" class="btn-success" onclick="downloadDossierFile(${client.id}, ${index}, ${i})">Télécharger</button>
        </span>
      </div>
    `).join('')
    : '<div class="details-empty">Aucun document.</div>';

  body.innerHTML = `
    <div class="details-grid">
      ${detailsHtml}
    </div>
    <div class="details-files">
      <h3><i class="fa-solid fa-paperclip"></i> Documents</h3>
      ${filesHtml}
    </div>
    <div class="details-actions">
      <button type="button" class="btn-primary" onclick="downloadDossierSummary(${client.id}, ${index})">
        <i class="fa-solid fa-file-arrow-down"></i> Télécharger fiche dossier
      </button>
      <button type="button" class="btn-danger" onclick="deleteDossier(${client.id}, ${index})" ${canDeleteData() ? '' : 'disabled'}>
        <i class="fa-solid fa-trash"></i> Supprimer dossier
      </button>
    </div>
  `;

  modal.style.display = 'flex';
}

function getDossierByIds(clientId, index){
  const client = AppState.clients.find(c=>c.id == clientId);
  if(!client) return null;
  const dossier = client.dossiers[index];
  if(!dossier) return null;
  return { client, dossier };
}

function viewDossierFile(clientId, index, fileIndex){
  const data = getDossierByIds(clientId, index);
  if(!data) return;
  const file = data.dossier.files?.[fileIndex];
  openStoredFile(file);
}

function downloadDossierFile(clientId, index, fileIndex){
  const data = getDossierByIds(clientId, index);
  if(!data) return;
  const file = data.dossier.files?.[fileIndex];
  downloadStoredFile(file);
}

function downloadDossierSummary(clientId, index){
  const data = getDossierByIds(clientId, index);
  if(!data) return;
  const { client, dossier } = data;
  const files = Array.isArray(dossier.files) ? dossier.files : [];
  const lines = [
    `Client: ${client.name || '-'}`,
    `Debiteur: ${dossier.debiteur || '-'}`,
    `Reference Client: ${dossier.referenceClient || '-'}`,
    `Date affectation: ${dossier.dateAffectation || '-'}`,
    `Procedure: ${dossier.procedure || '-'}`,
    `Montant: ${dossier.montant || '-'}`,
    `Ville: ${dossier.ville || '-'}`,
    `Adresse: ${dossier.adresse || '-'}`,
    `WW: ${dossier.ww || '-'}`,
    `Marque: ${dossier.marque || '-'}`,
    `Type: ${dossier.type || '-'}`,
    `Statut: ${dossier.statut || 'En cours'}`,
    `Avancement: ${dossier.avancement || '-'}`,
    `Note: ${dossier.note || '-'}`,
    '',
    'Documents:'
  ];
  files.forEach((f, i)=>{
    lines.push(`${i + 1}. ${f?.name || 'Sans nom'} (${Math.round(Number(f?.size || 0) / 1024)} KB)`);
  });
  if(!files.length) lines.push('Aucun document');

  const blob = new Blob([lines.join('\n')], { type: 'text/plain;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `dossier_${client.id}_${index + 1}.txt`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ================== DILIGENCE ==================
function getDiligenceRows(){
  const rows = [];
  getVisibleClients().forEach(c=>{
    c.dossiers.forEach((d, di)=>{
      const procedures = normalizeProcedures(d);
      procedures.forEach(proc=>{
        if(proc !== 'SFDC' && proc !== 'Injonction') return;
        const details = d?.procedureDetails?.[proc] || {};
        const tribunal = String(details.tribunal || '').trim();
        rows.push({
          clientId: c.id,
          dossierIndex: di,
          clientName: c.name || '',
          dossier: d,
          procedure: proc,
          details,
          tribunal
        });
      });
    });
  });
  return rows;
}

function syncDiligenceProcedureFilter(rows){
  const select = $('diligenceProcedureFilter');
  if(!select) return;
  const set = new Set();
  rows.forEach(r=>{
    if(r.procedure) set.add(String(r.procedure));
  });
  const sorted = [...set].sort((a,b)=>a.localeCompare(b, 'fr'));
  select.innerHTML = `<option value="all">Toutes</option>${sorted.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('')}`;
  if(filterDiligenceProcedure !== 'all' && !set.has(filterDiligenceProcedure)){
    filterDiligenceProcedure = 'all';
  }
  select.value = filterDiligenceProcedure;
}

function renderDiligence(){
  const body = $('diligenceBody');
  const count = $('diligenceCount');
  if(!body) return;
  const allRows = getDiligenceRows();
  syncDiligenceProcedureFilter(allRows);
  const q = $('diligenceSearchInput')?.value?.toLowerCase() || '';

  const rows = allRows.filter(row=>{
    if(filterDiligenceProcedure !== 'all' && row.procedure !== filterDiligenceProcedure) return false;
    if(!q) return true;
    const haystack = buildSuiviSearchHaystack(
      row.clientName,
      row.dossier,
      [row.procedure],
      row.tribunal ? [row.tribunal] : []
    );
    return haystack.includes(q);
  });

  if(count){
    const label = filterDiligenceProcedure === 'all' ? 'toutes les procédures' : `procédure: ${filterDiligenceProcedure}`;
    count.textContent = `${rows.length} ligne(s) diligence (${label})`;
  }

  if(!rows.length){
    body.innerHTML = '<tr><td colspan="11" class="diligence-empty">Aucun dossier SFDC/Injonction trouvé.</td></tr>';
    return;
  }

  body.innerHTML = rows.map(row=>`
    <tr>
      <td>${escapeHtml(row.clientName || '-')}</td>
      <td>${escapeHtml(row.dossier?.debiteur || '-')}</td>
      <td>${escapeHtml(row.details?.dateDepot || '-')}</td>
      <td>${escapeHtml(row.details?.referenceClient || row.dossier?.referenceClient || '-')}</td>
      <td>${escapeHtml(row.details?.attOrdOrOrdOk || '-')}</td>
      <td>${escapeHtml(row.details?.executionNo || '-')}</td>
      <td>${escapeHtml(row.dossier?.ville || '-')}</td>
      <td>${escapeHtml(row.details?.attDelegationOuDelegat || '-')}</td>
      <td>${escapeHtml(row.details?.huissier || '-')}</td>
      <td>${escapeHtml(row.details?.sort || '-')}</td>
      <td>${escapeHtml(row.tribunal || '-')}</td>
    </tr>
  `).join('');
}

// ================== EQUIPE ==================
function getClientNameById(id){
  return AppState.clients.find(c=>c.id === id)?.name || '-';
}

function getSelectedTeamClientIds(){
  return [...document.querySelectorAll('#teamClientsList input[type="checkbox"]:checked')]
    .map(i=>Number(i.value))
    .filter(v=>Number.isFinite(v));
}

function renderTeamClientCheckboxes(selectedIds = []){
  const box = $('teamClientsList');
  if(!box) return;
  if(!AppState.clients.length){
    box.innerHTML = '<div class="diligence-empty">Aucun client pour le moment.</div>';
    return;
  }
  box.innerHTML = AppState.clients.map(c=>{
    const checked = selectedIds.includes(c.id) ? 'checked' : '';
    return `<label class="team-client-item"><input type="checkbox" value="${c.id}" ${checked}> ${escapeHtml(c.name)}</label>`;
  }).join('');
}

function updateTeamClientSelectorState(){
  const role = $('teamRole')?.value || 'viewer';
  const wrap = $('teamClientsWrap');
  if(!wrap) return;
  const disabled = role !== 'viewer';
  wrap.style.opacity = disabled ? '0.6' : '1';
  wrap.querySelectorAll('input[type="checkbox"]').forEach(i=> i.disabled = disabled);
}

function addClientFromTeam(){
  if(!canManageTeam()) return alert('Accès refusé');
  const input = $('teamNewClientName');
  if(!input) return;
  const name = input.value.trim();
  if(!name) return alert('Nom client obligatoire');

  const selectedBefore = getSelectedTeamClientIds();
  const existing = AppState.clients.find(c=>c.name.trim().toLowerCase() === name.toLowerCase());
  const clientId = existing ? existing.id : Date.now();

  if(!existing){
    AppState.clients.push({ id: clientId, name, dossiers: [] });
  }

  const selectedAfter = [...new Set([...selectedBefore, clientId])];
  renderTeamClientCheckboxes(selectedAfter);
  updateTeamClientSelectorState();
  input.value = '';

  renderClients();
  updateClientDropdown();
  renderDashboard();
  renderSuivi();
  renderAudience();
  renderDiligence();
}

function resetTeamForm(){
  editingTeamUserId = null;
  if($('teamUsername')) $('teamUsername').value = '';
  if($('teamPassword')) $('teamPassword').value = '';
  if($('teamRole')) $('teamRole').value = 'viewer';
  if($('teamNewClientName')) $('teamNewClientName').value = '';
  renderTeamClientCheckboxes([]);
  updateTeamClientSelectorState();
  if($('teamSaveBtn')) $('teamSaveBtn').innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Enregistrer';
}

function saveTeamUser(){
  if(!canManageTeam()) return alert('Accès refusé');
  const username = $('teamUsername')?.value?.trim() || '';
  const password = $('teamPassword')?.value?.trim() || '';
  const role = $('teamRole')?.value || 'viewer';
  if(!username) return alert('Username obligatoire');
  if(!editingTeamUserId && !password) return alert('Mot de passe obligatoire');

  const selectedClientIds = role === 'manager' ? [] : getSelectedTeamClientIds();
  const finalClientIds = role === 'viewer' ? selectedClientIds : [];
  if(role === 'viewer' && finalClientIds.length === 0){
    return alert('Choisir au moins un client pour ce compte client');
  }

  const usernameTaken = USERS.some(u=>u.username === username && u.id !== editingTeamUserId);
  if(usernameTaken) return alert('Username déjà utilisé');

  if(editingTeamUserId){
    const user = USERS.find(u=>u.id === editingTeamUserId);
    if(!user) return;
    user.username = username;
    if(password) user.password = password;
    user.role = role;
    user.clientIds = finalClientIds;
  }else{
    USERS.push({
      id: Date.now(),
      username,
      password,
      role,
      clientIds: finalClientIds
    });
  }
  renderEquipe();
  resetTeamForm();
}

function editTeamUser(userId){
  if(!canManageTeam()) return;
  const user = USERS.find(u=>u.id === userId);
  if(!user) return;
  editingTeamUserId = userId;
  if($('teamUsername')) $('teamUsername').value = user.username;
  if($('teamPassword')) $('teamPassword').value = '';
  if($('teamRole')) $('teamRole').value = user.role;
  renderTeamClientCheckboxes(Array.isArray(user.clientIds) ? user.clientIds : []);
  updateTeamClientSelectorState();
  if($('teamSaveBtn')) $('teamSaveBtn').innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Mettre à jour';
}

function deleteTeamUser(userId){
  if(!canManageTeam()) return;
  const user = USERS.find(u=>u.id === userId);
  if(!user) return;
  const managerCount = USERS.filter(u=>u.role === 'manager').length;
  if(user.role === 'manager' && managerCount <= 1){
    return alert('Impossible de supprimer le dernier manager');
  }
  if(currentUser?.id === userId){
    return alert('Impossible de supprimer l’utilisateur connecté');
  }
  USERS = USERS.filter(u=>u.id !== userId);
  renderEquipe();
  if(editingTeamUserId === userId) resetTeamForm();
}

function renderEquipe(){
  const panel = $('teamManagerPanel');
  const locked = $('teamLocked');
  const body = $('teamUsersBody');
  if(!panel || !locked || !body) return;

  if(!canManageTeam()){
    panel.style.display = 'none';
    locked.style.display = '';
    body.innerHTML = '';
    return;
  }

  panel.style.display = '';
  locked.style.display = 'none';
  if(!editingTeamUserId) {
    renderTeamClientCheckboxes([]);
    updateTeamClientSelectorState();
  }

  body.innerHTML = USERS.map(u=>{
    const roleLabel = u.role === 'viewer' ? 'client' : u.role;
    const clients = (Array.isArray(u.clientIds) ? u.clientIds : [])
      .map(getClientNameById)
      .join(', ') || '-';
    return `
      <tr>
        <td>${escapeHtml(u.username)}</td>
        <td>${escapeHtml(roleLabel)}</td>
        <td>${escapeHtml(clients)}</td>
        <td>
          <button class="btn-primary" onclick="editTeamUser(${u.id})"><i class="fa-solid fa-pen"></i></button>
          <button class="btn-danger" onclick="deleteTeamUser(${u.id})"><i class="fa-solid fa-trash"></i></button>
        </td>
      </tr>
    `;
  }).join('');
}

function normalizeProcedures(d){
  let list = [];
  if(Array.isArray(d.procedureList)) list = d.procedureList;
  else if(typeof d.procedureList === 'string') list = d.procedureList.split(',');

  if(!list.length){
    if(Array.isArray(d.procedure)) list = d.procedure;
    else if(typeof d.procedure === 'string') list = d.procedure.split(',');
  }

  if(!list.length && d.procedure) list = [String(d.procedure)];

  const fromDetails = Object.keys(d.procedureDetails || {});
  list = list.concat(fromDetails);

  const cleaned = list.map(v=>String(v).trim()).filter(Boolean);
  return [...new Set(cleaned)];
}

function isAudienceProcedure(procName){
  const value = String(procName || '').trim().toLowerCase();
  if(!value) return false;
  return value !== 'sfdc' && value !== 'injonction';
}

function parseDateForAge(value){
  if(!value) return null;
  const raw = String(value).trim();
  if(!raw) return null;
  const isoInText = raw.match(/(\d{4}-\d{2}-\d{2})/);
  if(isoInText){
    return parseDateForAge(isoInText[1]);
  }
  const m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(m){
    const y = Number(m[1]);
    const mo = Number(m[2]) - 1;
    const d = Number(m[3]);
    return new Date(y, mo, d);
  }
  const fr = raw.match(/^(\d{2})[/-](\d{2})[/-](\d{4})$/);
  if(fr){
    const d = Number(fr[1]);
    const mo = Number(fr[2]) - 1;
    const y = Number(fr[3]);
    return new Date(y, mo, d);
  }
  const dt = new Date(raw);
  if(Number.isNaN(dt.getTime())) return null;
  return dt;
}

function toAgeDays(value){
  const dt = parseDateForAge(value);
  if(!dt) return null;
  const now = new Date();
  const start = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()).getTime();
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const days = Math.floor((end - start) / (1000 * 60 * 60 * 24));
  return days >= 0 ? days : 0;
}

function getDashboardAttSortRows(){
  const rows = [];
  AppState.clients.forEach((c, ci)=>{
    if(!canViewClient(c)) return;
    c.dossiers.forEach((d, di)=>{
      const procKeys = normalizeProcedures(d);
      procKeys.forEach(procKey=>{
        if(!isAudienceProcedure(procKey)) return;
        const p = getAudienceProcedure(ci, di, procKey);
        if(String(p?.color || '') !== 'blue') return;
        const audienceDate = String(p?.audience || '').trim() || '-';
        rows.push({
          client: c.name || '-',
          debiteur: d.debiteur || '-',
          procedure: procKey || '-',
          audienceDate
        });
      });
    });
  });
  return rows;
}

function getDashboardCalendarEvents(){
  const byDate = {};
  AppState.clients.forEach((c, ci)=>{
    if(!canViewClient(c)) return;
    c.dossiers.forEach((d, di)=>{
      const procKeys = normalizeProcedures(d);
      procKeys.forEach(procKey=>{
        if(!isAudienceProcedure(procKey)) return;
        const p = getAudienceProcedure(ci, di, procKey);
        const dt = parseDateForAge(p?.audience || '');
        if(!dt) return;
        const y = dt.getFullYear();
        const m = String(dt.getMonth() + 1).padStart(2, '0');
        const day = String(dt.getDate()).padStart(2, '0');
        const key = `${y}-${m}-${day}`;
        if(!byDate[key]) byDate[key] = [];
        byDate[key].push({
          client: c.name || '-',
          procedure: procKey || '-',
          debiteur: d.debiteur || '-'
        });
      });
    });
  });
  return byDate;
}

function renderDashboardCalendar(){
  const label = $('calendarMonthLabel');
  const grid = $('dashboardCalendarGrid');
  if(!label || !grid) return;

  const year = dashboardCalendarCursor.getFullYear();
  const month = dashboardCalendarCursor.getMonth();
  label.textContent = dashboardCalendarCursor.toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' });

  const eventsByDate = getDashboardCalendarEvents();
  const headers = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'];
  const firstDay = new Date(year, month, 1);
  const firstWeekday = (firstDay.getDay() + 6) % 7;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const today = new Date();
  const todayKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;

  let html = headers.map(h=>`<div class="dashboard-calendar-weekday">${h}</div>`).join('');
  for(let i = 0; i < firstWeekday; i++){
    html += '<div class="dashboard-calendar-day is-empty"></div>';
  }
  for(let day = 1; day <= daysInMonth; day++){
    const key = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    const events = eventsByDate[key] || [];
    const classes = ['dashboard-calendar-day'];
    if(key === todayKey) classes.push('is-today');
    if(events.length) classes.push('has-event');
    const tooltip = events
      .map(e=>`${e.client} / ${e.procedure} / ${e.debiteur}`)
      .join(' | ');
    html += `
      <div class="${classes.join(' ')}" title="${escapeAttr(tooltip)}">
        <span class="day-num">${day}</span>
        ${events.length ? `<span class="day-count">${events.length}</span>` : ''}
      </div>
    `;
  }
  grid.innerHTML = html;
}

// ================== DASHBOARD ==================
function renderDashboard(){
  const visibleClients = getVisibleClients();
  $('totalClients').innerText = visibleClients.length;
  let enCours = 0;
  let clotureCount = 0;
  const attSortRows = getDashboardAttSortRows();

  visibleClients.forEach(c=>{
    c.dossiers.forEach(d=>{
      enCours += 1;
      if(d.statut === 'Clôture') clotureCount += 1;
    });
  });

  $('dossiersEnCours').innerText = enCours;
  $('dossiersTermines').innerText = clotureCount;
  if($('dossiersAttSort')) $('dossiersAttSort').innerText = attSortRows.length;
  renderDashboardCalendar();
}

// ================== AUDIENCE ==================
function renderAudience(){
  const body = $('audienceBody');
  body.innerHTML='';

  const allRows = getAudienceRows();
  syncAudienceFilterOptions(allRows);
  const rows = allRows.filter(row=>{
    const tribunal = (row.p.tribunal || '').trim();
    if(filterAudienceProcedure !== 'all' && row.procKey !== filterAudienceProcedure) return false;
    if(filterAudienceTribunal !== 'all' && tribunal !== filterAudienceTribunal) return false;
    return true;
  });

  rows.forEach(row=>{
    const { c, d, procKey, p, color, key, draft } = row;
    const canEdit = canEditClient(c) && canEditData();
    const safeColor = ['blue', 'green', 'red', 'yellow', 'purple-dark', 'purple-light'].includes(color) ? color : '';
    const procKeyEncoded = encodeURIComponent(String(procKey));
    const keyEncoded = encodeURIComponent(String(key));
    const isChecked = !!safeColor;
    body.innerHTML += `
      <tr class="color-${safeColor}">
        <td>
          <input type="checkbox" class="audience-check"
            data-ci="${row.ci}"
            data-di="${row.di}"
            data-proc-key="${procKeyEncoded}"
            ${isChecked ? 'checked' : ''}
            ${canEdit ? '' : 'disabled'}
            onchange="setAudienceColorEncoded(${row.ci},${row.di},'${procKeyEncoded}', this.checked)">
        </td>
        <td>${escapeHtml(c.name)}</td>
        <td>${escapeHtml(d.debiteur||'-')}</td>
        <td><input value="${escapeAttr(draft.refDossier||p.referenceClient||'')}" ${canEdit ? '' : 'readonly'} oninput="updateAudienceDraftFromEncoded('${keyEncoded}','refDossier',this.value)"></td>
        <td><input value="${escapeAttr(draft.dateAudience||p.audience||'')}" ${canEdit ? '' : 'readonly'} oninput="updateAudienceDraftFromEncoded('${keyEncoded}','dateAudience',this.value)"></td>
        <td><input value="${escapeAttr(draft.juge||p.juge||'')}" ${canEdit ? '' : 'readonly'} oninput="updateAudienceDraftFromEncoded('${keyEncoded}','juge',this.value)"></td>
        <td><input value="${escapeAttr(draft.sort||p.sort||'')}" ${canEdit ? '' : 'readonly'} oninput="updateAudienceDraftFromEncoded('${keyEncoded}','sort',this.value)"></td>
        <td>${escapeHtml(p.tribunal||'-')}</td>
        <td>${escapeHtml(procKey||'-')}</td>
        <td>${escapeHtml(p.dateDepot||'-')}</td>
      </tr>
    `;
  });
}

function syncAudienceFilterOptions(rows){
  const procedureSelect = $('filterAudienceProcedure');
  const tribunalSelect = $('filterAudienceTribunal');
  if(!procedureSelect || !tribunalSelect) return;

  const procedureSet = new Set();
  const tribunalSet = new Set();
  rows.forEach(row=>{
    if(row.procKey) procedureSet.add(String(row.procKey));
    const tribunal = String(row.p.tribunal || '').trim();
    if(tribunal) tribunalSet.add(tribunal);
  });

  const procedures = [...procedureSet].sort((a, b)=>a.localeCompare(b, 'fr'));
  const tribunaux = [...tribunalSet].sort((a, b)=>a.localeCompare(b, 'fr'));

  procedureSelect.innerHTML = `<option value="all">Toutes</option>${procedures.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('')}`;
  tribunalSelect.innerHTML = `<option value="all">Tous</option>${tribunaux.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('')}`;

  if(filterAudienceProcedure !== 'all' && !procedureSet.has(filterAudienceProcedure)){
    filterAudienceProcedure = 'all';
  }
  if(filterAudienceTribunal !== 'all' && !tribunalSet.has(filterAudienceTribunal)){
    filterAudienceTribunal = 'all';
  }

  procedureSelect.value = filterAudienceProcedure;
  tribunalSelect.value = filterAudienceTribunal;
}

function getAudienceRows(){
  const q = $('filterAudience')?.value?.toLowerCase() || '';
  const rows = [];
  AppState.clients.forEach((c, ci)=>{
    if(!canViewClient(c)) return;
    c.dossiers.forEach((d, di)=>{
      let procKeys = Object.keys(d.procedureDetails || {});
      if(!procKeys.length){
        procKeys = normalizeProcedures(d);
      }
      if(!procKeys.length) return;
      procKeys.forEach(procKey=>{
        if(!isAudienceProcedure(procKey)) return;
        const p = getAudienceProcedure(ci, di, procKey);
        const color = p.color || '';
        if(filterAudienceColor !== 'all' && color !== filterAudienceColor) return;
        const key = makeAudienceDraftKey(ci, di, procKey);
        const draft = audienceDraft[key] || {};
        const haystack = buildAudienceSearchHaystack(c.name, d, procKey, p, draft);
        if(q && !haystack.includes(q)) return;
        rows.push({ c, d, procKey, p, color, key, draft, ci, di });
      });
    });
  });
  return rows;
}

function exportAudienceXLS(){
  const headers = [
    'Client',
    'Adversaire',
    'N° Dossier',
    'Tribunal',
    'Instruction',
    'Sort'
  ];
  const audienceRows = getAudienceRows();
  const dateAudience = audienceRows
    .map(r => r.draft.dateAudience || r.p.audience || '')
    .find(v => v && String(v).trim()) || '';

  const rows = audienceRows.map(r=>{
    const p = r.p;
    const d = r.d;
    const draft = r.draft;
    const sortValue = draft.sort || p.sort || '';
    return [
      r.c.name || '',
      d.debiteur || '',
      draft.refDossier || p.referenceClient || '',
      p.tribunal || '',
      sortValue,
      ''
    ];
  });

  const styles = `
    <style>
      body { font-family: Arial, sans-serif; }
      .title { text-align: center; font-size: 20px; font-weight: bold; color: #1e3a8a; margin: 8px 0 2px; }
      .subtitle { text-align: center; font-size: 14px; font-weight: 600; color: #334155; margin: 0 0 12px; }
      table { border-collapse: collapse; width: 100%; }
      th, td { border: 1px solid #d1d5db; padding: 10px 12px; font-size: 23px; }
      th { background: #1e40af; color: #ffffff; font-weight: bold; font-size: 23px; }
      tr:nth-child(even) td { background: #f1f5f9; }
    </style>
  `;

  const thead = `<tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join('')}</tr>`;
  const tbody = rows.map(r=>`<tr>${r.map(c=>`<td>${escapeHtml(String(c))}</td>`).join('')}</tr>`).join('');
  const html = `
    <html>
      <head>${styles}</head>
      <body>
        <div class="title">CABINET ARAQUI HOUSSAINI</div>
        ${dateAudience ? `<div class="subtitle">Date d'audience : ${escapeHtml(String(dateAudience))}</div>` : `<div class="subtitle">Date d'audience : -</div>`}
        <table>
          <thead>${thead}</thead>
          <tbody>${tbody}</tbody>
        </table>
      </body>
    </html>
  `;

  const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'audience_export.xls';
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(str){
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setAudienceColor(ci, di, procKey, checked){
  const client = AppState.clients?.[ci];
  if(!canEditData() || !canEditClient(client)) return;
  const p = getAudienceProcedure(ci, di, procKey);
  const allowed = new Set(['blue', 'green', 'red', 'yellow', 'purple-dark', 'purple-light']);
  if(!checked){
    p.color = '';
    renderDashboard();
    renderAudience();
    return;
  }
  if(selectedAudienceColor === 'all' || !allowed.has(selectedAudienceColor)){
    renderDashboard();
    renderAudience();
    return;
  }
  p.color = selectedAudienceColor;
  renderDashboard();
  renderAudience();
}

function setAudienceColorEncoded(ci, di, procKeyEncoded, checked){
  setAudienceColor(ci, di, decodeURIComponent(String(procKeyEncoded)), checked);
}

function getAudienceProcedure(ci, di, procKey){
  const dossier = AppState.clients?.[ci]?.dossiers?.[di];
  if(!dossier) return {};
  if(!dossier.procedureDetails) dossier.procedureDetails = {};
  if(!dossier.procedureDetails[procKey]) dossier.procedureDetails[procKey] = {};
  return dossier.procedureDetails[procKey];
}

function updateAudienceDraft(key, field, value){
  if(!canEditData()) return;
  if(!audienceDraft[key]) audienceDraft[key] = {};
  audienceDraft[key][field] = value;
}

function updateAudienceDraftFromEncoded(keyEncoded, field, value){
  updateAudienceDraft(decodeURIComponent(String(keyEncoded)), field, value);
}

function saveAllAudience(){
  if(!canEditData()) return alert('Accès refusé');
  Object.entries(audienceDraft).forEach(([key, data])=>{
    const { ci, di, procKey } = parseAudienceDraftKey(key);
    const p = getAudienceProcedure(ci, di, procKey);
    if(data.refDossier!==undefined) p.referenceClient = data.refDossier;
    if(data.dateAudience!==undefined) p.audience = data.dateAudience;
    if(data.juge!==undefined) p.juge = data.juge;
    if(data.sort!==undefined) p.sort = data.sort;
  });

  audienceDraft = {};
  renderDashboard();
  renderAudience();
  renderSuivi();
  alert('Toutes les modifications ont été enregistrées ✅');
}

// ================== PROCEDURE DETAILS ==================
function collectProcedureDraft(){
  const draft = {};
  document.querySelectorAll('#procedureDetails .proc-card').forEach(card=>{
    const name = card.querySelector('h4')?.innerText.trim();
    if(!name) return;
    draft[name] = {};
    card.querySelectorAll('input, select').forEach(fieldEl=>{
      draft[name][fieldEl.dataset.field] = fieldEl.value;
    });
  });
  return draft;
}

function updateInjonctionNotificationDateVisibility(card){
  if(!card) return;
  const notifSelect = card.querySelector('select[data-field="notificationStatus"]');
  const dateWrap = card.querySelector('.notif-date-wrap');
  if(!notifSelect || !dateWrap) return;
  const value = String(notifSelect.value || '').trim().toLowerCase();
  const show = value === 'notifier' || value === 'nb';
  dateWrap.style.display = show ? '' : 'none';
}

function renderProcedureDetails(forceList){
  const container = $('procedureDetails');
  const draft = collectProcedureDraft();
  container.innerHTML='';
  const selected = Array.isArray(forceList) && forceList.length
    ? forceList.slice()
    : [...document.querySelectorAll('.proc-check:checked')].map(cb=>cb.value);
  if(!forceList || !forceList.length) selected.push(...customProcedures);
  const activeLabels = [...document.querySelectorAll('.checkbox-group label.active')]
    .map(l=>l.dataset.proc)
    .filter(Boolean);
  activeLabels.forEach(p=>{
    if(!selected.includes(p)) selected.push(p);
  });
  for(let i=selected.length-1;i>=0;i--){
    if(selected[i] === 'Autre') selected.splice(i,1);
  }
  Object.keys(draft).forEach(p=>{
    if(!selected.includes(p)) selected.push(p);
  });
  if(!forceList || !forceList.length){
    const prev = editingDossier
      ? AppState.clients.find(c=>c.id == editingDossier.clientId)?.dossiers?.[editingDossier.index]
      : null;
    const prevList = normalizeProcedures(prev || {});
    prevList.forEach(p=>{
      if(!selected.includes(p)) selected.push(p);
    });
    editingOriginalProcedures.forEach(p=>{
      if(!selected.includes(p)) selected.push(p);
    });
  }

  const finalList = [...new Set(selected.map(v=>String(v).trim()).filter(Boolean))];

  finalList.forEach(proc=>{
    if(!proc || !String(proc).trim()) return;
    let procClass = '';
    if(proc === 'ASS') procClass = 'proc-ass';
    if(proc === 'Restitution') procClass = 'proc-restitution';
    if(proc === 'Nantissement') procClass = 'proc-nantissement';
    if(proc === 'SFDC') procClass = 'proc-sfdc';
    if(proc === 'Injonction') procClass = 'proc-injonction';
    if(customProcedures.includes(proc)) procClass = 'proc-autre';

    const div = document.createElement('div');
    div.className = `proc-card ${procClass}`;
    let fieldsHtml = `
      <input type="text" data-field="dateDepot" placeholder="Date d'affectation">
      <input type="text" data-field="depotLe" placeholder="Dépôt le">
      <input type="text" data-field="referenceClient" placeholder="Référence dossier">
      <input type="text" data-field="audience" placeholder="Audience">
      <input type="text" data-field="juge" placeholder="Juge">
      <input type="text" data-field="sort" placeholder="Sort">
      <input type="text" data-field="tribunal" placeholder="Tribunal">
    `;
    if(proc === 'SFDC'){
      fieldsHtml = `
        <input type="text" data-field="dateDepot" placeholder="Date d'affectation">
        <input type="text" data-field="depotLe" placeholder="Dépôt le">
        <input type="text" data-field="referenceClient" placeholder="Référence dossier">
        <input type="text" data-field="attOrdOrOrdOk" placeholder="att ord / ord ok">
        <input type="text" data-field="executionNo" placeholder="Execution N°">
        <input type="text" data-field="attDelegationOuDelegat" placeholder="att delegation ou delegat">
        <input type="text" data-field="huissier" placeholder="Huissier">
        <input type="text" data-field="sort" placeholder="Sort">
        <input type="text" data-field="tribunal" placeholder="Tribunal">
      `;
    }
    if(proc === 'Injonction'){
      fieldsHtml = `
        <input type="text" data-field="dateDepot" placeholder="Date d'affectation">
        <input type="text" data-field="depotLe" placeholder="Dépôt le">
        <input type="text" data-field="referenceClient" placeholder="Référence dossier">
        <input type="text" data-field="attOrdOrOrdOk" placeholder="att ord / ord ok">
        <input type="text" data-field="notificationNo" placeholder="Notification N°">
        <select data-field="notificationStatus">
          <option value="att plie avec tr">att plie avec tr</option>
          <option value="att notif">att notif</option>
          <option value="notifier">notifier</option>
          <option value="NB">NB</option>
        </select>
        <div class="notif-date-wrap">
          <input type="date" data-field="dateNotification" placeholder="Date notification">
        </div>
        <select data-field="certificatNonAppelStatus">
          <option value="">Certificat non appel</option>
          <option value="att certificat non appel">att certificat non appel</option>
          <option value="certificat non appel ok">certificat non appel ok</option>
        </select>
        <input type="text" data-field="executionNo" placeholder="Execution N°">
        <input type="text" data-field="huissier" placeholder="Huissier">
        <input type="text" data-field="sort" placeholder="Sort">
        <input type="text" data-field="tribunal" placeholder="Tribunal">
      `;
    }
    const title = document.createElement('h4');
    title.textContent = proc;
    const grid = document.createElement('div');
    grid.className = 'proc-grid';
    grid.innerHTML = fieldsHtml;
    div.appendChild(title);
    div.appendChild(grid);
    container.appendChild(div);
    if(draft[proc]){
      div.querySelectorAll('input, select').forEach(fieldEl=>{
        const key = fieldEl.dataset.field;
        if(key && draft[proc][key] !== undefined) fieldEl.value = draft[proc][key];
      });
    }
    if(proc === 'Injonction'){
      const notifSelect = div.querySelector('select[data-field="notificationStatus"]');
      if(notifSelect){
        notifSelect.addEventListener('change', ()=>updateInjonctionNotificationDateVisibility(div));
      }
      updateInjonctionNotificationDateVisibility(div);
    }
  });

  if(!forceList || !forceList.length){
    suppressProcedureChange = true;
    finalList.forEach(p=>{
      const cb = document.querySelector(`.proc-check[value="${p}"]`);
      if(cb){
        cb.checked = true;
        const label = cb.closest('label');
        if(label) label.classList.add('active');
      }
    });
    suppressProcedureChange = false;
  }
}

function addCustomProcedure(){
  const input = $('procedureCustom');
  if(!input) return;
  const value = input.value.trim();
  if(!value) return;
  if(customProcedures.includes(value)) return;
  customProcedures.push(value);
  input.value = '';
  renderCustomProcedures();
  renderProcedureDetails();
}

function removeCustomProcedure(value){
  customProcedures = customProcedures.filter(v=>v !== value);
  renderCustomProcedures();
  renderProcedureDetails();
}

function renderCustomProcedures(){
  const container = $('customProcedures');
  if(!container) return;
  container.innerHTML = '';
  customProcedures.forEach(v=>{
    const chip = document.createElement('span');
    chip.className = 'custom-proc custom-red';
    chip.appendChild(document.createTextNode(`${v} `));
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = 'x';
    btn.addEventListener('click', ()=>removeCustomProcedure(v));
    chip.appendChild(btn);
    container.appendChild(chip);
  });
}

document.addEventListener('DOMContentLoaded', initApplication);
