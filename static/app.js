let current=null;let data=[];
async function load(){
  const qs=new URLSearchParams({instalacao:fInst.value,medidor:fMd.value,nome_cliente:fNome.value,status:fStatus.value,operador:(typeof fOperador!=='undefined'?fOperador.value:''),data_ini:(typeof fDataIni!=='undefined'?fDataIni.value:''),data_fim:(typeof fDataFim!=='undefined'?fDataFim.value:'')});
  data=await (await fetch('/api/ars?'+qs)).json();
  count.textContent=data.length+' registros encontrados';
  rows.innerHTML=data.map(r=>`<tr>
<td><input type="checkbox" data-id="${r.id}"></td>
<td>${r.created_at||'-'}</td>
<td>${r.instalacao||'-'}</td>
<td>${r.medidor||'-'}</td>
<td>${r.nome_cliente||'-'}</td>
<td>${r.operador_nome||r.operador_usuario||'-'}</td>
<td>${r.filename||'-'}</td>
<td><span class="badge ${r.status==='Erro OCR'?'Erro':r.status}">${r.status||'Pendente'}</span></td>
<td><div class="act"><button class="blue" onclick="view(${r.id})" title="Visualizar">Ver</button><button onclick="location.href='/api/download/${r.id}'" title="Baixar">Baixar</button><button onclick="quickStatus(${r.id},'Conferido')" title="Marcar como conferido">OK</button><button onclick="completeBaseRow(${r.id})" title="Buscar na base">Base</button><button class="red" onclick="del(${r.id})" title="Excluir">Excluir</button></div></td>
</tr>`).join('');
  if(current){ const still=data.find(x=>x.id===current.id); if(still){ current=still; pName.value=still.filename; pInst.value=still.instalacao||''; pMd.value=still.medidor||''; pNome.value=still.nome_cliente||''; } }
}
function clearFilters(){fInst.value='';fMd.value='';fNome.value='';if(typeof fOperador!=='undefined')fOperador.value='';if(typeof fDataIni!=='undefined')fDataIni.value='';if(typeof fDataFim!=='undefined')fDataFim.value='';fStatus.value='Todos';load()}
function setProgress(title,text,pct){progressTitle.textContent=title;progressText.textContent=text;progressFill.style.width=pct+'%';progressPct.textContent=pct+'%';}
function showProgress(title,text){progressOverlay.classList.add('open');setProgress(title,text,0)}
function hideProgress(){setTimeout(()=>progressOverlay.classList.remove('open'),250)}
async function uploadFiles(files){
  if(!files || !files.length) return;
  const valid=[...files].filter(file=>(file.name||'').toLowerCase().match(/\.(pdf|jpg|jpeg|png)$/));
  if(!valid.length) return;
  const label = valid.length===1 ? valid[0].name : valid.length+' arquivo(s) selecionado(s)';
  pdfLabel.textContent=label; folderLabel.textContent=label;
  showProgress('Importando arquivos','Preparando importação...',0);
  let total=0;
  for(let i=0;i<valid.length;i++){
    const file=valid[i];
    setProgress('Importando arquivos','Enviando '+file.name, Math.round((i/valid.length)*90));
    let fd=new FormData(); fd.append('file',file);
    await fetch('/api/upload',{method:'POST',body:fd}); total++;
  }
  setProgress('Importando arquivos','Atualizando registros...',95);
  document.getElementById('file').value='';
  document.getElementById('folder').value='';
  await load();
  setProgress('Importação concluída',total+' arquivo(s) importado(s).',100);
  if(!silent) hideProgress();
}
async function uploadFiles(files){
  if(!files || !files.length) return;
  let total=0;
  for(const file of files){
    const name=(file.name||'').toLowerCase();
    if(!name.match(/\.(pdf|jpg|jpeg|png)$/)) continue;
    let fd=new FormData(); fd.append('file',file);
    await fetch('/api/upload',{method:'POST',body:fd}); total++;
  }
  document.getElementById('file').value='';
  document.getElementById('folder').value='';
  await load();
  if(total) alert(total+' arquivo(s) importado(s).');
}
async function uploadFile(file){ return uploadFiles(file?[file]:[]); }
async function uploadBase(file){
  if(!file)return;
  baseFileLabel.textContent=file.name;
  showProgress('Carregando base XLSX','Lendo planilha base...',10);
  let fd=new FormData();fd.append('file',file);
  let r=await (await fetch('/api/upload-base-cache',{method:'POST',body:fd})).json();
  setProgress('Carregando base XLSX','Montando lista de abas...',60);
  sheetSelect.innerHTML='<option value="TODAS">Todas as abas</option>'+r.sheets.map(s=>`<option>${s}</option>`).join('');
  baseInfo.textContent='Base carregada: '+r.sheets.length+' aba(s). Cabeçalho esperado na linha 5, a partir da coluna B.';
  await loadColumns();
  setProgress('Base carregada','Selecione a aba/colunas e clique em Processar Base.',100);
  hideProgress();
}
async function loadColumns(){
  colInst.innerHTML='<option value="">Automático</option>'; colMd.innerHTML='<option value="">Automático</option>'; colNome.innerHTML='<option value="">Automático</option>';
  if(sheetSelect.value==='TODAS') return;
  let r=await (await fetch('/api/base/columns?sheet='+encodeURIComponent(sheetSelect.value))).json();
  for(const c of r.columns){ if(!c) continue; colInst.innerHTML+=`<option>${c}</option>`; colMd.innerHTML+=`<option>${c}</option>`; colNome.innerHTML+=`<option>${c}</option>`; }
}
async function processBase(){
  showProgress('Processando base XLSX','Importando dados para pesquisa...',5);
  let fd=new FormData();fd.append('sheet',sheetSelect.value);fd.append('col_instalacao',colInst.value);fd.append('col_medidor',colMd.value);fd.append('col_nome_cliente',colNome.value);
  let r=await (await fetch('/api/import-base',{method:'POST',body:fd})).json();
  setProgress('Base processada',(r.count||0)+' registro(s) importado(s).',100);
  hideProgress();
}
function view(id){current=data.find(x=>x.id===id);document.querySelector('.preview').classList.add('open');viewer.src='/api/view/'+id;pName.value=current.filename;pInst.value=current.instalacao||'';pMd.value=current.medidor||'';pNome.value=current.nome_cliente||''}
function closePreview(){document.querySelector('.preview').classList.remove('open');viewer.src='';current=null}
async function saveRename(){if(!current)return;const id=current.id;let fd=new FormData();fd.append('instalacao',pInst.value);fd.append('medidor',pMd.value);fd.append('nome_cliente',pNome.value);await fetch('/api/rename/'+id,{method:'POST',body:fd});await load();const row=data.find(x=>x.id===id);if(row){view(id)}else{closePreview()}}
async function completeBase(){
  if(!current)return;
  let fd=new FormData();
  fd.append('instalacao',pInst.value);
  fd.append('medidor',pMd.value);
  fd.append('nome_cliente',pNome.value);
  let r=await (await fetch('/api/base/complete/'+current.id,{method:'POST',body:fd})).json();
  if(!r.ok) alert(r.message||'Não encontrado na base');
  await load();view(current.id)
}
async function completeBaseRow(id){let fd=new FormData();fd.append('nome_cliente','');let r=await (await fetch('/api/base/complete/'+id,{method:'POST',body:fd})).json();if(!r.ok) alert(r.message||'Não encontrado na base');await load()}
async function setStatus(s){if(!current)return;await quickStatus(current.id,s);await load();view(current.id)}
async function quickStatus(id,s){let fd=new FormData();fd.append('status',s);await fetch('/api/status/'+id,{method:'POST',body:fd});load()}
function downloadCurrent(){if(current)location.href='/api/download/'+current.id}
async function del(id){
  if(!confirm('Excluir este AR?')) return;
  const resp = await fetch('/api/ars/'+id,{method:'DELETE'});
  let data = {};
  try{ data = await resp.json(); }catch(e){}
  if(data.api_message && !data.api_deleted && data.api_message !== 'Sem vínculo com API'){
    alert('Excluído do painel, mas não consegui excluir da API: ' + data.api_message);
  }
  if(current && current.id===id) closePreview();
  load();
}
function downloadSelected(){document.querySelectorAll('tbody input[type="checkbox"][data-id]:checked').forEach(c=>window.open('/api/download/'+c.dataset.id,'_blank'))}
load();


// V1.0.0.7 - menu retrátil + tema claro/escuro + upload com progresso preservado
function toggleMenu(){ document.body.classList.toggle('menu-collapsed'); localStorage.setItem('arPanelMenuCollapsed', document.body.classList.contains('menu-collapsed')?'1':'0'); }
function setThemeMode(mode){
  document.body.classList.toggle('dark', mode==='dark');
  localStorage.setItem('arPanelTheme', mode);
  document.querySelectorAll('.themePill').forEach(b=>b.classList.remove('active'));
  const btn = [...document.querySelectorAll('.themePill')].find(b => (mode==='light' && b.textContent.includes('☀')) || (mode==='dark' && b.textContent.includes('☾')));
  if(btn) btn.classList.add('active');
}
(function initVisualPrefs(){
  if(localStorage.getItem('arPanelMenuCollapsed')==='1') document.body.classList.add('menu-collapsed');
  setThemeMode(localStorage.getItem('arPanelTheme') || 'light');
})();

// sobrescreve a função duplicada anterior e mantém overlay de progresso
uploadFiles = async function(files){
  if(!files || !files.length) return;
  const valid=[...files].filter(file=>(file.name||'').toLowerCase().match(/\.(pdf|jpg|jpeg|png)$/));
  if(!valid.length) return;
  const label = valid.length===1 ? valid[0].name : valid.length+' arquivo(s) selecionado(s)';
  pdfLabel.textContent=label; folderLabel.textContent=label;
  showProgress('Importando arquivos','Preparando importação...',0);
  let total=0;
  for(let i=0;i<valid.length;i++){
    const file=valid[i];
    setProgress('Importando arquivos','Enviando '+file.name, Math.round((i/valid.length)*90));
    let fd=new FormData(); fd.append('file',file);
    await fetch('/api/upload',{method:'POST',body:fd}); total++;
  }
  setProgress('Importando arquivos','Atualizando registros...',95);
  document.getElementById('file').value='';
  document.getElementById('folder').value='';
  await load();
  setProgress('Importação concluída',total+' arquivo(s) importado(s).',100);
  hideProgress();
}


// V1.0.0.8 - menu lateral retrátil por mouse hover
function enableHoverMenu(){
  const side = document.getElementById('sidebar');
  if(!side) return;
  if(!document.body.classList.contains('menu-hover-enabled')){
    document.body.classList.add('menu-hover-enabled','menu-collapsed');
  }
  side.addEventListener('mouseenter', () => {
    document.body.classList.remove('menu-collapsed');
  });
  side.addEventListener('mouseleave', () => {
    document.body.classList.add('menu-collapsed');
  });
}
(function initHoverMenu(){
  // Sempre inicia recolhido com DS; abre somente ao passar o mouse.
  localStorage.setItem('arPanelMenuCollapsed','1');
  enableHoverMenu();
})();


// V1.0.1.0 - ajustes finos: temas, overlay com ESC, OK fecha e navegação anterior/próximo
function toggleThemeMenu(){
  const el=document.getElementById('themeMenu');
  if(el) el.classList.toggle('open');
}
function setAccentTheme(theme){
  document.body.classList.remove('accent-green','accent-blue','accent-rose');
  document.body.classList.add('accent-'+theme);
  localStorage.setItem('arPanelAccent', theme);
  const el=document.getElementById('themeMenu');
  if(el) el.classList.remove('open');
}
(function initAccentTheme(){
  setAccentTheme(localStorage.getItem('arPanelAccent') || 'green');
})();
function currentIndex(){
  if(!current || !data || !data.length) return -1;
  return data.findIndex(x=>String(x.id)===String(current.id));
}
function prevRecord(){
  const i=currentIndex();
  if(i>0) view(data[i-1].id);
}
function nextRecord(){
  const i=currentIndex();
  if(i>=0 && i<data.length-1) view(data[i+1].id);
}
document.addEventListener('keydown', function(ev){
  if(ev.key === 'Escape') closePreview();
  if(document.querySelector('.preview.open')){
    if(ev.key === 'ArrowLeft') prevRecord();
    if(ev.key === 'ArrowRight') nextRecord();
  }
});
// Sobrescreve OK para confirmar e fechar o overlay
setStatus = async function(s){
  if(!current)return;
  const id=current.id;
  let fd=new FormData();
  fd.append('status',s);
  await fetch('/api/status/'+id,{method:'POST',body:fd});
  await load();
  if(s==='Conferido') closePreview();
  else {
    const row=data.find(x=>String(x.id)===String(id));
    if(row) view(id); else closePreview();
  }
}
// Reforço para salvar nome sem perder o overlay e sem esconder ações
saveRename = async function(){
  if(!current)return;
  const id=current.id;
  let fd=new FormData();
  fd.append('instalacao',pInst.value);
  fd.append('medidor',pMd.value);
  fd.append('nome_cliente',pNome.value);
  await fetch('/api/rename/'+id,{method:'POST',body:fd});
  await load();
  const row=data.find(x=>String(x.id)===String(id));
  if(row){ view(id); } else { closePreview(); }
}



function normalizeApiUrl(url){
  let u=(url||'').trim().replace(/\/+$/,'');
  if(!u) return 'https://dsystem-ar-api.onrender.com';
  if(u === 'dsystem-ar-api.onrender.com') return 'https://dsystem-ar-api.onrender.com';
  if(u.startsWith('http://dsystem-ar-api.onrender.com')) return 'https://dsystem-ar-api.onrender.com';
  if(/^\d+\.\d+\.\d+\.\d+/.test(u)) return 'http://' + u;
  if((u.includes('192.168.') || u.includes('10.17.') || u.includes('localhost')) && u.startsWith('https://')) return u.replace('https://','http://');
  return u;
}

function defaultApiUrl(){
  return 'https://dsystem-ar-api.onrender.com';
}

async function receberApi(){
  const apiAtual = localStorage.getItem('dsystem_ar_api_url') || defaultApiUrl();
  const apiUrl = prompt('URL da API DSYSTEM AR:', apiAtual);
  if(!apiUrl) return;

  const cleanUrl = normalizeApiUrl(apiUrl);
  localStorage.setItem('dsystem_ar_api_url', cleanUrl);

  showProgress('Recebendo da API','Consultando registros...',5);
  try{
    const fd = new FormData();
    fd.append('api_url', cleanUrl);
    setProgress('Recebendo da API','Baixando arquivos da API...',35);

    const resp = await fetch('/api/receber-api', {method:'POST', body:fd});
    const data = await resp.json();

    if(!resp.ok){
      throw new Error(data.detail || JSON.stringify(data));
    }

    setProgress('Recebendo da API','Atualizando painel...',90);
    await load();
    setProgress('Recebimento concluído',`${data.importados} importado(s), ${data.ignorados} já existente(s).`,100);
    hideProgress();

    if(data.erros && data.erros.length){
      alert('Recebido com alguns avisos:\n' + data.erros.join('\n'));
    }
  }catch(e){
    hideProgress();
    alert('Falha ao receber da API: ' + (e.message || e));
  }
}

async function atualizarComApi(silent=false){
  const apiUrl = localStorage.getItem('dsystem_ar_api_url') || defaultApiUrl();

  if(apiUrl){
    if(!silent) showProgress('Atualizando','Verificando novos registros na API...',10);
    try{
      const fd = new FormData();
      fd.append('api_url', normalizeApiUrl(apiUrl));
      if(!silent) setProgress('Atualizando','Recebendo novos registros...',45);
      await fetch('/api/receber-api', {method:'POST', body:fd});
      if(!silent) setProgress('Atualizando','Atualizando painel...',90);
    }catch(e){
      console.warn('Falha ao atualizar pela API:', e);
      if(!silent) alert('Falha ao atualizar pela API: ' + (e.message || e));
    }

    await load();

    if(!silent){
      setProgress('Atualização concluída','Painel atualizado.',100);
      hideProgress();
    }
  }else{
    await load();
  }
}

function toggleAllRows(master){
  document.querySelectorAll('tbody input[type="checkbox"][data-id]').forEach(c=>{c.checked=master.checked;});
}





let autoRefreshTimer = null;
const AUTO_REFRESH_ENABLED_KEY = 'dsystem_ar_auto_refresh_enabled';
const AUTO_REFRESH_MINUTES_KEY = 'dsystem_ar_auto_refresh_minutes';
const AUTO_REFRESH_COLLAPSED_KEY = 'dsystem_ar_auto_refresh_collapsed';

function getAutoRefreshEnabled(){
  return localStorage.getItem(AUTO_REFRESH_ENABLED_KEY) === '1';
}

function getAutoRefreshMinutes(){
  const v = Number(localStorage.getItem(AUTO_REFRESH_MINUTES_KEY) || '5');
  return [5,10,15].includes(v) ? v : 5;
}

function updateAutoRefreshStatus(){
  const el = document.getElementById('autoRefreshStatus');
  if(!el) return;
  if(getAutoRefreshEnabled()){
    el.textContent = 'Ativo: verificando uploads a cada ' + getAutoRefreshMinutes() + ' minuto(s).';
  }else{
    el.textContent = 'Automático desativado.';
  }
}

function saveAutoRefreshConfig(){
  const enabled = document.getElementById('autoRefreshEnabled');
  const minutes = document.getElementById('autoRefreshMinutes');

  if(enabled) localStorage.setItem(AUTO_REFRESH_ENABLED_KEY, enabled.checked ? '1' : '0');
  if(minutes) localStorage.setItem(AUTO_REFRESH_MINUTES_KEY, minutes.value || '5');

  updateAutoRefreshStatus();
  startAutoRefreshTimer();
}

function applyAutoRefreshConfig(){
  const enabled = document.getElementById('autoRefreshEnabled');
  const minutes = document.getElementById('autoRefreshMinutes');

  if(enabled) enabled.checked = getAutoRefreshEnabled();
  if(minutes) minutes.value = String(getAutoRefreshMinutes());

  applyAutoRefreshCollapse();
  updateAutoRefreshStatus();
  startAutoRefreshTimer();
}

function startAutoRefreshTimer(){
  if(autoRefreshTimer){
    clearInterval(autoRefreshTimer);
    autoRefreshTimer = null;
  }

  if(!getAutoRefreshEnabled()) return;

  const ms = getAutoRefreshMinutes() * 60 * 1000;
  autoRefreshTimer = setInterval(async ()=>{
    try{
      await atualizarComApi(true);
      updateAutoRefreshStatus();
    }catch(e){
      console.warn('Atualização automática falhou:', e);
    }
  }, ms);
}

function toggleAutoRefreshPanel(){
  const box = document.getElementById('autoRefreshPanelFields');
  const btn = document.getElementById('btnToggleAutoRefreshPanel');
  if(!box || !btn) return;

  const collapsed = !box.classList.contains('collapsed');
  box.classList.toggle('collapsed', collapsed);
  btn.textContent = collapsed ? 'EXIBIR' : 'RECOLHER';
  localStorage.setItem(AUTO_REFRESH_COLLAPSED_KEY, collapsed ? '1' : '0');
}

function applyAutoRefreshCollapse(){
  const box = document.getElementById('autoRefreshPanelFields');
  const btn = document.getElementById('btnToggleAutoRefreshPanel');
  if(!box || !btn) return;

  const collapsed = localStorage.getItem(AUTO_REFRESH_COLLAPSED_KEY) === '1';
  box.classList.toggle('collapsed', collapsed);
  btn.textContent = collapsed ? 'EXIBIR' : 'RECOLHER';
}

const PANEL_API_KEY='dsystem_ar_api_url';
const PANEL_SESSION_KEY='dsystem_ar_panel_session_api_v1';

function getPanelApiUrl(){
  let saved=(localStorage.getItem(PANEL_API_KEY)||'').replace(/\/+$/,'').trim();

  // Corrige URLs antigas/salvas erradas
  if(!saved || saved.includes('192.168.') || saved.includes('10.17.') || saved.includes('localhost') || saved.startsWith('http://dsystem-ar-api.onrender.com')){
    saved='https://dsystem-ar-api.onrender.com';
    localStorage.setItem(PANEL_API_KEY, saved);
  }

  return saved;
}


function getUserInitials(user){
  if(!user) return 'DS';
  const source = (user.nome || user.usuario || '').trim();
  if(!source) return 'DS';
  const parts = source.split(/\s+/).filter(Boolean);
  if(parts.length >= 2){
    return (parts[0][0] + parts[1][0]).toUpperCase();
  }
  const cleaned = source.replace(/[^A-Za-zÀ-ÿ0-9]/g, '');
  return cleaned.substring(0,2).toUpperCase() || 'DS';
}

function applyUserInitials(){
  const el = document.getElementById('userInitialsAvatar');
  if(!el) return;
  const user = currentPanelUser ? currentPanelUser() : null;
  el.textContent = getUserInitials(user);
}

function currentPanelUser(){
  return JSON.parse(localStorage.getItem(PANEL_SESSION_KEY)||'null');
}

function savePanelSession(user){
  localStorage.setItem(PANEL_SESSION_KEY, JSON.stringify(user));
}

function panelApi(path, options={}){
  return fetch(getPanelApiUrl()+path, options).then(async resp=>{
    const text=await resp.text();
    let data;
    try{ data=JSON.parse(text); }catch{ data=text; }
    if(!resp.ok) throw new Error((data&&data.detail)?data.detail:(typeof data==='string'?data:JSON.stringify(data)));
    return data;
  });
}

async function renderPanelUsers(){
  const wrap=document.getElementById('panelUsers');
  if(!wrap) return;
  wrap.innerHTML='<p>Carregando usuários...</p>';
  try{
    const users=await panelApi('/api/users');
    wrap.innerHTML=users.map(u=>`
      <div class="panel-user-row">
        <div>
          <strong>${u.nome}</strong><br>
          <small>${u.usuario} · ${u.perfil}</small>
        </div>
        <div class="panel-user-actions">
          <button class="pass" onclick="changePanelPassword('${u.usuario}')">Senha</button>
          ${u.usuario==='admin'?'':`<button class="del" onclick="deletePanelUser('${u.usuario}')">Excluir</button>`}
        </div>
      </div>
    `).join('');
  }catch(e){
    wrap.innerHTML='<p>Erro ao carregar usuários.</p>';
    alert('Falha ao carregar usuários na API: '+e.message+'\n\nSe a API estiver dormindo no Render, abra https://dsystem-ar-api.onrender.com e tente novamente.');
  }
}

async function createPanelUser(){
  const current=currentPanelUser();
  if(!current || current.perfil!=='admin'){
    alert('Apenas admin pode criar usuários.');
    return;
  }
  const usuario=document.getElementById('newUser').value.trim();
  const nome=document.getElementById('newName').value.trim();
  const senha=document.getElementById('newPass').value.trim();
  const perfil=document.getElementById('newRole').value;
  if(!usuario || !nome || !senha){
    alert('Preencha todos os campos.');
    return;
  }
  try{
    await panelApi('/api/users', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({usuario,nome,senha,perfil})
    });
    document.getElementById('newUser').value='';
    document.getElementById('newName').value='';
    document.getElementById('newPass').value='';
    await renderPanelUsers();
    alert('Usuário criado.');
  }catch(e){
    alert('Falha ao criar usuário: '+e.message);
  }
}

async function deletePanelUser(usuario){
  if(!confirm('Excluir usuário?')) return;
  try{
    await panelApi('/api/users/'+encodeURIComponent(usuario), {method:'DELETE'});
    await renderPanelUsers();
  }catch(e){
    alert('Falha ao excluir usuário: '+e.message);
  }
}

async function changePanelPassword(usuario){
  const nova=prompt('Nova senha:');
  if(!nova) return;
  try{
    await panelApi('/api/admin/change-password', {
      method:'PUT',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({usuario,nova_senha:nova})
    });
    alert('Senha alterada.');
  }catch(e){
    alert('Falha ao trocar senha: '+e.message);
  }
}

function toggleCreatePanelUser(){
  const box=document.getElementById('createPanelUserFields');
  const btn=document.getElementById('btnToggleCreatePanelUser');
  if(!box || !btn) return;
  const collapsed=box.classList.toggle('collapsed');
  btn.textContent=collapsed?'EXIBIR':'RECOLHER';
  localStorage.setItem('dsystem_ar_panel_create_collapsed', collapsed?'1':'0');
}

function applyCreatePanelUserCollapse(){
  const box=document.getElementById('createPanelUserFields');
  const btn=document.getElementById('btnToggleCreatePanelUser');
  if(!box || !btn) return;
  const collapsed=localStorage.getItem('dsystem_ar_panel_create_collapsed')==='1';
  box.classList.toggle('collapsed', collapsed);
  btn.textContent=collapsed?'EXIBIR':'RECOLHER';
}

async function openConfigPanel(){
  const current=currentPanelUser();
  if(!current || current.perfil!=='admin'){
    alert('Apenas admin pode acessar configuração.');
    return;
  }
  document.getElementById('configOverlay').classList.remove('hidden');
  applyCreatePanelUserCollapse();
  applyAutoRefreshConfig();
  await renderPanelUsers();
}

function closeConfigPanel(){
  document.getElementById('configOverlay').classList.add('hidden');
}


function localPanelLogin(usuario, senha){
  if(usuario === 'admin' && senha === 'admin123'){
    return {usuario:'admin', nome:'Administrador', perfil:'admin', local_fallback:true};
  }
  return null;
}
function isApiConnectionError(err){
  const msg = String((err && err.message) || err || '').toLowerCase();
  return msg.includes('failed to fetch') || msg.includes('network') || msg.includes('load failed');
}

async function doPanelLogin(){
  const usuario=document.getElementById('loginUser').value.trim();
  const senha=document.getElementById('loginPass').value;

  if(!usuario || !senha){
    document.getElementById('loginError').textContent='Informe usuário e senha.';
    return;
  }

  document.getElementById('loginError').textContent='Conectando...';

  try{
    const user=await panelApi('/api/login', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({usuario,senha})
    });
    savePanelSession(user);
    document.getElementById('loginError').textContent='';
    document.getElementById('loginOverlay').style.display='none';
    applyUserInitials();
    startAutoRefreshTimer();
  }catch(e){
    const fallback = localPanelLogin(usuario, senha);
    if(fallback && isApiConnectionError(e)){
      savePanelSession(fallback);
      document.getElementById('loginError').textContent='';
      document.getElementById('loginOverlay').style.display='none';
    applyUserInitials();
      alert('Entrou em modo local. A API não respondeu agora. Para criar usuários/sincronizar, verifique a API Render.');
      return;
    }
    document.getElementById('loginError').textContent=e.message||'Usuário ou senha inválidos.';
  }
}

function logoutPanel(){
  if(confirm('Sair do painel?')){
    localStorage.removeItem(PANEL_SESSION_KEY);
    if(autoRefreshTimer){clearInterval(autoRefreshTimer);autoRefreshTimer=null;}
    document.getElementById('loginPass').value='';
    applyUserInitials();
    document.getElementById('loginOverlay').style.display='flex';
  }
}

document.getElementById('btnLoginPanel').onclick=doPanelLogin;
document.getElementById('loginPass').addEventListener('keydown',e=>{ if(e.key==='Enter') doPanelLogin(); });

(function initPanelLogin(){
  const session=currentPanelUser();
  if(session){
    document.getElementById('loginOverlay').style.display='none';
    applyUserInitials();
    startAutoRefreshTimer();
  }
})();


try{ applyUserInitials(); }catch(e){}
