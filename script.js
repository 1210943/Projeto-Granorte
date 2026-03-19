// ─── FIREBASE ────────────────────────────────────────────────────────────────
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged }
  from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";
import { getFirestore, doc, getDoc, setDoc, onSnapshot }
  from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

const _fbApp = initializeApp({
  apiKey: "AIzaSyAYlN7EZvk2KN7O8IvG8YMj638KGcWhLlA",
  authDomain: "projeto-final-507ed.firebaseapp.com",
  projectId: "projeto-final-507ed",
  storageBucket: "projeto-final-507ed.firebasestorage.app",
  messagingSenderId: "995940038628",
  appId: "1:995940038628:web:c203a98f9143a9eb6539f1"
});
const _auth = getAuth(_fbApp);
const _db   = getFirestore(_fbApp);
const _DOC  = ()=> doc(_db,'wms','estado');

let _saveTimeout=null;
function saveState(){
  if(_saveTimeout)clearTimeout(_saveTimeout);
  _saveTimeout=setTimeout(async()=>{
    try{ await setDoc(_DOC(),S); }
    catch(e){ console.error('Firestore save:',e); }
  },800);
}
async function loadState(){
  try{
    const snap=await getDoc(_DOC());
    return snap.exists()?snap.data():{...DEFAULT_STATE};
  }catch(e){ return{...DEFAULT_STATE}; }
}
// ─────────────────────────────────────────────────────────────────────────────

let _chegadaId=null,_alocarId=null;

const DEFAULT_STATE={
  depositos:[],baias:[],descargas:[],fornecedores:[],operadores:[],
  turnoAtivo:{turno:'A',operadorId:''},
  baixas:[],recebimentos:[],movimentacoes:[],
  capDiaria:160,toleranciaMin:30,
  emailAlerta:"",horarioAlerta:"07:00",alertaAtivo:false,
  ocorrencias:[],
  acertosPeso:[],
  conferencias:[],
  materiais:[],transferencias:[],
  turnos:[
    {id:'t1',nome:'Turno A',inicio:'06:00',fim:'14:00',cor:'blue'},
    {id:'t2',nome:'Turno B',inicio:'14:00',fim:'22:00',cor:'green'},
    {id:'t3',nome:'Turno C',inicio:'22:00',fim:'06:00',cor:'amber'}
  ]
};

let S={...DEFAULT_STATE}; // preenchido após login pelo _boot()



function uid(){return 'x'+Math.random().toString(36).substr(2,9)}
function hoje(){return new Date().toISOString().slice(0,10)}
function fmtDate(d){if(!d)return'—';const[y,m,day]=d.split('-');return`${day}/${m}/${y}`}
function fmtDT(d){if(!d)return'—';return new Date(d).toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'})}

function fmtKg(v){
  // valores internos em toneladas → converter para kg (*1000)
  const n=Number(v)||0;
  const kg=n*1000;
  return kg.toLocaleString('pt-BR',{maximumFractionDigits:0})+'kg';
}
function fmtKgInput(v){
  // formata enquanto o usuário digita: remove não-dígitos e aplica pontos de milhar
  const raw=String(v).replace(/\D/g,'');
  if(!raw)return '';
  return Number(raw).toLocaleString('pt-BR');
}
function parseKg(v){
  // converte string formatada "10.000" para número 10000
  return Number(String(v).replace(/\./g,'').replace(',','.').replace(/\D/g,''))||0;
}
function stars(n){return'★'.repeat(n)+'☆'.repeat(5-n)}
function getCapDia(data){return S.descargas.filter(d=>d.data===data).reduce((s,d)=>s+Number(d.toneladas),0)}
function opNome(id){const o=S.operadores.find(x=>x.id===id);return o?o.nome:'—'}
function turnoLabel(t){const found=S.turnos?S.turnos.find(x=>x.id===t||x.nome===t):null;if(found)return found.nome+' ('+found.inicio+'–'+found.fim+')';return t||'—';}

function getAtrasados(){
  const hj=hoje();const agora=new Date();const tol=S.toleranciaMin||30;
  return S.descargas.filter(d=>{
    if(d.status!=='pendente'||d.data!==hj)return false;
    const[hh,mm]=d.hora.split(':').map(Number);
    const prev=new Date();prev.setHours(hh,mm,0,0);
    return(agora-prev)/60000>tol;
  });
}

// Map old page names to module + subtab
var _pageToModule={
  'agendamento':['loginterna','agendamento'],
  'recebimento':['loginterna','recebimento'],
  'depositos':['loginterna','depositos'],
  'movimentacoes':['loginterna','movimentacoes'],
  'transferencias':['loginterna','transferencias'],
  'ocorrencias':['loginterna','ocorrencias'],
  'fornecedores':['pcp','fornecedores'],
  'estoque':['pcp','estoque'],
  'baixa':['pcp','baixa'],
  'materiais':['pcp','materiais'],
  'conferencia':['pcp','conferencia'],
  'pcp_acerto':['pcp','pcp'],
};

function showPage(p){
  // Redirect legacy page names to their module
  if(_pageToModule[p]){
    var mod=_pageToModule[p][0];var tab=_pageToModule[p][1];
    showPage(mod);
    setTimeout(function(){showSubTab(mod,tab);},10);
    return;
  }
  document.querySelectorAll('.page').forEach(function(el){el.classList.remove('active');});
  document.querySelectorAll('.nav-tab').forEach(function(el){el.classList.remove('active');});
  var pageEl=document.getElementById('page-'+p);
  if(pageEl)pageEl.classList.add('active');
  var tabs=['dashboard','kpis','loginterna','pcp','relatorios','config'];
  var idx=tabs.indexOf(p);
  if(idx>=0)document.querySelectorAll('.nav-tab')[idx].classList.add('active');
  render();
}
function openModal(id){
  document.getElementById(id).classList.remove('hidden');
  if(id==='modal-ocorrencia'){abrirModalOcorrencia(null);return;}
  if(id==='modal-acerto-peso'){abrirModalAcerto(null);return;}
  if(id==='modal-conferencia'){abrirModalConferencia(null);return;}
  if(id==='modal-agendar'){
    populateAgBaias();populateSelForn();populateSelOp('ag-operador');
    const sm=document.getElementById('ag-material-sel');
    if(sm){sm.innerHTML='<option value="">Selecionar cadastrado</option>';S.materiais.forEach(function(m){sm.innerHTML+='<option value="'+m.nome+'">'+m.nome+'</option>';});}
  }
  if(id==='modal-baia'){const s=document.getElementById('baia-dep');s.innerHTML='<option value="">— Selecionar —</option>';S.depositos.forEach(d=>s.innerHTML+=`<option value="${d.id}">${d.nome}</option>`);}
  if(id==='modal-transferencia'){
    const sf=document.getElementById('tr-forn-sel');
    sf.innerHTML='<option value="">— Selecionar cadastrado —</option>';
    S.fornecedores.forEach(f=>sf.innerHTML+=`<option value="${f.id}">${f.nome}</option>`);
    const sm=document.getElementById('tr-material');
    sm.innerHTML='<option value="">— Selecionar —</option>';
    S.materiais.forEach(m=>sm.innerHTML+=`<option value="${m.nome}">${m.nome}</option>`);
    const sb=document.getElementById('tr-baia');
    sb.innerHTML='<option value="">— Sem sugestao —</option>';
    S.baias.filter(b=>!b.estoque).forEach(b=>{const dep=S.depositos.find(d=>d.id===b.dep);sb.innerHTML+=`<option value="${b.nome}">${b.nome}${dep?' ('+dep.nome+')':''}</option>`;});
  }
}
function closeModal(id){document.getElementById(id).classList.add('hidden')}

function populateSelForn(){
  const s=document.getElementById('ag-forn-sel');if(!s)return;
  s.innerHTML='<option value="">— Selecionar cadastrado —</option>';
  S.fornecedores.forEach(f=>s.innerHTML+=`<option value="${f.id}">${f.nome}</option>`);
}
function preencherForn(){
  const fId=document.getElementById('ag-forn-sel').value;if(!fId)return;
  const f=S.fornecedores.find(x=>x.id===fId);
  if(f){document.getElementById('ag-fornecedor').value=f.nome;if(f.material&&!document.getElementById('ag-material').value)document.getElementById('ag-material').value=f.material;}
}
function populateSelOp(selId){
  const s=document.getElementById(selId);if(!s)return;
  s.innerHTML='<option value="">— Selecionar —</option>';
  S.operadores.forEach(o=>s.innerHTML+=`<option value="${o.id}">${o.nome} (${o.cargo})</option>`);
  if(S.turnoAtivo.operadorId)s.value=S.turnoAtivo.operadorId;
}
function populateAgBaias(){
  const s=document.getElementById('ag-baia');if(!s)return;
  s.innerHTML='<option value="">— Sem sugestão —</option>';
  S.baias.filter(b=>!b.estoque).forEach(b=>{const dep=S.depositos.find(d=>d.id===b.dep);s.innerHTML+=`<option value="${b.id}">${b.nome}${dep?' ('+dep.nome+')':''}${b.material?' — '+b.material:''}</option>`;});
}
function populateBaixaBaias(){
  const s=document.getElementById('baixa-baia');if(!s)return;
  s.innerHTML='<option value="">— Selecionar baia —</option>';
  S.baias.filter(b=>b.estoques&&b.estoques.length>0).forEach(b=>{const tot=b.estoques.reduce((s,e)=>s+e.qtdAtual,0);s.innerHTML+=`<option value="${b.id}">${b.nome} — ${b.estoques.length} produto(s) (${fmtKg(tot)})</option>`;});
}

function checkCapModal(){
  const data=document.getElementById('ag-data').value;
  const ton=parseKg(document.getElementById('ag-toneladas').value)||0;
  const el=document.getElementById('ag-cap-alert');if(!data||!ton){el.innerHTML='';return;}
  const atual=getCapDia(data);const total=atual+ton;const cap=S.capDiaria||160;
  if(total>cap)el.innerHTML=`<div class="alert-bar alert-danger">Capacidade de ${fmtKg(cap)} será ultrapassada (${fmtKg(total)} previsto).</div>`;
  else el.innerHTML=`<div class="alert-bar alert-success">Disponível em ${fmtDate(data)}: ${fmtKg(cap-atual)}.</div>`;
}

function agendarDescarga(){
  const f=document.getElementById('ag-fornecedor').value.trim();
  const m=document.getElementById('ag-material').value.trim();
  const t=parseKg(document.getElementById('ag-toneladas').value);
  const d=document.getElementById('ag-data').value;
  const h=document.getElementById('ag-hora').value;
  if(!f||!m||!t||!d||!h){alert('Preencha os campos obrigatórios (*).');return;}
  const unid='kg';
  const tonVal=Number(t)||0;
  S.descargas.push({id:uid(),fornecedor:f,material:m,toneladas:tonVal,unidade:unid,qtdOriginal:Number(t),data:d,hora:h,lote:document.getElementById('ag-lote').value,baia:document.getElementById('ag-baia').value,obs:document.getElementById('ag-obs').value,status:'pendente',nf:document.getElementById('ag-nf').value,operador:document.getElementById('ag-operador').value});
  ['ag-fornecedor','ag-material','ag-toneladas','ag-data','ag-hora','ag-lote','ag-obs','ag-nf'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});// ag-unidade removed;const _ms=document.getElementById('ag-material-sel');if(_ms)_ms.value='';
  document.getElementById('ag-baia').value='';document.getElementById('ag-forn-sel').value='';document.getElementById('ag-cap-alert').innerHTML='';
  saveState();closeModal('modal-agendar');render();
}

function abrirChegada(dcId){
  const dc=S.descargas.find(d=>d.id===dcId);if(!dc)return;
  _chegadaId=dcId;
  const baia=dc.baia?S.baias.find(b=>b.id===dc.baia):null;
  const isAtrasado=!!getAtrasados().find(d=>d.id===dcId);
  document.getElementById('chegada-resumo').innerHTML=`
    <div style="background:var(--bg3);border-radius:var(--radius);padding:10px 12px">
      ${isAtrasado?`<div class="alert-bar alert-danger" style="margin-bottom:8px">Caminhão com atraso — previsto ${dc.hora}.</div>`:''}
      <div style="font-size:13px;font-weight:600;margin-bottom:6px">${dc.material} — ${dc.fornecedor}</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:12px">
        <div><span style="color:var(--text2)">Agendado: </span>${fmtDate(dc.data)} ${dc.hora}</div>
        <div><span style="color:var(--text2)">Previsto: </span><strong>${fmtKg(dc.toneladas)}</strong></div>
        <div><span style="color:var(--text2)">Lote: </span>${dc.lote||'—'}</div>
        <div><span style="color:var(--text2)">NF: </span>${dc.nf||'—'}</div>
        <div><span style="color:var(--text2)">Baia sugerida: </span>${baia?baia.nome:'Não definida'}</div>
      </div>
    </div>`;
  document.getElementById('ch-ton').value=dc.toneladas;
  document.getElementById('ch-lote').value=dc.lote||'';
  document.getElementById('ch-nf').value=dc.nf||'';
  document.getElementById('ch-hora').value=new Date().toTimeString().slice(0,5);
  ['ch-placa','ch-motorista','ch-obs'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('ch-div-alert').innerHTML='';
  populateSelOp('ch-operador');
  openModal('modal-chegada');
}
function checkDiv(){
  const dc=S.descargas.find(d=>d.id===_chegadaId);if(!dc)return;
  const real=parseFloat(document.getElementById('ch-ton').value)||0;
  const diff=real-dc.toneladas;const el=document.getElementById('ch-div-alert');
  if(!real){el.innerHTML='';return;}
  if(Math.abs(diff)>0.1)el.innerHTML=`<div class="alert-bar alert-warning" style="margin-top:6px">Divergência: ${diff>0?'+':''}${fmtKg(diff)} (previsto ${fmtKg(dc.toneladas)}, real ${fmtKg(real)}).</div>`;
  else el.innerHTML=`<div class="alert-bar alert-success" style="margin-top:6px">Quantidade confere.</div>`;
}
function confirmarChegada(){
  const dc=S.descargas.find(d=>d.id===_chegadaId);if(!dc)return;
  const tonReal=Number(document.getElementById('ch-ton').value);
  if(!tonReal){alert('Informe a quantidade real.');return;}
  const opId=document.getElementById('ch-operador').value;
  dc.recebimento={tonReal,loteReal:document.getElementById('ch-lote').value||dc.lote,horaReal:document.getElementById('ch-hora').value,placa:document.getElementById('ch-placa').value,motorista:document.getElementById('ch-motorista').value,nfReal:document.getElementById('ch-nf').value||dc.nf,obs:document.getElementById('ch-obs').value,divergencia:Math.abs(tonReal-dc.toneladas)>0.1,operadorId:opId,chegadaAt:new Date().toISOString()};
  dc.status='chegou';
  saveState();closeModal('modal-chegada');render();
  setTimeout(()=>abrirAlocacao(_chegadaId),200);
}

function abrirAlocacao(dcId){
  const dc=S.descargas.find(d=>d.id===dcId);if(!dc)return;
  _alocarId=dcId;const r=dc.recebimento;
  const baiaSug=dc.baia?S.baias.find(b=>b.id===dc.baia):null;
  document.getElementById('alocar-resumo').innerHTML=`
    <div style="background:var(--bg3);border-radius:var(--radius);padding:10px 12px;font-size:12px">
      <div style="font-size:13px;font-weight:600;margin-bottom:5px">${dc.material} — ${dc.fornecedor}</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px">
        <div><span style="color:var(--text2)">Chegada: </span>${r.horaReal}</div>
        <div><span style="color:var(--text2)">Real: </span><strong>${fmtKg(r.tonReal)}</strong></div>
        <div><span style="color:var(--text2)">NF: </span>${r.nfReal||'—'}</div>
        <div><span style="color:var(--text2)">Lote: </span>${r.loteReal||'—'}</div>
      </div>
    </div>`;
  let sugHtml='';
  if(baiaSug){const livre=!baiaSug.estoque;const dep=S.depositos.find(d=>d.id===baiaSug.dep);
    sugHtml=`<div class="alert-bar ${livre?'alert-success':'alert-warning'}" style="margin-bottom:10px"><div><strong>Baia sugerida: ${baiaSug.nome}</strong>${dep?' ('+dep.nome+')':''}<br><span style="font-size:11px">${livre?'Livre e disponível.':'Já possui estoque — selecione outra.'}</span></div></div>`;}
  document.getElementById('alocar-sug').innerHTML=sugHtml;
  const sel=document.getElementById('alocar-baia-sel');
  sel.innerHTML='<option value="">— Selecionar —</option>';
  S.baias.forEach(b=>{if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}const dep=S.depositos.find(d=>d.id===b.dep);const isSug=b.id===dc.baia;const ocupada=b.estoques&&b.estoques.length>0;sel.innerHTML+=`<option value="${b.id}"${isSug?' selected':''}>${b.nome}${dep?' ('+dep.nome+')':''}${ocupada?' ['+b.estoques.length+' prod.]':' [livre]'}${b.material?' — '+b.material:''} ${isSug?'★':''}</option>`;});
  document.getElementById('alocar-baia-info').innerHTML='';document.getElementById('alocar-alert').innerHTML='';document.getElementById('alocar-obs').value='';
  updateAlocarInfo();openModal('modal-alocar');
}
function updateAlocarInfo(){
  const baiaId=document.getElementById('alocar-baia-sel').value;
  const el=document.getElementById('alocar-baia-info');if(!baiaId){el.innerHTML='';return;}
  const b=S.baias.find(x=>x.id===baiaId);if(!b)return;
  const dep=S.depositos.find(d=>d.id===b.dep);
  const dc=S.descargas.find(d=>d.id===_alocarId);
  const ton=dc&&dc.recebimento?dc.recebimento.tonReal:0;
  const pct=b.cap>0?Math.min(100,Math.round(ton/b.cap*100)):0;
  const excede=b.cap>0&&ton>b.cap;
  el.innerHTML=`<div style="background:var(--bg3);border-radius:var(--radius);padding:9px 11px;font-size:12px;margin-bottom:8px">
    <div style="display:flex;justify-content:space-between;margin-bottom:3px"><span style="color:var(--text2)">Depósito</span><span>${dep?dep.nome:'—'}</span></div>
    <div style="display:flex;justify-content:space-between;margin-bottom:3px"><span style="color:var(--text2)">Localização</span><span>${b.loc||'—'}</span></div>
    <div style="display:flex;justify-content:space-between;margin-bottom:3px"><span style="color:var(--text2)">Capacidade</span><span>${fmtKg(b.cap)}</span></div>
    <div style="display:flex;justify-content:space-between;margin-bottom:3px"><span style="color:var(--text2)">Carga a alocar</span><span style="font-weight:600">${fmtKg(ton)}</span></div>
    <div class="progress-bar"><div class="progress-fill" style="width:${pct}%;background:${excede?'#E24B4A':pct>80?'#EF9F27':'#1D9E75'}"></div></div>
    <div style="margin-top:3px;color:${excede?'#A32D2D':'var(--text2)'}">${excede?'Excede capacidade!':pct+'% da capacidade'}</div>
  </div>`;
  document.getElementById('alocar-alert').innerHTML=excede?`<div class="alert-bar alert-danger">Carga (${fmtKg(ton)}) excede a capacidade (${fmtKg(b.cap)}).</div>`:'';
}
function confirmarAlocacao(){
  const dc=S.descargas.find(d=>d.id===_alocarId);if(!dc)return;
  const baiaId=document.getElementById('alocar-baia-sel').value;
  if(!baiaId){alert('Selecione a baia.');return;}
  const baia=S.baias.find(b=>b.id===baiaId);if(!baia)return;
  const r=dc.recebimento;
  if(baia.cap>0&&r.tonReal>baia.cap&&!confirm('Excede capacidade. Confirmar mesmo assim?'))return;
  const baiaSug=dc.baia?S.baias.find(b=>b.id===dc.baia):null;
  if(!baia.estoques){baia.estoques=baia.estoque?[baia.estoque]:[];delete baia.estoque;}
  if(baia.estoques.length>0){
    if(!confirm('Esta baia ja possui '+baia.estoques.length+' produto(s) armazenado(s).\nDeseja adicionar mais um produto misturando os estoques?'))return;
  }
  baia.estoques.push({id:uid(),fornecedorNome:dc.fornecedor,fornecedor:dc.fornecedor,lote:r.loteReal||dc.lote,qtdTotal:r.tonReal,qtdAtual:r.tonReal,qtdUsada:0,nf:r.nfReal||dc.nf,bigBags:0,dataEntrada:hoje()});
  dc.status='armazenado';
  dc.alocacao={baiaId,baiaName:baia.nome,seguiuSugestao:dc.baia===baiaId,baiaSugName:baiaSug?baiaSug.nome:null,obs:document.getElementById('alocar-obs').value,alocadoAt:new Date().toISOString()};
  const opNomeStr=opNome(r.operadorId||S.turnoAtivo.operadorId);
  S.movimentacoes.push({id:uid(),tipo:'entrada',baiaId:baia.id,baiaNome:baia.nome,desc:`Entrada: ${dc.material} — ${dc.fornecedor} (${fmtKg(r.tonReal)}, Lote ${r.loteReal||'—'})`,data:hoje(),operador:opNomeStr,nf:r.nfReal||dc.nf||'—'});
  S.recebimentos.push({id:uid(),dcId:dc.id,fornecedor:dc.fornecedor,material:dc.material,tonPrev:dc.toneladas,tonReal:r.tonReal,lote:r.loteReal||dc.lote,nf:r.nfReal||dc.nf,horaChegada:r.horaReal,placa:r.placa,motorista:r.motorista,baiaNome:baia.nome,baiaSugNome:baiaSug?baiaSug.nome:null,seguiuSugestao:dc.baia===baiaId,divergencia:r.divergencia,data:dc.data,operador:opNomeStr,turno:turnoLabel(S.turnoAtivo.turno),finalizadoAt:new Date().toISOString()});
  saveState();closeModal('modal-alocar');render();
}

function excluirDescarga(id){if(!confirm('Remover agendamento?'))return;S.descargas=S.descargas.filter(d=>d.id!==id);saveState();render();}

function toggleBaixaQtd(){document.getElementById('baixa-qtd-wrap').style.display=document.getElementById('baixa-tipo').value==='parcial'?'':'none';}
function registrarBaixa(){
  const baiaId=document.getElementById('baixa-baia').value;
  const op=document.getElementById('baixa-op').value.trim();
  const nf=document.getElementById('baixa-nf').value.trim();
  const lote=document.getElementById('baixa-lote').value.trim();
  const tipo=document.getElementById('baixa-tipo').value;
  const qtdInput=Number(document.getElementById('baixa-qtd').value)||0;
  const opId=document.getElementById('baixa-operador').value;
  const fb=document.getElementById('baixa-feedback');
  if(!baiaId||!op){fb.innerHTML='<div class="alert-bar alert-warning">Selecione a baia e informe a OP.</div>';return;}
  const baia=S.baias.find(b=>b.id===baiaId);
  if(!baia.estoques){baia.estoques=baia.estoque?[baia.estoque]:[];delete baia.estoque;}
  if(!baia||!baia.estoques||baia.estoques.length===0){fb.innerHTML='<div class="alert-bar alert-warning">Baia sem estoque.</div>';return;}
  // use first estoque entry for baixa (select if multiple)
  const baiaEstoque=baia.estoques[0];
  const opNomeStr=opNome(opId)||'—';
  if(tipo==='parcial'){
    if(!qtdInput||qtdInput>baiaEstoque.qtdAtual){fb.innerHTML=`<div class="alert-bar alert-warning">Qtd inválida (disponível: ${fmtKg(baiaEstoque.qtdAtual)}).</div>`;return;}
    S.baixas.push({id:uid(),baiaId,baiaName:baia.nome,op,nf,lote:lote||baiaEstoque.lote||'—',tipo:'parcial',qtd:qtdInput,material:baiaEstoque.fornecedorNome,data:hoje(),fornecedor:baiaEstoque.fornecedor,operador:opNomeStr,turno:turnoLabel(S.turnoAtivo.turno)});
    S.movimentacoes.push({id:uid(),tipo:'baixa',baiaId:baia.id,baiaNome:baia.nome,desc:`Baixa parcial: ${fmtKg(qtdInput)} — ${op}`,data:hoje(),operador:opNomeStr,nf:nf||'—'});
    baiaEstoque.qtdUsada+=qtdInput;baiaEstoque.qtdAtual-=qtdInput;
    if(baiaEstoque.qtdAtual<=0)baia.estoques=baia.estoques.filter(e=>e.id!==baiaEstoque.id);
  } else {
    const qtd=baiaEstoque.qtdAtual;
    S.baixas.push({id:uid(),baiaId,baiaName:baia.nome,op,nf,lote:lote||baiaEstoque.lote||'—',tipo:'total',qtd,material:baiaEstoque.fornecedorNome,data:hoje(),fornecedor:baiaEstoque.fornecedor,operador:opNomeStr,turno:turnoLabel(S.turnoAtivo.turno)});
    S.movimentacoes.push({id:uid(),tipo:'baixa',baiaId:baia.id,baiaNome:baia.nome,desc:`Baixa total: ${fmtKg(qtd)} — ${op}`,data:hoje(),operador:opNomeStr,nf:nf||'—'});
    baia.estoques=baia.estoques.filter(e=>e.id!==baiaEstoque.id);
  }
  fb.innerHTML='<div class="alert-bar alert-success">Baixa registrada com sucesso.</div>';
  ['baixa-op','baixa-nf','baixa-lote','baixa-qtd'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('baixa-tipo').value='total';toggleBaixaQtd();
  saveState();populateBaixaBaias();render();setTimeout(()=>fb.innerHTML='',3500);
}

function salvarFornecedor(){
  const n=document.getElementById('fn-nome').value.trim();if(!n){alert('Informe o nome.');return;}
  S.fornecedores.push({id:uid(),nome:n,razao:document.getElementById('fn-razao').value,cnpj:document.getElementById('fn-cnpj').value,tel:document.getElementById('fn-tel').value,material:document.getElementById('fn-material').value,aval:Number(document.getElementById('fn-aval').value),obs:document.getElementById('fn-obs').value});
  ['fn-nome','fn-razao','fn-cnpj','fn-tel','fn-material','fn-obs'].forEach(id=>document.getElementById(id).value='');
  saveState();closeModal('modal-fornecedor');render();
}
function verFornecedor(id){
  const f=S.fornecedores.find(x=>x.id===id);if(!f)return;
  document.getElementById('forn-det-title').textContent=f.nome;
  const entregas=S.recebimentos.filter(r=>r.fornecedor===f.nome);
  const divs=entregas.filter(r=>r.divergencia).length;
  document.getElementById('forn-det-body').innerHTML=`
    <div class="info-row"><span class="info-label">Razão social</span><span>${f.razao||'—'}</span></div>
    <div class="info-row"><span class="info-label">CNPJ</span><span>${f.cnpj||'—'}</span></div>
    <div class="info-row"><span class="info-label">Telefone</span><span>${f.tel||'—'}</span></div>
    <div class="info-row"><span class="info-label">Material habitual</span><span>${f.material||'—'}</span></div>
    <div class="info-row"><span class="info-label">Avaliação</span><span style="color:#EF9F27">${stars(f.aval)}</span></div>
    <div class="info-row"><span class="info-label">Total de entregas confirmadas</span><span>${entregas.length}</span></div>
    <div class="info-row"><span class="info-label">Divergências registradas</span><span style="color:${divs>0?'#A32D2D':'#0F6E56'}">${divs}</span></div>
    ${f.obs?`<div class="divider"></div><div style="font-size:12px;color:var(--text2)">${f.obs}</div>`:''}`;
  openModal('modal-forn-det');
}
function excluirFornecedor(id){if(!confirm('Remover fornecedor?'))return;S.fornecedores=S.fornecedores.filter(f=>f.id!==id);saveState();render();}

function salvarDeposito(){
  const n=document.getElementById('dep-nome').value.trim();if(!n){alert('Informe o nome.');return;}
  S.depositos.push({id:uid(),nome:n,local:document.getElementById('dep-local').value,cap:Number(document.getElementById('dep-cap').value)||0});
  ['dep-nome','dep-local','dep-cap'].forEach(id=>document.getElementById(id).value='');
  saveState();closeModal('modal-deposito');render();
}
function salvarBaia(){
  const n=document.getElementById('baia-nome').value.trim();const dep=document.getElementById('baia-dep').value;
  if(!n||!dep){alert('Informe nome e depósito.');return;}
  S.baias.push({id:uid(),dep,nome:n,cap:parseKg(document.getElementById('baia-cap').value)||0,material:document.getElementById('baia-mat').value,loc:document.getElementById('baia-loc').value,estoques:[]});
  ['baia-nome','baia-cap','baia-mat','baia-loc'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('baia-dep').value='';saveState();closeModal('modal-baia');render();
}
function abrirBaiaDet(id){
  const b=S.baias.find(x=>x.id===id);if(!b)return;
  const dep=S.depositos.find(d=>d.id===b.dep);
  document.getElementById('modal-baia-det')._baiaId=id;
  document.getElementById('baia-det-title').textContent=`Baia ${b.nome}`;
  let html=`<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px">
    <div class="metric-card"><div class="metric-label">Depósito</div><div style="font-size:13px;font-weight:600">${dep?dep.nome:'—'}</div></div>
    <div class="metric-card"><div class="metric-label">Localização</div><div style="font-size:13px;font-weight:600">${b.loc||'—'}</div></div>
    <div class="metric-card"><div class="metric-label">Capacidade</div><div style="font-size:13px;font-weight:600">${b.cap?fmtKg(b.cap):'—'}</div></div>
    <div class="metric-card"><div class="metric-label">Material sugerido</div><div style="font-size:13px;font-weight:600">${b.material||'Qualquer'}</div></div>
  </div>`;
  if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
  if(b.estoques&&b.estoques.length>0){const e=b.estoques[0];const pct=b.cap>0?Math.min(100,Math.round(e.qtdAtual/b.cap*100)):0;
    html+=`<div class="divider"></div><div class="section-title" style="margin-bottom:8px">Estoque atual</div>
    <div class="info-row"><span class="info-label">Nome do fornecedor</span><span>${e.fornecedorNome}</span></div>
    <div class="info-row"><span class="info-label">Fornecedor (razão)</span><span>${e.fornecedor}</span></div>
    <div class="info-row"><span class="info-label">Número de lote</span><span>${e.lote||'—'}</span></div>
    <div class="info-row"><span class="info-label">NF</span><span>${e.nf||'—'}</span></div>
    <div class="info-row"><span class="info-label">Quantidade total</span><span>${fmtKg(e.qtdTotal)}</span></div>
    <div class="info-row"><span class="info-label">Quantidade atual</span><span style="font-weight:600">${fmtKg(e.qtdAtual)}</span></div>
    <div class="info-row"><span class="info-label">Quantidade utilizada</span><span>${fmtKg(e.qtdUsada)}</span></div>
    <div class="progress-bar" style="margin-top:8px"><div class="progress-fill" style="width:${pct}%;background:${pct>80?'#E24B4A':pct>50?'#EF9F27':'#1D9E75'}"></div></div>
    <div style="font-size:11px;color:var(--text2);margin-top:3px">${pct}% da capacidade</div>`;
  } else html+=`<div class="divider"></div><div class="empty-state">Baia livre</div>`;
  const movs=S.movimentacoes.filter(m=>m.baiaId===id);
  if(movs.length){html+=`<div class="divider"></div><div style="font-size:12px;font-weight:600;margin-bottom:6px">Últimas movimentações</div>`;
    html+=movs.slice(-6).reverse().map(m=>`<div class="mov-item"><div style="display:flex;gap:8px"><div style="width:8px;height:8px;border-radius:50%;background:${m.tipo==='entrada'?'#1D9E75':'#E24B4A'};flex-shrink:0;margin-top:3px"></div><div><div>${m.desc}</div><div style="color:var(--text2);font-size:11px">${fmtDate(m.data)} · ${m.operador}${m.nf&&m.nf!=='—'?' · NF: '+m.nf:''}</div></div></div></div>`).join('');}
  document.getElementById('baia-det-body').innerHTML=html;
  openModal('modal-baia-det');
}

function salvarTurno(){
  S.turnoAtivo={turno:document.getElementById('cfg-turno').value,operadorId:document.getElementById('cfg-operador').value};
  S.toleranciaMin=Number(document.getElementById('cfg-tolerancia').value)||30;
  saveState();renderTurnoBar();render();
}
function addTurno(){
  const n=document.getElementById('turno-nome').value.trim();
  if(!n){alert('Informe o nome do turno.');return;}
  const t={id:uid(),nome:n,inicio:document.getElementById('turno-inicio').value,fim:document.getElementById('turno-fim').value,cor:document.getElementById('turno-cor').value};
  S.turnos.push(t);
  document.getElementById('turno-nome').value='';
  saveState();renderConfig();
}
function removerTurno(id){
  if(!confirm('Remover turno?'))return;
  S.turnos=S.turnos.filter(t=>t.id!==id);
  if(S.turnoAtivo.turno===id)S.turnoAtivo.turno='';
  saveState();renderConfig();renderTurnoBar();
}
function salvarCap(){
  const v=parseKg(document.getElementById('cfg-cap').value);
  if(!v||v<1){alert('Informe uma capacidade válida.');return;}
  S.capDiaria=v;saveState();render();
  alert('Capacidade diária salva: '+fmtKg(v));
}
function addOperador(){
  const n=document.getElementById('op-nome').value.trim();if(!n){alert('Informe o nome.');return;}
  S.operadores.push({id:uid(),nome:n,cargo:document.getElementById('op-cargo').value});
  document.getElementById('op-nome').value='';document.getElementById('op-cargo').value='';
  saveState();render();
}
function removerOperador(id){if(!confirm('Remover operador?'))return;S.operadores=S.operadores.filter(o=>o.id!==id);saveState();render();}

function exportarDados(){
  const blob=new Blob([JSON.stringify(S,null,2)],{type:'application/json'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='VittiaAfford_backup_'+hoje()+'.json';a.click();
}
function importarDados(event){
  const file=event.target.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{const data=JSON.parse(e.target.result);
      if(!data.descargas&&!data.baias){alert('Arquivo inválido.');return;}
      if(!confirm('Isso substituirá todos os dados atuais. Confirmar?'))return;
      S=data;saveState();render();alert('Dados importados com sucesso!');
    }catch(err){alert('Erro ao importar: '+err.message);}
  };
  reader.readAsText(file);event.target.value='';
}
function resetarDados(){
  if(!confirm('Apagar TODOS os dados? Esta ação não pode ser desfeita.'))return;
  S={...DEFAULT_STATE,capDiaria:160,toleranciaMin:30,depositos:[],baias:[],descargas:[],fornecedores:[],operadores:[],turnoAtivo:{turno:'',operadorId:''},baixas:[],recebimentos:[],movimentacoes:[],materiais:[],transferencias:[],turnos:[{id:'t1',nome:'Turno A',inicio:'06:00',fim:'14:00',cor:'blue'},{id:'t2',nome:'Turno B',inicio:'14:00',fim:'22:00',cor:'green'},{id:'t3',nome:'Turno C',inicio:'22:00',fim:'06:00',cor:'amber'}]};
  saveState();render();alert('Dados resetados.');
}

function renderTurnoBar(){
  const op=S.operadores.find(o=>o.id===S.turnoAtivo.operadorId);
  document.getElementById('turno-bar').innerHTML=`Turno ativo: <strong>${turnoLabel(S.turnoAtivo.turno)}</strong> &nbsp;·&nbsp; Operador: <strong>${op?op.nome:'Não definido'}</strong> &nbsp;·&nbsp; Capacidade diária: <strong>${fmtKg(S.capDiaria||160)}</strong>`;
}

function render(){
  autoAvancarDescargas();
  renderTurnoBar();renderDashboard();if(document.getElementById('page-kpis').classList.contains('active'))renderKPIs();updateModuleBadges();renderConfig();updateBadge();renderDashKPIMini();
}
function updateBadge(){
  const p=S.descargas.filter(d=>d.status==='pendente'||d.status==='chegou').length;
  const b=document.getElementById('badge-rec');b.style.display=p>0?'inline-flex':'none';b.textContent=p;
}

function renderDashboard(){
  const cap=S.capDiaria||160;
  const baias=S.baias;
  const livre=baias.filter(b=>!b.estoque).length;
  const ocup=baias.filter(b=>b.estoque).length;
  const hj=hoje();
  const capHj=getCapDia(hj);
  const pct=Math.min(100,Math.round(capHj/cap*100));
  const cor=pct>=100?'#E24B4A':pct>=75?'#EF9F27':'#1D9E75';
  const atrasados=getAtrasados();
  const pendRec=S.descargas.filter(d=>d.status==='pendente').length;
  const capDisp=Math.max(0,cap-capHj);
  const totalTonMat=baias.reduce((s,b)=>s+(b.estoque?b.estoque.qtdAtual:0),0);

  const alertEl=document.getElementById('dash-alerts');
  let alerts='';
  if(capHj>cap)alerts+=`<div class="alert-bar alert-danger" style="margin-bottom:10px">&#9888; Capacidade ultrapassada hoje: ${fmtKg(capHj)} de ${fmtKg(cap)}.</div>`;
  if(atrasados.length)alerts+=`<div class="alert-bar alert-warning" style="margin-bottom:10px">&#9200; ${atrasados.length} caminhao(oes) com atraso: ${atrasados.map(d=>d.fornecedor+' (prev. '+d.hora+')').join(' | ')}</div>`;
  alertEl.innerHTML=alerts;

  document.getElementById('metrics-row').innerHTML=`
    <div class="metric-card"><div class="metric-label">Total em estoque</div><div class="metric-value">${fmtKg(totalTonMat)}</div><div class="metric-sub">em estoque</div></div>
    <div class="metric-card"><div class="metric-label">Disponivel p/ recebimento</div><div class="metric-value" style="color:${capDisp>0?'#1D9E75':'#E24B4A'}">${fmtKg(capDisp)}</div><div class="metric-sub">de ${fmtKg(cap)} capacidade diaria</div></div>
    <div class="metric-card"><div class="metric-label">Baias livres / ocupadas</div><div class="metric-value"><span style="color:#1D9E75">${livre}</span><span style="font-size:14px;color:var(--text2)"> / <span style="color:#378ADD">${ocup}</span></span></div><div class="metric-sub">${baias.length} baias total</div></div>
    <div class="metric-card"><div class="metric-label">Atrasos / Aguardando</div><div class="metric-value" style="color:${(atrasados.length+pendRec)>0?'#854F0B':'inherit'}">${atrasados.length} / ${pendRec}</div><div class="metric-sub">alertas / recebimento</div></div>`;

  renderCapBar(); // renders cap-bar + notif-lista for selected day

  const notifs=[];
  atrasados.forEach(d=>notifs.push({tipo:'danger',msg:'Atraso: '+d.fornecedor+' — '+d.material+' (previsto '+d.hora+')'}));
  if(capHj/cap>=0.8&&capHj<=cap)notifs.push({tipo:'warning',msg:'Capacidade hoje: '+pct+'% utilizada ('+fmtKg(capHj)+'/'+fmtKg(cap)+')'});
  S.baias.filter(b=>b.estoque&&b.cap>0&&b.estoque.qtdAtual/b.cap>0.9).forEach(b=>notifs.push({tipo:'warning',msg:'Baia '+b.nome+' quase cheia ('+Math.round(b.estoque.qtdAtual/b.cap*100)+'%)'}));
  S.descargas.filter(d=>d.data===hj&&d.status==='pendente'&&!atrasados.find(a=>a.id===d.id)).forEach(d=>notifs.push({tipo:'info',msg:'Descarga prevista: '+d.material+' — '+d.fornecedor+' as '+d.hora}));
  S.transferencias.filter(t=>t.dataTransf&&t.dataTransf===hj).forEach(t=>notifs.push({tipo:'info',msg:'Transferencia hoje: '+(t.material||t.fornecedor)+' — NF '+t.nf}));
  const nEl=document.getElementById('notif-lista');
  if(!notifs.length){nEl.innerHTML='<div style="font-size:12px;color:var(--text3);padding-top:6px">Sem notificacoes no momento.</div>';}
  else nEl.innerHTML='<div style="font-size:11px;font-weight:600;color:var(--text2);margin-bottom:6px;padding-top:8px;border-top:1px solid var(--border)">Notificacoes ('+notifs.length+')</div>'+notifs.map(n=>'<div class="notif-item"><div class="notif-dot" style="background:'+({danger:'#E24B4A',warning:'#EF9F27',info:'#378ADD',success:'#1D9E75'}[n.tipo])+'"></div><div style="font-size:12px">'+n.msg+'</div></div>').join('');

  // Indicadores de materiais
  const matEl=document.getElementById('materiais-indicadores');
  const matStock={};
  baias.filter(b=>b.estoque).forEach(b=>{
    const key=b.estoque.fornecedorNome||'Desconhecido';
    if(!matStock[key])matStock[key]={qtd:0};
    matStock[key].qtd+=b.estoque.qtdAtual;
  });
  const matCad={};
  S.materiais.forEach(m=>{matCad[m.nome]={...m,qtdEstoque:matStock[m.nome]?matStock[m.nome].qtd:0};});
  Object.keys(matStock).forEach(k=>{if(!matCad[k])matCad[k]={nome:k,codigo:'—',qtdEstoque:matStock[k].qtd};});
  const matKeys=Object.keys(matCad);
  const maxQtd=matKeys.reduce((mx,k)=>Math.max(mx,matCad[k].qtdEstoque),0.1);
  if(!matKeys.length){matEl.innerHTML='<div style="font-size:12px;color:var(--text3)">Nenhum material em estoque. Cadastre na aba Materiais.</div>';}
  else{
    const totalMat=matKeys.reduce((s,k)=>s+matCad[k].qtdEstoque,0);
    matEl.innerHTML='<div style="display:flex;justify-content:space-between;font-size:11px;color:var(--text2);margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid var(--border)"><span>Total: <strong>'+fmtKg(totalMat)+'</strong></span><span>'+matKeys.length+' materiais</span></div>'+matKeys.map(k=>{
      const m=matCad[k];const pctM=Math.min(100,Math.round(m.qtdEstoque/maxQtd*100));
      return '<div class="mat-indicator"><div style="flex:1;min-width:0"><div style="font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+m.nome+'</div><div style="font-size:11px;color:var(--text2)">'+(m.codigo||'—')+'</div></div><div style="display:flex;align-items:center;gap:8px;flex-shrink:0"><div class="mat-bar-wrap"><div class="mat-bar" style="width:'+pctM+'%;background:'+(pctM>70?'#1D9E75':pctM>30?'#EF9F27':'#378ADD')+'"></div></div><div style="font-size:13px;font-weight:600;min-width:46px;text-align:right">'+fmtKg(m.qtdEstoque)+'</div></div></div>';
    }).join('');
  }

  // Calendario
  renderCalendarioDash();

  // Proximas descargas — unifica descargas + transferencias
  const todasProx=getAllProximasDescargas();
  const sMap={pendente:'badge-amber',chegou:'badge-blue',armazenado:'badge-green',transferencia:'badge-amber'};
  const sLabel={pendente:'Pendente',chegou:'Chegou',armazenado:'Armazenado',transferencia:'Transferencia'};
  const pEl=document.getElementById('proximas');
  if(!todasProx.length){pEl.innerHTML='<div class="empty-state">Sem descargas agendadas</div>';}
  else pEl.innerHTML=todasProx.slice(0,10).map(d=>'<div style="display:flex;align-items:center;gap:8px;padding:7px 0;border-bottom:1px solid var(--border)"><div style="min-width:40px;font-size:11px;font-weight:600;color:var(--blue)">'+(d.hora||'—')+'</div><div style="flex:1;min-width:0"><div style="font-size:12px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+d.material+' — '+d.fornecedor+'</div><div style="font-size:11px;color:var(--text2)">'+fmtDate(d.data)+(d.toneladas?' · '+fmtKg(d.toneladas):'')+(d.nf?' · '+d.nf:'')+'</div></div><span class="badge '+(sMap[d._tipo]||'badge-gray')+'" style="flex-shrink:0">'+(sLabel[d._tipo]||d._tipo)+'</span></div>').join('');

  // Estoque por baia
  const matB={};S.baias.filter(b=>b.estoque).forEach(b=>{matB[b.id]={qtd:b.estoque.qtdAtual,lote:b.estoque.lote,forn:b.estoque.fornecedorNome,baia:b.nome,nf:b.estoque.nf};});
  const keysB=Object.keys(matB);
  document.getElementById('estoque-resumo').innerHTML=keysB.length?'<div style="overflow-x:auto"><table><thead><tr><th>Baia</th><th>Fornecedor / Lote</th><th>NF</th><th>Qtd atual (kg)</th></tr></thead><tbody>'+keysB.map(k=>'<tr><td style="font-weight:600">'+matB[k].baia+'</td><td>'+matB[k].forn+'<br><span style="font-size:11px;color:var(--text2)">'+(matB[k].lote||'—')+'</span></td><td><span class="tag">'+(matB[k].nf||'—')+'</span></td><td style="font-weight:600">'+fmtKg(matB[k].qtd)+'</td></tr>').join('')+'</tbody></table></div>':'<div class="empty-state">Nenhum estoque alocado nas baias</div>';
}

function renderAgenda(){
  const sorted=[...S.descargas].sort((a,b)=>a.data+a.hora>b.data+b.hora?1:-1);
  const days={};sorted.forEach(d=>{if(!days[d.data])days[d.data]=[];days[d.data].push(d);});
  const hj=hoje();const el=document.getElementById('agenda-lista');const cap=S.capDiaria||160;
  const atrasados=getAtrasados();
  const sMap={pendente:'badge-amber',chegou:'badge-blue',armazenado:'badge-green'};
  const sLabel={pendente:'Pendente',chegou:'Chegou',armazenado:'Armazenado'};
  if(!sorted.length){el.innerHTML='<div class="empty-state">Nenhuma descarga agendada</div>';return;}
  el.innerHTML=Object.keys(days).map(day=>{
    const capDia=getCapDia(day);const over=capDia>cap;
    const amanha=new Date(Date.now()+86400000).toISOString().slice(0,10);
    return`<div style="margin-bottom:14px">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:7px">
        <div style="font-size:11px;font-weight:600;color:var(--text2);text-transform:uppercase">${day===hj?'Hoje, ':day===amanha?'Amanhã, ':''}${fmtDate(day)}</div>
        ${over?`<span class="badge badge-red">Excedido: ${fmtKg(capDia)}/${fmtKg(cap)}</span>`:`<span class="badge badge-gray">${fmtKg(capDia)}/${fmtKg(cap)}</span>`}
      </div>
      ${days[day].map(d=>{
        const baia=d.baia?S.baias.find(b=>b.id===d.baia):null;
        const isAtr=!!atrasados.find(a=>a.id===d.id);
        return`<div style="border:1px solid ${isAtr?'#f0bbbb':'var(--border)'};border-left:3px solid ${isAtr?'#E24B4A':d.status==='armazenado'?'#1D9E75':d.status==='chegou'?'#378ADD':'#EF9F27'};border-radius:var(--radius-lg);padding:9px 12px;margin-bottom:6px;background:${isAtr?'var(--red-bg)':'var(--bg)'}">
          <div style="display:flex;align-items:flex-start;gap:8px">
            <div style="min-width:40px;font-size:12px;font-weight:600;color:${isAtr?'#A32D2D':'var(--blue)'};padding-top:1px">${d.hora}</div>
            <div style="flex:1">
              <div style="font-size:13px;font-weight:600">${d.material} — ${d.fornecedor}</div>
              <div style="font-size:11px;color:var(--text2);margin-top:1px">${fmtKg(d.toneladas)}${d.lote?' · '+d.lote:''}${d.nf?' · '+d.nf:''}${baia?' · Baia sugerida: '+baia.nome:''}${d.operador?' · '+opNome(d.operador):''}${isAtr?' · ⚠ ATRASADO':''}</div>
            </div>
            <div style="display:flex;gap:5px;align-items:center;flex-wrap:wrap">
              <span class="badge ${sMap[d.status]||'badge-gray'}">${sLabel[d.status]||d.status}</span>
              ${d._reprogramado?'<span class="badge badge-amber" title="Reprogramada de '+fmtDate(d._naoConformeOrigem||'—')+'">Reprogramada</span>':''}
              ${d.status==='pendente'?`<button class="btn btn-warning btn-sm" onclick="abrirChegada('${d.id}')">Chegou</button>`:''}
              ${d.status==='chegou'?`<button class="btn btn-success btn-sm" onclick="abrirAlocacao('${d.id}')">Alocar</button>`:''}
              ${d.status==='pendente'?`<button class="btn btn-danger btn-sm" onclick="excluirDescarga('${d.id}')">✕</button>`:''}
            </div>
          </div>
        </div>`;
      }).join('')}
    </div>`;
  }).join('');
}

function renderRecebimento(){
  const pend=S.descargas.filter(d=>d.status==='pendente'||d.status==='chegou').sort((a,b)=>a.data+a.hora>b.data+b.hora?1:-1);
  const atrasados=getAtrasados();
  const pEl=document.getElementById('rec-pendentes');
  if(!pend.length){pEl.innerHTML='<div class="empty-state" style="padding:20px">Nenhuma descarga aguardando recebimento.</div>';}
  else pEl.innerHTML=pend.map(d=>{
    const baia=d.baia?S.baias.find(b=>b.id===d.baia):null;
    const isAtr=!!atrasados.find(a=>a.id===d.id);
    const isChegou=d.status==='chegou';
    return`<div class="rec-card ${isAtr?'atrasado':d.status}">
      <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:10px">
        <div>
          <div style="font-size:13px;font-weight:600">${d.material} — ${d.fornecedor}</div>
          <div style="font-size:11px;color:var(--text2);margin-top:2px">${fmtDate(d.data)} ${d.hora} · ${fmtKg(d.toneladas)}${baia?' · Sugerida: '+baia.nome:''}${d.nf?' · NF: '+d.nf:''}</div>
          ${isAtr?`<span class="badge badge-red" style="margin-top:4px">⚠ Atrasado — previsto ${d.hora}</span>`:''}
          ${isChegou?`<div style="margin-top:5px"><span class="badge badge-blue">Chegou ${d.recebimento.horaReal} · ${fmtKg(d.recebimento.tonReal)}${d.recebimento.placa?' · '+d.recebimento.placa:''}</span>${d.recebimento.divergencia?' <span class="badge badge-red">Divergência</span>':''}</div>`:''}
        </div>
        <div style="display:flex;gap:5px;flex-shrink:0">
          ${!isChegou?`<button class="btn btn-warning btn-sm" onclick="abrirChegada('${d.id}')">Confirmar chegada</button>`:''}
          ${isChegou?`<button class="btn btn-success btn-sm" onclick="abrirAlocacao('${d.id}')">Alocar na baia</button>`:''}
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:5px;margin-top:8px;font-size:11px">
        <div style="display:flex;align-items:center;gap:4px"><div class="step-dot ${isChegou?'done':'active'}">${isChegou?'✓':'1'}</div><span>Chegada</span></div>
        <div class="step-line"></div>
        <div style="display:flex;align-items:center;gap:4px"><div class="step-dot ${isChegou?'active':''}">${'2'}</div><span>Alocação</span></div>
        <div class="step-line"></div>
        <div style="display:flex;align-items:center;gap:4px"><div class="step-dot">3</div><span>Concluído</span></div>
      </div>
    </div>`;
  }).join('');
  const hist=[...S.recebimentos].reverse();
  const hEl=document.getElementById('rec-historico');
  if(!hist.length){hEl.innerHTML='<div class="empty-state" style="padding:16px">Nenhum recebimento finalizado.</div>';return;}
  hEl.innerHTML=hist.map(r=>`
    <div class="rec-card" style="cursor:pointer" onclick="verRecDet('${r.id}')">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:8px">
        <div><div style="font-size:13px;font-weight:600">${r.material} — ${r.fornecedor}</div>
        <div style="font-size:11px;color:var(--text2)">${fmtDate(r.data)} · ${fmtKg(r.tonReal)} · ${r.baiaNome}${r.nf?' · NF: '+r.nf:''} · ${r.operador}</div></div>
        <div style="display:flex;gap:4px;flex-wrap:wrap">${r.divergencia?'<span class="badge badge-amber">Divergência</span>':''} ${!r.seguiuSugestao&&r.baiaSugNome?'<span class="badge badge-blue">Baia alterada</span>':''} <span class="badge badge-green">Concluído</span></div>
      </div>
    </div>`).join('');
}
function verRecDet(id){
  const r=S.recebimentos.find(x=>x.id===id);if(!r)return;
  document.getElementById('rec-det-title').textContent=`${r.material} — ${r.fornecedor}`;
  document.getElementById('rec-det-body').innerHTML=`
    <div class="info-row"><span class="info-label">Data</span><span>${fmtDate(r.data)}</span></div>
    <div class="info-row"><span class="info-label">Hora de chegada</span><span>${r.horaChegada||'—'}</span></div>
    <div class="info-row"><span class="info-label">Kg previstas</span><span>${fmtKg(r.tonPrev)}</span></div>
    <div class="info-row"><span class="info-label">Kg reais</span><span style="font-weight:600">${fmtKg(r.tonReal)}</span></div>
    <div class="info-row"><span class="info-label">Divergência</span><span class="badge ${r.divergencia?'badge-red':'badge-green'}">${r.divergencia?'Sim':'Não'}</span></div>
    <div class="info-row"><span class="info-label">Lote</span><span>${r.lote||'—'}</span></div>
    <div class="info-row"><span class="info-label">NF</span><span>${r.nf||'—'}</span></div>
    <div class="info-row"><span class="info-label">Placa</span><span>${r.placa||'—'}</span></div>
    <div class="info-row"><span class="info-label">Motorista</span><span>${r.motorista||'—'}</span></div>
    <div class="info-row"><span class="info-label">Operador</span><span>${r.operador||'—'}</span></div>
    <div class="info-row"><span class="info-label">Baia sugerida</span><span>${r.baiaSugNome||'—'}</span></div>
    <div class="info-row"><span class="info-label">Baia utilizada</span><span style="font-weight:600">${r.baiaNome}</span></div>
    <div class="info-row"><span class="info-label">Seguiu sugestão</span><span class="badge ${r.seguiuSugestao?'badge-green':'badge-amber'}">${r.seguiuSugestao?'Sim':'Não'}</span></div>
    <div class="info-row"><span class="info-label">Finalizado em</span><span>${fmtDT(r.finalizadoAt)}</span></div>`;
  openModal('modal-rec-det');
}

function renderDepositos(){
  const el=document.getElementById('depositos-lista');
  if(!S.depositos.length){el.innerHTML='<div class="empty-state">Nenhum depósito cadastrado. Clique em "+ Depósito" para começar.</div>';return;}
  el.innerHTML=S.depositos.map(dep=>{
    const baias=S.baias.filter(b=>b.dep===dep.id);
    const livre=baias.filter(b=>!b.estoque).length;const ocup=baias.filter(b=>b.estoque).length;
    return`<div class="card" style="margin-bottom:8px">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
        <div><div style="font-size:13px;font-weight:600">${dep.nome}</div><div style="font-size:11px;color:var(--text2)">${dep.local||'—'}${dep.cap?' · Capacidade: '+fmtKg(dep.cap):''}</div></div>
        <div style="display:flex;gap:5px"><span class="badge badge-green">${livre} livres</span><span class="badge badge-blue">${ocup} ocupadas</span></div>
      </div>
      ${baias.length?`<div class="baia-grid">${baias.map(b=>{const _tot=baiaTotalAtual(b);
      var cls;
      if(b.status==='inutilizada')cls='inutilizada';
      else if(b.status==='manutencao')cls='manutencao';
      else cls=baiaTemEstoque(b)?(b.cap>0&&_tot/b.cap>.9?'alerta':'ocupada'):'livre';
      const _e0=baiaPrimEstoque(b);return`<div class="baia-card ${cls}" onclick="abrirBaiaDet('${b.id}')">
        <div class="baia-name">${b.nome}${b.estoques&&b.estoques.length>1?' ('+b.estoques.length+'p)':''}</div>
        <div class="baia-info">${b.status==='inutilizada'?'⚠ Inutilizada':b.status==='manutencao'?'🔧 Manutenção':_e0?_e0.fornecedorNome:'Livre'}${_e0&&b.status!=='inutilizada'&&b.status!=='manutencao'?`<br>${fmtKg(_tot)}/${fmtKg(b.cap)}`:''}</div>
        ${b.material&&!baiaTemEstoque(b)&&!b.status?`<div class="baia-info" style="margin-top:2px">Sug: ${b.material}</div>`:''}
      </div>`;}).join('')}</div>`:'<div class="empty-state" style="padding:12px">Nenhuma baia neste depósito — clique em "+ Baia".</div>'}
    </div>`;
  }).join('');
}

function renderFornecedores(){
  const el=document.getElementById('forn-lista');
  if(!S.fornecedores.length){el.innerHTML='<div class="empty-state">Nenhum fornecedor cadastrado.</div>';return;}
  el.innerHTML=`<div style="overflow-x:auto"><table><thead><tr><th>Nome</th><th>CNPJ</th><th>Material</th><th>Avaliação</th><th>Entregas</th><th></th></tr></thead><tbody>${S.fornecedores.map(f=>{
    const entregas=S.recebimentos.filter(r=>r.fornecedor===f.nome).length;
    return`<tr>
      <td><div style="font-weight:600;cursor:pointer;color:var(--blue)" onclick="verFornecedor('${f.id}')">${f.nome}</div><div style="font-size:11px;color:var(--text2)">${f.razao||''}</div></td>
      <td style="font-size:12px">${f.cnpj||'—'}</td>
      <td><span class="tag">${f.material||'—'}</span></td>
      <td style="color:#EF9F27">${stars(f.aval)}</td>
      <td style="font-size:12px">${entregas}</td>
      <td><button class="btn btn-sm btn-danger" onclick="excluirFornecedor('${f.id}')">✕</button></td>
    </tr>`;
  }).join('')}</tbody></table></div>`;
}

function renderEstoque(){
  var el=document.getElementById('estoque-lista');if(!el)return;
  S.baias.forEach(function(b){if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}});
  // populate filter dropdowns
  var fDep=document.getElementById('est-filtro-dep');
  var fMat=document.getElementById('est-filtro-mat');
  if(fDep){
    var curDep=fDep.value;
    fDep.innerHTML='<option value="">Todos os depósitos</option>';
    S.depositos.forEach(function(d){fDep.innerHTML+='<option value="'+d.id+'"'+(d.id===curDep?' selected':'')+'>'+d.nome+'</option>';});
  }
  if(fMat){
    var curMat=fMat.value;
    var matSet={};
    S.baias.forEach(function(b){(b.estoques||[]).forEach(function(e){if(e.fornecedorNome)matSet[e.fornecedorNome]=true;});});
    S.materiais.forEach(function(m){matSet[m.nome]=true;});
    fMat.innerHTML='<option value="">Todos os materiais</option>';
    Object.keys(matSet).sort().forEach(function(m){fMat.innerHTML+='<option value="'+m+'"'+(m===curMat?' selected':'')+'>'+m+'</option>';});
  }
  // apply filters
  var busca=(document.getElementById('est-busca')?document.getElementById('est-busca').value:'').toLowerCase().trim();
  var filtDep=fDep?fDep.value:'';
  var filtMat=fMat?fMat.value:'';
  var com=S.baias.filter(function(b){return b.estoques&&b.estoques.length>0;});
  // summary cards
  var totalTon=0,totalLotes=0,totalBaias=com.length;
  com.forEach(function(b){b.estoques.forEach(function(e){totalTon+=e.qtdAtual;totalLotes++;});});
  var matCount={};
  com.forEach(function(b){b.estoques.forEach(function(e){var k=e.fornecedorNome||'Outro';if(!matCount[k])matCount[k]=0;matCount[k]+=e.qtdAtual;});});
  var resumoEl=document.getElementById('est-resumo-cards');
  if(resumoEl)resumoEl.innerHTML=
    '<div class="metric-card"><div class="metric-label">Total em estoque</div><div class="metric-value" style="font-size:20px">'+fmtKg(totalTon)+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Baias ocupadas</div><div class="metric-value" style="font-size:20px">'+totalBaias+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Lotes ativos</div><div class="metric-value" style="font-size:20px">'+totalLotes+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Materiais distintos</div><div class="metric-value" style="font-size:20px">'+Object.keys(matCount).length+'</div></div>';
  // view mode
  var viewProduto=document.getElementById('btn-view-estoque')&&document.getElementById('btn-view-estoque')._viewProduto;
  if(viewProduto){
    renderEstoquePorProduto(com, busca, filtDep, filtMat, el);
    return;
  }
  // Default: table view with filters
  if(!com.length){
    el.innerHTML='<div class="empty-state">Nenhum estoque alocado.</div><div style="margin-top:12px"><button class="btn btn-primary btn-sm" onclick="openEntradaManual()">Lançar estoque manual</button></div>';
    return;
  }
  var rows='';
  com.forEach(function(b){
    if(filtDep&&b.dep!==filtDep)return;
    var dep=S.depositos.find(function(d){return d.id===b.dep;});
    b.estoques.forEach(function(e,idx){
      // apply search filter
      var searchStr=(b.nome+' '+(dep?dep.nome:'')+' '+e.fornecedorNome+' '+(e.material||'')+' '+(e.lote||'')+' '+(e.nf||'')).toLowerCase();
      if(busca&&!searchStr.includes(busca))return;
      if(filtMat&&e.fornecedorNome!==filtMat)return;
      var pct=b.cap>0?Math.min(100,Math.round(e.qtdAtual/b.cap*100)):0;
      rows+='<tr>'
        +'<td style="font-weight:600">'+b.nome+(b.estoques.length>1?' <span class="tag" style="font-size:10px">'+(idx+1)+'/'+b.estoques.length+'</span>':'')+(dep?'<br><span style="font-size:11px;color:var(--text2)">'+dep.nome+'</span>':'')+'</td>'
        +'<td>'+e.fornecedorNome+'</td>'
        +'<td style="color:var(--blue);font-size:12px">'+(e.material||e.fornecedorNome||'—')+'</td>'
        +'<td><span class="tag">'+(e.lote||'—')+'</span></td>'
        +'<td><span class="tag">'+(e.nf||'—')+'</span></td>'
        +'<td>'+e.qtdTotal+'t</td>'
        +'<td style="font-weight:600">'+fmtKg(e.qtdAtual)+'</td>'
        +'<td style="color:var(--text2)">'+fmtKg(e.qtdUsada)+'</td>'
        +'<td><div style="display:flex;align-items:center;gap:6px">'
          +'<input type="number" min="0" step="1" value="'+(e.bigBags||0)+'" data-bid="'+b.id+'" data-eid="'+e.id+'" onchange="salvarBigBag(this)" style="width:60px;padding:3px 6px;font-size:12px" title="Big Bags">'
          +'<span style="font-size:11px;color:var(--text2)">BB</span>'
        +'</div></td>'
        +'<td style="min-width:80px"><div class="progress-bar"><div class="progress-fill" style="width:'+pct+'%;background:'+(pct>80?'#E24B4A':pct>50?'#EF9F27':'#1D9E75')+'"></div></div>'
          +'<div style="font-size:11px;color:var(--text2);margin-top:2px">'+pct+'%</div></td>'
        +'<td>'
          +'<button class="btn btn-sm" data-bid="'+b.id+'" data-eid="'+e.id+'" onclick="abrirBaiaDet(this.dataset.bid)" style="font-size:11px">Ver baia</button>'
          +'<button class="btn btn-sm btn-danger" data-bid="'+b.id+'" data-eid="'+e.id+'" onclick="removerEstoqueBaia(this.dataset.bid,this.dataset.eid)" style="font-size:11px">Remover</button>'
        +'</td>'
        +'</tr>';
    });
  });
  if(!rows){
    el.innerHTML='<div class="empty-state">Nenhum estoque encontrado com os filtros aplicados.</div>';
    return;
  }
  el.innerHTML='<div style="overflow-x:auto"><table><thead><tr>'
    +'<th>Baia</th><th>Fornecedor / Material</th><th>Produto</th><th>Lote</th><th>NF</th>'
    +'<th>Qtd total</th><th>Qtd atual</th><th>Qtd usada</th>'
    +'<th>Big Bags</th><th>Ocupação</th><th></th>'
    +'</tr></thead><tbody>'+rows+'</tbody></table></div>';
}

function renderEstoquePorProduto(com, busca, filtDep, filtMat, el){
  // Group all stock by material/product name
  var grupos={};
  com.forEach(function(b){
    if(filtDep&&b.dep!==filtDep)return;
    var dep=S.depositos.find(function(d){return d.id===b.dep;});
    b.estoques.forEach(function(e){
      var mat=e.fornecedorNome||'Desconhecido';
      if(filtMat&&mat!==filtMat)return;
      var searchStr=(mat+' '+(e.material||'')+' '+(e.lote||'')+' '+b.nome+' '+(e.nf||'')+' '+(dep?dep.nome:'')).toLowerCase();
      if(busca&&!searchStr.includes(busca))return;
      if(!grupos[mat])grupos[mat]={mat:mat,totalAtual:0,totalTotal:0,totalUsado:0,totalBB:0,lotes:[]};
      grupos[mat].totalAtual+=e.qtdAtual;
      grupos[mat].totalTotal+=e.qtdTotal;
      grupos[mat].totalUsado+=e.qtdUsada;
      grupos[mat].totalBB+=(e.bigBags||0);
      grupos[mat].lotes.push({baia:b.nome,dep:dep?dep.nome:'',lote:e.lote||'—',nf:e.nf||'—',qtdAtual:e.qtdAtual,qtdTotal:e.qtdTotal,bigBags:e.bigBags||0,bid:b.id,eid:e.id});
    });
  });
  var keys=Object.keys(grupos).sort();
  if(!keys.length){el.innerHTML='<div class="empty-state">Nenhum estoque encontrado com os filtros aplicados.</div>';return;}
  el.innerHTML=keys.map(function(mat){
    var g=grupos[mat];
    var pctGlobal=g.totalTotal>0?Math.min(100,Math.round(g.totalAtual/g.totalTotal*100)):0;
    var cor=pctGlobal>80?'#E24B4A':pctGlobal>50?'#EF9F27':'#1D9E75';
    return '<div style="border:1px solid var(--border);border-radius:var(--radius-lg);margin-bottom:10px;overflow:hidden">'
      // header
      +'<div style="display:flex;align-items:center;gap:12px;padding:12px 16px;background:var(--bg3);cursor:pointer" onclick="toggleProdutoGroup(this)">'
        +'<div style="flex:1">'
          +'<div style="font-size:14px;font-weight:600">'+mat+'</div>'
          +'<div style="font-size:12px;color:var(--text2)">'+g.lotes.length+' lote(s) &nbsp;·&nbsp; '+g.totalBB+' big bags &nbsp;·&nbsp; Total atual: <strong>'+fmtKg(g.totalAtual)+'</strong></div>'
        +'</div>'
        +'<div style="display:flex;align-items:center;gap:12px;flex-shrink:0">'
          +'<div style="text-align:right">'
            +'<div style="font-size:18px;font-weight:600;color:'+cor+'">'+fmtKg(g.totalAtual)+'</div>'
            +'<div style="font-size:11px;color:var(--text2)">de '+fmtKg(g.totalTotal)+' total</div>'
          +'</div>'
          +'<span style="font-size:14px;color:var(--text3)">▾</span>'
        +'</div>'
      +'</div>'
      // lotes table
      +'<div class="prod-group-body" style="padding:0">'
        +'<table style="width:100%;border-collapse:collapse;font-size:12px">'
          +'<thead><tr style="background:var(--bg3)">'
            +'<th style="padding:6px 12px;text-align:left;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border)">Produto</th>'
            +'<th style="padding:6px 12px;text-align:left;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border)">Baia</th>'
            +'<th style="padding:6px 12px;text-align:left;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border)">Depósito</th>'
            +'<th style="padding:6px 12px;text-align:left;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border)">Lote</th>'
            +'<th style="padding:6px 12px;text-align:left;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border)">NF</th>'
            +'<th style="padding:6px 12px;text-align:left;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border)">Qtd atual</th>'
            +'<th style="padding:6px 12px;text-align:left;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border)">Big Bags</th>'
            +'<th style="padding:6px 12px;border-bottom:1px solid var(--border)"></th>'
          +'</tr></thead>'
          +'<tbody>'
          +g.lotes.map(function(l){
            return '<tr style="border-bottom:1px solid var(--border)">'
              +'<td style="padding:8px 12px;font-weight:600;color:var(--blue)">'+mat+'</td>'
              +'<td style="padding:8px 12px;font-weight:600">'+l.baia+'</td>'
              +'<td style="padding:8px 12px;color:var(--text2)">'+l.dep+'</td>'
              +'<td style="padding:8px 12px"><span class="tag">'+l.lote+'</span></td>'
              +'<td style="padding:8px 12px"><span class="tag">'+l.nf+'</span></td>'
              +'<td style="padding:8px 12px;font-weight:600">'+fmtKg(l.qtdAtual)+'</td>'
              +'<td style="padding:8px 12px">'+l.bigBags+'</td>'
              +'<td style="padding:8px 12px">'
                +'<button class="btn btn-sm" data-bid="'+l.bid+'" onclick="abrirBaiaDet(this.dataset.bid)" style="font-size:11px">Ver baia</button>'
              +'</td>'
              +'</tr>';
          }).join('')
          +'</tbody>'
        +'</table>'
      +'</div>'
    +'</div>';
  }).join('');
}

function toggleProdutoGroup(header){
  var body=header.nextElementSibling;
  var arrow=header.querySelector('span:last-child');
  if(body.style.display==='none'){
    body.style.display='';
    if(arrow)arrow.textContent='▾';
  } else {
    body.style.display='none';
    if(arrow)arrow.textContent='▸';
  }
}

function toggleEstoqueView(){
  var btn=document.getElementById('btn-view-estoque');
  if(!btn)return;
  btn._viewProduto=!btn._viewProduto;
  btn.textContent=btn._viewProduto?'📋 Visão por Baia':'📦 Visão por Produto';
  btn.style.fontWeight=btn._viewProduto?'600':'400';
  renderEstoque();
}
function salvarBigBag(input){
  const bid=input.dataset.bid, eid=input.dataset.eid;
  const baia=S.baias.find(function(b){return b.id===bid;});if(!baia)return;
  const est=baia.estoques.find(function(e){return e.id===eid;});if(!est)return;
  est.bigBags=parseInt(input.value)||0;
  saveState();
}

function removerEstoqueBaia(bid,eid){
  if(!confirm('Remover este lote da baia?'))return;
  const baia=S.baias.find(function(b){return b.id===bid;});if(!baia)return;
  baia.estoques=baia.estoques.filter(function(e){return e.id!==eid;});
  saveState();render();
}

function renderMov(){
  const sel=document.getElementById('filtro-baia-mov');if(!sel)return;
  const curVal=sel.value;
  sel.innerHTML='<option value="">Todas as baias</option>';
  S.baias.forEach(b=>sel.innerHTML+=`<option value="${b.id}"${b.id===curVal?' selected':''}>${b.nome}</option>`);
  const filtro=sel.value;
  const movs=S.movimentacoes.filter(m=>!filtro||m.baiaId===filtro).slice().reverse();
  const el=document.getElementById('mov-lista');
  if(!movs.length){el.innerHTML='<div class="empty-state">Nenhuma movimentação registrada.</div>';return;}
  const cMap={entrada:'#1D9E75',baixa:'#E24B4A',ajuste:'#EF9F27'};
  const lMap={entrada:'Entrada',baixa:'Baixa',ajuste:'Ajuste'};
  el.innerHTML=`<div style="overflow-x:auto"><table><thead><tr><th>Data</th><th>Baia</th><th>Tipo</th><th>Descrição</th><th>NF</th><th>Operador</th></tr></thead><tbody>${movs.map(m=>`
    <tr>
      <td style="font-size:12px;white-space:nowrap">${fmtDate(m.data)}</td>
      <td style="font-weight:600">${m.baiaNome}</td>
      <td><span class="badge" style="background:${cMap[m.tipo]||'#888'}22;color:${cMap[m.tipo]||'#888'}">${lMap[m.tipo]||m.tipo}</span></td>
      <td style="font-size:12px">${m.desc}</td>
      <td style="font-size:12px"><span class="tag">${m.nf&&m.nf!=='—'?m.nf:'—'}</span></td>
      <td style="font-size:12px">${m.operador||'—'}</td>
    </tr>`).join('')}</tbody></table></div>`;
}

function renderBaixa(){
  populateBaixaBaias();populateSelOp('baixa-operador');
  const el=document.getElementById('historico-baixas');
  if(!S.baixas.length){el.innerHTML='<div class="empty-state">Nenhuma baixa registrada.</div>';return;}
  el.innerHTML=[...S.baixas].reverse().map(b=>`
    <div style="padding:8px 0;border-bottom:1px solid var(--border)">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:2px">
        <span style="font-size:13px;font-weight:600">${b.baiaName}</span>
        <span class="badge ${b.tipo==='total'?'badge-red':'badge-amber'}">${b.tipo}</span>
        <span style="font-size:11px;color:var(--text2);margin-left:auto">${fmtDate(b.data)}</span>
      </div>
      <div style="font-size:11px;color:var(--text2)">${b.fornecedor} · ${fmtKg(b.qtd)} · OP: <strong style="color:var(--text)">${b.op}</strong>${b.nf?' · NF: '+b.nf:''}${b.lote&&b.lote!=='—'?' · Lote: '+b.lote:''} · ${b.turno||''} · ${b.operador||'—'}</div>
    </div>`).join('');
}

function renderConfig(){
  // operadores
  const s=document.getElementById('cfg-operador');
  if(s){s.innerHTML='<option value="">— Selecionar —</option>';S.operadores.forEach(o=>s.innerHTML+=`<option value="${o.id}"${o.id===S.turnoAtivo.operadorId?' selected':''}>${o.nome}</option>`);}
  // turnos dropdown
  const ct=document.getElementById('cfg-turno');
  if(ct){
    const cur=S.turnoAtivo.turno;
    ct.innerHTML='<option value="">— Selecionar —</option>';
    (S.turnos||[]).forEach(t=>ct.innerHTML+=`<option value="${t.id}"${t.id===cur?' selected':''}>${t.nome} (${t.inicio}–${t.fim})</option>`);
  }
  // lista de turnos cadastrados
  const tl=document.getElementById('turnos-lista');
  if(tl){
    const corMap={blue:'#185FA5',green:'#0F6E56',amber:'#854F0B',red:'#A32D2D',purple:'#534AB7'};
    const bgMap={blue:'var(--blue-bg)',green:'var(--green-bg)',amber:'var(--amber-bg)',red:'var(--red-bg)',purple:'#EEEDFE'};
    if(!S.turnos||!S.turnos.length){tl.innerHTML='<div style="font-size:12px;color:var(--text3)">Nenhum turno cadastrado.</div>';}
    else tl.innerHTML=S.turnos.map(t=>`<div style="display:flex;align-items:center;justify-content:space-between;padding:7px 0;border-bottom:1px solid var(--border)"><div style="display:flex;align-items:center;gap:8px"><div style="width:10px;height:10px;border-radius:50%;background:${corMap[t.cor]||'#888'};flex-shrink:0"></div><div><div style="font-size:13px;font-weight:500">${t.nome}</div><div style="font-size:11px;color:var(--text2)">${t.inicio} – ${t.fim}</div></div></div><button class="btn btn-sm btn-danger" onclick="removerTurno('${t.id}')">✕</button></div>`).join('');
  }
  const cc=document.getElementById('cfg-cap');if(cc)cc.value=S.capDiaria||160;
  const ct2=document.getElementById('cfg-tolerancia');if(ct2){ct2.value=S.toleranciaMin||30;document.getElementById('tol-val').textContent=(S.toleranciaMin||30)+' min';}
  const emailEl=document.getElementById('cfg-email');if(emailEl)emailEl.value=S.emailAlerta||'';
  const hEl=document.getElementById('cfg-horario-alerta');if(hEl)hEl.value=S.horarioAlerta||'07:00';
  const aEl=document.getElementById('cfg-alerta-ativo');if(aEl)aEl.checked=!!S.alertaAtivo;
  const el=document.getElementById('op-lista');if(!el)return;
  if(!S.operadores.length){el.innerHTML='<div class="empty-state" style="padding:12px">Nenhum operador cadastrado.</div>';return;}
  el.innerHTML=S.operadores.map(o=>`<div style="display:flex;align-items:center;justify-content:space-between;padding:7px 0;border-bottom:1px solid var(--border)"><div><div style="font-size:13px">${o.nome}</div><div style="font-size:11px;color:var(--text2)">${o.cargo||'—'}</div></div><button class="btn btn-sm btn-danger" onclick="removerOperador('${o.id}')">✕</button></div>`).join('');
}

function converter(){} // conversor removido

function csvEscape(v){if(v===null||v===undefined)return'';const s=String(v);return s.includes(',')||s.includes('"')||s.includes('\n')?'"'+s.replace(/"/g,'""')+'"':s;}
function downloadCSV(filename,rows){
  const csv=rows.map(r=>r.map(csvEscape).join(',')).join('\n');
  const blob=new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=filename;a.click();
}

function gerarRelatorio(tipo){
  var heads=[],rows=[];
  if(tipo==='entradas'){
    heads=['Data','Hora Chegada','Fornecedor','Material','Lote','NF','Baia','Turno','Operador','Prev.(kg)','Real(kg)','Divergência','Baia Sugerida','Seguiu Sug.','Placa','Motorista'];
    rows=S.recebimentos.map(function(r){return[fmtDate(r.data),r.horaChegada||'—',r.fornecedor,r.material,r.lote||'—',r.nf||'—',r.baiaNome,r.turno||'—',r.operador||'—',fmtKg(r.tonPrev),fmtKg(r.tonReal),r.divergencia?'Sim':'Não',r.baiaSugNome||'—',r.seguiuSugestao?'Sim':'Não',r.placa||'—',r.motorista||'—'];});
  }else if(tipo==='saidas'){
    heads=['Data','Baia','Fornecedor','Material','Lote','NF','OP','Tipo','Turno','Operador','Qtd Baixada'];
    rows=S.baixas.map(function(b){return[fmtDate(b.data),b.baiaName,b.fornecedor,b.material||'—',b.lote||'—',b.nf||'—',b.op,b.tipo,b.turno||'—',b.operador||'—',fmtKg(b.qtd)];});
  }else if(tipo==='ops'){
    heads=['OP','Data','Baia','Fornecedor','Material','Lote','NF','Tipo','Quantidade','Turno','Operador'];
    rows=S.baixas.map(function(b){return[b.op,fmtDate(b.data),b.baiaName,b.fornecedor,b.material||'—',b.lote||'—',b.nf||'—',b.tipo,fmtKg(b.qtd),b.turno||'—',b.operador||'—'];});
  }else if(tipo==='estoque'){
    heads=['Depósito','Baia','Fornecedor','Material','Lote','NF','Qtd Total','Qtd Atual','Qtd Utilizada','Cap. Baia','% Ocupação'];
    rows=S.baias.filter(function(b){return b.estoque;}).map(function(b){
      var dep=S.depositos.find(function(d){return d.id===b.dep;});var e=b.estoque;
      var pct=b.cap>0?(e.qtdAtual/b.cap*100).toFixed(1)+'%':'—';
      return[dep?dep.nome:'—',b.nome,e.fornecedorNome,e.fornecedor,e.lote||'—',e.nf||'—',fmtKg(e.qtdTotal),fmtKg(e.qtdAtual),fmtKg(e.qtdUsada),fmtKg(b.cap),pct];
    });
  }else if(tipo==='turnos'){
    heads=['Data','Turno','Operador','Descargas','Qtd. Recebida','Divergências'];
    var byTurno={};
    S.recebimentos.forEach(function(r){var k=r.data+'_'+(r.turno||'—');if(!byTurno[k])byTurno[k]={data:r.data,turno:r.turno||'—',op:r.operador||'—',rec:0,ton:0,divs:0};byTurno[k].rec++;byTurno[k].ton+=Number(r.tonReal)||0;if(r.divergencia)byTurno[k].divs++;});
    rows=Object.values(byTurno).map(function(t){return[fmtDate(t.data),t.turno,t.op,t.rec,fmtKg(t.ton),t.divs];});
  }else if(tipo==='ocorrencias'){
    heads=['Data','Tipo','Gravidade','Fornecedor','NF','Status','Descrição','Ação','Operador'];
    rows=(S.ocorrencias||[]).map(function(o){return[fmtDate(o.data),o.tipoLabel||o.tipo,o.gravidadeLabel||o.gravidade,o.fornecedor||'—',o.nf||'—',o.status,o.descricao,o.acao||'—',o.operador||'—'];});
  }else if(tipo==='movimentacoes'){
    heads=['Data','Baia','Tipo','Descrição','NF','Operador'];
    rows=S.movimentacoes.map(function(m){return[fmtDate(m.data),m.baiaNome,m.tipo,m.desc,m.nf||'—',m.operador||'—'];});
  }
  if(!heads.length){alert('Tipo de relatório inválido.');return;}
  var tipoLabel={'entradas':'Entradas','saidas':'Saídas','ops':'Ordens de Produção','estoque':'Estoque','turnos':'Turnos','ocorrencias':'Ocorrências','movimentacoes':'Movimentações'}[tipo]||tipo;
  var dt=new Date().toLocaleDateString('pt-BR');

  // ── Gerar XLSX via SpreadsheetML (abre nativamente no Excel) ──
  function xmlEsc(v){
    if(v===null||v===undefined)return'';
    return String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }
  var xmlRows='';
  // Linha de título
  xmlRows+='<Row ss:StyleID="s_title"><Cell ss:MergeAcross="'+(heads.length-1)+'" ss:StyleID="s_title"><Data ss:Type="String">VittiaAfford — Relatório de '+xmlEsc(tipoLabel)+' — Gerado em: '+xmlEsc(dt)+'</Data></Cell></Row>';
  // Cabeçalho
  xmlRows+='<Row ss:StyleID="s_head">';
  heads.forEach(function(h){xmlRows+='<Cell ss:StyleID="s_head"><Data ss:Type="String">'+xmlEsc(h)+'</Data></Cell>';});
  xmlRows+='</Row>';
  // Dados
  rows.forEach(function(r,ri){
    var style=ri%2===0?'s_even':'s_odd';
    xmlRows+='<Row>';
    r.forEach(function(c){xmlRows+='<Cell ss:StyleID="'+style+'"><Data ss:Type="String">'+xmlEsc(c)+'</Data></Cell>';});
    xmlRows+='</Row>';
  });
  // Linha de total
  xmlRows+='<Row ss:StyleID="s_total"><Cell ss:MergeAcross="'+(heads.length-1)+'" ss:StyleID="s_total"><Data ss:Type="String">Total de registros: '+rows.length+'</Data></Cell></Row>';

  var xml='<?xml version="1.0" encoding="UTF-8"?>'
    +'<?mso-application progid="Excel.Sheet"?>'
    +'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"'
    +' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"'
    +' xmlns:x="urn:schemas-microsoft-com:office:excel">'
    +'<Styles>'
    +'<Style ss:ID="s_title"><Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="0"/>'
    +'<Font ss:Bold="1" ss:Size="13" ss:Color="#1a1a1a"/>'
    +'<Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>'
    +'<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2" ss:Color="#1a1a1a"/></Borders></Style>'
    +'<Style ss:ID="s_head"><Alignment ss:Horizontal="Left" ss:Vertical="Center"/>'
    +'<Font ss:Bold="1" ss:Size="10" ss:Color="#FFFFFF"/>'
    +'<Interior ss:Color="#1D3557" ss:Pattern="Solid"/>'
    +'<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0f2236"/></Borders></Style>'
    +'<Style ss:ID="s_even"><Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="0"/>'
    +'<Font ss:Size="10" ss:Color="#1a1a1a"/>'
    +'<Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>'
    +'<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#e0e0e0"/></Borders></Style>'
    +'<Style ss:ID="s_odd"><Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="0"/>'
    +'<Font ss:Size="10" ss:Color="#1a1a1a"/>'
    +'<Interior ss:Color="#F2F6FB" ss:Pattern="Solid"/>'
    +'<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#e0e0e0"/></Borders></Style>'
    +'<Style ss:ID="s_total"><Alignment ss:Horizontal="Right" ss:Vertical="Center"/>'
    +'<Font ss:Bold="1" ss:Size="10" ss:Color="#666666"/>'
    +'<Interior ss:Color="#F5F5F5" ss:Pattern="Solid"/></Style>'
    +'</Styles>'
    +'<Worksheet ss:Name="'+xmlEsc(tipoLabel)+'">'
    +'<Table>'+xmlRows+'</Table>'
    +'<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">'
    +'<FreezePanes/><SplitHorizontal>2</SplitHorizontal><TopRowBottomPane>2</TopRowBottomPane>'
    +'</WorksheetOptions>'
    +'</Worksheet>'
    +'</Workbook>';

  var blob=new Blob([xml],{type:'application/vnd.ms-excel;charset=utf-8'});
  var a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='VittiaAfford_'+tipoLabel+'_'+hoje()+'.xls';
  a.click();
}


function getAllProximasDescargas(){
  const hj=hoje();
  const dc=S.descargas.filter(d=>d.data>=hj&&d.status!=='armazenado').map(d=>({...d,_tipo:d.status||'pendente'}));
  const tr=S.transferencias.filter(t=>t.dataTransf&&t.dataTransf>=hj).map(t=>({
    id:t.id,fornecedor:t.fornecedor,material:t.material||t.mp||'Transferência',
    toneladas:t.qtd||(t.qtdKg?t.qtdKg/1000:null),
    data:t.dataTransf,hora:t.hora||'08:00',nf:t.nf,status:'transferencia',_tipo:'transferencia'
  }));
  return [...dc,...tr].sort((a,b)=>{
    if(a.data!==b.data)return a.data>b.data?1:-1;
    return (a.hora||'00:00')>(b.hora||'00:00')?1:-1;
  });
}

function renderCalendarioDash(){
  var el=document.getElementById('calendario-dash');if(!el)return;
  var now=new Date();
  var viewYear=parseInt(el.getAttribute('data-year')||now.getFullYear());
  var viewMonth=parseInt(el.getAttribute('data-month')||now.getMonth());
  el.setAttribute('data-year',viewYear);el.setAttribute('data-month',viewMonth);
  var firstDay=new Date(viewYear,viewMonth,1);
  var lastDay=new Date(viewYear,viewMonth+1,0);
  var startDow=(firstDay.getDay()+6)%7;
  var monthNames=['Janeiro','Fevereiro','Marco','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  var evMap={};
  var allProx=getAllProximasDescargas();
  allProx.forEach(function(d){if(!evMap[d.data])evMap[d.data]=[];evMap[d.data].push(d);});
  var hj=hoje();
  var html='<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">';
  html+='<button class="btn btn-sm" onclick="calNav(-1)" style="padding:2px 8px">&#8592;</button>';
  html+='<span style="font-size:13px;font-weight:600">'+monthNames[viewMonth]+' '+viewYear+'</span>';
  html+='<button class="btn btn-sm" onclick="calNav(1)" style="padding:2px 8px">&#8594;</button>';
  html+='</div>';
  if(_selectedDay){
    html+='<div style="font-size:11px;color:#185FA5;margin-bottom:6px;text-align:center">';
    html+='Exibindo: '+fmtDate(_selectedDay);
    html+=' <button class="btn btn-sm" onclick="clearCalSel()" style="font-size:10px;padding:1px 6px">Hoje</button></div>';
  }
  html+='<div class="cal-grid">';
  ['Seg','Ter','Qua','Qui','Sex','Sab','Dom'].forEach(function(d){html+='<div class="cal-header-day">'+d+'</div>';});
  var totalCells=Math.ceil((startDow+lastDay.getDate())/7)*7;
  for(var i=0;i<totalCells;i++){
    var dayNum=i-startDow+1;
    var isValid=dayNum>=1&&dayNum<=lastDay.getDate();
    var dateStr=isValid?viewYear+'-'+String(viewMonth+1).padStart(2,'0')+'-'+String(dayNum).padStart(2,'0'):'';
    var isHoje=dateStr===hj;
    var isSel=dateStr===_selectedDay;
    var evs=evMap[dateStr]||[];
    var cls='cal-day'+(isSel?' cal-selected':isHoje?' hoje':evs.length?' tem-descarga':'')+(isValid?'':' outro-mes');
    html+='<div class="'+cls+'"'+(isValid?' onclick="selectCalDay(\''+dateStr+'\')" style="cursor:pointer"':'')+' title="'+(isValid?fmtDate(dateStr):'')+'">';
    if(isValid){
      html+='<div class="cal-day-num">'+dayNum+'</div>';
      evs.slice(0,2).forEach(function(e){
        html+='<div class="cal-evento'+(e._tipo==='transferencia'?' transf':'')+'">'+e.material.substring(0,8)+'</div>';
      });
      if(evs.length>2)html+='<div style="font-size:9px;color:var(--text2)">+'+(evs.length-2)+'</div>';
    }
    html+='</div>';
  }
  html+='</div>';
  el.innerHTML=html;
  renderCapBar();
}

var _selectedDay=null;

function selectCalDay(dateStr){
  _selectedDay=dateStr;
  renderCalendarioDash();
}

function clearCalSel(){
  _selectedDay=null;
  renderCalendarioDash();
}

function calNav(dir){
  var el=document.getElementById('calendario-dash');if(!el)return;
  var y=parseInt(el.getAttribute('data-year'));
  var m=parseInt(el.getAttribute('data-month'))+dir;
  if(m<0){m=11;y--;}if(m>11){m=0;y++;}
  el.setAttribute('data-year',y);el.setAttribute('data-month',m);
  renderCalendarioDash();
}

function showCalDay(dateStr){selectCalDay(dateStr);}

function renderCapBar(){
  var cap=S.capDiaria||160;
  var day=_selectedDay||hoje();
  var capDia=getCapDia(day);
  var pct=Math.min(100,Math.round(capDia/cap*100));
  var cor=pct>=100?'#E24B4A':pct>=75?'#EF9F27':'#1D9E75';
  var capDisp=Math.max(0,cap-capDia);
  var isToday=day===hoje();
  var dayLabel=isToday?'hoje':'em '+fmtDate(day);
  var capEl=document.getElementById('cap-bar');
  if(capEl){
    capEl.innerHTML=''
      +'<div style="display:flex;align-items:baseline;gap:6px;margin-bottom:5px">'
        +'<span style="font-size:18px;font-weight:600">'+fmtKg(capDia)+'</span>'
        +'<span style="font-size:12px;color:var(--text2)">de '+fmtKg(cap)+' '+dayLabel+'</span>'
        +'<span class="badge '+(pct>=100?'badge-red':pct>=75?'badge-amber':'badge-green')+'" style="margin-left:auto">'+pct+'%</span>'
      +'</div>'
      +'<div class="progress-bar" style="height:7px"><div class="progress-fill" style="width:'+pct+'%;background:'+cor+'"></div></div>'
      +'<div style="font-size:11px;color:var(--text2);margin-top:6px">'+(pct>=100?'Capacidade esgotada':'Disponivel: '+fmtKg(capDisp))+'</div>';
  }
  var notifEl=document.getElementById('notif-lista');
  if(!notifEl)return;
  var notifs=[];
  var evsDia=getAllProximasDescargas().filter(function(d){return d.data===day;});
  if(isToday){
    var atrasados=getAtrasados();
    atrasados.forEach(function(d){notifs.push({tipo:'danger',msg:'Atraso: '+d.fornecedor+' - '+d.material+' (prev. '+d.hora+')'});});
    if(capDia/cap>=0.8&&capDia<=cap)notifs.push({tipo:'warning',msg:'Capacidade hoje: '+pct+'% utilizada'});
    S.baias.forEach(function(b){
      if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
      var tot=b.estoques.reduce(function(s,e){return s+e.qtdAtual;},0);
      if(b.cap>0&&tot/b.cap>0.9)notifs.push({tipo:'warning',msg:'Baia '+b.nome+' quase cheia ('+Math.round(tot/b.cap*100)+'%)'});
    });
  }
  evsDia.forEach(function(d){
    notifs.push({tipo:'info',msg:(isToday?'':fmtDate(d.data)+' ')+d.hora+' - '+d.material+' ('+d.fornecedor+')'+(d.toneladas?' '+fmtKg(d.toneladas):'')});
  });
  if(!notifs.length){
    notifEl.innerHTML='<div style="font-size:12px;color:var(--text3);padding-top:6px">Sem eventos'+(isToday?' no momento':' em '+fmtDate(day))+'.</div>';
    return;
  }
  var cmap={danger:'#E24B4A',warning:'#EF9F27',info:'#378ADD',success:'#1D9E75'};
  notifEl.innerHTML='<div style="font-size:11px;font-weight:600;color:var(--text2);margin-bottom:6px;padding-top:8px;border-top:1px solid var(--border)">'+(isToday?'Hoje':'Descargas em '+fmtDate(day))+' ('+notifs.length+')</div>'
    +notifs.map(function(n){return '<div class="notif-item"><div class="notif-dot" style="background:'+cmap[n.tipo]+'"></div><div style="font-size:12px">'+n.msg+'</div></div>';}).join('');
}

function salvarMaterial(){
  const n=document.getElementById('mat-nome').value.trim();if(!n){alert('Informe o nome.');return;}
  S.materiais.push({id:uid(),nome:n,unidade:'kg',codigo:document.getElementById('mat-codigo').value,categoria:document.getElementById('mat-categoria').value,obs:document.getElementById('mat-obs').value});
  ['mat-nome','mat-codigo','mat-categoria','mat-obs'].forEach(id=>document.getElementById(id).value='');
  saveState();closeModal('modal-material');render();
}

function preencherFornTransf(){
  const fId=document.getElementById('tr-forn-sel').value;if(!fId)return;
  const f=S.fornecedores.find(x=>x.id===fId);
  if(f)document.getElementById('tr-fornecedor').value=f.nome;
}

function verificarDataTransf(){
  const dt=document.getElementById('tr-data-transf').value;
  const dm=document.getElementById('tr-data-matriz').value;
  const el=document.getElementById('tr-data-alert');
  if(!dt){el.innerHTML='';return;}
  if(dm&&dt<dm){el.innerHTML='<div class="alert-bar alert-warning">Data de transferência anterior à data de chegada na matriz.</div>';}
  else if(dt>=hoje()){el.innerHTML='<div class="alert-bar alert-info">Esta transferência aparecerá automaticamente nas próximas descargas do dashboard.</div>';}
  else el.innerHTML='';
}

function salvarTransferencia(){
  const f=document.getElementById('tr-fornecedor').value.trim();
  const nf=document.getElementById('tr-nf').value.trim();
  if(!f||!nf){alert('Informe fornecedor e número da NF.');return;}
  const unid='kg';
  const qtdRaw=parseFloat(document.getElementById('tr-qtd').value)||0;
  const qtdTon=qtdRaw; // input in tons, stored as tons
  const t={
    id:uid(),fornecedor:f,nf,
    mp:document.getElementById('tr-mp').value,
    mapa:document.getElementById('tr-mapa').value,
    material:document.getElementById('tr-material').value,
    qtd:qtdTon,unidadeOriginal:unid,qtdOriginal:qtdRaw,
    dataMatriz:document.getElementById('tr-data-matriz').value,
    dataTransf:document.getElementById('tr-data-transf').value,
    hora:document.getElementById('tr-hora').value||'08:00',
    baia:document.getElementById('tr-baia').value,
    analisado:document.getElementById('tr-analisado').checked,
    obs:document.getElementById('tr-obs').value,
    criadoEm:new Date().toISOString()
  };
  t.recebido=false;
  S.transferencias.push(t);
  ['tr-fornecedor','tr-nf','tr-mp','tr-mapa','tr-qtd','tr-obs'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('tr-analisado').checked=false;
  document.getElementById('tr-data-alert').innerHTML='';
  saveState();closeModal('modal-transferencia');render();
}

function toggleAnalisado(id){
  const t=S.transferencias.find(x=>x.id===id);if(!t)return;
  t.analisado=!t.analisado;saveState();render();
}

function excluirTransferencia(id){if(!confirm('Remover?'))return;S.transferencias=S.transferencias.filter(t=>t.id!==id);saveState();render();}

function renderMateriais(){
  const el=document.getElementById('materiais-lista');if(!el)return;
  if(!S.materiais.length){el.innerHTML='<div class="empty-state">Nenhum material cadastrado. Clique em "+ Material" para adicionar.</div>';return;}
  el.innerHTML='<table><thead><tr><th>Código</th><th>Nome</th><th>Categoria</th><th>Unidade</th><th>Em estoque (ton)</th><th>Observações</th><th></th></tr></thead><tbody>'+S.materiais.map(m=>{
    const qtd=S.baias.filter(b=>b.estoque&&b.estoque.fornecedorNome===m.nome).reduce((s,b)=>s+b.estoque.qtdAtual,0);
    return'<tr><td><span class="tag">'+(m.codigo||'—')+'</span></td><td style="font-weight:600">'+m.nome+'</td><td style="color:var(--text2)">'+(m.categoria||'—')+'</td><td>'+m.unidade+'</td><td style="font-weight:600">'+fmtKg(qtd)+'</td><td style="font-size:12px;color:var(--text2)">'+(m.obs||'—')+'</td><td><button class="btn btn-sm btn-danger" onclick="excluirMaterial(\''+m.id+'\')">&#x2715;</button></td></tr>';
  }).join('')+'</tbody></table>';
}

function excluirMaterial(id){if(!confirm('Remover material?'))return;S.materiais=S.materiais.filter(m=>m.id!==id);saveState();render();}

function renderTransferencias(){
  var el=document.getElementById('transferencias-lista');if(!el)return;
  var aEl=document.getElementById('transf-alerts');
  var hj=hoje();
  var atrasadas=S.transferencias.filter(function(t){return t.dataTransf&&t.dataTransf<hj&&!t.recebido;});
  if(aEl)aEl.innerHTML=atrasadas.length
    ?('<div class="alert-bar alert-warning" style="margin-bottom:10px">'+atrasadas.length+' transferencia(s) com data vencida nao recebida(s).</div>')
    :'';
  var pend=S.transferencias.filter(function(t){return !t.recebido;}).length;
  var badge=document.getElementById('badge-transf');
  if(badge){badge.style.display=pend>0?'inline-flex':'none';badge.textContent=pend;}
  if(!S.transferencias.length){el.innerHTML='<div class="empty-state">Nenhuma transferencia registrada.</div>';return;}
  var sorted=[].concat(S.transferencias).sort(function(a,b){
    if(!!a.recebido!==!!b.recebido)return a.recebido?1:-1;
    var da=a.dataTransf||a.dataMatriz||'9999';
    var db=b.dataTransf||b.dataMatriz||'9999';
    return da>db?1:-1;
  });
  el.innerHTML=sorted.map(function(t){
    var isRec=!!t.recebido;
    var vencida=!!(t.dataTransf&&t.dataTransf<hj&&!isRec);
    var tid=t.id;
    var cls='transf-card'+(isRec?' recebido':t.analisado?' analisado':t.dataTransf?' pendente-analise':' sem-data');
    var badges=(isRec?'<span class="badge badge-blue">Recebido</span>'
              :(t.analisado?'<span class="badge badge-green">Analisado</span>'
              :'<span class="badge badge-amber">Nao analisado</span>'))
      +(vencida?'<span class="badge badge-red">Vencida</span>':'');
    var dataHtml=t.dataTransf
      ?('<strong style="color:'+(isRec?'#185FA5':(t.dataTransf>=hj?'#1D9E75':'#E24B4A'))+'">'+fmtDate(t.dataTransf)+'</strong>')
      :'<span style="color:var(--text3)">Nao definida</span>';
    var recInfo=isRec
      ?('<div style="margin-top:4px;font-size:11px;color:#185FA5">Recebido em '
        +fmtDate(t.recebidoEm?t.recebidoEm.slice(0,10):'')
        +(t.recebidoBaia?' - Baia: <strong>'+t.recebidoBaia+'</strong>':'')
        +(t.recebidoTon?' - '+fmtKg(t.recebidoTon):'')
        +(t.recebidoOperador?' - '+t.recebidoOperador:'')
        +'</div>')
      :'';
    var btnRec=isRec
      ?('<button class="btn btn-sm" data-id="'+tid+'" onclick="desfazerRecebimentoTransf(this.dataset.id)" style="font-size:11px">Desfazer receb.</button>')
      :('<button class="btn btn-sm btn-success" data-id="'+tid+'" onclick="abrirConfirmarRecebTransf(this.dataset.id)" style="font-size:11px;white-space:nowrap">Confirmar receb.</button>');
    var btnAnal=!isRec
      ?('<button class="btn btn-sm '+(t.analisado?'btn-warning':'')+'" data-id="'+tid+'" onclick="toggleAnalisado(this.dataset.id)" style="font-size:11px">'+(t.analisado?'Desmarcar':'Marcar analisado')+'</button>')
      :'';
    return '<div class="'+cls+'">'
      +'<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:10px">'
        +'<div style="flex:1;min-width:0">'
          +'<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:4px">'
            +'<span style="font-size:13px;font-weight:600">'+(t.material||'Material')+' - '+t.fornecedor+'</span>'+badges
          +'</div>'
          +'<div style="font-size:12px;color:var(--text2);display:flex;flex-wrap:wrap;gap:12px;margin-bottom:3px">'
            +'<span>NF: <strong>'+t.nf+'</strong></span>'
            +(t.mp?'<span>MP: '+t.mp+'</span>':'')
            +(t.mapa?'<span>Mapa: '+t.mapa+'</span>':'')
            +(t.qtd?'<span>Qtd: '+fmtKg(t.qtd)+'</span>':'')
          +'</div>'
          +'<div style="font-size:11px;color:var(--text3)">'
            +'Chegada na matriz: '+(fmtDate(t.dataMatriz)||'nao informada')
            +' - Transferencia: '+dataHtml+(t.hora?' as '+t.hora:'')
            +(t.baia?' - Baia sugerida: '+t.baia:'')
          +'</div>'
          +recInfo
        +'</div>'
        +'<div style="display:flex;flex-direction:column;gap:4px;flex-shrink:0">'
          +btnRec+btnAnal
          +('<button class="btn btn-sm" data-id="'+tid+'" onclick="abrirEditarTransf(this.dataset.id)" style="font-size:11px">Editar</button>')
          +('<button class="btn btn-sm btn-danger" data-id="'+tid+'" onclick="excluirTransferencia(this.dataset.id)" style="font-size:11px">Remover</button>')
        +'</div>'
      +'</div>'
      +'</div>';
  }).join('');
}


function abrirEditarTransf(id){
  var t=S.transferencias.find(function(x){return x.id===id;});if(!t)return;
  document.getElementById('edit-transf-info').innerHTML='<strong>'+(t.material||'Material')+'</strong> - '+t.fornecedor+' | NF: '+t.nf+(t.mapa?' | Mapa: '+t.mapa:'');
  document.getElementById('edit-tr-fornecedor').value=t.fornecedor||'';
  document.getElementById('edit-tr-material').value=t.material||'';
  document.getElementById('edit-tr-nf').value=t.nf||'';
  document.getElementById('edit-tr-mp').value=t.mp||'';
  document.getElementById('edit-tr-mapa').value=t.mapa||'';
  document.getElementById('edit-tr-qtd').value=t.qtdOriginal||t.qtd||'';
  // edit-tr-unidade removed;
  document.getElementById('edit-tr-data-matriz').value=t.dataMatriz||'';
  document.getElementById('edit-tr-data-transf').value=t.dataTransf||'';
  document.getElementById('edit-tr-hora').value=t.hora||'08:00';
  document.getElementById('edit-tr-analisado').checked=!!t.analisado;
  document.getElementById('edit-tr-obs').value=t.obs||'';
  var sb=document.getElementById('edit-tr-baia');
  sb.innerHTML='<option value="">Sem sugestao</option>';
  S.baias.filter(function(b){return !b.estoque;}).forEach(function(b){
    var dep=S.depositos.find(function(d){return d.id===b.dep;});
    sb.innerHTML+='<option value="'+b.nome+'"'+(b.nome===t.baia?' selected':'')+'>'+b.nome+(dep?' ('+dep.nome+')':'')+'</option>';
  });
  document.getElementById('modal-editar-transf')._editId=id;
  openModal('modal-editar-transf');
}

function salvarEdicaoTransf(){
  var id=document.getElementById('modal-editar-transf')._editId;
  var t=S.transferencias.find(function(x){return x.id===id;});if(!t)return;
  var unid='kg';
  var qtdRaw=parseFloat(document.getElementById('edit-tr-qtd').value)||t.qtd;
  t.fornecedor=document.getElementById('edit-tr-fornecedor').value.trim()||t.fornecedor;
  t.material=document.getElementById('edit-tr-material').value.trim();
  t.nf=document.getElementById('edit-tr-nf').value.trim();
  t.mp=document.getElementById('edit-tr-mp').value.trim();
  t.mapa=document.getElementById('edit-tr-mapa').value.trim();
  t.qtdOriginal=qtdRaw;t.unidadeOriginal=unid;
  t.qtd=qtdRaw; // input in tons, stored as tons
  t.dataMatriz=document.getElementById('edit-tr-data-matriz').value;
  t.dataTransf=document.getElementById('edit-tr-data-transf').value;
  t.hora=document.getElementById('edit-tr-hora').value||'08:00';
  t.baia=document.getElementById('edit-tr-baia').value;
  t.analisado=document.getElementById('edit-tr-analisado').checked;
  t.obs=document.getElementById('edit-tr-obs').value;
  saveState();closeModal('modal-editar-transf');render();
}

function abrirConfirmarRecebTransf(id){
  var t=S.transferencias.find(function(x){return x.id===id;});if(!t)return;
  document.getElementById('modal-confirmar-receb-transf')._transfId=id;
  document.getElementById('conf-transf-resumo').innerHTML=
    '<div style="font-size:13px;font-weight:600;margin-bottom:5px">'+(t.material||'Material')+' - '+t.fornecedor+'</div>'
    +'<div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:12px">'
    +'<span>NF: <strong>'+t.nf+'</strong></span>'
    +(t.mp?'<span>MP: '+t.mp+'</span>':'')
    +(t.mapa?'<span>Mapa: <strong>'+t.mapa+'</strong></span>':'')
    +(t.qtd?'<span>Qtd prevista: <strong>'+fmtKg(t.qtd)+'</strong></span>':'')
    +'<span>Transf.: <strong>'+fmtDate(t.dataTransf)+'</strong></span>'
    +'</div>';
  document.getElementById('conf-hora-real').value=new Date().toTimeString().slice(0,5);
  document.getElementById('conf-ton-real').value=t.qtd?t.qtd.toFixed(3):'';
  document.getElementById('conf-placa').value='';
  document.getElementById('conf-motorista').value='';
  document.getElementById('conf-obs').value='';
  var sb=document.getElementById('conf-baia');
  sb.innerHTML='<option value="">Selecionar baia</option>';
  S.baias.filter(function(b){return !b.estoque;}).forEach(function(b){
    var dep=S.depositos.find(function(d){return d.id===b.dep;});
    var isSug=b.nome===t.baia;
    sb.innerHTML+='<option value="'+b.id+'"'+(isSug?' selected':'')+'>'+b.nome+(dep?' ('+dep.nome+')':'')+(isSug?' (sugerida)':'')+'</option>';
  });
  var so=document.getElementById('conf-operador');
  so.innerHTML='<option value="">Selecionar</option>';
  S.operadores.forEach(function(o){
    so.innerHTML+='<option value="'+o.id+'"'+(o.id===S.turnoAtivo.operadorId?' selected':'')+'>'+o.nome+'</option>';
  });
  openModal('modal-confirmar-receb-transf');
}

function confirmarRecebimentoTransf(){
  var id=document.getElementById('modal-confirmar-receb-transf')._transfId;
  var t=S.transferencias.find(function(x){return x.id===id;});if(!t)return;
  var baiaId=document.getElementById('conf-baia').value;
  var tonReal=parseFloat(document.getElementById('conf-ton-real').value);
  if(!baiaId){alert('Selecione a baia de destino.');return;}
  if(!tonReal||tonReal<=0){alert('Informe a quantidade real recebida.');return;}
  var baia=S.baias.find(function(b){return b.id===baiaId;});if(!baia)return;
  var opId=document.getElementById('conf-operador').value;
  var opNomeStr=opNome(opId)||'Nao definido';
  var horaReal=document.getElementById('conf-hora-real').value;
  var placa=document.getElementById('conf-placa').value;
  var motorista=document.getElementById('conf-motorista').value;
  var obsConf=document.getElementById('conf-obs').value;
  // Aloca estoque na baia
  if(!baia.estoques){baia.estoques=baia.estoque?[baia.estoque]:[];delete baia.estoque;}
  if(baia.estoques.length>0){
    if(!confirm('Esta baia ja possui '+baia.estoques.length+' produto(s).\nDeseja misturar estoques?'))return;
  }
  baia.estoques.push({id:uid(),fornecedorNome:t.fornecedor,fornecedor:t.fornecedor,lote:t.mp||'—',qtdTotal:tonReal,qtdAtual:tonReal,qtdUsada:0,nf:t.nf,bigBags:0,dataEntrada:hoje()});
  // Marca transferencia como recebida
  t.recebido=true;t.recebidoEm=new Date().toISOString();
  t.recebidoBaia=baia.nome;t.recebidoTon=tonReal;t.recebidoOperador=opNomeStr;
  // Cria registro no historico de recebimentos
  var dcId=uid();
  var diverge=Math.abs(tonReal-(t.qtd||tonReal))>0.01;
  S.descargas.push({
    id:dcId,fornecedor:t.fornecedor,material:t.material||'Transferencia',
    toneladas:t.qtd||tonReal,data:hoje(),hora:t.hora||'08:00',
    lote:t.mp||'—',baia:baiaId,obs:obsConf,status:'armazenado',nf:t.nf,
    recebimento:{tonReal:tonReal,loteReal:t.mp||'—',horaReal:horaReal,placa:placa,motorista:motorista,nfReal:t.nf,obs:obsConf,divergencia:diverge,operadorId:opId},
    alocacao:{baiaId:baiaId,baiaName:baia.nome,seguiuSugestao:baia.nome===t.baia,baiaSugName:t.baia||null,obs:obsConf,alocadoAt:new Date().toISOString()}
  });
  S.recebimentos.push({
    id:uid(),dcId:dcId,fornecedor:t.fornecedor,material:t.material||'Transferencia',
    tonPrev:t.qtd||tonReal,tonReal:tonReal,lote:t.mp||'—',nf:t.nf,
    horaChegada:horaReal,placa:placa,motorista:motorista,baiaNome:baia.nome,
    baiaSugNome:t.baia||null,seguiuSugestao:baia.nome===t.baia,divergencia:diverge,
    data:hoje(),operador:opNomeStr,turno:turnoLabel(S.turnoAtivo.turno),
    finalizadoAt:new Date().toISOString(),origemTransferencia:true,mapaNro:t.mapa||null
  });
  S.movimentacoes.push({
    id:uid(),tipo:'entrada',baiaId:baia.id,baiaNome:baia.nome,
    desc:'Entrada via transferencia: '+(t.material||'—')+' - '+t.fornecedor+' ('+fmtKg(tonReal)+', NF '+t.nf+(t.mapa?', Mapa '+t.mapa:'')+')',
    data:hoje(),operador:opNomeStr,nf:t.nf
  });
  saveState();closeModal('modal-confirmar-receb-transf');render();
  alert('Recebimento confirmado!\nBaia '+baia.nome+' alocada com '+fmtKg(tonReal)+'.\nRegistro criado na aba Recebimento.');
}

function desfazerRecebimentoTransf(id){
  if(!confirm('Desfazer recebimento? O estoque da baia sera removido.'))return;
  var t=S.transferencias.find(function(x){return x.id===id;});if(!t)return;
  var baia=S.baias.find(function(b){return b.nome===t.recebidoBaia;});
  if(baia&&baia.estoque&&baia.estoque.nf===t.nf)baia.estoque=null;
  S.recebimentos=S.recebimentos.filter(function(r){return !(r.origemTransferencia&&r.nf===t.nf);});
  S.descargas=S.descargas.filter(function(d){return !(d.nf===t.nf&&d.status==='armazenado'&&d.material===(t.material||'Transferencia'));});
  t.recebido=false;
  delete t.recebidoEm;delete t.recebidoBaia;delete t.recebidoTon;delete t.recebidoOperador;
  saveState();render();
}


function baiaPrimEstoque(b){
  // backward compat: support both old .estoque and new .estoques[]
  if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
  return b.estoques.length>0?b.estoques[0]:null;
}
function baiaTemEstoque(b){
  if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
  return b.estoques&&b.estoques.length>0;
}
function baiaTotalAtual(b){
  if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
  return (b.estoques||[]).reduce(function(s,e){return s+e.qtdAtual;},0);
}

function syncMatSel(){
  var sel=document.getElementById('ag-material-sel');
  var inp=document.getElementById('ag-material');
  if(sel&&inp&&sel.value)inp.value=sel.value;
}

function syncEmForn(){var s=document.getElementById('em-forn-sel');var i=document.getElementById('em-fornecedor');if(s&&i&&s.value)i.value=s.value;}
function syncEmMat(){var s=document.getElementById('em-mat-sel');var i=document.getElementById('em-material');if(s&&i&&s.value)i.value=s.value;}

function openEntradaManual(){
  var sb=document.getElementById('em-baia');
  sb.innerHTML='<option value="">Selecionar baia</option>';
  S.baias.forEach(function(b){var dep=S.depositos.find(function(d){return d.id===b.dep;});sb.innerHTML+='<option value="'+b.id+'">'+b.nome+(dep?' ('+dep.nome+')':'')+'</option>';});
  var sf=document.getElementById('em-forn-sel');
  sf.innerHTML='<option value="">Selecionar cadastrado</option>';
  S.fornecedores.forEach(function(f){sf.innerHTML+='<option value="'+f.nome+'">'+f.nome+'</option>';});
  var sm=document.getElementById('em-mat-sel');
  sm.innerHTML='<option value="">Selecionar cadastrado</option>';
  S.materiais.forEach(function(m){sm.innerHTML+='<option value="'+m.nome+'">'+m.nome+'</option>';});
  var today=new Date().toISOString().slice(0,10);
  document.getElementById('em-data').value=today;
  ['em-fornecedor','em-material','em-lote','em-nf','em-obs'].forEach(function(id){document.getElementById(id).value='';});
  document.getElementById('em-qtd').value='';
  document.getElementById('em-bigbags').value='0';
  document.getElementById('em-baia').value='';
  openModal('modal-entrada-manual');
}

function lancarEstoqueManual(){
  var baiaId=document.getElementById('em-baia').value;
  var forn=document.getElementById('em-fornecedor').value.trim();
  var qtdRaw=parseFloat(document.getElementById('em-qtd').value);
  if(!baiaId){alert('Selecione a baia.');return;}
  if(!forn){alert('Informe o fornecedor.');return;}
  if(!qtdRaw||qtdRaw<=0){alert('Informe a quantidade.');return;}
  var baia=S.baias.find(function(b){return b.id===baiaId;});if(!baia)return;
  if(!baia.estoques){baia.estoques=baia.estoque?[baia.estoque]:[];delete baia.estoque;}
  if(baia.estoques.length>0){
    if(!confirm('Esta baia ja possui '+baia.estoques.length+' produto(s) armazenado(s).\nDeseja adicionar mais um produto?'))return;
  }
  var unid='kg';
  var qtdTon=qtdRaw; // input in tons, stored as tons
  var mat=document.getElementById('em-material').value.trim();
  var lote=document.getElementById('em-lote').value.trim();
  var nf=document.getElementById('em-nf').value.trim();
  var bb=parseInt(document.getElementById('em-bigbags').value)||0;
  var dtEntrada=document.getElementById('em-data').value||hoje();
  var novoEst={
    id:uid(),fornecedorNome:forn,fornecedor:forn,material:mat,
    lote:lote||'—',qtdTotal:qtdTon,qtdAtual:qtdTon,qtdUsada:0,
    nf:nf||'—',bigBags:bb,dataEntrada:dtEntrada,manual:true
  };
  baia.estoques.push(novoEst);
  S.movimentacoes.push({
    id:uid(),tipo:'entrada',baiaId:baia.id,baiaNome:baia.nome,
    desc:'Entrada manual: '+(mat||forn)+' - '+forn+' ('+fmtKg(qtdTon)+(lote?' Lote '+lote:'')+')',
    data:dtEntrada,operador:opNome(S.turnoAtivo.operadorId)||'Manual',nf:nf||'—'
  });
  saveState();closeModal('modal-entrada-manual');render();
  alert('Estoque lancado com sucesso!\n'+fmtKg(qtdTon)+' alocados na baia '+baia.nome+'.');
}

function autoAvancarDescargas(){
  // Move descargas nao conformes (pendentes de dias anteriores) para hoje
  var hj=hoje();
  var amanha=proximoDia(hj);
  var alterou=false;
  S.descargas.forEach(function(d){
    if(d.status==='pendente'&&d.data<hj){
      // Calcular proximo dia util a partir do dia original
      var novoDia=hj; // padrao: hoje
      // Se ja passou de hoje tambem, usa amanha
      if(d.data<hj){
        var msg='';
        if(!d._avisoNaoConformeData||d._avisoNaoConformeData!==hj){
          d._avisoNaoConformeData=hj;
          d._naoConformeOrigem=d._naoConformeOrigem||d.data;
        }
        d.data=hj;
        d._reprogramado=true;
        alterou=true;
      }
    }
  });
  if(alterou)saveState();
  return alterou;
}

function proximoDia(dateStr){
  var d=new Date(dateStr+'T00:00:00');
  d.setDate(d.getDate()+1);
  return d.toISOString().slice(0,10);
}

function montarCorpoEmail(){
  var hj=hoje();
  var atrasados=S.descargas.filter(function(d){
    return d.status==='pendente'&&d.data===hj&&getAtrasados().find(function(a){return a.id===d.id;});
  });
  var reprogramados=S.descargas.filter(function(d){return d._reprogramado&&d.data===hj;});
  var pendentes=S.descargas.filter(function(d){return d.status==='pendente'&&d.data===hj;});
  var linhas=[];
  linhas.push('=== ALERTA VittiaAfford - '+fmtDate(hj)+' ===');
  linhas.push('');
  if(atrasados.length){
    linhas.push('** DESCARGAS COM ATRASO ('+atrasados.length+'):');
    atrasados.forEach(function(d){
      linhas.push('  - '+d.hora+' | '+d.fornecedor+' | '+d.material+' | '+fmtKg(d.toneladas)+' | NF: '+(d.nf||'—'));
    });
    linhas.push('');
  }
  if(reprogramados.length){
    linhas.push('** REPROGRAMADAS (nao conformes de dias anteriores) ('+reprogramados.length+'):');
    reprogramados.forEach(function(d){
      linhas.push('  - Origem: '+fmtDate(d._naoConformeOrigem||'—')+' | '+d.fornecedor+' | '+d.material+' | '+fmtKg(d.toneladas));
    });
    linhas.push('');
  }
  if(pendentes.length){
    linhas.push('** PENDENTES HOJE ('+pendentes.length+'):');
    pendentes.forEach(function(d){
      linhas.push('  - '+d.hora+' | '+d.fornecedor+' | '+d.material+' | '+fmtKg(d.toneladas));
    });
    linhas.push('');
  }
  var cap=S.capDiaria||160;
  var capHj=getCapDia(hj);
  linhas.push('Capacidade hoje: '+fmtKg(capHj)+' / '+fmtKg(cap)+' ('+Math.round(capHj/cap*100)+'%)');
  if(capHj>cap)linhas.push('ATENCAO: Capacidade diaria EXCEDIDA!');
  linhas.push('');
  linhas.push('Gerado automaticamente pelo VittiaAfford em '+new Date().toLocaleString('pt-BR'));
  return linhas.join('\n');
}

function enviarAlertaEmail(){
  var dest=S.emailAlerta||'';
  if(!dest){alert('Configure o email de destino nas Configuracoes.');return;}
  var corpo=montarCorpoEmail();
  var hj=hoje();
  var assunto='VittiaAfford - Alerta de Descargas '+fmtDate(hj);
  var uri='mailto:'+encodeURIComponent(dest)
    +'?subject='+encodeURIComponent(assunto)
    +'&body='+encodeURIComponent(corpo);
  window.location.href=uri;
}

function verificarHorarioEmail(){
  var horario=S.horarioAlerta;
  if(!horario||!S.emailAlerta)return;
  var agora=new Date();
  var hh=String(agora.getHours()).padStart(2,'0');
  var mm=String(agora.getMinutes()).padStart(2,'0');
  var horaAtual=hh+':'+mm;
  if(horaAtual===horario){
    var key='email_sent_'+hoje()+'_'+horario;
    if(!sessionStorage.getItem(key)){
      sessionStorage.setItem(key,'1');
      enviarAlertaEmail();
    }
  }
}

function salvarEmailConfig(){
  S.emailAlerta=document.getElementById('cfg-email').value.trim();
  S.horarioAlerta=document.getElementById('cfg-horario-alerta').value;
  S.alertaAtivo=document.getElementById('cfg-alerta-ativo').checked;
  saveState();
  var msg=S.alertaAtivo&&S.emailAlerta&&S.horarioAlerta
    ?'Alerta configurado: '+S.emailAlerta+' as '+S.horarioAlerta
    :'Configuracao salva.';
  alert(msg);
}

function verPreviewEmail(){
  var el=document.getElementById('email-preview');
  if(el.style.display==='none'){
    el.textContent=montarCorpoEmail();
    el.style.display='block';
  } else {
    el.style.display='none';
  }
}


/* ── OCORRÊNCIAS ─────────────────────────────────────── */
var TIPO_LABEL={carga_avariada:'Carga avariada',divergencia_peso:'Divergência de peso',doc_incorreta:'Doc. incorreta',atraso:'Atraso',qualidade:'Qualidade',outro:'Outro'};
var GRAV_LABEL={baixa:'Baixa',media:'Média',alta:'Alta',critica:'Crítica'};
var STATUS_LABEL={aberta:'Aberta',em_analise:'Em análise',resolvida:'Resolvida'};
var STATUS_BADGE={aberta:'badge-red',em_analise:'badge-amber',resolvida:'badge-green'};

function syncOcForn(){
  var s=document.getElementById('oc-forn-sel');
  var i=document.getElementById('oc-fornecedor');
  if(s&&i&&s.value)i.value=s.value;
}

function abrirModalOcorrencia(id){
  var oc=id?S.ocorrencias.find(function(o){return o.id===id;}):null;
  document.getElementById('ocorr-modal-title').textContent=oc?'Editar Ocorrência':'Registrar Ocorrência';
  document.getElementById('modal-ocorrencia')._editId=oc?id:null;
  var sf=document.getElementById('oc-forn-sel');
  sf.innerHTML='<option value="">Selecionar cadastrado</option>';
  S.fornecedores.forEach(function(f){sf.innerHTML+='<option value="'+f.nome+'">'+f.nome+'</option>';});
  var sd=document.getElementById('oc-descarga');
  sd.innerHTML='<option value="">Nenhuma específica</option>';
  S.recebimentos.slice(-30).reverse().forEach(function(r){
    sd.innerHTML+='<option value="'+r.id+'">'+fmtDate(r.data)+' - '+r.fornecedor+' ('+fmtKg(r.tonReal)+')</option>';
  });
  var so=document.getElementById('oc-operador');
  so.innerHTML='<option value="">Selecionar</option>';
  S.operadores.forEach(function(o){so.innerHTML+='<option value="'+o.nome+'">'+o.nome+'</option>';});
  if(oc){
    document.getElementById('oc-fornecedor').value=oc.fornecedor||'';
    document.getElementById('oc-tipo').value=oc.tipo||'outro';
    document.getElementById('oc-gravidade').value=oc.gravidade||'media';
    document.getElementById('oc-data').value=oc.data||'';
    document.getElementById('oc-nf').value=oc.nf||'';
    document.getElementById('oc-descricao').value=oc.descricao||'';
    document.getElementById('oc-acao').value=oc.acao||'';
    document.getElementById('oc-status').value=oc.status||'aberta';
    document.getElementById('oc-operador').value=oc.responsavel||'';
  } else {
    ['oc-fornecedor','oc-nf','oc-descricao','oc-acao'].forEach(function(id){document.getElementById(id).value='';});
    document.getElementById('oc-data').value=hoje();
    document.getElementById('oc-tipo').value='outro';
    document.getElementById('oc-gravidade').value='media';
    document.getElementById('oc-status').value='aberta';
    document.getElementById('oc-operador').value='';
  }
  openModal('modal-ocorrencia');
}

function salvarOcorrencia(){
  var forn=document.getElementById('oc-fornecedor').value.trim();
  var desc=document.getElementById('oc-descricao').value.trim();
  if(!forn||!desc){alert('Informe o fornecedor e a descrição.');return;}
  var editId=document.getElementById('modal-ocorrencia')._editId;
  var oc={
    id:editId||uid(),
    fornecedor:forn,
    tipo:document.getElementById('oc-tipo').value,
    gravidade:document.getElementById('oc-gravidade').value,
    data:document.getElementById('oc-data').value||hoje(),
    nf:document.getElementById('oc-nf').value.trim(),
    descricao:desc,
    acao:document.getElementById('oc-acao').value.trim(),
    status:document.getElementById('oc-status').value,
    responsavel:document.getElementById('oc-operador').value,
    descargaId:document.getElementById('oc-descarga').value,
    criadoEm:editId?(S.ocorrencias.find(function(o){return o.id===editId;})||{}).criadoEm||new Date().toISOString():new Date().toISOString()
  };
  if(editId){
    var idx=S.ocorrencias.findIndex(function(o){return o.id===editId;});
    if(idx>=0)S.ocorrencias[idx]=oc;
  } else {
    S.ocorrencias.push(oc);
  }
  saveState();closeModal('modal-ocorrencia');render();
}

function excluirOcorrencia(id){
  if(!confirm('Remover ocorrência?'))return;
  S.ocorrencias=S.ocorrencias.filter(function(o){return o.id!==id;});
  saveState();render();
}

function alterarStatusOcorr(id,novoStatus){
  var oc=S.ocorrencias.find(function(o){return o.id===id;});
  if(!oc)return;
  oc.status=novoStatus;
  if(novoStatus==='resolvida')oc.resolvidaEm=new Date().toISOString();
  saveState();render();
}

function verDetalheOcorr(id){
  var oc=S.ocorrencias.find(function(o){return o.id===id;});if(!oc)return;
  document.getElementById('ocorr-det-title').textContent=TIPO_LABEL[oc.tipo]||oc.tipo;
  var rec=oc.descargaId?S.recebimentos.find(function(r){return r.id===oc.descargaId;}):null;
  document.getElementById('ocorr-det-body').innerHTML=
    '<div class="info-row"><span class="info-label">Fornecedor</span><span>'+oc.fornecedor+'</span></div>'
    +'<div class="info-row"><span class="info-label">Tipo</span><span>'+(TIPO_LABEL[oc.tipo]||oc.tipo)+'</span></div>'
    +'<div class="info-row"><span class="info-label">Gravidade</span><span><span class="badge '+(oc.gravidade==='critica'||oc.gravidade==='alta'?'badge-red':oc.gravidade==='media'?'badge-amber':'badge-green')+'">'+(GRAV_LABEL[oc.gravidade]||oc.gravidade)+'</span></span></div>'
    +'<div class="info-row"><span class="info-label">Status</span><span><span class="badge '+(STATUS_BADGE[oc.status]||'badge-gray')+'">'+(STATUS_LABEL[oc.status]||oc.status)+'</span></span></div>'
    +'<div class="info-row"><span class="info-label">Data</span><span>'+fmtDate(oc.data)+'</span></div>'
    +'<div class="info-row"><span class="info-label">NF</span><span>'+(oc.nf||'—')+'</span></div>'
    +(rec?'<div class="info-row"><span class="info-label">Descarga</span><span>'+fmtDate(rec.data)+' - '+fmtKg(rec.tonReal)+'</span></div>':'')
    +'<div class="info-row"><span class="info-label">Responsável</span><span>'+(oc.responsavel||'—')+'</span></div>'
    +'<div class="divider"></div>'
    +'<div style="margin-bottom:8px"><div class="form-label" style="margin-bottom:4px">Descrição</div><div style="font-size:13px">'+oc.descricao+'</div></div>'
    +(oc.acao?'<div><div class="form-label" style="margin-bottom:4px">Ação tomada</div><div style="font-size:13px">'+oc.acao+'</div></div>':'')
    +(oc.resolvidaEm?'<div class="info-row" style="margin-top:8px"><span class="info-label">Resolvida em</span><span>'+fmtDT(oc.resolvidaEm)+'</span></div>':'');
  openModal('modal-ocorr-detalhe');
}

function renderOcorrencias(){
  // no page guard needed
  var abertas=S.ocorrencias.filter(function(o){return o.status!=='resolvida';});
  var criticas=S.ocorrencias.filter(function(o){return o.gravidade==='critica';});
  var badge=document.getElementById('badge-ocorr');
  if(badge){badge.style.display=abertas.length>0?'inline-flex':'none';badge.textContent=abertas.length;}
  document.getElementById('ocorr-resumo').innerHTML=
    '<div class="metric-card"><div class="metric-label">Total de ocorrências</div><div class="metric-value">'+S.ocorrencias.length+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Em aberto</div><div class="metric-value" style="color:'+(abertas.length?'#A32D2D':'#0F6E56')+'">'+abertas.length+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Críticas</div><div class="metric-value" style="color:'+(criticas.length?'#A32D2D':'inherit')+'">'+criticas.length+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Resolvidas</div><div class="metric-value" style="color:#0F6E56">'+S.ocorrencias.filter(function(o){return o.status==='resolvida';}).length+'</div></div>';
  var tipoFiltro=document.getElementById('ocorr-filtro-tipo')?document.getElementById('ocorr-filtro-tipo').value:'';
  var statusFiltro=document.getElementById('ocorr-filtro-status')?document.getElementById('ocorr-filtro-status').value:'';
  var sorted=[].concat(S.ocorrencias).sort(function(a,b){
    var gOrder={critica:0,alta:1,media:2,baixa:3};
    if(a.status==='resolvida'!==b.status==='resolvida')return a.status==='resolvida'?1:-1;
    return (gOrder[a.gravidade]||2)-(gOrder[b.gravidade]||2);
  });
  var abEl=document.getElementById('ocorr-abertas');
  var abList=sorted.filter(function(o){return o.status!=='resolvida';}).slice(0,5);
  if(!abList.length){abEl.innerHTML='<div class="empty-state">Nenhuma ocorrência em aberto.</div>';}
  else abEl.innerHTML=abList.map(function(o){return renderOcorrCard(o,true);}).join('');
  var rankForn={};
  S.ocorrencias.forEach(function(o){
    if(!rankForn[o.fornecedor])rankForn[o.fornecedor]={total:0,criticas:0,abertas:0};
    rankForn[o.fornecedor].total++;
    if(o.gravidade==='critica')rankForn[o.fornecedor].criticas++;
    if(o.status!=='resolvida')rankForn[o.fornecedor].abertas++;
  });
  var rankArr=Object.keys(rankForn).map(function(k){return {nome:k,...rankForn[k]};}).sort(function(a,b){return b.total-a.total;});
  var rankEl=document.getElementById('ocorr-ranking');
  if(!rankArr.length){rankEl.innerHTML='<div class="empty-state">Nenhuma ocorrência registrada.</div>';}
  else rankEl.innerHTML='<table><thead><tr><th>Fornecedor</th><th>Total</th><th>Críticas</th><th>Em aberto</th></tr></thead><tbody>'
    +rankArr.slice(0,8).map(function(f){
      return '<tr>'
        +'<td style="font-weight:600">'+f.nome+'</td>'
        +'<td><span class="badge '+(f.total>5?'badge-red':f.total>2?'badge-amber':'badge-gray')+'">'+f.total+'</span></td>'
        +'<td style="color:'+(f.criticas>0?'#A32D2D':'inherit')+'">'+f.criticas+'</td>'
        +'<td style="color:'+(f.abertas>0?'#854F0B':'#0F6E56')+'">'+f.abertas+'</td>'
        +'</tr>';
    }).join('')+'</tbody></table>';
  var filtered=sorted.filter(function(o){
    return(!tipoFiltro||o.tipo===tipoFiltro)&&(!statusFiltro||o.status===statusFiltro);
  });
  var histEl=document.getElementById('ocorr-historico');
  if(!filtered.length){histEl.innerHTML='<div class="empty-state">Nenhuma ocorrência encontrada.</div>';return;}
  histEl.innerHTML=filtered.map(function(o){return renderOcorrCard(o,false);}).join('');
}

function renderOcorrCard(o,compact){
  var gravBadge='<span class="badge '+(o.gravidade==='critica'||o.gravidade==='alta'?'badge-red':o.gravidade==='media'?'badge-amber':'badge-green')+'">'+(GRAV_LABEL[o.gravidade]||o.gravidade)+'</span>';
  var statusBadge='<span class="badge '+(STATUS_BADGE[o.status]||'badge-gray')+'">'+(STATUS_LABEL[o.status]||o.status)+'</span>';
  var btns='<button class="btn btn-sm" data-id="'+o.id+'" onclick="verDetalheOcorr(this.dataset.id)" style="font-size:11px">Ver</button>'
    +'<button class="btn btn-sm" data-id="'+o.id+'" onclick="abrirModalOcorrencia(this.dataset.id)" style="font-size:11px">Editar</button>';
  if(o.status!=='resolvida')btns+='<button class="btn btn-sm btn-success" data-id="'+o.id+'" onclick="alterarStatusOcorr(this.dataset.id,\'resolvida\')" style="font-size:11px">Resolver</button>';
  if(!compact)btns+='<button class="btn btn-sm btn-danger" data-id="'+o.id+'" onclick="excluirOcorrencia(this.dataset.id)" style="font-size:11px">✕</button>';
  return '<div class="ocorr-card '+o.gravidade+'">'
    +'<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px">'
      +'<div style="flex:1;min-width:0">'
        +'<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:4px">'
          +'<span style="font-size:13px;font-weight:600">'+(TIPO_LABEL[o.tipo]||o.tipo)+'</span>'
          +gravBadge+statusBadge
        +'</div>'
        +'<div style="font-size:12px;color:var(--text2);margin-bottom:3px">'+o.fornecedor+(o.nf?' · NF: '+o.nf:'')+(o.responsavel?' · '+o.responsavel:'')+'</div>'
        +'<div style="font-size:11px;color:var(--text3)">'+fmtDate(o.data)+' · '+(compact?o.descricao.slice(0,80):o.descricao)+(compact&&o.descricao.length>80?'...':'')+'</div>'
      +'</div>'
      +'<div style="display:flex;flex-direction:column;gap:4px;flex-shrink:0">'+btns+'</div>'
    +'</div>'
    +'</div>';
}

/* ── KPIs ─────────────────────────────────────────────── */
var _kpiCharts={};

function destroyChart(id){
  if(_kpiCharts[id]){_kpiCharts[id].destroy();delete _kpiCharts[id];}
}

function diasAtras(n){
  var d=new Date();d.setDate(d.getDate()-n);return d.toISOString().slice(0,10);
}

function semanaKey(dateStr){
  var d=new Date(dateStr+'T00:00:00');
  var dow=d.getDay();var diff=d.getDate()-(dow===0?6:dow-1);
  var mon=new Date(d.setDate(diff));
  return mon.toISOString().slice(0,10);
}

function renderKPIs(){
  if(!document.getElementById('page-kpis').classList.contains('active'))return;
  var periodo=parseInt(document.getElementById('kpi-periodo').value)||30;
  var desde=diasAtras(periodo);
  var recsFiltro=S.recebimentos.filter(function(r){return r.data>=desde;});
  var baixasFiltro=S.baixas.filter(function(b){return b.data>=desde;});
  var totalEntrada=recsFiltro.reduce(function(s,r){return s+r.tonReal;},0);
  var totalSaida=baixasFiltro.reduce(function(s,b){return s+b.qtd;},0);
  var totalDiverg=recsFiltro.filter(function(r){return r.divergencia;}).length;
  var txDiv=recsFiltro.length?Math.round(totalDiverg/recsFiltro.length*100):0;
  var nBaias=S.baias.length;
  var nOcup=S.baias.filter(function(b){return b.estoques&&b.estoques.length>0;}).length;
  var ocAbertas=S.ocorrencias.filter(function(o){return o.status!=='resolvida';}).length;
  document.getElementById('kpi-cards').innerHTML=
    '<div class="metric-card"><div class="metric-label">Entradas no período</div><div class="metric-value">'+fmtKg(totalEntrada)+'</div><div class="metric-sub">'+recsFiltro.length+' recebimentos</div></div>'
    +'<div class="metric-card"><div class="metric-label">Saídas no período</div><div class="metric-value">'+fmtKg(totalSaida)+'</div><div class="metric-sub">'+baixasFiltro.length+' baixas</div></div>'
    +'<div class="metric-card"><div class="metric-label">Taxa de divergência</div><div class="metric-value" style="color:'+(txDiv>10?'#A32D2D':txDiv>5?'#854F0B':'#0F6E56')+'">'+txDiv+'%</div><div class="metric-sub">'+totalDiverg+' de '+recsFiltro.length+' receb.</div></div>'
    +'<div class="metric-card"><div class="metric-label">Ocorrências abertas</div><div class="metric-value" style="color:'+(ocAbertas>0?'#A32D2D':'#0F6E56')+'">'+ocAbertas+'</div><div class="metric-sub">Baias ocup.: '+nOcup+'/'+nBaias+'</div></div>';
  // Chart 1: Entradas por semana
  var semEnt={};
  recsFiltro.forEach(function(r){var k=semanaKey(r.data);if(!semEnt[k])semEnt[k]=0;semEnt[k]+=r.tonReal;});
  var semKeys=Object.keys(semEnt).sort();
  destroyChart('entradas');
  var ctxE=document.getElementById('chart-entradas');
  if(ctxE){
    _kpiCharts['entradas']=new Chart(ctxE,{type:'bar',data:{labels:semKeys.map(function(k){return fmtDate(k);}),datasets:[{label:'Kg recebidas',data:semKeys.map(function(k){return Math.round(semEnt[k]*10)/10;}),backgroundColor:'#1D9E7566',borderColor:'#1D9E75',borderWidth:1.5,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'#00000011'}}}}});
  }
  // Chart 2: Saídas por semana
  var semSai={};
  baixasFiltro.forEach(function(b){var k=semanaKey(b.data);if(!semSai[k])semSai[k]=0;semSai[k]+=b.qtd;});
  var semSKeys=Object.keys(semSai).sort();
  destroyChart('saidas');
  var ctxS=document.getElementById('chart-saidas');
  if(ctxS){
    _kpiCharts['saidas']=new Chart(ctxS,{type:'bar',data:{labels:semSKeys.map(function(k){return fmtDate(k);}),datasets:[{label:'Kg baixadas',data:semSKeys.map(function(k){return Math.round(semSai[k]*10)/10;}),backgroundColor:'#E24B4A66',borderColor:'#E24B4A',borderWidth:1.5,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'#00000011'}}}}});
  }
  // Chart 3: Ranking fornecedores
  var fornTon={};
  recsFiltro.forEach(function(r){if(!fornTon[r.fornecedor])fornTon[r.fornecedor]=0;fornTon[r.fornecedor]+=r.tonReal;});
  var fornArr=Object.keys(fornTon).map(function(k){return{nome:k,ton:fornTon[k]};}).sort(function(a,b){return b.ton-a.ton;}).slice(0,8);
  destroyChart('fornecedores');
  var ctxF=document.getElementById('chart-fornecedores');
  if(ctxF&&fornArr.length){
    _kpiCharts['fornecedores']=new Chart(ctxF,{type:'bar',data:{labels:fornArr.map(function(f){return f.nome.length>14?f.nome.slice(0,14)+'…':f.nome;}),datasets:[{label: 'Kg',data:fornArr.map(function(f){return Math.round(f.ton*10)/10;}),backgroundColor:'#378ADD66',borderColor:'#378ADD',borderWidth:1.5,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,grid:{color:'#00000011'}}}}});
  }
  // Chart 4: Divergências por fornecedor
  var fornDiv={};
  recsFiltro.forEach(function(r){if(!fornDiv[r.fornecedor])fornDiv[r.fornecedor]={total:0,div:0};fornDiv[r.fornecedor].total++;if(r.divergencia)fornDiv[r.fornecedor].div++;});
  var divArr=Object.keys(fornDiv).filter(function(k){return fornDiv[k].total>=2;}).map(function(k){return{nome:k,pct:Math.round(fornDiv[k].div/fornDiv[k].total*100)};}).sort(function(a,b){return b.pct-a.pct;}).slice(0,8);
  destroyChart('divergencias');
  var ctxD=document.getElementById('chart-divergencias');
  if(ctxD&&divArr.length){
    _kpiCharts['divergencias']=new Chart(ctxD,{type:'bar',data:{labels:divArr.map(function(d){return d.nome.length>14?d.nome.slice(0,14)+'…':d.nome;}),datasets:[{label:'%',data:divArr.map(function(d){return d.pct;}),backgroundColor:divArr.map(function(d){return d.pct>20?'#E24B4A99':d.pct>10?'#EF9F2799':'#1D9E7566';}),borderColor:divArr.map(function(d){return d.pct>20?'#E24B4A':d.pct>10?'#EF9F27':'#1D9E75';}),borderWidth:1.5,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,max:100,grid:{color:'#00000011'}}}}});
  }
  // Chart 5: Ocupação das baias
  var baiaNomes=[],baiaOcup=[];
  S.baias.forEach(function(b){
    if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
    if(b.cap>0){
      var tot=b.estoques.reduce(function(s,e){return s+e.qtdAtual;},0);
      baiaNomes.push(b.nome);
      baiaOcup.push(Math.min(100,Math.round(tot/b.cap*100)));
    }
  });
  destroyChart('baias');
  var ctxB=document.getElementById('chart-baias');
  if(ctxB&&baiaNomes.length){
    _kpiCharts['baias']=new Chart(ctxB,{type:'bar',data:{labels:baiaNomes,datasets:[{label:'% ocupação',data:baiaOcup,backgroundColor:baiaOcup.map(function(p){return p>85?'#E24B4A99':p>60?'#EF9F2799':'#1D9E7566';}),borderColor:baiaOcup.map(function(p){return p>85?'#E24B4A':p>60?'#EF9F27':'#1D9E75';}),borderWidth:1.5,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,max:100,grid:{color:'#00000011'}}}}});
  }
  // Chart 6: Distribuição por material (donut)
  var matTon={};
  S.baias.forEach(function(b){
    (b.estoques||[]).forEach(function(e){
      var k=e.fornecedorNome||'Outros';
      if(!matTon[k])matTon[k]=0;
      matTon[k]+=e.qtdAtual;
    });
  });
  var matKeys=Object.keys(matTon);
  var colors=['#1D9E75','#378ADD','#EF9F27','#E24B4A','#534AB7','#D85A30','#639922','#185FA5'];
  destroyChart('materiais-pie');
  var ctxM=document.getElementById('chart-materiais-pie');
  if(ctxM&&matKeys.length){
    _kpiCharts['materiais-pie']=new Chart(ctxM,{type:'doughnut',data:{labels:matKeys,datasets:[{data:matKeys.map(function(k){return Math.round(matTon[k]*10)/10;}),backgroundColor:colors.slice(0,matKeys.length),borderWidth:2,borderColor:'#ffffff'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{font:{size:11},boxWidth:12}}}}});
  }
  // Tabela por turno
  var turnoStats={};
  recsFiltro.forEach(function(r){
    var t=r.turno||'Não definido';
    if(!turnoStats[t])turnoStats[t]={rec:0,ton:0,div:0};
    turnoStats[t].rec++;turnoStats[t].ton+=r.tonReal;
    if(r.divergencia)turnoStats[t].div++;
  });
  var turnoEl=document.getElementById('kpi-turnos-table');
  var tKeys=Object.keys(turnoStats);
  if(!tKeys.length){turnoEl.innerHTML='<div class="empty-state">Sem dados de recebimento no período.</div>';return;}
  turnoEl.innerHTML='<table><thead><tr><th>Turno</th><th>Recebimentos</th><th>Ton. recebidas</th><th>Divergências</th><th>Taxa diverg.</th></tr></thead><tbody>'
    +tKeys.map(function(t){
      var ts=turnoStats[t];var txD=Math.round(ts.div/ts.rec*100);
      return '<tr>'
        +'<td style="font-weight:600">'+t+'</td>'
        +'<td>'+ts.rec+'</td>'
        +'<td>'+fmtKg(ts.ton)+'</td>'
        +'<td style="color:'+(ts.div>0?'#A32D2D':'#0F6E56')+'">'+ts.div+'</td>'
        +'<td><span class="badge '+(txD>10?'badge-red':txD>5?'badge-amber':'badge-green')+'">'+txD+'%</span></td>'
        +'</tr>';
    }).join('')+'</tbody></table>';

  // Chart: Descargas por turno
  if(_kpiCharts['turnos']){_kpiCharts['turnos'].destroy();delete _kpiCharts['turnos'];}
  var ctx5=document.getElementById('chart-turnos');
  if(ctx5){
    var turnoMap={};
    recs.forEach(function(r){var t=r.turno||'Nao definido';if(!turnoMap[t])turnoMap[t]=0;turnoMap[t]+=r.tonReal;});
    var tKeys=Object.keys(turnoMap);
    if(tKeys.length){_kpiCharts['turnos']=new Chart(ctx5,{type:'doughnut',data:{labels:tKeys,datasets:[{data:tKeys.map(function(k){return Math.round(turnoMap[k]*10)/10;}),backgroundColor:['#378ADD88','#1D9E7588','#EF9F2788','#534AB788','#E24B4A88'],borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{boxWidth:10,font:{size:10}}}}}});}
    else{ctx5.parentElement.innerHTML='<div class="empty-state" style="padding:30px 10px">Sem dados no periodo</div>';}
  }
  // Ranking fornecedores
  var rkEl=document.getElementById('ranking-fornecedores');if(!rkEl)return;
  if(rkEl){
    var fMap={};
    recs.forEach(function(r){if(!fMap[r.fornecedor])fMap[r.fornecedor]={ton:0,n:0,dv:0};fMap[r.fornecedor].ton+=r.tonReal;fMap[r.fornecedor].n++;if(r.divergencia)fMap[r.fornecedor].dv++;});
    var fKeys=Object.keys(fMap).sort(function(a,b){return fMap[b].ton-fMap[a].ton;}).slice(0,8);
    if(!fKeys.length){rkEl.innerHTML='<div class="empty-state">Sem dados no periodo</div>';}
    else{var mx=fMap[fKeys[0]].ton;rkEl.innerHTML=fKeys.map(function(k,i){var f=fMap[k];var pct=Math.round(f.ton/mx*100);var txD=Math.round(f.dv/f.n*100);return '<div style="margin-bottom:9px"><div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px"><span><b>'+(i+1)+'.</b> '+(k.length>18?k.slice(0,17)+'...':k)+'</span><span style="color:var(--text2)">'+fmtKg(f.ton)+' '+(txD>0?'<span style="color:#A32D2D">'+txD+'%div</span>':'<span style="color:#0F6E56">ok</span>')+'</span></div><div style="height:5px;border-radius:3px;background:var(--border)"><div style="width:'+pct+'%;height:100%;border-radius:3px;background:'+(txD>15?'#E24B4A':txD>5?'#EF9F27':'#1D9E75')+'"></div></div></div>';}).join('');}
  }

}


function atualizarStatusOcorr(id,novoStatus){
  var oc=S.ocorrencias.find(function(x){return x.id===id;});if(!oc)return;
  oc.status=novoStatus;oc.atualizadoEm=new Date().toISOString();
  saveState();renderOcorrencias();
}
function editarOcorrencia(id){
  var oc=S.ocorrencias.find(function(x){return x.id===id;});if(!oc)return;
  openModal('modal-ocorrencia');
  setTimeout(function(){
    document.getElementById('ocorr-tipo').value=oc.tipo||'';
    document.getElementById('ocorr-gravidade').value=oc.gravidade||'media';
    document.getElementById('ocorr-fornecedor').value=oc.fornecedor||'';
    document.getElementById('ocorr-data').value=oc.data||hoje();
    document.getElementById('ocorr-desc').value=oc.descricao||'';
    document.getElementById('ocorr-acao').value=oc.acao||'';
    document.getElementById('ocorr-status').value=oc.status||'aberta';
    if(document.getElementById('ocorr-modal-title'))document.getElementById('ocorr-modal-title').textContent='Editar Ocorrencia';
    _ocorrEditId=id;
  },80);
}

/* ── DASHBOARD KPI MINI ───────────────────────────────── */
var _dashCharts={};
function renderDashKPIMini(){
  if(!document.getElementById('page-dashboard').classList.contains('active'))return;
  var desde=diasAtras(30);
  var recs=S.recebimentos.filter(function(r){return r.data>=desde;});
  var baixas=S.baixas.filter(function(b){return b.data>=desde;});
  var divs=recs.filter(function(r){return r.divergencia;}).length;
  var txDiv=recs.length?Math.round(divs/recs.length*100):0;
  var ocAb=S.ocorrencias.filter(function(o){return o.status!=='resolvida';}).length;
  var mini=document.getElementById('dash-kpi-mini');
  if(mini)mini.innerHTML=
    '<div class="metric-card"><div class="metric-label">Entradas (30d)</div><div class="metric-value" style="font-size:18px">'+fmtKg(recs.reduce(function(s,r){return s+r.tonReal;},0))+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Saídas (30d)</div><div class="metric-value" style="font-size:18px">'+fmtKg(baixas.reduce(function(s,b){return s+b.qtd;},0))+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Taxa divergência</div><div class="metric-value" style="font-size:18px;color:'+(txDiv>10?'#A32D2D':txDiv>5?'#854F0B':'#0F6E56')+'">'+txDiv+'%</div></div>'
    +'<div class="metric-card"><div class="metric-label">Ocorrências abertas</div><div class="metric-value" style="font-size:18px;color:'+(ocAb>0?'#A32D2D':'#0F6E56')+'">'+ocAb+'</div></div>';
  // Mini chart entradas/saidas por semana
  var semEnt={},semSai={};
  recs.forEach(function(r){var k=semanaKey(r.data);if(!semEnt[k])semEnt[k]=0;semEnt[k]+=r.tonReal;});
  baixas.forEach(function(b){var k=semanaKey(b.data);if(!semSai[k])semSai[k]=0;semSai[k]+=b.qtd;});
  var allKeys=[].concat(Object.keys(semEnt),Object.keys(semSai)).filter(function(v,i,a){return a.indexOf(v)===i;}).sort();
  if(_dashCharts.semana){_dashCharts.semana.destroy();delete _dashCharts.semana;}
  var ctxS=document.getElementById('dash-chart-semana');
  if(ctxS&&allKeys.length){
    _dashCharts.semana=new Chart(ctxS,{type:'bar',data:{labels:allKeys.map(function(k){return fmtDate(k);}),datasets:[
      {label:'Entrada',data:allKeys.map(function(k){return Math.round((semEnt[k]||0)*10)/10;}),backgroundColor:'#1D9E7566',borderColor:'#1D9E75',borderWidth:1,borderRadius:3},
      {label:'Saída',data:allKeys.map(function(k){return Math.round((semSai[k]||0)*10)/10;}),backgroundColor:'#E24B4A55',borderColor:'#E24B4A',borderWidth:1,borderRadius:3}
    ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:{size:10},boxWidth:10}}},scales:{y:{beginAtZero:true,grid:{color:'#00000011'},ticks:{font:{size:10}}},x:{ticks:{font:{size:9}}}}}});
  }
  // Mini chart ocupação baias
  S.baias.forEach(function(b){if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}});
  var bns=[],boc=[];
  S.baias.filter(function(b){return b.cap>0;}).forEach(function(b){
    var tot=b.estoques.reduce(function(s,e){return s+e.qtdAtual;},0);
    bns.push(b.nome);boc.push(Math.min(100,Math.round(tot/b.cap*100)));
  });
  if(_dashCharts.ocup){_dashCharts.ocup.destroy();delete _dashCharts.ocup;}
  var ctxO=document.getElementById('dash-chart-ocup');
  if(ctxO&&bns.length){
    _dashCharts.ocup=new Chart(ctxO,{type:'bar',data:{labels:bns,datasets:[{label:'Ocupação %',data:boc,backgroundColor:boc.map(function(p){return p>85?'#E24B4A99':p>60?'#EF9F2799':'#378ADD66';}),borderColor:boc.map(function(p){return p>85?'#E24B4A':p>60?'#EF9F27':'#378ADD';}),borderWidth:1,borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:{size:10},boxWidth:10}}},scales:{y:{beginAtZero:true,max:100,grid:{color:'#00000011'},ticks:{font:{size:10}}},x:{ticks:{font:{size:9}}}}}});
  }
}

/* ── PCP / ACERTO DE PESO ─────────────────────────────── */
function syncApProduto(){
  var s=document.getElementById('ap-produto-sel');
  var i=document.getElementById('ap-produto');
  if(s&&i&&s.value)i.value=s.value;
}

function calcAcerto(){
  var bb=parseInt(document.getElementById('ap-bb-qtd').value)||0;
  var kgBB=parseFloat(document.getElementById('ap-bb-kg').value)||1000;
  var real=parseFloat(document.getElementById('ap-qtd-real').value)||0;
  var totalBB=bb*kgBB;
  document.getElementById('ap-bb-total').value=totalBB.toFixed(1)+' kg';
  if(real>0&&bb>0){
    var diff=real-totalBB;
    // Divergência em kg
    document.getElementById('ap-diferenca').value=(diff>0?'+':'')+diff.toFixed(1)+' kg';
    // Kg real médio por big bag
    var kgRealPorBB=real/bb;
    document.getElementById('ap-pct-diverg').value=kgRealPorBB.toFixed(2)+' kg/BB';
    // Alerta baseado na divergência em kg
    var absDiff=Math.abs(diff);
    var alEl=document.getElementById('ap-alerta-div');
    var pct=totalBB>0?Math.round(absDiff/totalBB*1000)/10:0;
    if(absDiff>0){
      var difPorBB=Math.round((kgRealPorBB-kgBB)*100)/100;
      var msg='Divergencia total: '+(diff>0?'+':'')+diff.toFixed(1)+' kg'
        +' &nbsp;|&nbsp; Por big bag: '+(difPorBB>0?'+':'')+difPorBB.toFixed(2)+' kg'
        +' &nbsp;|&nbsp; Peso real medio: '+kgRealPorBB.toFixed(2)+' kg (padrao: '+kgBB+' kg)';
      if(pct>5)alEl.innerHTML='<div class="alert-bar alert-danger" style="margin-top:6px">'+msg+'</div>';
      else if(pct>2)alEl.innerHTML='<div class="alert-bar alert-warning" style="margin-top:6px">'+msg+'</div>';
      else alEl.innerHTML='<div class="alert-bar alert-success" style="margin-top:6px">'+msg+'</div>';
    } else {
      alEl.innerHTML='<div class="alert-bar alert-success" style="margin-top:6px">Sem divergencia. Peso real medio: '+kgRealPorBB.toFixed(2)+' kg/BB (padrao: '+kgBB+' kg).</div>';
    }
    if(totalBB>0)calcAcertoPreviewEstoque(diff/1000);
  } else {
    document.getElementById('ap-diferenca').value='—';
    document.getElementById('ap-pct-diverg').value='—';
    document.getElementById('ap-alerta-div').innerHTML='';
    var prev=document.getElementById('ap-ajuste-preview');
    if(prev)prev.innerHTML='';
  }
}

function calcAcertoPreviewEstoque(diffTon){
  var baiaId=document.getElementById('ap-baia-ajuste')?document.getElementById('ap-baia-ajuste').value:'';
  var prev=document.getElementById('ap-ajuste-preview');
  if(!prev)return;
  if(!baiaId){prev.innerHTML='';return;}
  var baia=S.baias.find(function(b){return b.id===baiaId;});
  if(!baia){prev.innerHTML='';return;}
  if(!baia.estoques){baia.estoques=baia.estoque?[baia.estoque]:[];delete baia.estoque;}
  var tot=baia.estoques.reduce(function(s,e){return s+e.qtdAtual;},0);
  var novoTot=Math.max(0,tot+diffTon);
  prev.innerHTML='Estoque atual da baia: <strong>'+fmtKg(tot)+'</strong> &nbsp;→&nbsp; Apos ajuste: <strong style="color:'+(novoTot>tot?'#0F6E56':'#A32D2D')+'">'+fmtKg(novoTot)+'</strong> ('+(diffTon>0?'+':'')+fmtKg(diffTon)+')';
}

function onAcertoBaiaChange(){
  var baiaId=document.getElementById('ap-baia-ajuste')?document.getElementById('ap-baia-ajuste').value:'';
  var loteSel=document.getElementById('ap-lote-ajuste');
  if(!loteSel)return;
  loteSel.innerHTML='<option value="">Todos os lotes da baia</option>';
  if(!baiaId){calcAcerto();return;}
  var baia=S.baias.find(function(b){return b.id===baiaId;});
  if(!baia){calcAcerto();return;}
  if(!baia.estoques){baia.estoques=baia.estoque?[baia.estoque]:[];delete baia.estoque;}
  baia.estoques.forEach(function(e){
    loteSel.innerHTML+='<option value="'+e.id+'">'+e.lote+' — '+e.fornecedorNome+' ('+fmtKg(e.qtdAtual)+')</option>';
  });
  calcAcerto();
}

function abrirModalAcerto(id){
  var a=id?S.acertosPeso.find(function(x){return x.id===id;}):null;
  document.getElementById('acerto-modal-title').textContent=a?'Editar Acerto de Peso':'Registrar Acerto de Peso';
  document.getElementById('modal-acerto-peso')._editId=a?id:null;
  var sp=document.getElementById('ap-produto-sel');
  sp.innerHTML='<option value="">Selecionar cadastrado</option>';
  S.materiais.forEach(function(m){sp.innerHTML+='<option value="'+m.nome+'">'+m.nome+'</option>';});
  var so=document.getElementById('ap-operador');
  so.innerHTML='<option value="">Selecionar</option>';
  S.operadores.forEach(function(o){so.innerHTML+='<option value="'+o.nome+'">'+o.nome+'</option>';});
  var sb=document.getElementById('ap-baia-ajuste');
  if(sb){
    sb.innerHTML='<option value="">Nao ajustar estoque</option>';
    S.baias.forEach(function(b){
      if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
      if(b.estoques.length>0){var dep=S.depositos.find(function(d){return d.id===b.dep;});sb.innerHTML+='<option value="'+b.id+'">'+b.nome+(dep?' ('+dep.nome+')':'')+'</option>';}
    });
    sb.onchange=onAcertoBaiaChange;
  }
  if(a){
    document.getElementById('ap-mapa').value=a.mapa||'';
    document.getElementById('ap-lote').value=a.lote||'';
    document.getElementById('ap-data').value=a.data||'';
    document.getElementById('ap-produto').value=a.produto||'';
    document.getElementById('ap-bb-qtd').value=a.bbQtd||0;
    document.getElementById('ap-bb-kg').value=a.bbKg||1000;
    document.getElementById('ap-qtd-real').value=a.qtdReal||0;
    document.getElementById('ap-obs').value=a.obs||'';
    calcAcerto();
  } else {
    ['ap-mapa','ap-lote','ap-produto','ap-obs'].forEach(function(id){document.getElementById(id).value='';});
    document.getElementById('ap-data').value=hoje();
    document.getElementById('ap-bb-qtd').value='';
    document.getElementById('ap-bb-kg').value=1000;
    document.getElementById('ap-qtd-real').value='';
    document.getElementById('ap-bb-total').value='';
    document.getElementById('ap-diferenca').value='';
    document.getElementById('ap-pct-diverg').value='';
    document.getElementById('ap-alerta-div').innerHTML='';
  }
  openModal('modal-acerto-peso');
}

function salvarAcertoPeso(){
  var mapa=document.getElementById('ap-mapa').value.trim();
  var lote=document.getElementById('ap-lote').value.trim();
  var prod=document.getElementById('ap-produto').value.trim();
  var qtdReal=parseFloat(document.getElementById('ap-qtd-real').value)||0;
  if(!mapa||!lote||!prod){alert('Preencha: Nº do Mapa, Lote e Produto.');return;}
  if(!qtdReal){alert('Informe a quantidade real carregada.');return;}
  var bb=parseInt(document.getElementById('ap-bb-qtd').value)||0;
  var kgBB=parseFloat(document.getElementById('ap-bb-kg').value)||1000;
  var totalBB=bb*kgBB;
  var diff=qtdReal-totalBB;
  var pct=totalBB>0?Math.round(diff/totalBB*1000)/10:0;
  var editId=document.getElementById('modal-acerto-peso')._editId;
  var acerto={
    id:editId||uid(),
    mapa:mapa,lote:lote,produto:prod,
    data:document.getElementById('ap-data').value||hoje(),
    bbQtd:bb,bbKg:kgBB,totalBBkg:totalBB,
    qtdReal:qtdReal,diferenca:diff,pctDiverg:pct,
    obs:document.getElementById('ap-obs').value.trim(),
    operador:document.getElementById('ap-operador').value,
    criadoEm:new Date().toISOString()
  };
  if(editId){
    var idx=S.acertosPeso.findIndex(function(a){return a.id===editId;});
    if(idx>=0)S.acertosPeso[idx]=acerto;
  } else {
    S.acertosPeso.push(acerto);
  }
  // Aplicar ajuste no estoque se baia selecionada
  var baiaAjId=document.getElementById('ap-baia-ajuste')?document.getElementById('ap-baia-ajuste').value:'';
  var loteAjId=document.getElementById('ap-lote-ajuste')?document.getElementById('ap-lote-ajuste').value:'';
  if(baiaAjId){
    var baiaAj=S.baias.find(function(b){return b.id===baiaAjId;});
    if(baiaAj&&baiaAj.estoques){
      var diffTon=diff/1000; // convert kg to ton
      if(loteAjId){
        // ajusta lote especifico
        var estAj=baiaAj.estoques.find(function(e){return e.id===loteAjId;});
        if(estAj){
          var qtdAntes=estAj.qtdAtual;
          estAj.qtdAtual=Math.max(0,estAj.qtdAtual+diffTon);
          if(diffTon<0)estAj.qtdUsada=Math.max(0,estAj.qtdUsada-diffTon);
          S.movimentacoes.push({id:uid(),tipo:'ajuste',baiaId:baiaAj.id,baiaNome:baiaAj.nome,
            desc:'Acerto de peso (PCP): mapa '+mapa+', lote '+lote+' — '+(diffTon>0?'+':'')+fmtKg(diffTon)+' ('+acerto.produto+')',
            data:acerto.data,operador:acerto.operador||'—',nf:'—'});
        }
      } else {
        // ajusta proporcional em todos os lotes da baia
        var totBaia=baiaAj.estoques.reduce(function(s,e){return s+e.qtdAtual;},0);
        baiaAj.estoques.forEach(function(e){
          var proporcao=totBaia>0?e.qtdAtual/totBaia:1/baiaAj.estoques.length;
          e.qtdAtual=Math.max(0,e.qtdAtual+diffTon*proporcao);
        });
        S.movimentacoes.push({id:uid(),tipo:'ajuste',baiaId:baiaAj.id,baiaNome:baiaAj.nome,
          desc:'Acerto de peso (PCP): mapa '+mapa+', lote '+lote+' — '+(diffTon>0?'+':'')+fmtKg(diffTon)+' ('+acerto.produto+')',
          data:acerto.data,operador:acerto.operador||'—',nf:'—'});
      }
    }
  }
  saveState();closeModal('modal-acerto-peso');render();
}

function excluirAcerto(id){
  if(!confirm('Remover este acerto?'))return;
  S.acertosPeso=S.acertosPeso.filter(function(a){return a.id!==id;});
  saveState();render();
}

function renderPCP(){
  if(!document.getElementById('page-pcp').classList.contains('active'))return;
  var total=S.acertosPeso.length;
  var comDiv=S.acertosPeso.filter(function(a){return Math.abs(a.diferenca)>20;}).length;
  var txDiv=total?Math.round(comDiv/total*100):0;
  var totalBBkg=S.acertosPeso.reduce(function(s,a){return s+(a.bbQtd>0?a.qtdReal/a.bbQtd:0);},0);
  var mediaKgBB=S.acertosPeso.filter(function(a){return a.bbQtd>0;}).length
    ?Math.round(totalBBkg/S.acertosPeso.filter(function(a){return a.bbQtd>0;}).length*10)/10:0;
  document.getElementById('pcp-resumo-cards').innerHTML=
    '<div class="metric-card"><div class="metric-label">Total de acertos</div><div class="metric-value">'+total+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Com divergência (&gt;20 kg)</div><div class="metric-value" style="color:'+(comDiv>0?'#A32D2D':'#0F6E56')+'">'+comDiv+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Taxa divergência</div><div class="metric-value" style="color:'+(txDiv>10?'#A32D2D':txDiv>5?'#854F0B':'#0F6E56')+'">'+txDiv+'%</div></div>'
    +'<div class="metric-card"><div class="metric-label">Média kg/big bag</div><div class="metric-value">'+mediaKgBB+'</div><div class="metric-sub">base: '+document.getElementById('ap-bb-kg')?'1000':'1000'+' kg</div></div>';
  // Divergência por produto
  var prodDiv={};
  S.acertosPeso.forEach(function(a){
    if(!prodDiv[a.produto])prodDiv[a.produto]={count:0,divTotal:0,kgBBTotal:0,bbCount:0};
    prodDiv[a.produto].count++;
    prodDiv[a.produto].divTotal+=Math.abs(a.diferenca||0);  // abs kg
    if(a.bbQtd>0){prodDiv[a.produto].kgBBTotal+=a.qtdReal/a.bbQtd;prodDiv[a.produto].bbCount++;}
  });
  var prodKeys=Object.keys(prodDiv);
  var divEl=document.getElementById('pcp-diverg-lista');
  if(!prodKeys.length){divEl.innerHTML='<div class="empty-state">Sem dados de acerto.</div>';}
  else divEl.innerHTML=prodKeys.map(function(p){
    var pd=prodDiv[p];
    var mediaDivKg=Math.round(pd.divTotal/pd.count*10)/10;  // media abs kg
    var barPct=Math.min(100,mediaDivKg/10);  // 100% = 1000kg
    return '<div style="margin-bottom:10px">'
      +'<div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px">'
        +'<span style="font-weight:600">'+p+'</span>'
        +'<span class="badge '+(mediaDivKg>100?'badge-red':mediaDivKg>20?'badge-amber':'badge-green')+'">'+mediaDivKg.toFixed(1)+' kg média</span>'
      +'</div>'
      +'<div class="progress-bar"><div class="progress-fill" style="width:'+barPct+'%;background:'+(mediaDivKg>100?'#E24B4A':mediaDivKg>20?'#EF9F27':'#1D9E75')+'"></div></div>'
      +'<div style="font-size:11px;color:var(--text2);margin-top:2px">'+pd.count+' acertos registrados</div>'
      +'</div>';
  }).join('');
  // Média kg por big bag
  var bbEl=document.getElementById('pcp-bigbag-media');
  if(!prodKeys.length){bbEl.innerHTML='<div class="empty-state">Sem dados.</div>';}
  else bbEl.innerHTML='<table><thead><tr><th>Produto</th><th>Média kg/BB</th><th>Padrão</th><th>Variação</th><th>Acertos</th></tr></thead><tbody>'
    +prodKeys.map(function(p){
      var pd=prodDiv[p];
      var media=pd.bbCount>0?Math.round(pd.kgBBTotal/pd.bbCount*10)/10:0;
      var padrao=1000;
      var var_=Math.round((media-padrao)*10)/10;
      return '<tr>'
        +'<td style="font-weight:600">'+p+'</td>'
        +'<td style="font-weight:600">'+media+' kg</td>'
        +'<td style="color:var(--text2)">'+padrao+' kg</td>'
        +'<td style="color:'+(Math.abs(var_)>50?'#A32D2D':Math.abs(var_)>20?'#854F0B':'#0F6E56')+'">'+(var_>0?'+':'')+var_+' kg</td>'
        +'<td>'+pd.count+'</td>'
        +'</tr>';
    }).join('')+'</tbody></table>';
  // populate filtros
  var sp=document.getElementById('pcp-filtro-produto');
  if(sp){
    var curP=sp.value;
    sp.innerHTML='<option value="">Todos os produtos</option>';
    prodKeys.forEach(function(p){sp.innerHTML+='<option value="'+p+'"'+(p===curP?' selected':'')+'>'+p+'</option>';});
  }
  var sm=document.getElementById('pcp-filtro-mes');
  if(sm){
    var meses={};
    S.acertosPeso.forEach(function(a){if(a.data)meses[a.data.slice(0,7)]=true;});
    var curM=sm.value;
    sm.innerHTML='<option value="">Todos os meses</option>';
    Object.keys(meses).sort().reverse().forEach(function(m){sm.innerHTML+='<option value="'+m+'"'+(m===curM?' selected':'')+'>'+m+'</option>';});
  }
  var fprod=document.getElementById('pcp-filtro-produto')?document.getElementById('pcp-filtro-produto').value:'';
  var fmes=document.getElementById('pcp-filtro-mes')?document.getElementById('pcp-filtro-mes').value:'';
  var lista=[].concat(S.acertosPeso).filter(function(a){
    return(!fprod||a.produto===fprod)&&(!fmes||a.data.slice(0,7)===fmes);
  }).sort(function(a,b){return b.criadoEm>a.criadoEm?1:-1;});
  var histEl=document.getElementById('pcp-historico');
  if(!lista.length){histEl.innerHTML='<div class="empty-state">Nenhum acerto registrado.</div>';return;}
  histEl.innerHTML='<div style="overflow-x:auto"><table><thead><tr><th>Data</th><th>Mapa</th><th>Lote</th><th>Produto</th><th>Big Bags</th><th>Total BB (kg)</th><th>Real (kg)</th><th>Diferença</th><th>Divergência (kg)</th><th>Kg real / BB</th><th>Operador</th><th></th></tr></thead><tbody>'
    +lista.map(function(a){
      var absPct=Math.abs(a.pctDiverg);
      var gravCls=absPct>5?'badge-red':absPct>2?'badge-amber':'badge-green';
      var mediaKg=a.bbQtd>0?Math.round(a.qtdReal/a.bbQtd*10)/10:0;
      return '<tr>'
        +'<td style="white-space:nowrap">'+fmtDate(a.data)+'</td>'
        +'<td><span class="tag">'+a.mapa+'</span></td>'
        +'<td><span class="tag">'+a.lote+'</span></td>'
        +'<td style="font-weight:600">'+a.produto+'</td>'
        +'<td>'+a.bbQtd+'</td>'
        +'<td>'+a.totalBBkg.toFixed(1)+'</td>'
        +'<td style="font-weight:600">'+a.qtdReal.toFixed(1)+'</td>'
        +'<td style="color:'+(a.diferenca<0?'#A32D2D':a.diferenca>0?'#0F6E56':'inherit')+'">'+(a.diferenca>0?'+':'')+a.diferenca.toFixed(1)+'</td>'
        +'<td><span class="badge '+gravCls+'">'+(a.diferenca>0?'+':'')+a.diferenca.toFixed(1)+' kg</span></td>'
        +'<td style="font-weight:600;color:#1D9E75">'+mediaKg.toFixed(2)+' kg</td>'
        +'<td style="font-size:12px;color:var(--text2)">'+(a.operador||'—')+'</td>'
        +'<td style="white-space:nowrap">'
          +'<button class="btn btn-sm" data-id="'+a.id+'" onclick="abrirModalAcerto(this.dataset.id)" style="font-size:11px">Editar</button>'
          +'<button class="btn btn-sm btn-danger" data-id="'+a.id+'" onclick="excluirAcerto(this.dataset.id)" style="font-size:11px">✕</button>'
        +'</td>'
        +'</tr>';
    }).join('')+'</tbody></table></div>';
}

/* ── CONFERÊNCIA ──────────────────────────────────────── */
function preencherConfProduto(){
  var prod=document.getElementById('cf-produto-sel')?document.getElementById('cf-produto-sel').value:'';
  var sist=document.getElementById('cf-qtd-sistema');
  var det=document.getElementById('cf-baias-detalhe');
  if(!prod){if(sist)sist.value='';if(det)det.style.display='none';return;}
  // collect all estoque entries matching this product
  var entries=[];
  S.baias.forEach(function(b){
    if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
    b.estoques.forEach(function(e){
      if(e.fornecedorNome===prod||e.material===prod)entries.push({baia:b.nome,lote:e.lote,qtd:e.qtdAtual,nf:e.nf});
    });
  });
  var tot=entries.reduce(function(s,e){return s+e.qtd;},0);
  if(sist)sist.value=fmtKg(tot);
  if(det){
    if(entries.length>0){
      det.innerHTML='<strong>Lotes em estoque:</strong> '
        +entries.map(function(e){return e.baia+' / '+e.lote+' ('+fmtKg(e.qtd)+')';}).join(' &nbsp;|&nbsp; ');
      det.style.display='';
    } else {
      det.innerHTML='Nenhum estoque encontrado para este produto.';
      det.style.display='';
    }
  }
  calcConferencia();
}

// keep old name as alias for backward compat
function preencherConfBaia(){preencherConfProduto();}

function calcConferencia(){
  var prod=document.getElementById('cf-produto-sel')?document.getElementById('cf-produto-sel').value:'';
  var fisica=parseFloat(document.getElementById('cf-qtd-fisica').value)||0;
  var divEl=document.getElementById('cf-divergencia');
  var alertEl=document.getElementById('cf-diverg-alert');
  if(!prod||!fisica){if(divEl)divEl.value='—';if(alertEl)alertEl.innerHTML='';return;}
  // sum all matching stock
  var tot=0;
  S.baias.forEach(function(b){
    if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
    b.estoques.forEach(function(e){if(e.fornecedorNome===prod||e.material===prod)tot+=e.qtdAtual;});
  });
  var diff=fisica-tot;
  if(divEl)divEl.value=(diff>0?'+':'')+fmtKg(diff);
  var resEl=document.getElementById('cf-resultado');
  if(resEl)resEl.value=Math.abs(diff)>0.001?'divergencia':'conforme';
  if(alertEl){
    var absDiff=Math.abs(diff);
    var pct=tot>0?Math.round(absDiff/tot*1000)/10:0;
    if(absDiff>0.001)
      alertEl.innerHTML='<div class="alert-bar '+(pct>5?'alert-danger':'alert-warning')+'" style="margin-bottom:8px">Divergencia de '+(diff>0?'+':'')+fmtKg(diff)+' ('+pct+'%). Sistema: '+fmtKg(tot)+' / Conferido: '+fmtKg(fisica)+'</div>';
    else
      alertEl.innerHTML='<div class="alert-bar alert-success" style="margin-bottom:8px">Estoque conforme. Sistema e conferencia coincidem.</div>';
  }
}

function toggleConfDiverg(){}

function abrirModalConferencia(id){
  var c=id?S.conferencias.find(function(x){return x.id===id;}):null;
  document.getElementById('conf-modal-title').textContent=c?'Editar Conferencia':'Nova Conferencia';
  document.getElementById('modal-conferencia')._editId=c?id:null;
  // build product list from: materiais cadastrados + fornecedores em estoque
  var prodSet={};
  S.materiais.forEach(function(m){prodSet[m.nome]=true;});
  S.baias.forEach(function(b){
    if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
    b.estoques.forEach(function(e){if(e.fornecedorNome)prodSet[e.fornecedorNome]=true;if(e.material)prodSet[e.material]=true;});
  });
  var sp=document.getElementById('cf-produto-sel');
  sp.innerHTML='<option value="">Selecionar produto</option>';
  Object.keys(prodSet).sort().forEach(function(p){sp.innerHTML+='<option value="'+p+'"'+(c&&c.produto===p?' selected':'')+'>'+p+'</option>';});
  var so=document.getElementById('cf-operador');
  so.innerHTML='<option value="">Selecionar</option>';
  S.operadores.forEach(function(o){so.innerHTML+='<option value="'+o.nome+'"'+(c&&c.operador===o.nome?' selected':'')+'>'+o.nome+'</option>';});
  if(c){
    document.getElementById('cf-data').value=c.data||hoje();
    document.getElementById('cf-qtd-fisica').value=c.qtdFisica||'';
    document.getElementById('cf-obs').value=c.obs||'';
    document.getElementById('cf-resultado').value=c.resultado||'conforme';
    if(c.produto){preencherConfProduto();}
  } else {
    document.getElementById('cf-data').value=hoje();
    document.getElementById('cf-qtd-fisica').value='';
    document.getElementById('cf-qtd-sistema').value='';
    document.getElementById('cf-divergencia').value='—';
    document.getElementById('cf-obs').value='';
    document.getElementById('cf-resultado').value='conforme';
    var det=document.getElementById('cf-baias-detalhe');
    if(det)det.style.display='none';
    var alertEl=document.getElementById('cf-diverg-alert');
    if(alertEl)alertEl.innerHTML='';
  }
  openModal('modal-conferencia');
}

function salvarConferencia(){
  var prod=document.getElementById('cf-produto-sel')?document.getElementById('cf-produto-sel').value:'';
  var qtdFisica=parseFloat(document.getElementById('cf-qtd-fisica').value);
  if(!prod){alert('Selecione o produto.');return;}
  if(isNaN(qtdFisica)||qtdFisica<0){alert('Informe a quantidade conferida.');return;}
  // sum system qty for this product
  var qtdSist=0;var baiasInfo=[];
  S.baias.forEach(function(b){
    if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
    b.estoques.forEach(function(e){
      if(e.fornecedorNome===prod||e.material===prod){qtdSist+=e.qtdAtual;baiasInfo.push(b.nome+'/'+e.lote);}
    });
  });
  var diff=qtdFisica-qtdSist;
  var resultado=Math.abs(diff)>0.001?'divergencia':'conforme';
  var editId=document.getElementById('modal-conferencia')._editId;
  var conf={
    id:editId||uid(),
    produto:prod,
    baiasInfo:baiasInfo.join(', '),
    // keep baiaNome for compat with renderConfCard
    baiaNome:prod,
    data:document.getElementById('cf-data').value||hoje(),
    qtdSistema:qtdSist,qtdFisica:qtdFisica,diferenca:diff,
    resultado:resultado,
    status:resultado==='conforme'?'conforme':'divergencia',
    obs:document.getElementById('cf-obs').value.trim(),
    operador:document.getElementById('cf-operador').value,
    criadoEm:new Date().toISOString()
  };
  if(editId){
    var idx=S.conferencias.findIndex(function(c){return c.id===editId;});
    if(idx>=0)S.conferencias[idx]=conf;
  } else {
    S.conferencias.push(conf);
  }
  saveState();closeModal('modal-conferencia');render();
}

function abrirInvestigacao(id){
  var c=S.conferencias.find(function(x){return x.id===id;});if(!c)return;
  document.getElementById('modal-investigacao')._confId=id;
  document.getElementById('inv-resumo').innerHTML=
    '<div style="font-weight:600;margin-bottom:5px">'+c.baiaNome+' — '+fmtDate(c.data)+'</div>'
    +'<div style="font-size:12px;color:var(--text2)">Sistema: <strong>'+fmtKg(c.qtdSistema)+'</strong> &nbsp;·&nbsp; Conferido: <strong>'+fmtKg(c.qtdFisica)+'</strong> &nbsp;·&nbsp; Diferença: <strong style="color:'+(c.diferenca<0?'#A32D2D':'#0F6E56')+'">'+(c.diferenca>0?'+':'')+fmtKg(c.diferenca)+'</strong></div>';
  document.getElementById('inv-status').value=c.statusInv||'investigando';
  document.getElementById('inv-motivo').value=c.motivoInv||'';
  document.getElementById('inv-obs').value=c.obsInv||'';
  document.getElementById('inv-ajustar').checked=false;
  document.getElementById('inv-ajuste-wrap').style.display='none';
  document.getElementById('inv-motivo-wrap').style.display='';
  openModal('modal-investigacao');
}

function toggleInvAjuste(){
  var cb=document.getElementById('inv-ajustar');
  // no extra UI needed
}

function salvarInvestigacao(){
  var id=document.getElementById('modal-investigacao')._confId;
  var c=S.conferencias.find(function(x){return x.id===id;});if(!c)return;
  var status=document.getElementById('inv-status').value;
  var motivo=document.getElementById('inv-motivo').value;
  var obs=document.getElementById('inv-obs').value.trim();
  if(status==='resolvida'&&!obs){alert('Informe as observações com o motivo da divergência.');return;}
  c.statusInv=status;c.motivoInv=motivo;c.obsInv=obs;
  c.status=status;
  if(status==='resolvida'){
    c.resolvidaEm=new Date().toISOString();
    // Ajustar estoque se solicitado
    if(document.getElementById('inv-ajustar').checked){
      var baia=S.baias.find(function(b){return b.id===c.baiaId;});
      if(baia&&baia.estoques&&baia.estoques.length>0){
        var diff=c.diferenca; // fisica - sistema
        var e=baia.estoques[0];
        e.qtdAtual=Math.max(0,e.qtdAtual+diff);
        S.movimentacoes.push({id:uid(),tipo:'ajuste',baiaId:baia.id,baiaNome:baia.nome,
          desc:'Ajuste por conferência: '+(diff>0?'+':'')+fmtKg(diff)+' — '+obs,
          data:hoje(),operador:c.operador||'—',nf:'—'});
      }
    }
  }
  saveState();closeModal('modal-investigacao');render();
}

function excluirConferencia(id){
  if(!confirm('Remover esta conferência?'))return;
  S.conferencias=S.conferencias.filter(function(c){return c.id!==id;});
  saveState();render();
}

function renderConferencia(){
  // no page guard
  var hj=hoje();
  var abertas=S.conferencias.filter(function(c){return c.status==='divergencia'||c.status==='investigando';});
  // add produto filtro options
  var fp=document.getElementById('conf-filtro-status');
  var confHj=S.conferencias.filter(function(c){return c.data===hj;});
  var badge=document.getElementById('badge-conf');
  if(badge){badge.style.display=abertas.length>0?'inline-flex':'none';badge.textContent=abertas.length;}
  var total=S.conferencias.length;
  var comDiv=S.conferencias.filter(function(c){return c.resultado==='divergencia';}).length;
  var resolvidas=S.conferencias.filter(function(c){return c.status==='resolvida';}).length;
  document.getElementById('conf-resumo-cards').innerHTML=
    '<div class="metric-card"><div class="metric-label">Total conferências</div><div class="metric-value">'+total+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Com divergência</div><div class="metric-value" style="color:'+(comDiv>0?'#A32D2D':'#0F6E56')+'">'+comDiv+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Em investigação</div><div class="metric-value" style="color:'+(abertas.length>0?'#854F0B':'#0F6E56')+'">'+abertas.length+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Resolvidas</div><div class="metric-value" style="color:#0F6E56">'+resolvidas+'</div></div>';
  // Pendentes
  var pEl=document.getElementById('conf-pendentes');
  if(!abertas.length){pEl.innerHTML='<div class="empty-state">Nenhuma divergência em aberto.</div>';}
  else pEl.innerHTML=abertas.map(function(c){return renderConfCard(c);}).join('');
  // Hoje
  var hEl=document.getElementById('conf-hoje');
  if(!confHj.length){hEl.innerHTML='<div class="empty-state">Nenhuma conferência registrada hoje.</div>';}
  else hEl.innerHTML=confHj.map(function(c){return renderConfCard(c);}).join('');
  // Histórico com filtros
  var fStatus=document.getElementById('conf-filtro-status')?document.getElementById('conf-filtro-status').value:'';
  var fData=document.getElementById('conf-filtro-data')?document.getElementById('conf-filtro-data').value:'';
  var filtered=[].concat(S.conferencias).filter(function(c){
    return(!fStatus||c.status===fStatus)&&(!fData||c.data===fData);
  }).sort(function(a,b){return b.criadoEm>a.criadoEm?1:-1;});
  var histEl=document.getElementById('conf-historico');
  if(!filtered.length){histEl.innerHTML='<div class="empty-state">Nenhuma conferência encontrada.</div>';return;}
  histEl.innerHTML=filtered.map(function(c){return renderConfCard(c);}).join('');
}

function renderConfCard(c){
  var statusLabel={conforme:'Conforme',divergencia:'Divergência',investigando:'Investigando',resolvida:'Resolvida'};
  var statusBadge={conforme:'badge-green',divergencia:'badge-red',investigando:'badge-amber',resolvida:'badge-blue'};
  var difStr=(c.diferenca>0?'+':'')+fmtKg(c.diferenca);
  var btns='<button class="btn btn-sm" data-id="'+c.id+'" onclick="abrirModalConferencia(this.dataset.id)" style="font-size:11px">Editar</button>';
  if(c.status==='divergencia')btns+='<button class="btn btn-sm btn-warning" data-id="'+c.id+'" onclick="abrirInvestigacao(this.dataset.id)" style="font-size:11px">Investigar</button>';
  if(c.status==='investigando')btns+='<button class="btn btn-sm btn-success" data-id="'+c.id+'" onclick="abrirInvestigacao(this.dataset.id)" style="font-size:11px">Resolver</button>';
  btns+='<button class="btn btn-sm btn-danger" data-id="'+c.id+'" onclick="excluirConferencia(this.dataset.id)" style="font-size:11px">✕</button>';
  return '<div class="conf-card '+c.status+'">'
    +'<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px">'
      +'<div style="flex:1;min-width:0">'
        +'<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:4px">'
          +'<span style="font-size:13px;font-weight:600">'+(c.produto||c.baiaNome)+'</span>'
          +'<span class="badge '+(statusBadge[c.status]||'badge-gray')+'">'+(statusLabel[c.status]||c.status)+'</span>'
          +(c.resultado==='divergencia'?'<span class="badge badge-red" style="font-size:10px">Dif: '+difStr+'</span>':'')
        +'</div>'
        +'<div style="font-size:12px;color:var(--text2);margin-bottom:2px">'
          +fmtDate(c.data)
          +' &nbsp;·&nbsp; Sistema: '+fmtKg(c.qtdSistema)
          +' &nbsp;·&nbsp; Conferido: '+fmtKg(c.qtdFisica)
          +(c.baiasInfo?' &nbsp;·&nbsp; Baias: '+c.baiasInfo:'')+(c.operador?' &nbsp;·&nbsp; '+c.operador:'')
        +'</div>'
        +(c.obs?'<div style="font-size:11px;color:var(--text3)">'+c.obs+'</div>':'')
        +(c.obsInv?'<div style="font-size:11px;color:var(--text2);margin-top:3px;border-top:1px solid var(--border);padding-top:3px"><strong>Investigação:</strong> '+c.obsInv+(c.resolvidaEm?' (Resolvido em '+fmtDate(c.resolvidaEm.slice(0,10))+')':'')+'</div>':'')
      +'</div>'
      +'<div style="display:flex;flex-direction:column;gap:4px;flex-shrink:0">'+btns+'</div>'
    +'</div>'
    +'</div>';
}


function showSubTab(moduleId, tabKey){
  // Find all subsec divs for this module
  var prefix='subsec-'+moduleId+'-';
  document.querySelectorAll('[id^="'+prefix+'"]').forEach(function(el){
    el.style.display='none';
  });
  var target=document.getElementById(prefix+tabKey);
  if(target)target.style.display='';
  // Update subtab button states
  var btnPrefix='subtab-'+moduleId+'-';
  document.querySelectorAll('[id^="'+btnPrefix+'"]').forEach(function(btn){
    btn.classList.remove('subtab-active');
    btn.style.borderBottomColor='transparent';
    btn.style.color='var(--text2)';
    btn.style.fontWeight='400';
  });
  var activeBtn=document.getElementById(btnPrefix+tabKey);
  if(activeBtn){
    activeBtn.classList.add('subtab-active');
    activeBtn.style.borderBottomColor='var(--blue)';
    activeBtn.style.color='var(--blue)';
    activeBtn.style.fontWeight='600';
  }
  // trigger renders for sections that need it
  var renderMap={
    'recebimento': renderRecebimento,
    'movimentacoes': renderMov,
    'ocorrencias': renderOcorrencias,
    'conferencia': renderConferencia,
    'estoque': renderEstoque,
    'fornecedores': renderFornecedores,
    'baixa': function(){renderBaixa();},
    'materiais': renderMateriais,
    'transferencias': renderTransferencias,
    'depositos': renderDepositos,
    'agendamento': renderAgenda,
    'pcp': renderPCP,
    'consumo': renderKPIConsumo,
  };
  if(renderMap[tabKey])renderMap[tabKey]();
}

function updateModuleBadges(){
  // Log. Interna badge: rec pending + transf pending + ocorr open
  var recPend=S.descargas.filter(function(d){return d.status==='pendente'||d.status==='chegou';}).length;
  var transfPend=S.transferencias.filter(function(t){return !t.recebido;}).length;
  var ocorrOpen=S.ocorrencias.filter(function(o){return o.status!=='resolvida';}).length;
  var logTotal=recPend+transfPend+ocorrOpen;
  var bLog=document.getElementById('badge-loginterna');
  if(bLog){bLog.style.display=logTotal>0?'inline-flex':'none';bLog.textContent=logTotal;}
  // PCP badge: conferencias com divergencia
  var confPend=(S.conferencias||[]).filter(function(c){return c.status==='divergencia'||c.status==='investigando';}).length;
  var bPcp=document.getElementById('badge-pcp');
  if(bPcp){bPcp.style.display=confPend>0?'inline-flex':'none';bPcp.textContent=confPend;}
}


function abrirEditarBaia(id){
  var b=S.baias.find(function(x){return x.id===id;});if(!b)return;
  document.getElementById('modal-editar-baia')._baiaId=id;
  document.getElementById('edit-baia-nome').value=b.nome||'';
  document.getElementById('edit-baia-cap').value=b.cap||'';
  document.getElementById('edit-baia-mat').value=b.material||'';
  document.getElementById('edit-baia-loc').value=b.loc||'';
  // populate depositos
  var sd=document.getElementById('edit-baia-dep');
  sd.innerHTML='';
  S.depositos.forEach(function(d){sd.innerHTML+='<option value="'+d.id+'"'+(d.id===b.dep?' selected':'')+'>'+d.nome+'</option>';});
  // status buttons
  setBaiaStatusUI(b.status||'ativa');
  document.getElementById('edit-baia-status-alert').innerHTML='';
  // warn if has stock
  if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
  if(b.estoques&&b.estoques.length>0){
    document.getElementById('edit-baia-status-alert').innerHTML='<div class="alert-bar alert-warning" style="margin-bottom:10px">Esta baia possui estoque alocado. Ao inutilizar ou deletar, o estoque não será removido automaticamente.</div>';
  }
  closeModal('modal-baia-det');
  openModal('modal-editar-baia');
}

var _editBaiaStatus='ativa';

function setBaiaStatus(status){
  _editBaiaStatus=status;
  setBaiaStatusUI(status);
}

function setBaiaStatusUI(status){
  _editBaiaStatus=status;
  var btns={ativa:'btn-baia-ativa',inutilizada:'btn-baia-inutilizada',manutencao:'btn-baia-manutencao'};
  var desc={
    ativa:'Baia disponível para receber materiais.',
    inutilizada:'Baia fora de uso. Aparece destacada em cinza e não aparece como sugestão no agendamento.',
    manutencao:'Baia em manutenção. Aparece em azul e bloqueada para novos lançamentos.'
  };
  Object.keys(btns).forEach(function(s){
    var btn=document.getElementById(btns[s]);
    if(btn)btn.style.fontWeight=s===status?'700':'400';
    if(btn)btn.style.opacity=s===status?'1':'0.5';
  });
  var descEl=document.getElementById('edit-baia-status-desc');
  if(descEl)descEl.textContent=desc[status]||'';
}

function salvarEdicaoBaia(){
  var id=document.getElementById('modal-editar-baia')._baiaId;
  var b=S.baias.find(function(x){return x.id===id;});if(!b)return;
  var novoNome=document.getElementById('edit-baia-nome').value.trim();
  if(!novoNome){alert('Informe o nome da baia.');return;}
  b.nome=novoNome;
  b.dep=document.getElementById('edit-baia-dep').value||b.dep;
  b.cap=parseKg(document.getElementById('edit-baia-cap').value)||0;
  b.material=document.getElementById('edit-baia-mat').value.trim();
  b.loc=document.getElementById('edit-baia-loc').value.trim();
  b.status=_editBaiaStatus==='ativa'?null:_editBaiaStatus;
  saveState();closeModal('modal-editar-baia');render();
}

function excluirBaiaConfirm(){
  var id=document.getElementById('modal-editar-baia')._baiaId;
  var b=S.baias.find(function(x){return x.id===id;});if(!b)return;
  if(!b.estoques){b.estoques=b.estoque?[b.estoque]:[];delete b.estoque;}
  var temEstoque=b.estoques&&b.estoques.length>0;
  var msg=temEstoque
    ?'ATENÇÃO: A baia "'+b.nome+'" possui estoque alocado!\n\nTem certeza que deseja DELETAR permanentemente?\nO estoque vinculado também será removido.'
    :'Deletar permanentemente a baia "'+b.nome+'"?\n\nEsta ação não pode ser desfeita.';
  if(!confirm(msg))return;
  S.baias=S.baias.filter(function(x){return x.id!==id;});
  saveState();closeModal('modal-editar-baia');render();
}


var _kpiConsumoChart1=null, _kpiConsumoChart2=null;

function renderKPIConsumo(){
  var pcpEl=document.getElementById('page-pcp');
  if(!pcpEl||!pcpEl.classList.contains('active'))return;

  // Populate material filter
  var matSel=document.getElementById('kpi-consumo-mat');
  if(matSel){
    var cur=matSel.value;
    matSel.innerHTML='<option value="">Todos os materiais</option>';
    var mats={};
    S.baixas.forEach(function(b){if(b.material)mats[b.material]=true;});
    S.materiais.forEach(function(m){mats[m.nome]=true;});
    Object.keys(mats).sort().forEach(function(m){matSel.innerHTML+='<option value="'+m+'"'+(m===cur?' selected':'')+'>'+m+'</option>';});
    if(cur)matSel.value=cur;
  }

  var periodo=document.getElementById('kpi-consumo-periodo')?document.getElementById('kpi-consumo-periodo').value:'mes';
  var filtMat=matSel?matSel.value:'';

  // Date range
  var hoje2=hoje();
  var dataInicio;
  if(periodo==='semana'){
    var d=new Date(hoje2+'T00:00:00');
    d.setDate(d.getDate()-d.getDay()+1);
    dataInicio=d.toISOString().slice(0,10);
  } else if(periodo==='mes'){
    dataInicio=hoje2.slice(0,7)+'-01';
  } else if(periodo==='ano'){
    dataInicio=hoje2.slice(0,4)+'-01-01';
  } else {
    dataInicio='2000-01-01';
  }

  var baixasFiltro=S.baixas.filter(function(b){
    return b.data>=dataInicio&&(!filtMat||b.material===filtMat);
  });

  // Summary cards
  var totalKg=baixasFiltro.reduce(function(s,b){return s+(b.qtd||0);},0);
  var totalBaixas=baixasFiltro.length;
  var maisConsumido='—',maisQtd=0;
  var matQtd={};
  baixasFiltro.forEach(function(b){
    var m=b.material||'Desconhecido';
    if(!matQtd[m])matQtd[m]=0;
    matQtd[m]+=(b.qtd||0);
    if(matQtd[m]>maisQtd){maisQtd=matQtd[m];maisConsumido=m;}
  });
  var periodoLabel={semana:'Semana',mes:'Mês',ano:'Ano',tudo:'Total'}[periodo]||periodo;
  var estAtual=S.baias.reduce(function(s,b){return s+(b.estoques||[]).reduce(function(ss,e){return ss+e.qtdAtual;},0);},0);

  document.getElementById('kpi-consumo-cards').innerHTML=
    '<div class="metric-card"><div class="metric-label">Consumido no período</div><div class="metric-value" style="font-size:20px">'+fmtKg(totalKg)+'</div><div class="metric-sub">'+periodoLabel+': '+dataInicio+' a '+hoje2+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Nº de baixas</div><div class="metric-value" style="font-size:20px">'+totalBaixas+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Mais consumido</div><div class="metric-value" style="font-size:14px;font-weight:600">'+maisConsumido+'</div><div class="metric-sub">'+fmtKg(maisQtd)+'</div></div>'
    +'<div class="metric-card"><div class="metric-label">Estoque atual total</div><div class="metric-value" style="font-size:20px">'+fmtKg(estAtual)+'</div></div>';

  // Chart 1: consumo por semana
  var semConsumo={};
  baixasFiltro.forEach(function(b){
    var k=semanaKey(b.data);
    if(!semConsumo[k])semConsumo[k]=0;
    semConsumo[k]+=(b.qtd||0);
  });
  var semKeys=Object.keys(semConsumo).sort();
  if(_kpiConsumoChart1){_kpiConsumoChart1.destroy();_kpiConsumoChart1=null;}
  var ctx1=document.getElementById('chart-consumo-semana');
  if(ctx1&&semKeys.length){
    _kpiConsumoChart1=new Chart(ctx1,{type:'bar',
      data:{labels:semKeys.map(function(k){return fmtDate(k);}),
        datasets:[{label:'Consumo (ton)',
          data:semKeys.map(function(k){return Math.round(semConsumo[k]*1000)/1000;}),
          backgroundColor:'#E24B4A66',borderColor:'#E24B4A',borderWidth:1.5,borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{display:false}},
        scales:{y:{beginAtZero:true,grid:{color:'#00000011'}}}}});
  } else if(ctx1){
    ctx1.getContext('2d').clearRect(0,0,ctx1.width,ctx1.height);
  }

  // Chart 2: consumo por material (donut ou barra)
  var matKeys2=Object.keys(matQtd).sort(function(a,b){return matQtd[b]-matQtd[a];}).slice(0,8);
  if(_kpiConsumoChart2){_kpiConsumoChart2.destroy();_kpiConsumoChart2=null;}
  var ctx2=document.getElementById('chart-consumo-mat');
  var colors=['#E24B4A','#378ADD','#EF9F27','#1D9E75','#534AB7','#D85A30','#639922','#185FA5'];
  if(ctx2&&matKeys2.length){
    _kpiConsumoChart2=new Chart(ctx2,{type:'doughnut',
      data:{labels:matKeys2.map(function(k){return k.length>20?k.slice(0,20)+'…':k;}),
        datasets:[{data:matKeys2.map(function(k){return Math.round(matQtd[k]*1000)/1000;}),
          backgroundColor:colors.slice(0,matKeys2.length),borderWidth:2,borderColor:'#fff'}]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{position:'right',labels:{font:{size:10},boxWidth:12}}}}});
  }

  // Tabela por material
  var tabEl=document.getElementById('kpi-consumo-tabela');
  var allMats=Object.keys(matQtd).sort(function(a,b){return matQtd[b]-matQtd[a];});
  if(!allMats.length){tabEl.innerHTML='<div class="empty-state">Nenhuma baixa no período.</div>';}
  else tabEl.innerHTML='<table><thead><tr><th>Material</th><th>Consumido</th><th>Nº Baixas</th><th>Estoque atual</th><th>Cobertura estimada</th></tr></thead><tbody>'
    +allMats.map(function(m){
      var consumido=matQtd[m];
      var nBaixas=baixasFiltro.filter(function(b){return b.material===m;}).length;
      // estoque atual deste material
      var estMat=0;
      S.baias.forEach(function(b){(b.estoques||[]).forEach(function(e){if(e.fornecedorNome===m||e.material===m)estMat+=e.qtdAtual;});});
      // cobertura: diasNoPeriodo / consumido * estoque
      var diasPeriodo=Math.max(1,Math.round((new Date(hoje2)-new Date(dataInicio))/(1000*60*60*24)));
      var consumoDiario=consumido/diasPeriodo;
      var coberturaDias=consumoDiario>0?Math.round(estMat/consumoDiario):null;
      return '<tr>'
        +'<td style="font-weight:600">'+m+'</td>'
        +'<td>'+fmtKg(consumido)+'</td>'
        +'<td>'+nBaixas+'</td>'
        +'<td style="color:'+(estMat<consumido?'#A32D2D':'#0F6E56')+'">'+fmtKg(estMat)+'</td>'
        +'<td>'+(coberturaDias!==null?'<span class="badge '+(coberturaDias<7?'badge-red':coberturaDias<30?'badge-amber':'badge-green')+'">~'+coberturaDias+' dias</span>':'—')+'</td>'
        +'</tr>';
    }).join('')+'</tbody></table>';

  // Historico de baixas
  var histEl=document.getElementById('kpi-consumo-historico');
  var sorted=[].concat(baixasFiltro).sort(function(a,b){return b.data>a.data?1:-1;});
  if(!sorted.length){histEl.innerHTML='<div class="empty-state">Nenhuma baixa no período.</div>';}
  else histEl.innerHTML='<div style="overflow-x:auto"><table><thead><tr><th>Data</th><th>Material</th><th>Baia</th><th>Lote</th><th>NF</th><th>Quantidade</th><th>Tipo</th><th>Operador</th></tr></thead><tbody>'
    +sorted.map(function(b){return'<tr>'
      +'<td style="white-space:nowrap">'+fmtDate(b.data)+'</td>'
      +'<td style="font-weight:600">'+(b.material||'—')+'</td>'
      +'<td>'+(b.baiaName||'—')+'</td>'
      +'<td><span class="tag">'+(b.lote||'—')+'</span></td>'
      +'<td><span class="tag">'+(b.nf||'—')+'</span></td>'
      +'<td style="font-weight:600">'+b.qtd+'t</td>'
      +'<td><span class="badge '+(b.tipo==='total'?'badge-red':'badge-amber')+'">'+b.tipo+'</span></td>'
      +'<td style="font-size:12px;color:var(--text2)">'+(b.operador||'—')+'</td>'
      +'</tr>';}).join('')+'</tbody></table></div>';
}

// ─── BOOT ─────────────────────────────────────────────────────────────────────
async function _boot(user){
  // Mostra email do usuário logado no nav
  var el=document.getElementById('nav-user-email');
  if(el)el.textContent=user.email;

  // Carrega estado do Firestore
  S=await loadState();
  // Garantir campos novos em dados antigos
  if(!S.capDiaria)S.capDiaria=160;
  if(!S.toleranciaMin)S.toleranciaMin=30;
  if(!S.movimentacoes)S.movimentacoes=[];
  if(!S.materiais)S.materiais=[];
  if(S.emailAlerta===undefined)S.emailAlerta="";
  if(!S.ocorrencias)S.ocorrencias=[];
  if(!S.acertosPeso)S.acertosPeso=[];
  if(!S.conferencias)S.conferencias=[];
  if(S.horarioAlerta===undefined)S.horarioAlerta="07:00";
  if(S.alertaAtivo===undefined)S.alertaAtivo=false;
  if(!S.transferencias)S.transferencias=[];
  if(!S.turnos||!S.turnos.length)S.turnos=[{id:'t1',nome:'Turno A',inicio:'06:00',fim:'14:00',cor:'blue'},{id:'t2',nome:'Turno B',inicio:'14:00',fim:'22:00',cor:'green'},{id:'t3',nome:'Turno C',inicio:'22:00',fim:'06:00',cor:'amber'}];
  if(!S.recebimentos)S.recebimentos=[];
  if(!S.fornecedores)S.fornecedores=[];
  if(!S.operadores)S.operadores=[];

  // Escuta mudanças em tempo real (outros usuários)
  onSnapshot(_DOC(),function(snap){
    if(snap.exists()){S=snap.data();render();}
  });

  setInterval(function(){if(S.alertaAtivo)verificarHorarioEmail();},60000);
  autoAvancarDescargas();
  render();

  // Oculta tela de login, mostra o app
  document.getElementById('login-screen').style.display='none';
  document.getElementById('app-wrapper').style.display='block';

  // Auto-formatação de campos de quantidade (kg)
  document.addEventListener('input', function(e){
    if(e.target && e.target.classList.contains('input-kg')){
      var raw = e.target.value.replace(/\D/g,'');
      if(raw===''){e.target.value='';return;}
      var num = parseInt(raw,10);
      var pos = e.target.selectionStart;
      var oldLen = e.target.value.length;
      e.target.value = num.toLocaleString('pt-BR');
      var newLen = e.target.value.length;
      e.target.selectionStart = e.target.selectionEnd = pos + (newLen - oldLen);
    }
  });

}

// ─── EXPOR FUNÇÕES AO HTML (necessário por causa do type="module") ────────────
window.showPage                  = showPage;
window.showSubTab                = showSubTab;
window.openModal                 = openModal;
window.closeModal                = closeModal;
window.agendarDescarga           = agendarDescarga;
window.confirmarChegada          = confirmarChegada;
window.confirmarAlocacao         = confirmarAlocacao;
window.registrarBaixa            = registrarBaixa;
window.salvarFornecedor          = salvarFornecedor;
window.salvarDeposito            = salvarDeposito;
window.salvarBaia                = salvarBaia;
window.salvarMaterial            = salvarMaterial;
window.salvarTransferencia       = salvarTransferencia;
window.salvarOcorrencia          = salvarOcorrencia;
window.salvarAcertoPeso          = salvarAcertoPeso;
window.salvarConferencia         = salvarConferencia;
window.salvarInvestigacao        = salvarInvestigacao;
window.salvarEdicaoBaia          = salvarEdicaoBaia;
window.salvarEdicaoTransf        = salvarEdicaoTransf;
window.salvarCap                 = salvarCap;
window.salvarEmailConfig         = salvarEmailConfig;
window.salvarTurno               = salvarTurno;
window.addOperador               = addOperador;
window.addTurno                  = addTurno;
window.setBaiaStatus             = setBaiaStatus;
window.abrirEditarBaia           = abrirEditarBaia;
window.excluirBaiaConfirm        = excluirBaiaConfirm;
window.confirmarRecebimentoTransf= confirmarRecebimentoTransf;
window.lancarEstoqueManual       = lancarEstoqueManual;
window.openEntradaManual         = openEntradaManual;
window.exportarDados             = exportarDados;
window.importarDados             = importarDados;
window.resetarDados              = resetarDados;
window.gerarRelatorio            = gerarRelatorio;
window.enviarAlertaEmail         = enviarAlertaEmail;
window.verPreviewEmail           = verPreviewEmail;
window.toggleEstoqueView         = toggleEstoqueView;
window.renderKPIs                = renderKPIs;
window.renderKPIConsumo          = renderKPIConsumo;
window.renderEstoque             = renderEstoque;
window.renderMov                 = renderMov;
window.renderOcorrencias         = renderOcorrencias;
window.renderPCP                 = renderPCP;
window.renderConferencia         = renderConferencia;
window.calcAcerto                = calcAcerto;
window.calcConferencia           = calcConferencia;
window.checkCapModal             = checkCapModal;
window.checkDiv                  = checkDiv;
window.converter                 = converter;
window.preencherConfProduto      = preencherConfProduto;
window.preencherForn             = preencherForn;
window.preencherFornTransf       = preencherFornTransf;
window.syncApProduto             = syncApProduto;
window.syncEmForn                = syncEmForn;
window.syncEmMat                 = syncEmMat;
window.syncMatSel                = syncMatSel;
window.syncOcForn                = syncOcForn;
window.toggleBaixaQtd            = toggleBaixaQtd;
window.toggleInvAjuste           = toggleInvAjuste;
window.updateAlocarInfo          = updateAlocarInfo;
window.verificarDataTransf       = verificarDataTransf;
window.selectCalDay              = selectCalDay;
window.clearCalSel               = clearCalSel;
window.calNav                    = calNav;
window.showCalDay                = showCalDay;
window.abrirAlocacao             = abrirAlocacao;
window.abrirBaiaDet              = abrirBaiaDet;
window.abrirChegada              = abrirChegada;
window.abrirConfirmarRecebTransf = abrirConfirmarRecebTransf;
window.abrirEditarTransf         = abrirEditarTransf;
window.abrirInvestigacao         = abrirInvestigacao;
window.abrirModalAcerto          = abrirModalAcerto;
window.abrirModalConferencia     = abrirModalConferencia;
window.abrirModalOcorrencia      = abrirModalOcorrencia;
window.alterarStatusOcorr        = alterarStatusOcorr;
window.desfazerRecebimentoTransf = desfazerRecebimentoTransf;
window.excluirAcerto             = excluirAcerto;
window.excluirConferencia        = excluirConferencia;
window.excluirDescarga           = excluirDescarga;
window.excluirFornecedor         = excluirFornecedor;
window.excluirMaterial           = excluirMaterial;
window.excluirOcorrencia         = excluirOcorrencia;
window.excluirTransferencia      = excluirTransferencia;
window.removerEstoqueBaia        = removerEstoqueBaia;
window.removerOperador           = removerOperador;
window.removerTurno              = removerTurno;
window.toggleAnalisado           = toggleAnalisado;
window.toggleProdutoGroup        = toggleProdutoGroup;
window.verDetalheOcorr           = verDetalheOcorr;
window.verFornecedor             = verFornecedor;
window.verRecDet                 = verRecDet;
// ─────────────────────────────────────────────────────────────────────────────

// ─── INICIALIZAÇÃO ────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', function(){

  // Escuta estado de autenticação — NADA carrega antes disso
  onAuthStateChanged(_auth, function(user){
    if(user){
      _boot(user);
    } else {
      document.getElementById('login-screen').style.display='flex';
      document.getElementById('app-wrapper').style.display='none';
    }
  });

  // Botão login
  document.getElementById('btn-login').addEventListener('click', async function(){
    var email=document.getElementById('login-email').value.trim();
    var senha=document.getElementById('login-senha').value;
    var errEl=document.getElementById('login-erro');
    var btn=this;
    errEl.textContent='';
    btn.disabled=true; btn.textContent='Entrando...';
    try{
      await signInWithEmailAndPassword(_auth, email, senha);
    }catch(e){
      var msgs={
        'auth/invalid-credential':'E-mail ou senha incorretos.',
        'auth/user-not-found':'E-mail ou senha incorretos.',
        'auth/wrong-password':'E-mail ou senha incorretos.',
        'auth/too-many-requests':'Muitas tentativas. Aguarde alguns minutos.',
        'auth/invalid-email':'E-mail inválido.'
      };
      errEl.textContent=msgs[e.code]||'E-mail ou senha incorretos.';
      btn.disabled=false; btn.textContent='Entrar';
    }
  });

  // Enter no campo senha dispara login
  document.getElementById('login-senha').addEventListener('keydown', function(e){
    if(e.key==='Enter') document.getElementById('btn-login').click();
  });

  // Botão logout
  document.getElementById('btn-logout').addEventListener('click', function(){
    signOut(_auth);
  });

});
// ─────────────────────────────────────────────────────────────────────────────
