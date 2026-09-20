const SHEET_ID='1QpyTzbc2P0F-LrGw7fIe7m0MOE10rbVSFAyCiWrgCNI';
const ABAS={ATIVIDADES:'ATIVIDADES',ENTREGAS:'ENTREGAS',CORRECOES:'CORRECOES',MATERIAIS:'MATERIAIS',CONFIG:'CONFIG'};
const TZ='America/Sao_Paulo';

function doGet(e){
  const action=String((e&&e.parameter&&e.parameter.action)||'listar').trim();
  try{
    if(action==='listar') return json_(listar_(e.parameter.turma,e.parameter.componente));
    if(action==='atividade'){
      const atividade=atividade_(e.parameter.id);
      if(atividade&&!liberada_(atividade))throw new Error('Atividade ainda não liberada. Liberação em '+dataHoraBr_(atividade.liberacao)+'.');
      return json_({ok:true,atividade:atividade});
    }
    if(action==='ping') return json_({ok:true,sistema:'PE',versao:'1.3-data-hora',agora:agoraLocal_()});
    return json_({ok:false,erro:'Ação inválida'});
  }catch(err){return json_({ok:false,erro:String(err.message||err)})}
}

function doPost(e){
  try{
    const d=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    if(d.action==='salvarAtividade') return json_(salvarAtividade_(d));
    if(d.action==='salvarMaterial') return json_(salvarMaterial_(d));
    if(d.action==='enviarAtividade') return json_(enviarAtividade_(d));
    return json_({ok:false,erro:'Ação inválida'});
  }catch(err){return json_({ok:false,erro:String(err.message||err)})}
}

function ss_(){return SpreadsheetApp.openById(SHEET_ID)}
function sh_(nome){const sh=ss_().getSheetByName(nome);if(!sh)throw new Error('Aba não encontrada: '+nome);return sh}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}
function ensureHeader_(sh,nome){const lastCol=Math.max(1,sh.getLastColumn());const h=sh.getRange(1,1,1,lastCol).getValues()[0].map(String);if(!h.includes(nome))sh.getRange(1,lastCol+1).setValue(nome)}
function agoraLocal_(){return Utilities.formatDate(new Date(),TZ,"yyyy-MM-dd'T'HH:mm")}
function liberada_(a){return !a||!a.liberacao||agoraLocal_()>=normalizaDataHora_(a.liberacao,'00:00')}
function encerrada_(a){return !!(a&&a.prazo&&agoraLocal_()>normalizaDataHora_(a.prazo,'23:59'))}
function normalizaDataHora_(v,horaPadrao){if(!v)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,TZ,"yyyy-MM-dd'T'HH:mm");let s=String(v).trim();if(/^\d{4}-\d{2}-\d{2}$/.test(s))s+='T'+horaPadrao;return s.slice(0,16)}
function dataHoraBr_(v){if(!v)return'';const s=normalizaDataHora_(v,'00:00'),m=s.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);return m?`${m[3]}/${m[2]}/${m[1]} às ${m[4]}:${m[5]}`:String(v)}

function listar_(turma,componente){
  const atividades=rows_(ABAS.ATIVIDADES)
    .filter(x=>(!turma||eq_(x.TURMA,turma))&&(!componente||eq_(x.COMPONENTE,componente)))
    .sort((a,b)=>Number(a.ORDEM||999)-Number(b.ORDEM||999))
    .map(mapAtividade_)
    .filter(liberada_);
  const materiais=rows_(ABAS.MATERIAIS)
    .filter(x=>(!componente||eq_(x.COMPONENTE,componente)))
    .sort((a,b)=>Number(a.ORDEM||999)-Number(b.ORDEM||999))
    .map(m=>({id:m.ID_MATERIAL,titulo:m.TITULO,descricao:m.DESCRICAO,url:m.ARQUIVO_URL,tipo:m.TIPO,status:m.STATUS,publicadoEm:date_(m.PUBLICADO_EM)}));
  return {ok:true,atividades,materiais};
}

function atividade_(id){
  if(!id)return null;
  const a=rows_(ABAS.ATIVIDADES).find(x=>eq_(x.ID_ATIVIDADE,id));
  return a?mapAtividade_(a):null;
}

function mapAtividade_(a){
  const tipoParticipacao=String(a.TIPO_PARTICIPACAO||'INDIVIDUAL').toUpperCase()==='GRUPO'?'GRUPO':'INDIVIDUAL';
  return {
    id:a.ID_ATIVIDADE,turma:a.TURMA,componente:a.COMPONENTE,titulo:a.TITULO,descricao:a.DESCRICAO,
    orientacoes:a.ORIENTACOES||'',tipoEnvio:a.TIPO_ENVIO||'SEM_ENVIO',extensoes:a.EXTENSOES||'',
    maxArquivos:Number(a.MAX_ARQUIVOS||0),liberacao:normalizaDataHora_(a.LIBERACAO,'00:00'),prazo:normalizaDataHora_(a.PRAZO,'23:59'),materialUrl:a.MATERIAL_APOIO_URL||'',
    correcaoIA:a.CORRECAO_IA||'NAO',criterios:a.GABARITO_CRITERIOS||'',status:a.STATUS||'RASCUNHO',
    ordem:Number(a.ORDEM||999),tipoParticipacao:tipoParticipacao,
    maxAlunosGrupo:tipoParticipacao==='GRUPO'?Math.max(1,Number(a.MAX_ALUNOS_GRUPO||2)):1
  };
}

function salvarAtividade_(d){
  if(!d.id||!d.titulo)throw new Error('ID e título são obrigatórios');
  const liberacao=normalizaDataHora_(d.liberacao,'00:00');
  const prazo=normalizaDataHora_(d.prazo,'23:59');
  if(liberacao&&prazo&&liberacao>prazo)throw new Error('A liberação não pode ser posterior à finalização');
  const sh=sh_(ABAS.ATIVIDADES);
  ensureHeader_(sh,'LIBERACAO');
  const dados=sheetData_(sh),now=new Date();
  const atual=dados.rows.find(x=>eq_(x.ID_ATIVIDADE,d.id));
  const tipoParticipacao=String(d.tipoParticipacao||'INDIVIDUAL').toUpperCase()==='GRUPO'?'GRUPO':'INDIVIDUAL';
  const obj={
    ID_ATIVIDADE:String(d.id).trim(),TURMA:d.turma||'1º MTEC - Administração - Mairinque',COMPONENTE:d.componente||'PE',
    TITULO:d.titulo||'',DESCRICAO:d.descricao||'',TIPO_ENVIO:d.tipoEnvio||'SEM_ENVIO',
    EXTENSOES:d.extensoes||extensoesPadrao_(d.tipoEnvio),MAX_ARQUIVOS:Number(d.maxArquivos||0),LIBERACAO:liberacao,PRAZO:prazo,
    MATERIAL_APOIO_URL:d.materialUrl||'',CORRECAO_IA:d.correcaoIA||'NAO',GABARITO_CRITERIOS:d.criterios||'',
    STATUS:d.status||'RASCUNHO',ORDEM:Number(d.ordem||999),CRIADO_EM:atual&&atual.CRIADO_EM?atual.CRIADO_EM:now,
    ATUALIZADO_EM:now,ORIENTACOES:d.orientacoes||'',TIPO_PARTICIPACAO:tipoParticipacao,
    MAX_ALUNOS_GRUPO:tipoParticipacao==='GRUPO'?Math.max(1,Number(d.maxAlunosGrupo||2)):1
  };
  upsert_(sh,dados,'ID_ATIVIDADE',obj.ID_ATIVIDADE,obj);
  return {ok:true,id:obj.ID_ATIVIDADE};
}

function salvarMaterial_(d){
  if(!d.titulo||!d.url)throw new Error('Título e URL são obrigatórios');
  const sh=sh_(ABAS.MATERIAIS),dados=sheetData_(sh),now=new Date();
  const id=d.id||('MAT-'+Utilities.getUuid().slice(0,8).toUpperCase());
  const obj={ID_MATERIAL:id,COMPONENTE:d.componente||'PE',TITULO:d.titulo,DESCRICAO:d.descricao||'',ARQUIVO_URL:d.url,TIPO:d.tipo||'MATERIAL',ORDEM:Number(d.ordem||999),STATUS:d.status||'PUBLICADO',PUBLICADO_EM:now};
  upsert_(sh,dados,'ID_MATERIAL',id,obj);
  return {ok:true,id};
}

function enviarAtividade_(d){
  if(!d.idAtividade)throw new Error('Atividade não informada');
  const atividade=atividade_(d.idAtividade);
  if(!atividade)throw new Error('Atividade não encontrada');
  if(atividade.status!=='PUBLICADA')throw new Error('Atividade não está aberta para envio');
  if(!liberada_(atividade))throw new Error('Atividade ainda não liberada. Liberação em '+dataHoraBr_(atividade.liberacao)+'.');
  if(encerrada_(atividade))throw new Error('Prazo de entrega encerrado em '+dataHoraBr_(atividade.prazo)+'.');

  const emGrupo=atividade.tipoParticipacao==='GRUPO';
  let integrantes=Array.isArray(d.integrantes)?d.integrantes:[];
  integrantes=integrantes.map(x=>({nome:String((x&&x.nome)||'').trim(),email:String((x&&x.email)||'').trim()})).filter(x=>x.nome);
  if(!integrantes.length&&String(d.aluno||'').trim())integrantes=[{nome:String(d.aluno).trim(),email:String(d.email||'').trim()}];
  if(!integrantes.length)throw new Error('Informe ao menos um aluno');
  if(!emGrupo)integrantes=[integrantes[0]];
  if(emGrupo&&integrantes.length>atividade.maxAlunosGrupo)throw new Error('Este trabalho permite no máximo '+atividade.maxAlunosGrupo+' aluno(s) por grupo');
  const nomesUnicos=new Set(integrantes.map(x=>x.nome.toLowerCase()));
  if(nomesUnicos.size!==integrantes.length)throw new Error('Há nomes de alunos repetidos no grupo');

  const arquivos=Array.isArray(d.arquivos)?d.arquivos:[];
  if(atividade.tipoEnvio!=='FORMULARIO'&&atividade.tipoEnvio!=='TEXTO'&&atividade.tipoEnvio!=='SEM_ENVIO'){
    if(!arquivos.length)throw new Error('Selecione ao menos um arquivo');
    if(atividade.maxArquivos&&arquivos.length>atividade.maxArquivos)throw new Error('Máximo de '+atividade.maxArquivos+' arquivo(s)');
  }
  if((atividade.tipoEnvio==='FORMULARIO'||atividade.tipoEnvio==='TEXTO')&&!String(d.respostaTexto||'').trim())throw new Error('Digite a resposta da atividade');
  validaArquivos_(arquivos,atividade.extensoes);

  const cfg=config_();
  if(!cfg.PASTA_ENTREGAS_ID)throw new Error('PASTA_ENTREGAS_ID não configurada');
  const raiz=DriveApp.getFolderById(cfg.PASTA_ENTREGAS_ID);
  const pastaAtividade=getOrCreateFolder_(raiz,sanitize_(atividade.id+' - '+atividade.titulo));
  const idGrupo=emGrupo?'GRP-'+Utilities.getUuid().slice(0,10).toUpperCase():'';
  const nomePasta=emGrupo?(idGrupo+' - '+integrantes.map(x=>x.nome).join(', ')):integrantes[0].nome;
  const pastaEntrega=getOrCreateFolder_(pastaAtividade,sanitize_(nomePasta));

  const urls=[];
  arquivos.forEach((f,i)=>{
    const raw=String(f.data||'').replace(/^data:[^;]+;base64/,'').replace(/^,/,'');
    const bytes=Utilities.base64Decode(raw);
    const nome=sanitizeFile_(f.name||('arquivo-'+(i+1)));
    const blob=Utilities.newBlob(bytes,f.mime||MimeType.PLAIN_TEXT,nome);
    urls.push(pastaEntrega.createFile(blob).getUrl());
  });

  const sh=sh_(ABAS.ENTREGAS),dados=sheetData_(sh);
  const idEntrega='ENT-'+Utilities.getUuid().slice(0,12).toUpperCase();
  const dataEnvio=new Date();
  integrantes.forEach(integrante=>{
    const obj={
      ID_ENTREGA:idEntrega,ID_ATIVIDADE:atividade.id,ALUNO:integrante.nome,EMAIL:integrante.email,
      TURMA:atividade.turma||d.turma||'',ARQUIVOS_URL:urls.join('\n'),RESPOSTA_TEXTO:d.respostaTexto||'',
      DATA_ENVIO:dataEnvio,STATUS:'RECEBIDA',TENTATIVA:Number(d.tentativa||1),OBSERVACAO:d.observacao||'',ID_GRUPO:idGrupo
    };
    appendByHeaders_(sh,dados.headers,obj);
  });
  return {ok:true,idEntrega:idEntrega,idGrupo:idGrupo,integrantes:integrantes.length,arquivos:urls.length};
}

function validaArquivos_(arquivos,extensoes){const permitidas=String(extensoes||'').toLowerCase().split(',').map(x=>x.trim().replace(/^\./,'')).filter(Boolean);if(!permitidas.length)return;arquivos.forEach(f=>{const nome=String(f.name||'');const ext=nome.includes('.')?nome.split('.').pop().toLowerCase():'';if(!permitidas.includes(ext))throw new Error('Arquivo não permitido: '+nome)})}
function extensoesPadrao_(tipo){const m={FOTOS:'jpg,jpeg,png,webp',PDF:'pdf',WORD:'doc,docx',EXCEL:'xls,xlsx,xlsm,xltm',ARQUIVO:'pdf,doc,docx,xls,xlsx,xlsm,xltm,ppt,pptx,jpg,jpeg,png,webp'};return m[tipo]||''}
function config_(){const o={};rows_(ABAS.CONFIG).forEach(r=>o[String(r.CHAVE||'').trim()]=r.VALOR);return o}
function getOrCreateFolder_(pai,nome){const it=pai.getFoldersByName(nome);return it.hasNext()?it.next():pai.createFolder(nome)}
function sanitize_(s){return String(s||'').replace(/[\\/:*?"<>|#%{}~]/g,'-').replace(/\s+/g,' ').trim().slice(0,120)||'Sem nome'}
function sanitizeFile_(s){return String(s||'arquivo').replace(/[\\/:*?"<>|]/g,'-').slice(0,160)}
function sheetData_(sh){const v=sh.getDataRange().getValues();if(!v.length)return{headers:[],rows:[]};const headers=v[0].map(String);const rows=v.slice(1).filter(r=>r.some(c=>c!==''&&c!==null)).map((r,idx)=>{const o={__row:idx+2};headers.forEach((h,i)=>o[h]=r[i]);return o});return{headers,rows}}
function rows_(nome){return sheetData_(sh_(nome)).rows}
function appendByHeaders_(sh,headers,obj){sh.appendRow(headers.map(h=>obj[h]!==undefined?obj[h]:''))}
function upsert_(sh,dados,chave,valor,obj){const found=dados.rows.find(r=>eq_(r[chave],valor));const row=dados.headers.map(h=>obj[h]!==undefined?obj[h]:(found?found[h]:''));if(found)sh.getRange(found.__row,1,1,row.length).setValues([row]);else sh.appendRow(row)}
function eq_(a,b){return String(a??'').trim()===String(b??'').trim()}
function date_(v){if(!v)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,TZ,'yyyy-MM-dd');return String(v).slice(0,10)}

function autorizarDrive(){
  const cfg=config_();
  if(!cfg.PASTA_ENTREGAS_ID)throw new Error('PASTA_ENTREGAS_ID não configurada');
  const pasta=DriveApp.getFolderById(cfg.PASTA_ENTREGAS_ID);
  const teste=pasta.createFolder('TESTE_AUTORIZACAO_PE');
  teste.setTrashed(true);
  return 'OK';
}
