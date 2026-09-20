const SHEET_ID='1QpyTzbc2P0F-LrGw7fIe7m0MOE10rbVSFAyCiWrgCNI';
const TURMA='1º MTEC - Administração - Mairinque';
const COMPONENTE='PE';
const TZ='America/Sao_Paulo';

function doGet(){
  return HtmlService.createHtmlOutputFromFile('Admin')
    .setTitle('Painel do Professor - PE')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function ss_(){return SpreadsheetApp.openById(SHEET_ID)}
function sh_(nome){const sh=ss_().getSheetByName(nome);if(!sh)throw new Error('Aba não encontrada: '+nome);return sh}
function rows_(nome){const sh=sh_(nome);const lastRow=sh.getLastRow(),lastCol=sh.getLastColumn();if(lastRow<2||lastCol<1)return[];const v=sh.getRange(1,1,lastRow,lastCol).getValues();const h=v.shift();return v.filter(r=>r.some(c=>c!==''&&c!==null)).map(r=>Object.fromEntries(h.map((k,i)=>[String(k),r[i]])))}
function dateTimeInput_(v,horaPadrao){if(!v)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,TZ,"yyyy-MM-dd'T'HH:mm");let s=String(v).trim();if(/^\d{4}-\d{2}-\d{2}$/.test(s))s+='T'+horaPadrao;return s.slice(0,16)}
function dateTime_(v){if(!v)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,TZ,'dd/MM/yyyy HH:mm:ss');return String(v)}
function ensureHeader_(sh,nome){const lastCol=Math.max(1,sh.getLastColumn());const h=sh.getRange(1,1,1,lastCol).getValues()[0].map(String);if(!h.includes(nome)){sh.getRange(1,lastCol+1).setValue(nome)}}

function listarAtividades(){
  return rows_('ATIVIDADES').filter(x=>x.TURMA===TURMA&&x.COMPONENTE===COMPONENTE).map(a=>({
    id:a.ID_ATIVIDADE,titulo:a.TITULO,descricao:a.DESCRICAO,orientacoes:a.ORIENTACOES||'',tipoEnvio:a.TIPO_ENVIO,
    extensoes:a.EXTENSOES,maxArquivos:a.MAX_ARQUIVOS,liberacao:dateTimeInput_(a.LIBERACAO,'00:00'),prazo:dateTimeInput_(a.PRAZO,'23:59'),materialUrl:a.MATERIAL_APOIO_URL,
    correcaoIA:a.CORRECAO_IA,criterios:a.GABARITO_CRITERIOS,status:a.STATUS,ordem:a.ORDEM,
    tipoParticipacao:String(a.TIPO_PARTICIPACAO||'INDIVIDUAL').toUpperCase()==='GRUPO'?'GRUPO':'INDIVIDUAL',
    maxAlunosGrupo:Math.max(1,Number(a.MAX_ALUNOS_GRUPO||1))
  })).sort((a,b)=>(Number(a.ordem)||999)-(Number(b.ordem)||999));
}

function salvarAtividade(d){
  if(!d||!d.id||!d.titulo)throw new Error('ID e título são obrigatórios');
  const liberacao=dateTimeInput_(d.liberacao,'00:00');
  const prazo=dateTimeInput_(d.prazo,'23:59');
  if(liberacao&&prazo&&liberacao>prazo)throw new Error('A liberação não pode ser posterior à finalização');
  const sh=sh_('ATIVIDADES');
  ensureHeader_(sh,'LIBERACAO');
  const v=sh.getDataRange().getValues(),h=v[0],idx=h.indexOf('ID_ATIVIDADE');
  const now=new Date();
  const tipoParticipacao=String(d.tipoParticipacao||'INDIVIDUAL').toUpperCase()==='GRUPO'?'GRUPO':'INDIVIDUAL';
  const map={
    ID_ATIVIDADE:d.id,TURMA:TURMA,COMPONENTE:COMPONENTE,TITULO:d.titulo,DESCRICAO:d.descricao||'',ORIENTACOES:d.orientacoes||'',
    TIPO_ENVIO:d.tipoEnvio||'SEM_ENVIO',EXTENSOES:d.extensoes||'',MAX_ARQUIVOS:Number(d.maxArquivos||0),LIBERACAO:liberacao,PRAZO:prazo,
    MATERIAL_APOIO_URL:d.materialUrl||'',CORRECAO_IA:d.correcaoIA||'NAO',GABARITO_CRITERIOS:d.criterios||'',
    STATUS:d.status||'RASCUNHO',ORDEM:Number(d.ordem||999),CRIADO_EM:now,ATUALIZADO_EM:now,
    TIPO_PARTICIPACAO:tipoParticipacao,MAX_ALUNOS_GRUPO:tipoParticipacao==='GRUPO'?Math.max(1,Number(d.maxAlunosGrupo||2)):1
  };
  let found=0;
  for(let i=1;i<v.length;i++)if(String(v[i][idx])===String(d.id)){found=i+1;break}
  if(found){
    const old=Object.fromEntries(h.map((k,i)=>[k,v[found-1][i]]));
    map.CRIADO_EM=old.CRIADO_EM||now;
    const row=h.map(k=>map[k]!==undefined?map[k]:'');
    sh.getRange(found,1,1,row.length).setValues([row]);
  }else{
    const row=h.map(k=>map[k]!==undefined?map[k]:'');
    sh.appendRow(row);
  }
  return {ok:true,id:d.id};
}

function excluirAtividade(id){
  if(!id)throw new Error('ID da atividade não informado');
  const sh=sh_('ATIVIDADES');
  const v=sh.getDataRange().getValues();
  if(v.length<2)throw new Error('Nenhuma atividade cadastrada');
  const h=v[0];
  const idxId=h.indexOf('ID_ATIVIDADE');
  const idxTurma=h.indexOf('TURMA');
  const idxComp=h.indexOf('COMPONENTE');
  if(idxId<0)throw new Error('Coluna ID_ATIVIDADE não encontrada');
  for(let i=1;i<v.length;i++){
    const mesmoId=String(v[i][idxId])===String(id);
    const mesmaTurma=idxTurma<0||String(v[i][idxTurma])===TURMA;
    const mesmoComponente=idxComp<0||String(v[i][idxComp])===COMPONENTE;
    if(mesmoId&&mesmaTurma&&mesmoComponente){
      sh.deleteRow(i+1);
      return {ok:true,id:id};
    }
  }
  throw new Error('Atividade não encontrada');
}

function listarMateriais(){
  return rows_('MATERIAIS')
    .filter(x=>String(x.COMPONENTE||'')===COMPONENTE)
    .map(m=>({id:String(m.ID_MATERIAL||''),titulo:String(m.TITULO||''),descricao:String(m.DESCRICAO||''),url:String(m.ARQUIVO_URL||''),tipo:String(m.TIPO||'MATERIAL'),ordem:Number(m.ORDEM||999),status:String(m.STATUS||'OCULTO'),publicadoEm:dateTime_(m.PUBLICADO_EM)}))
    .sort((a,b)=>(Number(a.ordem)||999)-(Number(b.ordem)||999));
}

function salvarMaterial(d){
  if(!d||!d.titulo||!d.url)throw new Error('Título e URL são obrigatórios');
  const sh=sh_('MATERIAIS');
  const v=sh.getDataRange().getValues();
  const h=v[0];
  const idx=h.indexOf('ID_MATERIAL');
  if(idx<0)throw new Error('Coluna ID_MATERIAL não encontrada');
  const id=String(d.id||('MAT-'+Utilities.getUuid().slice(0,8).toUpperCase())).trim();
  const now=new Date();
  const map={ID_MATERIAL:id,COMPONENTE:COMPONENTE,TITULO:String(d.titulo||'').trim(),DESCRICAO:String(d.descricao||''),ARQUIVO_URL:String(d.url||'').trim(),TIPO:String(d.tipo||'MATERIAL'),ORDEM:Number(d.ordem||999),STATUS:String(d.status||'PUBLICADO'),PUBLICADO_EM:now};
  let found=0;
  for(let i=1;i<v.length;i++)if(String(v[i][idx])===id){found=i+1;break}
  if(found){const old=Object.fromEntries(h.map((k,i)=>[k,v[found-1][i]]));if(old.PUBLICADO_EM)map.PUBLICADO_EM=old.PUBLICADO_EM;const row=h.map(k=>map[k]!==undefined?map[k]:'');sh.getRange(found,1,1,row.length).setValues([row]);}
  else sh.appendRow(h.map(k=>map[k]!==undefined?map[k]:''));
  return {ok:true,id:id};
}

function excluirMaterial(id){
  if(!id)throw new Error('ID do material não informado');
  const sh=sh_('MATERIAIS');
  const v=sh.getDataRange().getValues();
  if(v.length<2)throw new Error('Nenhum material cadastrado');
  const h=v[0],idxId=h.indexOf('ID_MATERIAL'),idxComp=h.indexOf('COMPONENTE');
  if(idxId<0)throw new Error('Coluna ID_MATERIAL não encontrada');
  for(let i=1;i<v.length;i++){
    const mesmoId=String(v[i][idxId])===String(id);
    const mesmoComponente=idxComp<0||String(v[i][idxComp])===COMPONENTE;
    if(mesmoId&&mesmoComponente){sh.deleteRow(i+1);return {ok:true,id:id};}
  }
  throw new Error('Material não encontrado');
}

function listarEntregas(idAtividade){
  return rows_('ENTREGAS')
    .filter(x=>!idAtividade||String(x.ID_ATIVIDADE)===String(idAtividade))
    .map(e=>({
      id:String(e.ID_ENTREGA||''),idAtividade:String(e.ID_ATIVIDADE||''),aluno:String(e.ALUNO||''),email:String(e.EMAIL||''),
      arquivos:String(e.ARQUIVOS_URL||''),resposta:String(e.RESPOSTA_TEXTO||''),dataEnvio:dateTime_(e.DATA_ENVIO),
      status:String(e.STATUS||''),tentativa:Number(e.TENTATIVA||1),observacao:String(e.OBSERVACAO||''),idGrupo:String(e.ID_GRUPO||'')
    }));
}

function listarCorrecoes(idAtividade){
  return rows_('CORRECOES').filter(x=>!idAtividade||String(x.ID_ATIVIDADE)===String(idAtividade)).map(x=>{const o={};Object.keys(x).forEach(k=>o[k]=Object.prototype.toString.call(x[k])==='[object Date]'?dateTime_(x[k]):x[k]);return o;});
}
