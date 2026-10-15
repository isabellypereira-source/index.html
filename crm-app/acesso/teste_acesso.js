// Teste do bloqueio de acesso: carrega o projeto já "protegido" num sandbox com stubs do Apps Script.
const vm=require('vm'),fs=require('fs'),path=require('path');
const dir=process.argv[2];
const files=fs.readdirSync(dir).filter(f=>/\.(js|gs)$/.test(f));
function novo(email){
  const cache={};const cfg=[['Segmentos','Origens','Canais','Resultados','Motivos','Usuários (nome)','E-mail','','Configuração','Valor'],
   ['x','','','','','Isa','isabelly.pereira@bioedtech.com.br','','Meta v2: kg','100'],['','','','','','Rayssa','ray@hiconnexions.com','','Código de acesso: Rayssa','sha256:abc123:'+require('crypto').createHash('sha256').update('abc123:ABCD2345WXYZ1234').digest('hex')]];
  const sheet=(rows)=>({getRange:(r,c,nr,nc)=>{if(typeof r==='string'){const m=r.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);const c1=m[1].charCodeAt(0)-65,c2=m[3].charCodeAt(0)-65,r1=+m[2]-1,r2=+m[4];return {getValues:()=>rows.slice(r1,r2).map(x=>{const o=[];for(let i=c1;i<=c2;i++)o.push(x[i]===undefined?'':x[i]);return o}),setValues(){},setValue(){}}}return {getValues:()=>[[]],setValues(){},setValue(){}}},getLastRow:()=>10,getLastColumn:()=>10,appendRow(){},setFrozenRows(){}});
  const ss={getSheetByName:(n)=>n==='Config'?sheet(cfg):(n==='Acessos'?sheet([]):null),insertSheet:()=>sheet([])};
  const ctx={console,Logger:{log(){}},Utilities:{getUuid:()=>require('crypto').randomUUID(),sleep(){},formatDate:()=>'',DigestAlgorithm:{SHA_256:'s'},Charset:{UTF_8:'u'},computeDigest:(a,t)=>[...require('crypto').createHash('sha256').update(t).digest()].map(x=>x>127?x-256:x)},
   Session:{getActiveUser:()=>({getEmail:()=>email}),getEffectiveUser:()=>({getEmail:()=>'isabelly.pereira@bioedtech.com.br'})},
   CacheService:{getScriptCache:()=>({get:k=>cache[k]??null,put:(k,v)=>{cache[k]=v},remove:k=>{delete cache[k]}})},
   SpreadsheetApp:{getActiveSpreadsheet:()=>ss,openById:()=>ss},PropertiesService:{getScriptProperties:()=>({getProperty:()=>null,setProperty(){}})},
   LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},ScriptApp:{},UrlFetchApp:{},HtmlService:{},MailApp:{},GmailApp:{},module:undefined};
  vm.createContext(ctx);
  // ordem: Código, Inteligencia, Acesso
  const ordem=files.sort((a,b)=>(/^C/.test(a)?0:/^Int/.test(a)?1:2)-(/^C/.test(b)?0:/^Int/.test(b)?1:2));
  ordem.forEach(f=>vm.runInContext(fs.readFileSync(path.join(dir,f),'utf8'),ctx,{filename:f}));
  return ctx;
}
let falhas=0;const ok=(c,m)=>{console.log((c?'OK   ':'FALHA ')+m);if(!c)falhas++};
const src=files.filter(f=>!/^Acesso/.test(f)).map(f=>fs.readFileSync(path.join(dir,f),'utf8')).join('\n');
const publicas=[...src.matchAll(/^function ([A-Za-z0-9]+)\(/gm)].map(m=>m[1]);
const ISENTAS=['doGet','buscaSemanal','sincronizarWix','relatorioSemanal'];
// 1) estranho (e-mail vazio ou de outro domínio) NÃO passa em nenhuma função pública
for(const email of ['','alguem@gmail.com','x@bioedtech.com.br.evil.com']){
  const c=novo(email);let vazou=[];
  publicas.filter(n=>!ISENTAS.includes(n)).forEach(n=>{try{c[n]({},{},{})}catch(e){if(!/SEM_ACESSO/.test(String(e.message)))vazou.push(n+': '+e.message)}});
  ok(vazou.length===0,'e-mail "'+email+'": todas as '+(publicas.length-ISENTAS.length)+' funções públicas recusam'+(vazou.length?' | passaram: '+vazou.join('; '):''));
}
// 2) interno passa da barreira
{const c=novo('isabelly.pereira@bioedtech.com.br');let r='';try{c.getAll()}catch(e){r=e.message}ok(!/SEM_ACESSO/.test(r),'interno passa da barreira ('+(r||'rodou')+')');
 ok(c.getAcesso().modo==='interno','getAcesso interno')}
{const c=novo('');ok(c.getAcesso().modo==='externo','getAcesso externo (e-mail vazio)')}
// 3) fluxo do código
{const c=novo('');
 let e1='';try{c.entrarComCodigo('ERRADO-9999')}catch(e){e1=e.message}ok(/inválido/i.test(e1),'código errado recusado');
 let e2='';try{c.entrarComCodigo('')}catch(e){e2=e.message}ok(/inválido/i.test(e2),'código vazio recusado');
 const s=c.entrarComCodigo('abcd-2345-wxyz-1234');ok(s&&s.nome==='Rayssa'&&s.email==='ray@hiconnexions.com'&&s.token.length>=32,'código certo (minúsculas, com traços) entra como Rayssa');
 let e3='';try{c.apiExterna('token-falso','getAll',[])}catch(e){e3=e.message}ok(/SESSAO_EXPIRADA/.test(e3),'token falso recusado');
 let e4='';try{c.apiExterna(s.token,'setup',[])}catch(e){e4=e.message}ok(/não permitida/.test(e4),'função fora da lista (setup) recusada mesmo com sessão');
 let e5='';try{c.apiExterna(s.token,'relatorioSemanal',['x@x.com'])}catch(e){e5=e.message}ok(/não permitida/.test(e5),'relatorioSemanal fora da lista recusada');
 let e6='';try{c.apiExterna(s.token,'getAll',[])}catch(e){e6=e.message}ok(!/SEM_ACESSO|SESSAO|não permitida/.test(e6),'sessão válida passa da barreira em getAll ('+(e6||'rodou')+')');
 ok(c.ACESSO_EXT_&&c.ACESSO_EXT_.nome==='Rayssa'&&c.me_().nome==='Rayssa','me_() devolve a identidade da sessão');
 const s2=c.entrarComCodigo('ABCD-2345-WXYZ-1234');let e7='';try{c.apiExterna(s.token,'getAll',[])}catch(e){e7=e.message}ok(/SESSAO_SUBSTITUIDA/.test(e7),'novo login derruba a sessão anterior (outro aparelho com o mesmo código)');
 for(let i=0;i<21;i++)try{c.entrarComCodigo('ERRADO'+i+'XXXX')}catch(e){}
 let e8='';try{c.entrarComCodigo('ABCD-2345-WXYZ-1234')}catch(e){e8=e.message}ok(/Muitas tentativas/.test(e8),'bloqueia após muitas tentativas erradas');
}
console.log(falhas?('\n'+falhas+' FALHA(S)'):'\nTudo certo');process.exit(falhas?1:0);
