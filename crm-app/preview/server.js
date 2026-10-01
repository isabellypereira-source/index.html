// Servidor de PREVIEW local: roda o front real (Index.html) com um google.script.run simulado.
// Usa os mesmos cálculos do Apps Script (Inteligencia.js). Dados: REAL_DIR (reais, fora do git) ou test/sample.js
const http=require('http'),fs=require('fs'),path=require('path');
const M=require('../Inteligencia.js');
const REAL=process.env.REAL_DIR;
const HOJE=process.env.HOJE||'2026-10-01';
let mon,crm;
if(REAL&&fs.existsSync(path.join(REAL,'mon_real.json'))){mon=JSON.parse(fs.readFileSync(path.join(REAL,'mon_real.json')));crm=JSON.parse(fs.readFileSync(path.join(REAL,'crm_real.json')))}
else{mon=require('../test/sample.js').mon;crm={leads:[],atividades:[],users:[{nome:'Isa',email:''}],segmentos:[],origens:[],canais:[],resultados:[],motivos:[]}}
mon.listas={resultadosFup:['✅ Comprou','📅 Vai comprar (agendado)','⏸️ Adiou / pediu p/ voltar depois','🔇 Não respondeu','❌ Recusou'],canais:['WhatsApp','Ligação','E-mail','Visita','Instagram','Outro'],tiposContato:['Recompra','Amostra','Reativação'],
 transportes:[...new Set(mon.pedidos.map(p=>p.transporte).filter(Boolean))],origens:[...new Set(mon.pedidos.map(p=>p.origem).filter(Boolean))],tabelas:['Antiga','Nova'],faixas:['Auto','Varejo','Faixa 2','Faixa 3'],tiposCliente:['Ativo','Prospecção','Inativo'],ultimaTabela:'Antiga',
 motivosQueda:['Preço','Não girou / vendeu pouco','Concorrente','Frete / logística','Produto / qualidade','Fechou / mudou de foco','Sazonalidade / estoque alto','Sem resposta','Outro']};
let META={kg:300000,dias:365,inicio:'2026-01-01'};
function meta(){const d=s=>Date.UTC(+s.slice(0,4),+s.slice(5,7)-1,+s.slice(8,10))/864e5,H=d(HOJE),ini=d(META.inicio),fim=ini+META.dias;
 const ped=mon.pedidos.filter(p=>p.tipo==='Pedido'&&p.data);let kg=0,kg90=0,cl={};
 ped.forEach(p=>{const x=d(p.data);if(x>=ini&&x<=fim)kg+=p.kg;if(x>H-90&&x<=H){kg90+=p.kg;cl[p.cliente]=1}});
 const n=Object.keys(cl).length,kpc=n?kg90/n/3:0,dp=Math.max(1,H-ini),dr=Math.max(1,fim-H),rit=kg/dp,proj=kg+rit*dr,falta=Math.max(0,META.kg-proj);
 const nov=kpc>0?Math.ceil(falta/(kpc*(dr/30.4))):null;
 const iso=x=>new Date(x*864e5).toISOString().slice(0,10);
 return {metaKg:META.kg,periodoDias:META.dias,inicio:iso(ini),fim:iso(fim),kgAtingido:kg,progresso:META.kg?kg/META.kg:0,ritmoAtualDia:Math.round(rit*10)/10,kgPorClienteMes:Math.round(kpc*10)/10,novosClientesRealista:nov,novosClientesGordura:nov?Math.ceil(nov*1.3):null}}
const inteligencia=()=>M.montarInteligencia_(mon,HOJE,meta());
const H={
 getAll:()=>({me:{nome:'Isa',email:'isabelly.pereira@bioedtech.com.br'},leads:crm.leads,atividades:crm.atividades,users:crm.users,etapas:['Sugerido','Mapeado','Contato feito','Respondeu','Amostra enviada','Negociação','Ganho','Frio','Perdido'],segmentos:crm.segmentos,origens:crm.origens,canais:crm.canais,resultados:crm.resultados,motivos:crm.motivos,temChaveMaps:true,metaContatosSemana:30,metaGanhosMes:4,modelos:[],rodizio:crm.users.map(u=>u.nome)}),
 getInteligencia:()=>inteligencia(),
 metaProgresso:()=>meta(),
 salvarMeta:(kg,dias,ini)=>{META={kg:+kg,dias:+dias,inicio:ini||META.inicio};return meta()},
 registrarPedido:o=>{const q=['postas','c200','c300','d200','d300'].map(k=>+o[k]||0);const un=q.reduce((a,b)=>a+b,0);const kg=q[0]*.2+q[1]*.2+q[2]*.3+q[3]*.2+q[4]*.3;
   const valor=+o.valorAjustado||Math.round(un*38);mon.pedidos.push({data:o.data,cliente:o.cliente,tipo:'Pedido',postas:q[0],c200:q[1],c300:q[2],d200:q[3],d300:q[4],un,kg,entrega:o.entrega||'',transporte:o.transporte||'',origem:o.origem||'',obs:o.obs||'',tabela:o.tabela||'',valor,custo:valor*.45,margemR:valor*.55,margemPct:55});
   const c=mon.clientes.find(c=>c.nome===o.cliente);if(c&&c.tipo!=='Ativo')c.tipo='Ativo';const r=inteligencia();r.salvo={linha:99,un,kg,valorFinal:valor,margemR:valor*.55,margemPct:.55};return r},
 registrarFollowup:o=>{mon.fups.push({data:o.data,cliente:o.cliente,canal:o.canal,resultado:o.resultado,prox:o.prox,obs:o.obs,tipo:o.tipo});return inteligencia()},
 registrarSellout:o=>{mon.sellout.push({data:o.data,cliente:o.cliente,un:+o.un,obs:o.obs,autor:'Isa'});return inteligencia()},
 salvarFicha:o=>{mon.fichas[M.nrm_(o.cliente)]={cliente:o.cliente,ticketCardapio:+o.ticketCardapio||0,publico:o.publico,noCardapio:o.noCardapio,desde:o.desde,posicao:o.posicao,pedidosSemana:+o.pedidosSemana||0,motivo:o.motivo,acoes:o.acoes,notas:o.notas};return inteligencia()},
 salvarCliente:o=>{let c=mon.clientes.find(c=>c.nome===o.nome);if(!c){c={nome:o.nome};mon.clientes.push(c)}Object.assign(c,{tipo:o.tipo,cidade:o.cidade,ciclo:+o.ciclo||0,antec:+o.antec||0,obs:o.obs,motivo:o.motivo});return inteligencia()},
 registrarEntregaAmostra:o=>inteligencia()
};
const server=http.createServer((req,res)=>{
 if(req.method==='POST'&&req.url.startsWith('/rpc/')){let b='';req.on('data',c=>b+=c);req.on('end',()=>{const fn=req.url.slice(5);try{const out=H[fn]?H[fn](...JSON.parse(b||'[]')):{};res.setHeader('content-type','application/json');res.end(JSON.stringify({result:out}))}catch(e){res.end(JSON.stringify({error:String(e.message||e)}))}});return}
 if(req.url.startsWith('/Index')||req.url==='/'){let h=fs.readFileSync(path.join(__dirname,'../Index.html'),'utf8');
  const shim=`<script>(function(){function mk(ok,ko){return new Proxy({},{get:function(_,n){if(n==='withSuccessHandler')return function(f){return mk(f,ko)};if(n==='withFailureHandler')return function(f){return mk(ok,f)};
  return function(){var a=[].slice.call(arguments);fetch('/rpc/'+n,{method:'POST',body:JSON.stringify(a)}).then(function(r){return r.json()}).then(function(j){if(j.error){ko&&ko({message:j.error})}else{ok&&ok(j.result)}})}}})}
  window.google={script:{run:mk()}}})()</script>`;
  h=h.replace('<script>',shim+'<script>');res.setHeader('content-type','text/html; charset=utf-8');res.end(h);return}
 res.statusCode=404;res.end('nf')});
server.listen(process.env.PORT||8765,()=>console.log('preview em http://localhost:'+(process.env.PORT||8765)));
