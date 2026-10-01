// Gera um HTML único (sem Node, sem Apps Script) com o front + cálculos + dados simulados embutidos.
const fs=require('fs'),path=require('path');
const REAL=process.env.REAL_DIR,OUT=process.argv[2]||'Morphe_CRM_local.html';
let mon,crm;
if(REAL&&fs.existsSync(path.join(REAL,'mon_real.json'))){mon=JSON.parse(fs.readFileSync(path.join(REAL,'mon_real.json')));crm=JSON.parse(fs.readFileSync(path.join(REAL,'crm_real.json')))}
else{mon=require('../test/sample.js').mon;crm={leads:[],atividades:[],users:[{nome:'Isa',email:''}],segmentos:[],origens:[],canais:[],resultados:[],motivos:[]}}
const srv=fs.readFileSync(path.join(__dirname,'server.js'),'utf8');
const a=srv.indexOf('mon.listas='),b=srv.indexOf('const server=');
const body=srv.slice(a,b);
const inteligencia=fs.readFileSync(path.join(__dirname,'../Inteligencia.js'),'utf8');
const hojeTxt=process.env.HOJE||'2026-10-01';
const shim=`<script>
var module={exports:{}};var exports=module.exports;
${inteligencia}
</script><script>
(function(){
var M=module.exports,HOJE=${JSON.stringify(hojeTxt)};
var mon=${JSON.stringify(mon)},crm=${JSON.stringify(crm)};
var META={kg:300,periodo:'mensal'};
${body.replace(/^mon\.listas=/,'mon.listas=').replace(/let META=[^\n]*\n/,'')}
function mk(ok,ko){return new Proxy({},{get:function(_,n){if(n==='withSuccessHandler')return function(f){return mk(f,ko)};if(n==='withFailureHandler')return function(f){return mk(ok,f)};
 return function(){var a=[].slice.call(arguments);setTimeout(function(){try{var r=H[n]?H[n].apply(null,JSON.parse(JSON.stringify(a))):{};ok&&ok(JSON.parse(JSON.stringify(r===undefined?null:r)))}catch(e){ko?ko({message:String(e.message||e)}):alert(e.message||e)}},30)}}})}
window.google={script:{run:mk()}};
})();
</script>`;
let h=fs.readFileSync(path.join(__dirname,'../Index.html'),'utf8');
h=h.replace('<script>',()=>shim+'<script>');
fs.writeFileSync(OUT,h);console.log('ok',OUT,(h.length/1024|0)+' KB');
