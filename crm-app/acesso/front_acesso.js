/* ===================== Entrada: domínio (como sempre) ou código de acesso ===================== */
var EXT={on:false,tok:'',nome:'',cod:''};
function gsr(){return EXT.on?extRun(null,null):google.script.run}
function extRun(ok,ko){return new Proxy({},{get:function(_,n){
  if(n==='withSuccessHandler')return function(f){return extRun(f,ko)};
  if(n==='withFailureHandler')return function(f){return extRun(ok,f)};
  return function(){extCall(n,[].slice.call(arguments),ok,ko,0)}}})}
function extCall(n,a,ok,ko,t){google.script.run
  .withSuccessHandler(function(r){if(ok)ok(r)})
  .withFailureHandler(function(e){var m=String((e&&e.message)||e);
    if(t<1&&/SESSAO_/.test(m)&&EXT.cod){extLogin(EXT.cod).then(function(){extCall(n,a,ok,ko,t+1)}).catch(function(){if(ko)ko(e)})}
    else if(ko)ko(e)})
  .apiExterna(EXT.tok,n,a)}
function extLogin(cod){return new Promise(function(res,rej){google.script.run
  .withSuccessHandler(function(r){EXT.on=true;EXT.tok=r.token;EXT.nome=r.nome;EXT.cod=cod;try{localStorage.setItem('morphe_cod',cod)}catch(e){}res(r)})
  .withFailureHandler(function(e){rej(e)}).entrarComCodigo(cod)})}
function telaCodigo(msg){var t=document.getElementById('telaCodigo');if(!t){
  document.body.insertAdjacentHTML('beforeend','<div id="telaCodigo" style="position:fixed;inset:0;z-index:99999;background:#054A37;display:flex;align-items:center;justify-content:center;padding:16px;font-family:Inter,Arial,sans-serif">'
   +'<form id="fCodigo" style="background:#FBF6EA;border-radius:16px;padding:32px;max-width:380px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,.35)">'
   +'<div style="font-family:Anton,Impact,sans-serif;font-size:30px;color:#054A37;letter-spacing:.4px">Morphê <span style="color:#006E51">Comercial</span></div>'
   +'<p style="color:#4B5E55;font-size:14px;margin:10px 0 18px">Digite o seu código de acesso para entrar.</p>'
   +'<input id="inCodigo" autocomplete="off" autocapitalize="characters" spellcheck="false" placeholder="XXXX-XXXX-XXXX" style="width:100%;box-sizing:border-box;padding:12px 14px;font-size:16px;letter-spacing:1px;border:1px solid #D9C9A8;border-radius:10px;background:#fff">'
   +'<div id="erroCodigo" style="color:#B4380F;font-size:13px;min-height:20px;margin:8px 0"></div>'
   +'<button id="btCodigo" type="submit" style="width:100%;padding:12px;border:0;border-radius:10px;background:#FA5528;color:#fff;font-weight:700;font-size:15px;cursor:pointer">Entrar</button></form></div>');
  document.getElementById('fCodigo').onsubmit=function(ev){ev.preventDefault();var c=document.getElementById('inCodigo').value,b=document.getElementById('btCodigo');if(!c.trim())return;b.disabled=true;b.textContent='Entrando…';
    extLogin(c).then(function(){var x=document.getElementById('telaCodigo');if(x)x.remove();load()}).catch(function(e){b.disabled=false;b.textContent='Entrar';document.getElementById('erroCodigo').textContent=String((e&&e.message)||e).replace(/^Error:\s*/,'')})}}
  document.getElementById('erroCodigo').textContent=msg||'';var i=document.getElementById('inCodigo');if(i)i.focus()}
function entrarExterno(){var c='';try{c=localStorage.getItem('morphe_cod')||''}catch(e){}
  if(!c){telaCodigo('');return}
  extLogin(c).then(function(){load()}).catch(function(e){try{localStorage.removeItem('morphe_cod')}catch(x){}telaCodigo(String((e&&e.message)||e).replace(/^Error:\s*/,''))})}
function iniciar(){google.script.run.withSuccessHandler(function(a){if(a&&a.modo==='externo')entrarExterno();else load()}).withFailureHandler(function(){load()}).getAcesso()}
