"""Aplica o acesso externo por código sobre uma pasta de projeto do Apps Script (clone do clasp).
uso: python3 aplicar_acesso.py <pasta_do_projeto> <pasta_work_com_Index.html_e_Acesso.js>
Idempotente: pode rodar mais de uma vez."""
import re,sys,os,json,shutil
proj,work=sys.argv[1],sys.argv[2]
EXEMPTAS={'doGet','buscaSemanal','sincronizarWix','relatorioSemanal','onOpen','onEdit','onInstall'}
def cod_file():
    for n in ('Código.js','Codigo.js','Código.gs','Codigo.gs'):
        if os.path.exists(os.path.join(proj,n)): return n
    raise SystemExit('Código.js não encontrado em '+proj)
def guard(src,fname):
    n=0
    def f(m):
        nonlocal n
        nome=m.group(1)
        if nome.endswith('_') or nome in EXEMPTAS: return m.group(0)
        if src[m.end():m.end()+40].lstrip().startswith('exigir_();'): return m.group(0)
        n+=1; return m.group(0)+' exigir_();'
    out=re.sub(r'(?m)^function ([A-Za-z0-9_]+)\([^)]*\)\s*\{',f,src)
    print('  %s: %d função(ões) protegida(s)'%(fname,n)); return out
cf=cod_file(); P=lambda n:os.path.join(proj,n)
c=open(P(cf),encoding='utf8').read()
c=guard(c,cf)
# me_(): identidade da sessão por código
if 'ACESSO_EXT_' not in c.split('function me_()')[1][:300]:
    c=c.replace('function me_() {','function me_() {\n  if (ACESSO_EXT_) return { email: ACESSO_EXT_.email || \'\', nome: ACESSO_EXT_.nome };',1)
# relatorioSemanal: só exige acesso quando chamada com um e-mail (trigger passa o evento)
m=re.search(r'function relatorioSemanal\((\w+)\)\s*\{',c)
if m and "exigir_" not in c[m.end():m.end()+120]:
    c=c[:m.end()]+" if (typeof %s === 'string' && %s) exigir_();"%(m.group(1),m.group(1))+c[m.end():]
# doGet: migrações só quando é alguém do domínio
c=c.replace("try { rodarMigracoes_(); }","try { if (internoOk_()) rodarMigracoes_(); }",1)
# migração do código externo
if 'acesso_externo' not in c:
    c,k=re.subn(r'(var MIGRACOES_ = \[.*?\})(\s*\n\];)',lambda m:m.group(1)+",\n  { chave: 'mig_2026_10_15_acesso_externo', fn: function () { return migrarAcessoExterno_(); } }"+m.group(2),c,count=1,flags=re.S)
    assert k==1,'MIGRACOES_ não encontrado'
open(P(cf),'w',encoding='utf8').write(c)
# demais arquivos do projeto (menos o Acesso.js e o arquivo principal já tratado)
for fn in sorted(os.listdir(proj)):
    if fn==cf or fn=='Acesso.js' or not re.search(r'\.(js|gs)$',fn): continue
    t=open(P(fn),encoding='utf8').read(); open(P(fn),'w',encoding='utf8').write(guard(t,fn))
shutil.copy(os.path.join(work,'acesso','Acesso.js'),P('Acesso.js')) if os.path.exists(os.path.join(work,'acesso','Acesso.js')) else shutil.copy(os.path.join(work,'Acesso.js'),P('Acesso.js'))
shutil.copy(os.path.join(work,'Index.html'),P('Index.html'))
# manifesto: só o acesso do web app (escopos OAuth intactos)
mf=json.load(open(P('appsscript.json'),encoding='utf8')); mf.setdefault('webapp',{})['access']='ANYONE'
json.dump(mf,open(P('appsscript.json'),'w',encoding='utf8'),indent=2,ensure_ascii=False)
print('ok; escopos:',mf.get('oauthScopes','(não definidos)'))
