/* ===================== Acesso externo por código (pessoas fora do domínio) =====================
 * Quem é do domínio interno entra como sempre (o Google informa o e-mail).
 * Quem é de fora (ex.: Rayssa) precisa de um código pessoal, guardado na aba Config:
 *   coluna I = "Código de acesso: Nome"   coluna J = o código.
 * Para revogar, apague o código na planilha. Todo login fica registrado na aba "Acessos".
 * Toda função pública do app chama exigir_() e só roda para quem é do domínio ou tem sessão válida.
 */
var ACESSO_DOMINIO_ = '@bioedtech.com.br';
var ACESSO_EXT_ = null; // { nome, email } durante uma chamada autenticada por código (vale só para a execução atual)
var ACESSO_TTL_ = 21600; // 6 horas (máximo do CacheService)
var ACESSO_PREFIXO_ = 'Código de acesso: ';

function emailAtivo_() { try { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; } }
function internoOk_() { var e = emailAtivo_(); return e.length > ACESSO_DOMINIO_.length && e.slice(-ACESSO_DOMINIO_.length) === ACESSO_DOMINIO_; }
function exigir_() { if (ACESSO_EXT_ || internoOk_()) return; throw new Error('SEM_ACESSO: entre com o seu código de acesso.'); }

/** Primeira chamada da página: diz se precisa de código. Não devolve nenhum dado. */
function getAcesso() { return { modo: internoOk_() ? 'interno' : 'externo' }; }

/** Na planilha só fica o HASH do código ("sha256:sal:hash"); o código em si nunca é gravado. */
function normCodigo_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function hashCodigo_(sal, cod) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + normCodigo_(cod), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join('');
}
function confereCodigo_(valor, cod) { var p = String(valor).split(':'); return p.length === 3 && p[0] === 'sha256' && hashCodigo_(p[1], cod) === p[2]; }
function codigosExternos_() {
  var v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config').getRange('I2:J60').getValues(), out = [];
  v.forEach(function (r) {
    var k = String(r[0]).trim();
    if (k.indexOf(ACESSO_PREFIXO_) === 0 && String(r[1]).trim()) out.push({ nome: k.slice(ACESSO_PREFIXO_.length).trim(), valor: String(r[1]).trim() });
  });
  return out;
}
function emailDoUsuario_(nome) {
  var v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config').getRange('F2:G30').getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === nome) return String(v[i][1]).trim().toLowerCase();
  return '';
}
function logAcesso_(nome, evento) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName('Acessos');
    if (!s) { s = ss.insertSheet('Acessos'); s.appendRow(['Data', 'Pessoa', 'Evento']); s.setFrozenRows(1); }
    s.appendRow([new Date(), nome, evento]);
  } catch (e) {}
}

/** Valida o código e abre uma sessão de 6 h. Um novo login derruba a sessão anterior da mesma pessoa. */
function entrarComCodigo(codigo) {
  var cache = CacheService.getScriptCache(), falhas = +(cache.get('acs_falhas') || 0);
  if (falhas >= 20) throw new Error('Muitas tentativas. Aguarde 15 minutos e tente de novo.');
  var c = normCodigo_(codigo), achou = null;
  if (c.length >= 12) codigosExternos_().forEach(function (x) { if (confereCodigo_(x.valor, c)) achou = x; });
  if (!achou) {
    cache.put('acs_falhas', String(falhas + 1), 900); Utilities.sleep(1500);
    logAcesso_('(desconhecido)', 'código recusado');
    throw new Error('Código inválido.');
  }
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var info = { nome: achou.nome, email: emailDoUsuario_(achou.nome) };
  cache.put('sess_' + token, JSON.stringify(info), ACESSO_TTL_);
  cache.put('ativa_' + achou.nome, token, ACESSO_TTL_);
  logAcesso_(achou.nome, 'entrou');
  return { token: token, nome: achou.nome, email: info.email };
}

function sessaoExterna_(token) {
  var cache = CacheService.getScriptCache(), raw = token ? cache.get('sess_' + token) : null;
  if (!raw) throw new Error('SESSAO_EXPIRADA');
  var info = JSON.parse(raw);
  if (cache.get('ativa_' + info.nome) !== token) throw new Error('SESSAO_SUBSTITUIDA');
  // o código ainda existe? (permite revogar apagando a linha do Config)
  if (!codigosExternos_().some(function (x) { return x.nome === info.nome; })) throw new Error('SESSAO_EXPIRADA');
  return info;
}

/** Única porta de entrada para quem usa código: confere a sessão e chama a função pedida. */
var ACESSO_PERMITIDAS_ = ['getAll', 'getInteligencia', 'saveLead', 'moveLead', 'addActivity', 'reativarLead', 'bulkMove', 'metaProgresso', 'salvarMeta', 'salvarMetaKg', 'searchPlaces', 'importPlaces', 'addLeads', 'removerSugerido', 'distribuir', 'addTask', 'doneTask', 'relatorioParaMim', 'registrarSellout', 'salvarFicha', 'registrarAmostra', 'registrarPedido', 'registrarEntregaAmostra', 'registrarFollowup', 'salvarCliente', 'mesclarClientes', 'getAvisos', 'transferirLead'];
function apiExterna(token, fn, args) {
  var info = sessaoExterna_(token);
  if (ACESSO_PERMITIDAS_.indexOf(fn) < 0 || typeof globalThis[fn] !== 'function') throw new Error('Função não permitida: ' + fn);
  ACESSO_EXT_ = info;
  return globalThis[fn].apply(null, args || []);
}

/** Sementes de código (hash), preenchidas na publicação. Cria a linha no Config se ela ainda não existir. */
var ACESSO_SEMENTES_ = {};
function migrarAcessoExterno_() {
  var C = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config'), msg = [];
  Object.keys(ACESSO_SEMENTES_).forEach(function (nome) {
    var v = C.getRange('I2:J60').getValues(), chave = ACESSO_PREFIXO_ + nome, lin = -1, vazio = -1;
    for (var i = 0; i < v.length; i++) { if (String(v[i][0]).trim() === chave) lin = i; if (vazio < 0 && !v[i][0] && !v[i][1]) vazio = i; }
    if (lin >= 0) { msg.push(nome + ': já existe'); return; }
    if (vazio < 0) vazio = v.length;
    C.getRange(vazio + 2, 9, 1, 2).setValues([[chave, ACESSO_SEMENTES_[nome]]]);
    msg.push(nome + ': acesso criado na linha ' + (vazio + 2));
  });
  return msg.join('; ') || 'sem sementes';
}
