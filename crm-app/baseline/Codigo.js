// ================================================================== limparDuplicatasAtv — remove Contatos duplicados do mesmo dia/lead sem observação
function limparDuplicatasAtv() {
  var A = sh_('Atividades');
  var rows = A.getDataRange().getValues();
  var h = rows[0];
  var iData = h.indexOf('Data'), iLeadId = h.indexOf('Lead ID'), iTipo = h.indexOf('Tipo'),
      iTexto = h.indexOf('Texto'), iRes = h.indexOf('Resultado');
  // 1ª passagem: marca o primeiro Contato por leadId+dia como "visto"
  var seen = {}, toDelete = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][iTipo] !== 'Contato') continue;
    var lid = String(rows[i][iLeadId] || ''); if (!lid) continue;
    var dv = rows[i][iData]; var ds = dv instanceof Date ? day_(dv) : String(dv).slice(0, 10);
    var key = lid + '|' + ds;
    var txt = String(rows[i][iTexto] || '').trim(), res = String(rows[i][iRes] || '').trim();
    if (seen[key]) {
      // duplicata: remove se não tem texto nem resultado (clique duplo acidental)
      if (!txt && !res) toDelete.push(i + 1); // linha da planilha (1-based)
    } else {
      seen[key] = true;
    }
  }
  if (!toDelete.length) { Logger.log('Nenhuma duplicata sem observação.'); return 'Nenhuma duplicata encontrada.'; }
  // deleta de baixo pra cima pra não deslocar índices
  toDelete.sort(function(a, b) { return b - a; });
  toDelete.forEach(function(r) { A.deleteRow(r); });
  SpreadsheetApp.flush();
  var msg = toDelete.length + ' atividade(s) duplicada(s) sem observação removida(s).';
  Logger.log(msg);
  return msg;
}

// ================================================================== ajustes8 — detecta e corrige IDs duplicados nos Leads
function ajustes8() {
  var L = sh_('Leads');
  var rows = L.getDataRange().getValues();
  var h = rows[0];
  var idIdx = h.indexOf('ID'), nomeIdx = h.indexOf('Estabelecimento');
  if (idIdx < 0) throw new Error('Coluna ID não encontrada');
  var seen = {}, fixed = 0;
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][idIdx] || '');
    if (!id) continue;
    if (seen[id] !== undefined) {
      var novoId = newId_('L');
      L.getRange(i + 1, idIdx + 1).setValue(novoId);
      Logger.log('DUPLICATA CORRIGIDA: ' + rows[i][nomeIdx] + ' | linha ' + (i + 1) + ' | ' + id + ' -> ' + novoId);
      fixed++;
    } else {
      seen[id] = i;
    }
  }
  if (fixed > 0) SpreadsheetApp.flush();
  var msg = fixed > 0 ? fixed + ' duplicata(s) corrigida(s). O histórico de atividades fica no lead original; o lead duplicado começa limpo.' : 'Nenhuma duplicata encontrada.';
  Logger.log(msg);
  return msg;
}

// ================================================================== ajustes7 — cap 50 + desativa relatório semanal
function ajustes7() {
  // 1) Aumenta cap de sugestões por semana para 50
  var C = sh_('Config'), v = C.getRange('I2:J30').getValues(), capOk = false;
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'Máx. sugestões por semana') { C.getRange(i + 2, 10).setValue(50); capOk = true; break; }
  }
  if (!capOk) {
    for (var j = 0; j < v.length; j++) {
      if (!v[j][0]) { C.getRange(j + 2, 9, 1, 2).setValues([['Máx. sugestões por semana', 50]]); break; }
    }
  }
  // 2) Desativa gatilho do relatório semanal (a pedido — reativar depois com instalarMelhorias)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'relatorioSemanal') ScriptApp.deleteTrigger(t);
  });
  return 'ajustes7 ok';
}

/**
 * Morphê Comercial — mini CRM (Apps Script Web App)
 * Dados: abas Leads, Atividades, Config desta planilha.
 * Publicar: Implantar > Nova implantação > App da Web
 *   Executar como: Eu | Quem pode acessar: Qualquer pessoa em bioedtech.com.br
 */
var CFG = {
  SS_ID: '1Gb7rU_HOoS-vCg4f9SAYzuAinexOKjFZtDH27ATorp0',          // planilha do CRM (dados)
  CONTATOS_ID: '1HKewTq-2zYebCaNKR4WACodmHVsAqnTjHzKDLFrlfZs',     // planilha "contatos morphê" (importação)
  MONITOR_ID: '1sP8-D1TSDDmlualNFnb6i9-blS_iP8GXq_Uxk4SkO1M',      // Monitoramento Clientes B2B
  ETAPAS: ['Sugerido', 'Mapeado', 'Contato feito', 'Respondeu', 'Amostra enviada', 'Negociação', 'Ganho', 'Frio', 'Perdido'],
  LEAD_COLS: ['ID', 'Estabelecimento', 'Segmento', 'Cidade', 'Bairro', 'Contato (pessoa)', 'WhatsApp', 'Telefone', 'Instagram', 'E-mail', 'Site',
    'Google Maps', 'Place ID', 'Origem', 'Responsável', 'Etapa', 'Próximo passo', 'Data próximo passo', 'Último contato', 'Potencial (R$/mês)',
    'Motivo de perda', 'Observações', 'Criado em', 'Atualizado em', 'Etapa desde'],
  ATV_COLS: ['ID', 'Data', 'Lead ID', 'Estabelecimento', 'Autor', 'Tipo', 'Canal', 'Resultado', 'Texto', 'Menções']
};

function doGet() {
  try { rodarMigracoes_(); } catch (e) { Logger.log('doGet migrações: ' + e); }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Morphê · Comercial')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------------------------------ migrações automáticas
// Cada migração roda 1x só (flag em ScriptProperties). Adicionar uma nova migração:
// acrescentar {chave: 'mig_AAAA_MM_DD_nome', fn: suaFuncao} na lista abaixo.
var MIGRACOES_ = [
  { chave: 'mig_2026_10_01_ajustes8', fn: ajustes8 },
  { chave: 'mig_2026_10_01_limparDuplicatasAtv', fn: limparDuplicatasAtv }
];
function rodarMigracoes_() {
  var props = PropertiesService.getScriptProperties();
  var pendentes = MIGRACOES_.filter(function (m) { return !props.getProperty(m.chave); });
  if (!pendentes.length) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // outra requisição já está rodando as migrações
  try {
    pendentes.forEach(function (m) {
      if (props.getProperty(m.chave)) return; // outra requisição concluiu enquanto esperava o lock
      try {
        var res = m.fn();
        props.setProperty(m.chave, new Date().toISOString());
        Logger.log('Migração ' + m.chave + ' ok: ' + res);
      } catch (e) {
        Logger.log('Migração ' + m.chave + ' falhou (tenta de novo na próxima vez): ' + e);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------ helpers
function ss_() { return SpreadsheetApp.openById(CFG.SS_ID); }
function sh_(n) { return ss_().getSheetByName(n); }
function iso_(v) { return (v instanceof Date) ? Utilities.formatDate(v, 'America/Sao_Paulo', "yyyy-MM-dd'T'HH:mm:ss") : (v === null || v === undefined ? '' : v); }
function day_(v) { return (v instanceof Date) ? Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd') : (v || ''); }
function parseDay_(s) { if (!s) return ''; if (s instanceof Date) return s; var p = String(s).slice(0, 10).split('-'); if (p.length < 3) return ''; return new Date(+p[0], +p[1] - 1, +p[2]); }
function newId_(p) { return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
function norm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, ''); }
function cfgList_(col) {
  var v = sh_('Config').getRange(2, col, 60, 1).getValues().map(function (r) { return String(r[0]).trim(); }).filter(String);
  return v;
}
function users_() {
  var v = sh_('Config').getRange('F2:G30').getValues();
  return v.filter(function (r) { return r[0]; }).map(function (r) { return { nome: String(r[0]).trim(), email: String(r[1]).trim().toLowerCase() }; });
}
function setting_(key) {
  var v = sh_('Config').getRange('I2:J30').getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === key) return v[i][1];
  return '';
}
function me_() {
  var email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  var u = users_().filter(function (x) { return x.email && x.email === email; })[0];
  return { email: email, nome: u ? u.nome : '' };
}

// ------------------------------------------------------------------ leitura
function getAll() {
  var L = sh_('Leads'), A = sh_('Atividades');
  var lv = L.getDataRange().getValues(), lh = lv.shift();
  var leads = lv.filter(function (r) { return r[0]; }).map(function (r) { return serializeRow_(lh, r); });
  var av = A.getDataRange().getValues(), ah = av.shift();
  var atv = av.filter(function (r) { return r[0]; }).map(function (r) { return serializeRow_(ah, r); });
  var base = { me: me_(), leads: leads, atividades: atv };
  var cfg = cachedConfig_();
  Object.keys(cfg).forEach(function (k) { base[k] = cfg[k]; });
  return base;
}

// Config/listas que quase nunca mudam (Config, usuários, modelos, rodízio) — cacheadas 5 min
// pra não reler a aba Config em todo load. invalidarCacheConfig_() limpa na hora se precisar.
var cachedConfig_cache_ = null;
function cachedConfig_() {
  if (cachedConfig_cache_) return cachedConfig_cache_;
  var cache = CacheService.getScriptCache();
  var hit = cache.get('cfg_v1');
  if (hit) { cachedConfig_cache_ = JSON.parse(hit); return cachedConfig_cache_; }
  var data = {
    users: users_(), etapas: CFG.ETAPAS, segmentos: cfgList_(1), origens: cfgList_(2), canais: cfgList_(3), resultados: cfgList_(4), motivos: cfgList_(5),
    temChaveMaps: !!setting_('Google Maps API Key'), metaContatosSemana: setting_('Meta: contatos por semana') || 30,
    metaGanhosMes: setting_('Meta: novos clientes por mês') || 4,
    modelos: modelos_(), rodizio: rodizioLista_()
  };
  try { cache.put('cfg_v1', JSON.stringify(data), 300); } catch (e) { Logger.log('cachedConfig_ put: ' + e); }
  cachedConfig_cache_ = data;
  return data;
}
function invalidarCacheConfig_() { try { CacheService.getScriptCache().remove('cfg_v1'); } catch (e) {} cachedConfig_cache_ = null; }

// ------------------------------------------------------------------ escrita
function serializeRow_(h, row) {
  var o = {}; h.forEach(function (k, i) { o[k] = iso_(row[i]); }); return o;
}
function rowOf_(sheet, id) {
  var ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
/** Lê de volta 1 lead já escrito (1 leitura de linha, não a planilha inteira) p/ resposta enxuta. */
function leadLeve_(id) {
  var L = sh_('Leads'), r = rowOf_(L, id); if (r < 0) return null;
  var h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0];
  return serializeRow_(h, L.getRange(r, 1, 1, h.length).getValues()[0]);
}
function writeLead_(o) {
  var L = sh_('Leads'), h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0];
  var r = o.ID ? rowOf_(L, o.ID) : -1;
  var now = new Date();
  var cur = r > 0 ? L.getRange(r, 1, 1, h.length).getValues()[0] : h.map(function () { return ''; });
  var obj = {}; h.forEach(function (k, i) { obj[k] = cur[i]; });
  Object.keys(o).forEach(function (k) { if (h.indexOf(k) >= 0) obj[k] = o[k]; });
  if (!obj.ID) { obj.ID = newId_('L'); obj['Criado em'] = now; }
  if (!obj['Etapa desde'] || (r > 0 && o.Etapa && o.Etapa !== cur[h.indexOf('Etapa')])) obj['Etapa desde'] = now;
  obj['Atualizado em'] = now;
  ['Data próximo passo', 'Último contato'].forEach(function (k) { if (obj[k] && !(obj[k] instanceof Date)) obj[k] = parseDay_(obj[k]); });
  var row = h.map(function (k) { return obj[k] === undefined ? '' : obj[k]; });
  if (r > 0) L.getRange(r, 1, 1, h.length).setValues([row]); else L.appendRow(row);
  return obj.ID;
}
/** Cria a atividade e devolve o objeto já serializado (sem reler a planilha) p/ resposta enxuta. */
function logAtv_(a) {
  var A = sh_('Atividades');
  var me = me_();
  var id = newId_('A'), data = new Date(), autor = a.autor || me.nome || me.email, tipo = a.tipo || 'Comentário';
  var leadId = a.leadId || '', nome = a.nome || '', canal = a.canal || '', resultado = a.resultado || '', texto = a.texto || '', mencoes = a.mencoes || '';
  A.appendRow([id, data, leadId, nome, autor, tipo, canal, resultado, texto, mencoes]);
  return { ID: id, Data: iso_(data), 'Lead ID': leadId, Estabelecimento: nome, Autor: autor, Tipo: tipo, Canal: canal, Resultado: resultado, Texto: texto, 'Menções': mencoes };
}

function saveLead(o) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var novo = !o.ID;
    var id = writeLead_(o);
    var atvs = [];
    if (novo) atvs.push(logAtv_({ leadId: id, nome: o.Estabelecimento, tipo: 'Sistema', texto: 'Lead criado (' + (o.Origem || 'manual') + ')' }));
    if (o.Etapa === 'Ganho') enviarParaClientes_(o);
    return { leads: [leadLeve_(id)], atividades: atvs };
  } finally { lock.releaseLock(); }
}

function moveLead(id, etapa, motivo) {
  var t0 = Date.now(); // TEMP Fase 1: medir ganho de performance — remover depois de aprovado
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var L = sh_('Leads'), r = rowOf_(L, id); if (r < 0) throw new Error('Lead não encontrado');
    var h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0];
    var row = L.getRange(r, 1, 1, h.length).getValues()[0];
    var antes = row[h.indexOf('Etapa')];
    var upd = { ID: id, Etapa: etapa };
    if (etapa === 'Perdido' && motivo) upd['Motivo de perda'] = motivo;
    if (etapa === 'Mapeado' && !row[h.indexOf('Responsável')]) upd['Responsável'] = me_().nome;
    writeLead_(upd);
    var atv1 = logAtv_({ leadId: id, nome: row[h.indexOf('Estabelecimento')], tipo: 'Etapa', texto: antes + ' → ' + etapa + (motivo ? ' (' + motivo + ')' : '') });
    if (etapa === 'Ganho') {
      var o = {}; h.forEach(function (k, i) { o[k] = row[i]; }); enviarParaClientes_(o);
    }
    if (etapa === 'Amostra enviada' && antes !== etapa) {
      registrarAmostrasMonitoramento_(row[h.indexOf('Estabelecimento')], new Date(), null);
    }
    var out = { leads: [leadLeve_(id)], atividades: [atv1] };
    Logger.log('[PERF] moveLead: ' + (Date.now() - t0) + 'ms');
    return out;
  } finally { lock.releaseLock(); }
}

/** Registra contato/comentário. a = {leadId, tipo, canal, resultado, texto, proximoPasso, dataProx, novaEtapa, mencoes:[nomes]} */
function addActivity(a) {
  var t0 = Date.now(); // TEMP Fase 1: medir ganho de performance — remover depois de aprovado
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    // Leitura única da aba Atividades (reaproveitada pela regra de "1 contato/dia" e pela de "5+ sem resposta" —
    // antes eram 2 leituras completas da aba por ação)
    var atv, aLI, aTI, aDI, aRI;
    if (a.tipo === 'Contato' && a.leadId) {
      var A0 = sh_('Atividades'); atv = A0.getDataRange().getValues();
      var ah = atv[0];
      aLI = ah.indexOf('Lead ID'); aTI = ah.indexOf('Tipo'); aDI = ah.indexOf('Data'); aRI = ah.indexOf('Resultado');
      var hj = day_(new Date());
      for (var xi = 1; xi < atv.length; xi++) {
        if (String(atv[xi][aLI]) === String(a.leadId) && atv[xi][aTI] === 'Contato' && day_(atv[xi][aDI]) === hj) {
          if (!a.texto || !a.texto.trim()) throw new Error('Já há um contato registrado hoje para este lead. Escreva uma observação sobre o que mudou para registrar outro.');
          break;
        }
      }
    }
    var nome = '';
    if (a.leadId) {
      var L = sh_('Leads'), r = rowOf_(L, a.leadId);
      if (r < 0) throw new Error('Lead não encontrado: ' + a.leadId);
      var h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0];
      var leadRow = L.getRange(r, 1, 1, h.length).getValues()[0]; // 1 leitura da linha em vez de 2 getValue() separados
      nome = leadRow[h.indexOf('Estabelecimento')];
      var upd = { ID: a.leadId };
      if (a.tipo === 'Contato') upd['Último contato'] = new Date();
      if (a.proximoPasso !== undefined && a.proximoPasso !== null) upd['Próximo passo'] = a.proximoPasso;
      if (a.dataProx) upd['Data próximo passo'] = parseDay_(a.dataProx);
      var etapaAtual = leadRow[h.indexOf('Etapa')];
      if (a.novaEtapa && a.novaEtapa !== etapaAtual) upd.Etapa = a.novaEtapa;
      else if (a.tipo === 'Contato' && ['Sugerido', 'Mapeado'].indexOf(etapaAtual) >= 0) upd.Etapa = 'Contato feito';
      if (a.tipo === 'Contato' && /respondeu|interess|pediu|amostra|valores/i.test(a.resultado || '') && ['Sugerido', 'Mapeado', 'Contato feito'].indexOf(etapaAtual) >= 0 && !a.novaEtapa) upd.Etapa = 'Respondeu';
      if (upd.Etapa === 'Perdido' && a.motivo) upd['Motivo de perda'] = a.motivo;
      writeLead_(upd);
      var atvs = [];
      if (upd.Etapa && upd.Etapa !== etapaAtual) atvs.push(logAtv_({ leadId: a.leadId, nome: nome, tipo: 'Etapa', texto: etapaAtual + ' → ' + upd.Etapa }));
      if (upd.Etapa === 'Ganho') {
        var oGanho = {}; h.forEach(function (k, i) { oGanho[k] = leadRow[i]; }); enviarParaClientes_(oGanho);
      }
      // Regra: 5+ Contatos sem resposta positiva → Frio
      if (a.tipo === 'Contato') {
        var etapasPos = ['Respondeu','Amostra enviada','Negociação','Ganho','Perdido','Frio'];
        var etapaFinal = upd.Etapa || etapaAtual;
        if (etapasPos.indexOf(etapaFinal) < 0) {
          var nC = 1, temR = /respondeu|interess|pediu|amostra|valores/i.test(String(a.resultado||''));
          for (var ci2 = 1; ci2 < atv.length; ci2++) {
            if (String(atv[ci2][aLI]) === String(a.leadId) && atv[ci2][aTI] === 'Contato') {
              nC++;
              if (/respondeu|interess|pediu|amostra|valores/i.test(String(atv[ci2][aRI]||''))) temR = true;
            }
          }
          if (nC >= 5 && !temR) {
            writeLead_({ID: a.leadId, Etapa: 'Frio'});
            atvs.push(logAtv_({leadId: a.leadId, nome: nome, tipo: 'Etapa', texto: etapaFinal + ' → Frio ('+nC+' tentativas sem resposta)'}));
          }
        }
      }
      // Amostra enviada → registra a amostra no Monitoramento Clientes B2B (prospecção)
      if (upd.Etapa === 'Amostra enviada') {
        registrarAmostrasMonitoramento_(nome, new Date(), null);
      }
    }
    var men = (a.mencoes || []).join(', ');
    var atvPrincipal = logAtv_({ leadId: a.leadId, nome: nome, tipo: a.tipo, canal: a.canal, resultado: a.resultado, texto: a.texto, mencoes: men });
    notificar_(a.mencoes || [], nome, a.texto, a.leadId);
    var leadsOut = a.leadId ? [leadLeve_(a.leadId)] : [];
    var atividadesOut = (typeof atvs !== 'undefined' ? atvs : []).concat([atvPrincipal]);
    var out = { leads: leadsOut, atividades: atividadesOut };
    Logger.log('[PERF] addActivity: ' + (Date.now() - t0) + 'ms');
    return out;
  } finally { lock.releaseLock(); }
}

/** Reativa um lead Frio → volta para "Contato feito" */
function reativarLead(id) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var L = sh_('Leads'), r = rowOf_(L, id);
    if (r < 0) throw new Error('Lead não encontrado');
    var h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0];
    var etapaAtual = L.getRange(r, h.indexOf('Etapa') + 1).getValue();
    if (etapaAtual !== 'Frio') throw new Error('Lead não está como Frio');
    var nome = L.getRange(r, h.indexOf('Estabelecimento') + 1).getValue();
    writeLead_({ID: id, Etapa: 'Contato feito'});
    var atv1 = logAtv_({leadId: id, nome: nome, tipo: 'Etapa', texto: 'Frio → Contato feito (reativado)'});
    return { leads: [leadLeve_(id)], atividades: [atv1] };
  } finally { lock.releaseLock(); }
}

function notificar_(nomes, estab, texto, leadId) {
  if (!nomes.length) return;
  var me = me_();
  var url = ScriptApp.getService().getUrl() || '';
  users_().forEach(function (u) {
    if (nomes.indexOf(u.nome) >= 0 && u.email && u.email !== me.email) {
      try {
        MailApp.sendEmail({
          to: u.email,
          subject: '[Morphê Comercial] ' + (me.nome || me.email) + ' mencionou você' + (estab ? ' em ' + estab : ''),
          htmlBody: '<p><b>' + (me.nome || me.email) + '</b> escreveu' + (estab ? ' sobre <b>' + estab + '</b>' : ' no mural') + ':</p><blockquote>' +
            String(texto || '').replace(/</g, '&lt;') + '</blockquote><p><a href="' + url + '">Abrir o Morphê Comercial</a></p>'
        });
      } catch (e) {}
    }
  });
}

function bulkMove(ids, etapa, responsavel) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ids.forEach(function (id) { var u = { ID: id, Etapa: etapa }; if (responsavel) u['Responsável'] = responsavel; writeLead_(u); });
    var atv1 = logAtv_({ tipo: 'Sistema', texto: ids.length + ' lead(s) movidos para ' + etapa + (responsavel ? ' (resp.: ' + responsavel + ')' : '') });
    return { leads: ids.map(leadLeve_), atividades: [atv1] };
  } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------ Ganho -> Monitoramento Clientes B2B
function enviarParaClientes_(o) {
  try {
    var cl = SpreadsheetApp.openById(CFG.MONITOR_ID).getSheetByName('Clientes');
    var nomes = cl.getRange('A2:A60').getValues().map(function (r) { return norm_(r[0]); });
    var idx = nomes.indexOf(norm_(o.Estabelecimento));
    if (idx >= 0) {
      // já existe (ex.: entrou como prospecção por amostra) → promove para Ativo
      var rowExist = idx + 2, tipoAtual = cl.getRange(rowExist, 2).getValue();
      if (tipoAtual !== 'Ativo') cl.getRange(rowExist, 2).setValue('Ativo');
      return;
    }
    var i = nomes.indexOf('');
    if (i < 0) return;
    var cidade = o.Cidade ? (o.Cidade + (/\/[A-Z]{2}$/.test(o.Cidade) ? '' : '/SP')) : '';
    cl.getRange(i + 2, 1, 1, 6).setValues([[o.Estabelecimento, 'Ativo', cidade, 21, 3, 'Veio do Morphê Comercial · resp.: ' + (o['Responsável'] || '')]]);
  } catch (e) { Logger.log('enviarParaClientes_: ' + e); }
}

/** Amostra (prospecção) -> Monitoramento Clientes B2B.
 * Chamada 2x pela mesma amostra: 1ª com dataPedido (etapa "Amostra enviada"), depois com
 * dataEntrega (tarefa "mandar amostra" concluída). Casa com a linha "Amostra" pendente (sem
 * Data de entrega) do mesmo cliente; se não achar, cria uma nova. Não conta como cliente ativo. */
function registrarAmostrasMonitoramento_(nomeEstab, dataPedido, dataEntrega) {
  try {
    var ss = SpreadsheetApp.openById(CFG.MONITOR_ID);

    // Clientes: garante a linha do prospect (não baixa quem já é Ativo)
    var cl = ss.getSheetByName('Clientes');
    var clNomes = cl.getRange('A2:A60').getValues();
    var rowCl = -1, livreCl = -1;
    for (var i = 0; i < clNomes.length; i++) {
      if (!clNomes[i][0]) { if (livreCl < 0) livreCl = i; continue; }
      if (norm_(clNomes[i][0]) === norm_(nomeEstab)) { rowCl = i + 2; break; }
    }
    if (rowCl < 0 && livreCl >= 0) {
      cl.getRange(livreCl + 2, 1, 1, 5).setValues([[String(nomeEstab).trim(), 'Prospecção', '', 7, 0]]);
    } else if (rowCl > 0) {
      var tipoAtual = cl.getRange(rowCl, 2).getValue();
      if (tipoAtual !== 'Ativo' && tipoAtual !== 'Prospecção') cl.getRange(rowCl, 2).setValue('Prospecção');
    }

    // Pedidos B2B: acha a linha "Amostra" pendente deste cliente (sem Data de entrega) ou abre uma nova
    var pb = ss.getSheetByName('Pedidos B2B');
    var clientes = pb.getRange('B2:B1001').getValues();
    var tipos = pb.getRange('C2:C1001').getValues();
    var entregas = pb.getRange('K2:K1001').getValues();
    var rowPb = -1, livrePb = -1;
    for (var j = 0; j < clientes.length; j++) {
      if (!clientes[j][0]) { if (livrePb < 0) livrePb = j; continue; }
      if (norm_(clientes[j][0]) === norm_(nomeEstab) && tipos[j][0] === 'Amostra' && !entregas[j][0]) { rowPb = j + 2; break; }
    }
    if (rowPb < 0) { if (livrePb < 0) return; rowPb = livrePb + 2; }
    pb.getRange(rowPb, 2, 1, 2).setValues([[String(nomeEstab).trim(), 'Amostra']]);
    if (dataPedido) pb.getRange(rowPb, 1).setValue(dataPedido);
    if (dataEntrega) pb.getRange(rowPb, 11).setValue(dataEntrega);
  } catch (e) { Logger.log('registrarAmostrasMonitoramento_: ' + e); }
}

// ------------------------------------------------------------------ Meta de volume (kg)
// Soma "Peso total (kg)" (coluna J) de Pedidos B2B tipo "Pedido" (ignora Amostra) no Monitoramento.
function metaProgresso() {
  var kgMeta = +setting_('Meta: kg (período)') || 300000;
  var periodoDias = +setting_('Meta: período (dias)') || 365;
  var inicioStr = setting_('Meta: início do período');
  var inicio = inicioStr ? parseDay_(inicioStr) : new Date(new Date().getFullYear(), 0, 1);
  var fim = new Date(inicio.getTime() + periodoDias * 864e5);
  var pb = SpreadsheetApp.openById(CFG.MONITOR_ID).getSheetByName('Pedidos B2B');
  var dados = pb.getRange('A2:J1001').getValues(); // A=Data do pedido, B=Cliente, C=Tipo, J=Peso total (kg)
  var kg = 0, kg90 = 0, clientes90 = {};
  var hoje = new Date(), d90 = new Date(hoje.getTime() - 90 * 864e5);
  dados.forEach(function (r) {
    var d = r[0], cliente = r[1], tipo = r[2], peso = +r[9] || 0;
    if (!(d instanceof Date) || tipo !== 'Pedido') return;
    if (d >= inicio && d <= fim) kg += peso;
    if (d >= d90 && d <= hoje) { kg90 += peso; if (cliente) clientes90[cliente] = true; }
  });
  var nClientes90 = Object.keys(clientes90).length;
  var kgPorClienteMes = nClientes90 ? (kg90 / nClientes90) / 3 : 0; // janela de 90 dias ~ 3 meses

  var diasPassados = Math.max(1, Math.round((hoje - inicio) / 864e5));
  var diasRestantes = Math.max(1, (fim - hoje) / 864e5);
  var mesesRestantes = diasRestantes / 30.4;
  var ritmoAtual = kg / diasPassados;
  var kgProjetadoCarteiraAtual = kg + ritmoAtual * diasRestantes; // se a carteira atual mantiver o ritmo até o fim
  var faltaDeNovos = Math.max(0, kgMeta - kgProjetadoCarteiraAtual);
  var novosClientesRealista = kgPorClienteMes > 0 ? Math.ceil(faltaDeNovos / (kgPorClienteMes * mesesRestantes)) : null;
  var novosClientesGordura = novosClientesRealista !== null ? Math.ceil(novosClientesRealista * 1.3) : null; // +30% de margem
  return {
    metaKg: kgMeta, periodoDias: periodoDias, inicio: iso_(inicio), fim: iso_(fim),
    kgAtingido: kg, progresso: kgMeta ? kg / kgMeta : 0,
    ritmoAtualDia: Math.round(ritmoAtual * 10) / 10,
    kgPorClienteMes: Math.round(kgPorClienteMes * 10) / 10,
    novosClientesRealista: novosClientesRealista, novosClientesGordura: novosClientesGordura
  };
}
function salvarMeta(kg, periodoDias, inicio) {
  var C = sh_('Config'), v = C.getRange('I2:J30').getValues();
  function set(key, val) {
    for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === key) { C.getRange(i + 2, 10).setValue(val); return; }
    for (var j = 0; j < v.length; j++) if (!v[j][0]) { C.getRange(j + 2, 9, 1, 2).setValues([[key, val]]); v[j] = [key, val]; return; }
  }
  set('Meta: kg (período)', +kg || 300000);
  set('Meta: período (dias)', +periodoDias || 365);
  set('Meta: início do período', inicio || '');
  invalidarCacheConfig_();
  return metaProgresso();
}

// ------------------------------------------------------------------ Google Maps (Places API New)
function searchPlaces(query, cidade) {
  var key = setting_('Google Maps API Key');
  if (!key) throw new Error('Cadastre a chave do Google Maps na aba Config (célula J2).');
  var body = { textQuery: query + (cidade ? ' em ' + cidade : ''), languageCode: 'pt-BR', regionCode: 'BR', pageSize: 20 };
  var res = UrlFetchApp.fetch('https://places.googleapis.com/v1/places:searchText', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body),
    headers: {
      'X-Goog-Api-Key': key,
      'X-Goog-FieldMask': 'places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,places.internationalPhoneNumber,places.websiteUri,places.googleMapsUri,places.rating,places.userRatingCount,places.primaryTypeDisplayName,places.businessStatus'
    }
  });
  var js = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('Google Maps: ' + ((js.error && js.error.message) || res.getResponseCode()));
  var L = sh_('Leads'), lv = L.getDataRange().getValues(), lh = lv.shift();
  var iP = lh.indexOf('Place ID'), iN = lh.indexOf('Estabelecimento');
  var existP = {}, existN = {};
  lv.forEach(function (r) { if (r[iP]) existP[r[iP]] = 1; existN[norm_(r[iN])] = 1; });
  return (js.places || []).filter(function (p) { return p.businessStatus !== 'CLOSED_PERMANENTLY'; }).map(function (p) {
    var nome = p.displayName ? p.displayName.text : '';
    return {
      placeId: p.id, nome: nome, endereco: p.formattedAddress || '', telefone: p.nationalPhoneNumber || '', telIntl: p.internationalPhoneNumber || '',
      site: p.websiteUri || '', maps: p.googleMapsUri || '', nota: p.rating || '', avaliacoes: p.userRatingCount || 0,
      tipo: p.primaryTypeDisplayName ? p.primaryTypeDisplayName.text : '',
      jaExiste: !!(existP[p.id] || existN[norm_(nome)] || dupDe_(nome, '', p.nationalPhoneNumber || ''))
    };
  });
}

function importPlaces(list, segmento, cidade) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    list.forEach(function (p) {
      var cel = /^\(?\d{2}\)?\s?9\d{4}-?\d{4}$/.test(String(p.telefone).trim());
      writeLead_({
        Estabelecimento: p.nome, Segmento: segmento || '', Cidade: cidade || '', Bairro: p.endereco,
        WhatsApp: cel ? p.telefone : '', Telefone: cel ? '' : p.telefone, Site: p.site, 'Google Maps': p.maps, 'Place ID': p.placeId,
        Origem: 'Google Maps', Etapa: 'Sugerido',
        Observações: (p.tipo ? p.tipo + ' · ' : '') + (p.nota ? 'nota ' + p.nota + ' (' + p.avaliacoes + ' avaliações)' : '')
      });
    });
    logAtv_({ tipo: 'Sistema', texto: list.length + ' sugestão(ões) importadas do Google Maps' + (cidade ? ' · ' + cidade : '') });
    return getAll();
  } finally { lock.releaseLock(); }
}

/** Busca semanal automática: usa a lista "Buscas automáticas" da aba Config (colunas L:M). */
function buscaSemanal() {
  if (!setting_('Google Maps API Key')) return;
  var v = sh_('Config').getRange('L2:N40').getValues().filter(function (r) { return r[0]; });
  var limite = +setting_('Máx. sugestões por semana') || 50, total = 0;
  for (var i = 0; i < v.length && total < limite; i++) {
    try {
      var res = searchPlaces(v[i][0], v[i][1]).filter(function (p) { return !p.jaExiste && p.avaliacoes >= 5; }).slice(0, limite - total);
      if (res.length) { importPlaces(res, v[i][2] || '', v[i][1]); total += res.length; }
    } catch (e) { Logger.log(e); }
  }
}
function instalarBuscaSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'buscaSemanal') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('buscaSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
}

// ------------------------------------------------------------------ duplicados (comparação "inteligente")
var GENERICOS_ = ['restaurante','emporio','emporium','cozinha','vegana','vegano','vegan','veg','vegetariano','vegetariana','vegetal','natural','naturais',
  'produtos','cafe','cafeteria','bistro','bar','loja','mercado','supermercado','supermercados','hortifruti','comida','culinaria','plant','based','delicias',
  'casa','espaco','store','food','foods','market','saudavel','saudaveis','granel','atelie','organicos','organico','sao','paulo','campinas','maria','santa','verde','vida','grao','graos','terra','mundo','sabor','sabores'];
function tokens_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9 ]/g, ' ').split(/\s+/)
    .map(function (w) { return w.replace(/(.)\1+/g, '$1'); })
    .filter(function (w) { return w.length >= 4 && GENERICOS_.indexOf(w) < 0; });
}
function ig_(h) { h = String(h || '').trim().toLowerCase(); var m = h.match(/instagram\.com\/([^/?#\s]+)/); if (m) h = m[1]; return h.replace(/^@/, ''); }
function tel_(t) { var d = String(t || '').replace(/\D/g, ''); if (d.length > 11 && d.indexOf('55') === 0) d = d.slice(2); return d.length >= 10 ? d : ''; }
/** Retorna o nome do lead já existente que é o mesmo (nome igual/contido, mesmo @ ou mesmo telefone). */
function dupDe_(nome, insta, fone, ignorarId) {
  var L = sh_('Leads'), v = L.getDataRange().getValues(), h = v.shift();
  var iN = h.indexOf('Estabelecimento'), iI = h.indexOf('Instagram'), iW = h.indexOf('WhatsApp'), iT = h.indexOf('Telefone'), iE = h.indexOf('Etapa');
  var n = norm_(nome).replace(/(.)\1+/g, '$1'), g = ig_(insta), f = tel_(fone);
  for (var i = 0; i < v.length; i++) {
    var r = v[i]; if (!r[0] || r[0] === ignorarId) continue;
    var m = norm_(r[iN]).replace(/(.)\1+/g, '$1');
    if (n && m && (n === m || (n.length >= 6 && m.indexOf(n) >= 0) || (m.length >= 6 && n.indexOf(m) >= 0))) return r[iN];
    if (g && ig_(r[iI]) === g) return r[iN];
    if (f && (tel_(r[iW]) === f || tel_(r[iT]) === f)) return r[iN];
  }
  return '';
}

/** Adição rápida: lista de leads [{Estabelecimento, Instagram, WhatsApp, E-mail, Site, Cidade, Segmento, Etapa, Responsável, Observações, Origem, canal}] */
function addLeads(list) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var criados = [], ignorados = [], me = me_();
    list.forEach(function (o) {
      if (!o.Estabelecimento) return;
      var d = dupDe_(o.Estabelecimento, o.Instagram, o.WhatsApp);
      if (d) { ignorados.push(o.Estabelecimento + ' (já existe: ' + d + ')'); return; }
      var etapa = o.Etapa || 'Mapeado', contatado = ['Sugerido', 'Mapeado'].indexOf(etapa) < 0;
      var lead = {};
      Object.keys(o).forEach(function (k) { if (CFG.LEAD_COLS.indexOf(k) >= 0) lead[k] = o[k]; });
      lead.Etapa = etapa; lead['Responsável'] = o['Responsável'] || me.nome || '';
      lead.Origem = o.Origem || 'Prospecção ativa';
      if (!lead.Segmento) lead.Segmento = segGuess_(o.Estabelecimento);
      if (contatado) { lead['Último contato'] = new Date(); if (!lead['Próximo passo']) { lead['Próximo passo'] = 'Fazer follow-up'; lead['Data próximo passo'] = new Date(Date.now() + 3 * 864e5); } }
      var id = writeLead_(lead);
      logAtv_({ leadId: id, nome: o.Estabelecimento, tipo: 'Sistema', texto: 'Lead criado (adição rápida)' });
      if (contatado) logAtv_({ leadId: id, nome: o.Estabelecimento, tipo: 'Contato', canal: o.canal || (o.Instagram ? 'Instagram' : (o.WhatsApp ? 'WhatsApp' : 'Outro')), resultado: 'Enviado (aguardando)', texto: 'Primeiro contato' });
      criados.push(o.Estabelecimento);
    });
    var r = getAll(); r.criados = criados; r.ignorados = ignorados; return r;
  } finally { lock.releaseLock(); }
}

/** Remove um lead sugerido que é duplicado (apaga a linha). */
function removerSugerido(id) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { var L = sh_('Leads'), r = rowOf_(L, id); if (r > 0 && L.getRange(r, CFG.LEAD_COLS.indexOf('Etapa') + 1).getValue() === 'Sugerido') L.deleteRow(r); return getAll(); }
  finally { lock.releaseLock(); }
}


// ================================================================== MELHORIAS (fase "Agora")
// ---- modelos de mensagem (Config, colunas P:Q)
function modelos_() {
  try {
    var v = sh_('Config').getRange('P2:Q30').getValues();
    return v.filter(function (r) { return r[0] && r[1]; }).map(function (r) { return { nome: String(r[0]).trim(), texto: String(r[1]) }; });
  } catch (e) { return []; }
}

// ---- rodízio de leads sem dono (Config: "Rodízio")
function rodizioLista_() {
  var v = String(setting_('Rodízio (nomes, separados por vírgula)') || '').split(',').map(function (x) { return x.trim(); }).filter(String);
  var nomes = users_().map(function (u) { return u.nome; });
  v = v.filter(function (x) { return nomes.indexOf(x) >= 0; });
  return v.length ? v : nomes;
}
function proximoRodizio_() {
  var lista = rodizioLista_(), p = PropertiesService.getScriptProperties();
  var i = (+p.getProperty('RODIZIO_IDX') || 0) % lista.length;
  p.setProperty('RODIZIO_IDX', String(i + 1));
  return lista[i];
}
function distribuir(ids) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var cont = {};
    ids.forEach(function (id) { var n = proximoRodizio_(); cont[n] = (cont[n] || 0) + 1; writeLead_({ ID: id, 'Responsável': n }); });
    logAtv_({ tipo: 'Sistema', texto: ids.length + ' lead(s) sem dono distribuídos: ' + Object.keys(cont).map(function (k) { return k + ' ' + cont[k]; }).join(', ') });
    return getAll();
  } finally { lock.releaseLock(); }
}

// ---- tarefas (Atividades: Tipo "Tarefa", Resultado Aberta/Feita, responsável em Menções, coluna Prazo)
function addTask(t) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var A = sh_('Atividades'), me = me_(), nome = '';
    if (t.leadId) { var L = sh_('Leads'), r = rowOf_(L, t.leadId); if (r > 0) nome = L.getRange(r, CFG.LEAD_COLS.indexOf('Estabelecimento') + 1).getValue(); }
    A.appendRow([newId_('T'), new Date(), t.leadId || '', nome, me.nome || me.email, 'Tarefa', '', 'Aberta', t.texto || '', t.pessoa || me.nome || '', t.prazo ? parseDay_(t.prazo) : '', '']);
    if (t.pessoa && t.pessoa !== me.nome) notificar_([t.pessoa], nome, 'Nova tarefa para você: ' + (t.texto || '') + (t.prazo ? ' (prazo ' + t.prazo.split('-').reverse().join('/') + ')' : ''), t.leadId);
    return getAll();
  } finally { lock.releaseLock(); }
}
function doneTask(id, feita) {
  var t0 = Date.now(); // TEMP Fase 1: medir ganho de performance — remover depois de aprovado
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var A = sh_('Atividades'), r = rowOf_(A, id); if (r < 0) throw new Error('Tarefa não encontrada');
    var h = A.getRange(1, 1, 1, A.getLastColumn()).getValues()[0];
    var row = A.getRange(r, 1, 1, h.length).getValues()[0];
    var iRes = h.indexOf('Resultado'), iFeita = h.indexOf('Feita em');
    row[iRes] = feita === false ? 'Aberta' : 'Feita';
    if (iFeita >= 0) row[iFeita] = feita === false ? '' : new Date();
    A.getRange(r, 1, 1, h.length).setValues([row]); // 1 escrita em lote da linha inteira
    // Tarefa de "mandar amostra" concluída → registra a entrega no Monitoramento Clientes B2B
    if (feita !== false) {
      var nomeT = row[h.indexOf('Estabelecimento')], textoT = row[h.indexOf('Texto')];
      if (nomeT && /amostra/i.test(String(textoT || ''))) {
        registrarAmostrasMonitoramento_(nomeT, null, new Date());
      }
    }
    var out = { leads: [], atividades: [serializeRow_(h, row)] };
    Logger.log('[PERF] doneTask: ' + (Date.now() - t0) + 'ms');
    return out;
  } finally { lock.releaseLock(); }
}

// ---- formulário do site (Wix) -> CRM, de hora em hora
function sincronizarWix() {
  return; // desligado a pedido da Isa (28/09/2026)
  var antes = sh_('Leads').getLastRow();
  importarWix_();
  var L = sh_('Leads'), depois = L.getLastRow();
  if (depois <= antes) return;
  var H = CFG.LEAD_COLS, novos = L.getRange(antes + 1, 1, depois - antes, H.length).getValues();
  var url = ScriptApp.getService().getUrl() || '';
  novos.forEach(function (r) {
    var resp = r[H.indexOf('Responsável')], u = users_().filter(function (x) { return x.nome === resp && x.email; })[0];
    if (u) try {
      MailApp.sendEmail({ to: u.email, subject: '[Morphê Comercial] Novo lead do site: ' + r[1],
        htmlBody: '<p><b>' + r[1] + '</b> preencheu o formulário do site e caiu para você.</p><p>' + (r[H.indexOf('Observações')] || '') + '</p><p>Responder em até 24 h. <a href="' + url + '">Abrir o Morphê Comercial</a></p>' });
    } catch (e) {}
  });
}

// ---- relatório semanal por e-mail
function relatorioSemanal(so) {
  var d = getAll(), now = new Date(), seteDias = 7 * 864e5;
  var recente = function (s) { if (!s) return false; var x = new Date(s); return now - x <= seteDias && now - x >= 0; };
  var RESP = ['Respondeu', 'Interessado', 'Pediu valores / material', 'Pediu amostra', 'Encaminhou para outro contato/canal', 'Pediu para retornar depois', 'Sem interesse'];
  var C = d.atividades.filter(function (a) { return a.Tipo === 'Contato' && recente(a.Data); });
  var etapas = d.atividades.filter(function (a) { return a.Tipo === 'Etapa' && recente(a.Data); });
  var ganhos = etapas.filter(function (a) { return /→ Ganho/.test(a.Texto); }).map(function (a) { return a.Estabelecimento; });
  var amostras = etapas.filter(function (a) { return /→ Amostra enviada/.test(a.Texto); }).length;
  var fimHoje = new Date(); fimHoje.setHours(23, 59, 59, 999);
  var ativos = d.leads.filter(function (l) { return ['Sugerido', 'Ganho', 'Perdido'].indexOf(l.Etapa) < 0; });
  var td = 'style="padding:6px 10px;border-bottom:1px solid #EADFCB"';
  var linhas = d.users.map(function (u) {
    var c = C.filter(function (a) { return a.Autor === u.nome; });
    var r = c.filter(function (a) { return RESP.indexOf(a.Resultado) >= 0; }).length;
    var venc = ativos.filter(function (l) { return l['Responsável'] === u.nome && l['Data próximo passo'] && new Date(l['Data próximo passo']) <= fimHoje; }).length;
    var tar = d.atividades.filter(function (a) { return a.Tipo === 'Tarefa' && a.Resultado !== 'Feita' && a['Menções'] === u.nome; }).length;
    return '<tr><td ' + td + '><b>' + u.nome + '</b></td><td ' + td + '>' + c.length + '</td><td ' + td + '>' + r + '</td><td ' + td + '>' + venc + '</td><td ' + td + '>' + tar + '</td></tr>';
  }).join('');
  var novos = d.leads.filter(function (l) { return recente(l['Criado em']); }).length;
  var site = d.leads.filter(function (l) { return /site|formul/i.test(l.Origem) && !l['Último contato'] && ['Sugerido', 'Mapeado'].indexOf(l.Etapa) >= 0; });
  var resp = C.filter(function (a) { return RESP.indexOf(a.Resultado) >= 0; }).length;
  var url = ScriptApp.getService().getUrl() || '';
  var html = '<div style="font-family:Arial,sans-serif;color:#1d2b25"><h2 style="color:#1E5E4A;margin:0 0 4px">Morphê Comercial · resumo da semana</h2>' +
    '<p style="color:#6B6B6B;margin:0 0 16px">' + Utilities.formatDate(new Date(now - seteDias), 'America/Sao_Paulo', 'dd/MM') + ' a ' + Utilities.formatDate(now, 'America/Sao_Paulo', 'dd/MM/yyyy') + '</p>' +
    '<p><b>' + C.length + '</b> contatos · <b>' + resp + '</b> respostas · <b>' + novos + '</b> leads novos · <b>' + amostras + '</b> amostras · <b>' + ganhos.length + '</b> clientes ganhos' + (ganhos.length ? ' (' + ganhos.join(', ') + ')' : '') + '</p>' +
    '<table style="border-collapse:collapse;font-size:14px"><tr style="background:#1E5E4A;color:#fff"><th ' + td + '>Pessoa</th><th ' + td + '>Contatos</th><th ' + td + '>Respostas</th><th ' + td + '>Follow-ups vencidos</th><th ' + td + '>Tarefas abertas</th></tr>' + linhas + '</table>' +
    (site.length ? '<p style="color:#E5484D"><b>' + site.length + ' lead(s) do site sem resposta:</b> ' + site.map(function (l) { return l.Estabelecimento; }).join(', ') + '</p>' : '') +
    '<p style="margin-top:12px">Meta: ' + d.metaContatosSemana + ' contatos por semana · <a href="' + url + '">Abrir o Morphê Comercial</a></p></div>';
  var para = (typeof so === 'string' && so) ? [so] : d.users.map(function (u) { return u.email; }).filter(String);
  para.forEach(function (e) { MailApp.sendEmail({ to: e, subject: '[Morphê Comercial] Resumo da semana', htmlBody: html }); });
  return para.length;
}
function relatorioParaMim() { var e = me_().email; if (!e) throw new Error('Não consegui identificar seu e-mail'); relatorioSemanal(e); return e; }



// ------------------------------------------------------------------ SETUP (rodar 1 vez)
function setup() {
  var ss = ss_();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  var HDR = '#1E5E4A', FONT = 'Montserrat';
  function mk(name, cols, widths) {
    var s = ss.getSheetByName(name) || ss.insertSheet(name);
    s.clear();
    s.getRange(1, 1, 1, cols.length).setValues([cols]).setBackground(HDR).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
    s.setFrozenRows(1); s.setRowHeight(1, 34);
    if (widths) widths.forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
    s.getRange(1, 1, s.getMaxRows(), cols.length).setFontFamily(FONT);
    return s;
  }
  var L = mk('Leads', CFG.LEAD_COLS, [110, 220, 130, 130, 180, 130, 130, 120, 150, 180, 160, 120, 120, 130, 110, 130, 200, 110, 110, 110, 150, 260, 120, 120, 120]);
  var A = mk('Atividades', CFG.ATV_COLS, [110, 130, 110, 200, 100, 100, 100, 180, 380, 150]);
  var C = ss.getSheetByName('Config') || ss.insertSheet('Config'); C.clear();
  C.getRange('A1:N1').setValues([['Segmentos', 'Origens', 'Canais', 'Resultados de contato', 'Motivos de perda', 'Usuários (nome)', 'E-mail', '', 'Configuração', 'Valor', '', 'Buscas automáticas (termo)', 'Cidade', 'Segmento']])
    .setBackground(HDR).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
  function col(c, arr) { C.getRange(2, c, arr.length, 1).setValues(arr.map(function (x) { return [x]; })); }
  col(1, ['Empório / loja natural', 'Restaurante vegano/veg', 'Restaurante (geral)', 'Poke / oriental', 'Café / padaria', 'Hortifruti / mercado', 'Buffet / catering / chef', 'Distribuidor', 'Outro']);
  col(2, ['Prospecção ativa', 'Instagram (inbound)', 'Indicação', 'Google Maps', 'Pesquisa Claude', 'Evento / feira', 'Site', 'Outro']);
  col(3, ['WhatsApp', 'Instagram', 'E-mail', 'Ligação', 'Visita', 'Outro']);
  col(4, ['Enviado (aguardando)', 'Não respondeu', 'Respondeu', 'Interessado', 'Pediu valores / material', 'Pediu amostra', 'Encaminhou para outro contato/canal', 'Pediu para retornar depois', 'Sem interesse']);
  col(5, ['Sem interesse', 'Preço / margem', 'Caixa / vendas fracas', 'Não tem cozinha / não se aplica', 'Frete / logística', 'Não respondeu (esgotou tentativas)', 'Já tem fornecedor', 'Outro']);
  C.getRange('F2:G6').setValues([['Isa', 'isabelly.pereira@bioedtech.com.br'], ['Maria', ''], ['Barbara', 'barbara.santos@bioedtech.com.br'], ['Fernanda', 'fernanda.sviech@bioedtech.com.br'], ['Janaina', '']]);
  C.getRange('I2:J6').setValues([['Google Maps API Key', ''], ['Máx. sugestões por semana', 20], ['Meta: contatos por semana', 30], ['Meta: novos clientes por mês', 4], ['Dias para lead "parado"', 7]]);
  var buscas = [
    ['empório de produtos naturais', 'São Paulo, SP', 'Empório / loja natural'], ['restaurante vegano', 'São Paulo, SP', 'Restaurante vegano/veg'],
    ['empório vegano', 'São Paulo, SP', 'Empório / loja natural'], ['poke', 'São Paulo, SP', 'Poke / oriental'],
    ['empório de produtos naturais', 'Campinas, SP', 'Empório / loja natural'], ['restaurante vegano', 'Campinas, SP', 'Restaurante vegano/veg'],
    ['empório de produtos naturais', 'Ribeirão Preto, SP', 'Empório / loja natural'], ['restaurante vegano', 'Sorocaba, SP', 'Restaurante vegano/veg'],
    ['empório de produtos naturais', 'São José dos Campos, SP', 'Empório / loja natural'], ['empório de produtos naturais', 'Santos, SP', 'Empório / loja natural'],
    ['buffet vegano', 'São Paulo, SP', 'Buffet / catering / chef'], ['empório de produtos naturais', 'Jundiaí, SP', 'Empório / loja natural']];
  C.getRange(2, 12, buscas.length, 3).setValues(buscas);
  C.getRange('A2:N60').setBackground('#FFF8E1');
  C.getRange('I10').setValue('Como pegar a chave do Google Maps: console.cloud.google.com → criar projeto → ativar "Places API (New)" → Credenciais → Criar chave de API → colar em J2.').setWrap(true);
  C.setColumnWidths(1, 14, 160); C.setColumnWidth(9, 220); C.setColumnWidth(10, 260);
  C.getRange('A1:N60').setFontFamily(FONT);

  // validações na aba Leads
  function dv(colName, list) {
    var c = CFG.LEAD_COLS.indexOf(colName) + 1;
    L.getRange(2, c, 2000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build());
  }
  dv('Etapa', CFG.ETAPAS); dv('Segmento', cfgList_(1)); dv('Origem', cfgList_(2)); dv('Motivo de perda', cfgList_(5));
  dv('Responsável', users_().map(function (u) { return u.nome; }));

  var p1 = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1'); if (p1 && ss.getSheets().length > 1) ss.deleteSheet(p1);
  importarContatos_();
  importarSugestoesClaude_();
  instalarBuscaSemanal();
}

function segGuess_(n) {
  n = norm_(n);
  if (/poke|sushi|oriental|japa|asia|shiawase|jappo|marukai|suupaa|azuki|kibo|urizun|gohan|domo|towa|nami/.test(n)) return 'Poke / oriental';
  if (/hortifruti|sacolao|supermercad|mercado|frutaria|marche|hortisabor|frutas/.test(n)) return 'Hortifruti / mercado';
  if (/cafe|coffe|bakehouse|padaria|farinha|lattelie|drama/.test(n)) return 'Café / padaria';
  if (/buffet|catering|chef/.test(n)) return 'Buffet / catering / chef';
  if (/emporio|armazem|armazen|natural|cerealista|graos|grao|verde|saude|vida|campo|castanheira|nutri/.test(n)) return 'Empório / loja natural';
  if (/veg|vegan|vegetal|planta/.test(n)) return 'Restaurante vegano/veg';
  if (/restaurante|trattoria|bistro|bar|tavern|bodega|cozinha|panuzzo|sororoca|cais|insalata/.test(n)) return 'Restaurante (geral)';
  return 'Outro';
}

function importarContatos_() {
  var src = SpreadsheetApp.openById(CFG.CONTATOS_ID).getSheets()[0].getDataRange().getValues();
  src.shift();
  var clientes = ['fiftycafe', 'elmelocoton', 'manjericaopircicaba', 'emporionovossabores', 'nempeixe'];
  var L = sh_('Leads'), A = sh_('Atividades');
  var leads = [], atv = [], H = CFG.LEAD_COLS, now = new Date();
  src.forEach(function (r) {
    var nome = String(r[1] || '').trim(); if (!nome) return;
    var resp = String(r[0] || '').trim() || 'Isa';
    var canal = /insta/i.test(r[2]) ? 'Instagram' : (/whats/i.test(r[2]) ? 'WhatsApp' : String(r[2] || ''));
    var sit = String(r[3] || '').trim();
    var dt = r[4] instanceof Date ? r[4] : parseDay_(String(r[4] || '').split('/').reverse().join('-'));
    var email = (sit.match(/[\w.\-]+@[\w.\-]+\.\w+/) || [''])[0];
    var n = norm_(nome), etapa = 'Contato feito', motivo = '', res = sit ? 'Respondeu' : 'Enviado (aguardando)', prox = '', dprox = '';
    if (clientes.indexOf(n) >= 0) { etapa = 'Ganho'; res = 'Respondeu'; }
    else if (n === 'mercadinhovegano' || /amostra/i.test(sit)) { etapa = 'Amostra enviada'; }
    else if (/n[aã]o (tem|possui) interesse/i.test(sit)) { etapa = 'Perdido'; motivo = 'Sem interesse'; res = 'Sem interesse'; }
    else if (/n[aã]o tem cozinha/i.test(sit)) { etapa = 'Perdido'; motivo = 'Não tem cozinha / não se aplica'; res = 'Sem interesse'; }
    else if (/vendas fracas|caixa/i.test(sit)) { etapa = 'Respondeu'; res = 'Pediu para retornar depois'; prox = 'Retomar contato (pediram para voltar depois)'; dprox = /2 meses/.test(sit) ? new Date(now.getTime() + 50 * 864e5) : new Date(now.getTime() + 10 * 864e5); }
    else if (/analisando|valores|material|interessad|whatsapp|numero|compras/i.test(sit)) { etapa = 'Negociação'; res = /interessad/i.test(sit) ? 'Interessado' : 'Pediu valores / material'; }
    else if (email) { etapa = 'Respondeu'; res = 'Encaminhou para outro contato/canal'; }
    if (/n[aã]o respond/i.test(sit) && etapa === 'Contato feito') res = 'Não respondeu';
    var o = {}; H.forEach(function (k) { o[k] = ''; });
    o.ID = newId_('L'); o.Estabelecimento = nome; o.Segmento = segGuess_(nome); o['E-mail'] = email;
    o.Origem = (resp.indexOf('Maria') === 0 && canal === 'Instagram') ? 'Instagram (inbound)' : 'Prospecção ativa';
    o['Responsável'] = resp.replace(/\s+$/, ''); o.Etapa = etapa; o['Motivo de perda'] = motivo;
    o['Próximo passo'] = prox || (['Contato feito', 'Respondeu', 'Negociação', 'Amostra enviada'].indexOf(etapa) >= 0 ? 'Fazer follow-up' : '');
    o['Data próximo passo'] = dprox || (o['Próximo passo'] ? new Date(Math.max((dt ? dt.getTime() : now.getTime()) + 7 * 864e5, now.getTime())) : '');
    o['Último contato'] = dt || ''; o['Observações'] = sit; o['Criado em'] = dt || now; o['Atualizado em'] = now; o['Etapa desde'] = dt || now;
    leads.push(H.map(function (k) { return o[k]; }));
    // histórico: contato no canal original; se foi encaminhado p/ e-mail, 2º contato por e-mail
    var migrouEmail = !!email && /enviad|enviar|ok|feito/i.test(sit);
    atv.push([newId_('A'), dt || now, o.ID, nome, o['Responsável'], 'Contato', canal, migrouEmail ? 'Encaminhou para outro contato/canal' : res, 'Importado da planilha de contatos' + (sit ? ': ' + sit : ''), '']);
    if (migrouEmail) {
      var resEmail = /n[aã]o tem interesse/i.test(sit) ? 'Sem interesse' : 'Enviado (aguardando)';
      atv.push([newId_('A'), dt || now, o.ID, nome, o['Responsável'], 'Contato', 'E-mail', resEmail, 'E-mail enviado para ' + email, '']);
    }
  });
  if (leads.length) L.getRange(2, 1, leads.length, H.length).setValues(leads);
  if (atv.length) A.getRange(2, 1, atv.length, CFG.ATV_COLS.length).setValues(atv);
  L.getRange(2, H.indexOf('Data próximo passo') + 1, 2000, 1).setNumberFormat('dd/mm/yyyy');
  L.getRange(2, H.indexOf('Último contato') + 1, 2000, 1).setNumberFormat('dd/mm/yyyy');
  A.getRange(2, 2, 3000, 1).setNumberFormat('dd/mm/yyyy HH:mm');
}

function importarSugestoesClaude_() {
  var s = SUGESTOES_CLAUDE, H = CFG.LEAD_COLS, now = new Date(), rows = [];
  var L = sh_('Leads');
  var exist = {}; L.getRange(2, 2, Math.max(L.getLastRow() - 1, 1), 1).getValues().forEach(function (r) { exist[norm_(r[0])] = 1; });
  s.forEach(function (x) {
    if (exist[norm_(x[0])]) return;
    var o = {}; H.forEach(function (k) { o[k] = ''; });
    o.ID = newId_('L'); o.Estabelecimento = x[0]; o.Segmento = x[1]; o.Cidade = x[2]; o.Instagram = x[3]; o.WhatsApp = x[4]; o.Site = x[5];
    o.Origem = 'Pesquisa Claude'; o.Etapa = 'Sugerido'; o['Observações'] = x[6]; o['Criado em'] = now; o['Atualizado em'] = now; o['Etapa desde'] = now;
    rows.push(H.map(function (k) { return o[k]; }));
  });
  if (rows.length) L.getRange(L.getLastRow() + 1, 1, rows.length, H.length).setValues(rows);
}

// [nome, segmento, cidade, instagram, whatsapp (só se divulgado), site/fonte, observação]
var SUGESTOES_CLAUDE = [
  ['Ideal Veggie Cozinha Vegana', 'Restaurante vegano/veg', 'Campinas, SP (confirmar)', '@idealveggie', '', 'https://www.instagram.com/idealveggie/', 'Cozinha vegana. WhatsApp não divulgado na fonte.'],
  ['Empório Vegan Style', 'Empório / loja natural', '(confirmar cidade)', '@emporioveganstyle', '', 'https://www.instagram.com/emporioveganstyle/', 'Empório vegano. WhatsApp não divulgado na fonte.'],
  ['Bem Vegano Campinas', 'Restaurante vegano/veg', 'Campinas, SP', '@bemveganocampinas', '', 'https://www.instagram.com/bemveganocampinas/', 'WhatsApp não divulgado na fonte.'],
  ['Grão Empório Natural', 'Empório / loja natural', 'Campinas, SP (confirmar)', '@grao_emporionatural', '', 'https://www.instagram.com/grao_emporionatural/', 'Empório de produtos naturais.'],
  ['Espaço A Granel', 'Empório / loja natural', 'Campinas, SP', '@espacoagranel', '', 'https://www.instagram.com/espacoagranel/', 'Loja de produtos naturais em Campinas.'],
  ['Vita Natural Restaurante', 'Restaurante vegano/veg', '(confirmar cidade)', '@vitanatural.vegano', '', 'https://www.instagram.com/vitanatural.vegano/', 'Restaurante vegano.'],
  ['The Beet Ateliê Vegano', 'Restaurante vegano/veg', 'Campinas, SP', '@thebeetcampinas', '', 'https://www.instagram.com/thebeetcampinas/', 'Ateliê vegano.'],
  ['Vegano SP · Cozinha Vegetal', 'Restaurante vegano/veg', 'São Paulo, SP', '@veganosp', '', 'https://www.instagram.com/veganosp/', 'Cozinha vegetal em SP.'],
  ['Couve Flor Bistrô Veg', 'Restaurante vegano/veg', 'Sorocaba, SP', '@couveflorbistro', '', 'https://www.instagram.com/couveflorbistro/', 'Bistrô vegetariano/vegano.'],
  ['Naturall da Fazenda', 'Empório / loja natural', 'São Paulo, SP', '', '', 'https://www.facebook.com/naturalldafazenda/', 'Empório vegano/vegetariano. Instagram não confirmado.'],
  ['Com-Sciência Restaurante Vegetariano', 'Restaurante vegano/veg', 'São José dos Campos, SP', '', '', 'https://www.facebook.com/ComScienciaRestauranteVegetariano/', 'Instagram não confirmado.'],
  ["Veg'n'Rock Culinária Vegana", 'Restaurante vegano/veg', 'São José dos Campos, SP', '', '', 'https://www.facebook.com/vegnrock/', 'Instagram não confirmado.'],
  ['Quadrúpede Comida Vegana', 'Restaurante vegano/veg', 'Sorocaba, SP (confirmar)', '', '', 'https://www.facebook.com/quadrupedecomidavegana/', 'Instagram não confirmado.'],
  ['Sattwa Plant Based Bar', 'Restaurante vegano/veg', '(confirmar cidade)', '', '', 'https://wanderlog.com/place/details/11543704/', 'Restaurante e empório vegano.'],
  ['Sr. Shiitake Veg', 'Restaurante vegano/veg', 'Ribeirão Preto, SP', '', '', 'https://veganizze.com.br/locais/ribeirao-preto-sp', 'Restaurante oriental vegano.'],
  ['Q Delícia Vegan', 'Outro', 'Ribeirão Preto, SP', '', '', 'https://veganizze.com.br/locais/ribeirao-preto-sp', 'Massas congeladas veganas (delivery) — possível parceria/revenda.'],
  ['Maria Valente Delícias Veganas', 'Restaurante vegano/veg', 'Ribeirão Preto, SP', '', '', 'https://veganizze.com.br/locais/ribeirao-preto-sp', 'Restaurante vegano.'],
  ['Restaurante Vegetariano Apfel', 'Restaurante vegano/veg', 'São Paulo, SP', '', '', 'https://wanderlog.com/place/details/3683848/', 'Restaurante vegetariano tradicional em SP.']
];

function ajustarFuso() { ss_().setSpreadsheetTimeZone('America/Sao_Paulo'); }

// ------------------------------------------------------------------ AJUSTES 2 (rodar 1 vez)
function ajustes2() {
  var C = sh_('Config');
  // e-mails do time (para avisos de @menção)
  var u = C.getRange('F2:G30').getValues();
  u.forEach(function (r, i) {
    if (r[0] === 'Maria' && !r[1]) C.getRange(i + 2, 7).setValue('maria.malagutti@bioedtech.com.br');
    if (r[0] === 'Janaina' && !r[1]) C.getRange(i + 2, 7).setValue('janaina@bioedtech.com.br');
  });
  // origem "Site (formulário)"
  var orig = cfgList_(2); if (orig.indexOf('Site (formulário Wix)') < 0) C.getRange(orig.length + 2, 2).setValue('Site (formulário Wix)');

  // 1) Naturall da Fazenda: já está com a Isa -> tira dos sugeridos
  var L = sh_('Leads'), H = CFG.LEAD_COLS;
  var v = L.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (norm_(v[i][1]) === 'naturalldafazenda' && v[i][H.indexOf('Etapa')] === 'Sugerido') L.deleteRow(i + 1);

  // 2) Clientes conquistados pela Isa: Vegsim Moema, Vegsim Mocca, 4 Estrelas
  var prim = primeiroPedido_();
  [['Vegsim Moema', 'Restaurante vegano/veg', 'São Paulo, SP', ['vegsim', 'moema']],
   ['Vegsim Mocca', 'Restaurante vegano/veg', 'São Paulo, SP', ['vegsim', 'moo', 'moca', 'mocc']],
   ['4 Estrelas', '', '', ['estrela']]].forEach(function (c) {
    var dt = null;
    Object.keys(prim).forEach(function (k) {
      var ok = c[3][0] === 'estrela' ? k.indexOf('estrela') >= 0 : (k.indexOf('vegsim') >= 0 && c[3].slice(1).some(function (x) { return k.indexOf(x) >= 0; }));
      if (ok && (!dt || prim[k] < dt)) dt = prim[k];
    });
    var ex = dupDe_(c[0], '', '');
    if (ex) { var r = -1, vv = L.getDataRange().getValues(); for (var j = 1; j < vv.length; j++) if (vv[j][1] === ex) r = j + 1;
      if (r > 0) { L.getRange(r, H.indexOf('Etapa') + 1).setValue('Ganho'); L.getRange(r, H.indexOf('Responsável') + 1).setValue('Isa'); } return; }
    var when = dt || new Date();
    var id = writeLead_({ Estabelecimento: c[0], Segmento: c[1] || segGuess_(c[0]), Cidade: c[2], Origem: 'Prospecção ativa', 'Responsável': 'Isa', Etapa: 'Ganho',
      'Observações': 'Cliente ativo (conquistado pela Isa). Pedidos no Monitoramento Clientes B2B.', 'Último contato': when });
    var rr = rowOf_(L, id); L.getRange(rr, H.indexOf('Criado em') + 1).setValue(when); L.getRange(rr, H.indexOf('Etapa desde') + 1).setValue(when);
    sh_('Atividades').appendRow([newId_('A'), when, id, c[0], 'Isa', 'Etapa', '', '', '→ Ganho (primeiro pedido' + (dt ? ' em ' + Utilities.formatDate(dt, 'America/Sao_Paulo', 'dd/MM/yyyy') : '') + ')', '']);
  });

  // 3) Leads que vieram pelo formulário do site (planilha "Empresas que entraram em contato")
  importarWix_();
  SpreadsheetApp.flush();
}

function primeiroPedido_() {
  var out = {};
  try {
    var s = SpreadsheetApp.openById(CFG.MONITOR_ID).getSheetByName('Pedidos B2B');
    var v = s.getDataRange().getValues(), h = v.shift().map(function (x) { return norm_(x); });
    var iD = 0, iC = -1;
    h.forEach(function (x, i) { if (iC < 0 && x.indexOf('cliente') >= 0) iC = i; if (x === 'data' || x.indexOf('datadopedido') === 0) iD = i; });
    if (iC < 0) iC = 1;
    v.forEach(function (r) { var k = norm_(r[iC]), d = r[iD]; if (k && d instanceof Date && (!out[k] || d < out[k])) out[k] = d; });
  } catch (e) { Logger.log(e); }
  return out;
}

var WIX_ID = '1U7IEAfozfJ-jBCs9AjG3XtNA5jtw5PrKPmPwRpECY4c';
function importarWix_() {
  var v = SpreadsheetApp.openById(WIX_ID).getSheets()[0].getDataRange().getValues();
  var hr = -1; v.forEach(function (r, i) { if (hr < 0 && r.some(function (c) { return String(c).trim() === 'Nome'; })) hr = i; });
  if (hr < 0) return;
  var h = v[hr].map(function (x) { return String(x).trim(); });
  var col = function (n) { return h.indexOf(n); };
  var alias = { manjericao: 'Manjerição Pircicaba', melocoton: 'El Melocoton', fifty: 'Fifty cafe', novossabores: 'Emporio Novos sabores' };
  var L = sh_('Leads'), H = CFG.LEAD_COLS;
  v.slice(hr + 1).forEach(function (r) {
    var emp = String(r[col('Empresa')] || '').trim(); if (!emp) return;
    var nome = String(r[col('Nome')] || '').trim(), fone = String(r[col('Telefone')] || '').trim(), mail = String(r[col('Email')] || '').trim();
    var cargo = String(r[col('Cargo')] || '').trim(), interesse = String(r[col('Qual o seu interesse?')] || '').trim(), conh = String(r[col('Como você nos conheceu?')] || '').trim();
    var feito = r[col('Contato Feito?')] === true || String(r[col('Contato Feito?')]).toUpperCase() === 'TRUE';
    var alvo = ''; Object.keys(alias).forEach(function (k) { if (norm_(emp).indexOf(k) >= 0) alvo = alias[k]; });
    if (!alvo) alvo = dupDe_(emp, '', fone);
    var obs = 'Formulário do site: ' + interesse + (conh ? ' · conheceu por: ' + conh : '');
    var pessoa = nome ? nome.replace(/\b\w/g, function (c) { return c.toUpperCase(); }).replace(/\B\w/g, function (c) { return c.toLowerCase(); }) + (cargo ? ' (' + cargo + ')' : '') : '';
    if (alvo) {
      var vv = L.getDataRange().getValues(), rr = -1;
      for (var j = 1; j < vv.length; j++) if (vv[j][1] === alvo) rr = j + 1;
      if (rr < 0) return;
      var row = vv[rr - 1], set = function (k, val) { if (val && !row[H.indexOf(k)]) L.getRange(rr, H.indexOf(k) + 1).setValue(val); };
      set('Contato (pessoa)', pessoa); set('WhatsApp', fone); set('E-mail', mail);
      L.getRange(rr, H.indexOf('Origem') + 1).setValue('Site (formulário Wix)');
      var o = String(row[H.indexOf('Observações')] || ''); if (o.indexOf('Formulário do site') < 0) L.getRange(rr, H.indexOf('Observações') + 1).setValue((o ? o + ' | ' : '') + obs);
    } else {
      var id = writeLead_({ Estabelecimento: emp, Segmento: segGuess_(emp), Cidade: /^22/.test(fone.replace(/\D/g, '')) ? 'RJ (DDD 22)' : '', 'Contato (pessoa)': pessoa, WhatsApp: fone, 'E-mail': mail,
        Origem: 'Site (formulário Wix)', Etapa: feito ? 'Contato feito' : 'Mapeado', 'Responsável': proximoRodizio_(), 'Observações': obs,
        'Próximo passo': feito ? 'Fazer follow-up' : 'Responder quem entrou em contato pelo site', 'Data próximo passo': new Date() });
      logAtv_({ leadId: id, nome: emp, autor: 'Sistema', tipo: 'Sistema', texto: 'Lead recebido pelo formulário do site (' + interesse + ')' });
    }
  });
}


// ------------------------------------------------------------------ INSTALAR MELHORIAS (rodar 1 vez)
function instalarMelhorias() {
  var C = sh_('Config');
  // modelos de mensagem
  if (!C.getRange('P1').getValue()) {
    C.getRange('P1:Q1').setValues([['Modelos de mensagem (nome)', 'Texto — use {contato}, {estabelecimento}, {eu}']]).setBackground('#1E5E4A').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
    C.getRange('P2:Q6').setValues([
      ['1º contato', 'Oi{contato}! Aqui é {eu}, da Morphê 🌱 Somos o primeiro análogo vegetal de salmão em posta do Brasil. Posso mandar nossa tabela e uma amostra para o {estabelecimento}?'],
      ['Follow-up', 'Oi{contato}, tudo bem? Aqui é {eu}, da Morphê. Passando para saber se conseguiram ver nossa proposta do salmão vegetal. Posso ajudar com alguma dúvida?'],
      ['Pós-amostra', 'Oi{contato}! Conseguiram provar a amostra da Morphê? Queria muito saber o que acharam 😊'],
      ['Tabela / valores', 'Oi{contato}! Segue nossa tabela de valores e formatos (postas, cubos e desfiado). Qualquer dúvida sobre preparo ou pedido mínimo, é só me chamar. {eu} · Morphê'],
      ['Retomar contato', 'Oi{contato}, aqui é {eu}, da Morphê 🌱 Faz um tempinho que conversamos — o {estabelecimento} ainda tem interesse em testar nosso salmão vegetal?']]);
    C.getRange('P2:Q30').setBackground('#FFF8E1').setWrap(true).setFontFamily('Montserrat'); C.setColumnWidth(16, 160); C.setColumnWidth(17, 420);
  }
  // rodízio
  var cfg = C.getRange('I2:I30').getValues().map(function (r) { return String(r[0]); });
  if (cfg.indexOf('Rodízio (nomes, separados por vírgula)') < 0) {
    var row = cfg.indexOf('') + 2; C.getRange(row, 9, 1, 2).setValues([['Rodízio (nomes, separados por vírgula)', 'Isa, Maria, Barbara, Fernanda']]);
  }
  // tarefas: colunas extras em Atividades
  var A = sh_('Atividades'), h = A.getRange(1, 1, 1, A.getLastColumn()).getValues()[0];
  if (h.indexOf('Prazo') < 0) {
    var c0 = h.length + 1;
    A.getRange(1, c0, 1, 2).setValues([['Prazo', 'Feita em']]).setBackground('#1E5E4A').setFontColor('#FFFFFF').setFontWeight('bold');
    A.getRange(2, c0, 3000, 2).setNumberFormat('dd/mm/yyyy');
  }
  // gatilhos: site (1 em 1 hora) e relatório (segunda de manhã)
  ScriptApp.getProjectTriggers().forEach(function (t) { if (['sincronizarWix', 'relatorioSemanal'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sincronizarWix').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('relatorioSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  // ajustes pendentes (Vegsim, 4 Estrelas, e-mails do time, leads do site)
  ajustes2();
}

function desligarSite() { ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sincronizarWix') ScriptApp.deleteTrigger(t); }); }

// ------------------------------------------------------------------ AJUSTES 3 (rodar 1 vez): tira Empório Lagos + modelos da Isa
function ajustes3() {
  var L = sh_('Leads'), A = sh_('Atividades'), ids = [];
  var v = L.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (norm_(v[i][1]) === 'emporiolagos') { ids.push(String(v[i][0])); L.deleteRow(i + 1); }
  var a = A.getDataRange().getValues();
  for (var j = a.length - 1; j >= 1; j--) if (ids.indexOf(String(a[j][2])) >= 0 || norm_(a[j][3]) === 'emporiolagos') A.deleteRow(j + 1);
  var C = sh_('Config');
  C.getRange('P2:Q30').clearContent();
  C.getRange('P2:Q8').setValues([
    ['1º contato (Isa)', '{saudacao}, tudo bem?\nMeu nome é Isabelly, e queria aproveitar nosso contato para saber se vocês teriam interesse em trabalhar com o nosso análogo de salmão da Morphê (apresentado no vídeo abaixo). Estamos com uma demanda bem legal pelo produto e acredito que pode fazer bastante sentido para vocês também!\n\nhttps://www.instagram.com/reel/DbHVB46pSnL/?igsh=M245N216dnlpMGk5\n\nSe houver interesse, posso te apresentar as condições para fornecimento e conversamos sobre como podemos trabalhar juntos! 🧡\n\nGostaria de apresentar essa opção de ingrediente para a equipe de compras/cozinha. Vem fazendo muito sucesso em restaurantes como uma alternativa vegana!'],
    ['1º contato', '{saudacao}, tudo bem? Aqui é {eu}, da Morphê 🧡 Queria saber se vocês teriam interesse em trabalhar com o nosso análogo de salmão (apresentado no vídeo abaixo). Vem fazendo muito sucesso em restaurantes como uma alternativa vegana!\n\nhttps://www.instagram.com/reel/DbHVB46pSnL/?igsh=M245N216dnlpMGk5\n\nSe houver interesse, posso te apresentar as condições para fornecimento 🧡'],
    ['Follow-up', '{saudacao}{contato}, tudo bem? Aqui é {eu}, da Morphê 🧡 Passando para saber se conseguiram ver o vídeo do nosso análogo de salmão. Posso te mandar nossa apresentação comercial com os formatos (posta, cubos e desfiado) e as condições para o {estabelecimento}?'],
    ['Condições restaurante', 'Segue nossa apresentação comercial 🧡\nPara restaurantes: pedido mínimo de 5 unidades, 30% de desconto de 5 a 15 unidades e 35% acima de 15. Produto congelado, validade de 3 meses, entrega em até 5 dias úteis e frete grátis em São Paulo/SP.\nTemos posta (ideal para grelha e forno), cubos (pokes, bowls e espetinhos) e desfiado (recheios, patês e wraps, o mais vendido da linha). Posso te ajudar a montar o primeiro pedido?'],
    ['Condições revenda', 'Segue nossa apresentação comercial 🧡\nPara revenda: pedido mínimo de 10 unidades, com 15% de desconto de 5 a 10 unidades e 25% acima de 10, sobre o preço sugerido de varejo. Produto congelado, validade de 3 meses e entrega em até 7 dias úteis. Posso te ajudar a montar o primeiro pedido para o {estabelecimento}?'],
    ['Pós-amostra', '{saudacao}{contato}! Conseguiram provar a amostra do nosso análogo de salmão? Queria muito saber o que a equipe achou 🧡'],
    ['Retomar contato', '{saudacao}{contato}, aqui é {eu}, da Morphê 🧡 Faz um tempinho que conversamos. O {estabelecimento} ainda teria interesse em testar nosso análogo de salmão? Posso mandar uma amostra.']]);
  C.getRange('Q1').setValue('Texto — use {saudacao}, {contato}, {estabelecimento}, {eu}');
}


var INDEX_HTML_ANTIGO = "<!doctype html>\n<html lang=\"pt-BR\">\n<head>\n<meta charset=\"utf-8\">\n<base target=\"_top\">\n<link href=\"https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap\" rel=\"stylesheet\">\n<style>\n:root{--verde:#1E5E4A;--verde2:#2E7D5F;--lima:#A5C93A;--laranja:#F2572B;--creme:#FFF6E9;--linha:#EADFCB;--cinza:#6B6B6B;--txt:#1d2b25;--branco:#fff;--verm:#E5484D;--amar:#F2C94C;--azul:#4A7BF7}\n*{box-sizing:border-box}\nbody{margin:0;font-family:Montserrat,Arial,sans-serif;background:var(--creme);color:var(--txt);font-size:14px}\nheader{background:var(--verde);color:#fff;display:flex;align-items:center;gap:16px;padding:10px 20px;position:sticky;top:0;z-index:20;flex-wrap:wrap}\nheader .logo{font-weight:700;font-size:18px;letter-spacing:.3px}\nheader .logo span{color:var(--lima)}\nnav{display:flex;gap:4px;flex-wrap:wrap}\nnav button{background:transparent;border:0;color:#CFE3D6;font:600 13px Montserrat;padding:8px 12px;border-radius:8px;cursor:pointer;position:relative}\nnav button.on{background:rgba(255,255,255,.14);color:#fff}\nnav .badge{background:var(--laranja);color:#fff;border-radius:10px;padding:1px 6px;font-size:10px;margin-left:4px}\nheader .me{margin-left:auto;display:flex;align-items:center;gap:10px;font-size:14px;font-weight:700}\nheader select{font:700 15px Montserrat;border-radius:22px;border:2px solid #fff;padding:8px 14px;background:var(--lima);color:var(--verde);min-width:170px;cursor:pointer;box-shadow:0 2px 6px rgba(0,0,0,.2)}header select.vazio{background:var(--laranja);color:#fff;animation:pulse 1.5s infinite}@keyframes pulse{50%{box-shadow:0 0 0 6px rgba(242,87,43,.35)}}\nmain{padding:18px 20px;max-width:1600px;margin:0 auto}\nh2{margin:6px 0 14px;color:var(--verde);font-size:18px}\nh3{margin:0 0 10px;color:var(--laranja);font-size:13px;text-transform:uppercase;letter-spacing:.5px}\n.row{display:flex;gap:14px;flex-wrap:wrap}\n.card{background:#fff;border-radius:12px;padding:14px 16px;box-shadow:0 1px 0 var(--linha);border:1px solid var(--linha)}\n.kpi{flex:1;min-width:160px}\n.kpi .l{font-size:11px;color:var(--cinza);font-weight:700;text-transform:uppercase}\n.kpi .v{font-size:28px;font-weight:700;color:var(--verde);margin-top:4px}\n.kpi .s{font-size:11px;color:var(--cinza);margin-top:2px}\n.btn{background:var(--verde);color:#fff;border:0;border-radius:8px;padding:8px 14px;font:600 13px Montserrat;cursor:pointer}\n.btn.sec{background:#fff;color:var(--verde);border:1px solid var(--verde)}\n.btn.lar{background:var(--laranja)}\n.btn.sm{padding:5px 10px;font-size:12px}\n.btn:disabled{opacity:.5}\ninput,select,textarea{font:500 13px Montserrat;border:1px solid #d9cfbd;border-radius:8px;padding:7px 9px;background:#fff;color:var(--txt)}\ntextarea{width:100%;min-height:70px;resize:vertical}\n.filters{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;align-items:center}\n/* kanban */\n.board{display:flex;gap:12px;overflow-x:auto;padding-bottom:10px;align-items:flex-start}\n.col{background:#F4EBDA;border-radius:12px;min-width:250px;width:250px;flex-shrink:0;display:flex;flex-direction:column;max-height:calc(100vh - 190px)}\n.col .hd{padding:10px 12px;font-weight:700;font-size:13px;display:flex;justify-content:space-between;border-top:4px solid var(--c);border-radius:12px 12px 0 0}\n.col .hd small{color:var(--cinza);font-weight:600}\n.col .bd{padding:6px 8px 10px;overflow-y:auto;flex:1;min-height:60px}\n.col.over{outline:2px dashed var(--verde)}\n.lc{background:#fff;border-radius:10px;padding:9px 10px;margin-bottom:8px;cursor:pointer;border:1px solid var(--linha);box-shadow:0 1px 2px rgba(0,0,0,.04)}\n.lc:hover{border-color:var(--verde)}\n.lc.venc{border-left:3px solid var(--verm);background:#FFF8F8}\n.lc .n{font-weight:700;font-size:13px}\n.lc .m{font-size:11px;color:var(--cinza);margin-top:3px}\n.lc .f{display:flex;justify-content:space-between;align-items:center;margin-top:6px;font-size:11px}\n.av{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;background:var(--lima);color:var(--verde);font-weight:700;font-size:10px}\n.tag{display:inline-block;padding:2px 7px;border-radius:10px;font-size:10px;font-weight:700;background:#EEF5DC;color:var(--verde)}\n.tag.red{background:#F8D7DA;color:#a61b22}.tag.yel{background:#FFF3CD;color:#7a5b00}.tag.blue{background:#E3E8FF;color:#2849b8}.tag.gray{background:#eee;color:#555}\n/* tabela */\ntable{width:100%;border-collapse:collapse;background:#fff;border-radius:12px;overflow:hidden}\nth{background:var(--verde);color:#fff;text-align:left;font-size:12px;padding:8px;position:sticky;top:0}\ntd{border-bottom:1px solid var(--linha);padding:7px 8px;font-size:13px;vertical-align:top}\ntr.click:hover td{background:#FBF7EE;cursor:pointer}\n/* drawer */\n.ov{position:fixed;inset:0;background:rgba(0,0,0,.25);z-index:40;display:none}\n.drawer{position:fixed;right:0;top:0;bottom:0;width:min(560px,100%);background:#fff;z-index:50;box-shadow:-4px 0 20px rgba(0,0,0,.15);transform:translateX(100%);transition:.2s;display:flex;flex-direction:column}\n.drawer.open{transform:none}.ov.open{display:block}\n.drawer .top{background:var(--verde);color:#fff;padding:14px 18px}\n.drawer .top h2{color:#fff;margin:0 0 4px}\n.drawer .body{padding:14px 18px;overflow-y:auto;flex:1}\n.grid2{display:grid;grid-template-columns:1fr 1fr;gap:8px 10px}\n.grid2 label{font-size:11px;font-weight:700;color:var(--cinza);display:flex;flex-direction:column;gap:3px}\n.grid2 .full{grid-column:1/3}\n.acts{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}\n.acts a{text-decoration:none}\n.tl{border-left:2px solid var(--linha);margin-left:6px;padding-left:12px}\n.tl .it{margin-bottom:12px;position:relative}\n.tl .it:before{content:'';position:absolute;left:-18px;top:4px;width:10px;height:10px;border-radius:50%;background:var(--c,var(--lima))}\n.tl .h{font-size:11px;color:var(--cinza)}\n.tl .t{font-size:13px;white-space:pre-wrap}\n.box{background:#FBF7EE;border-radius:10px;padding:10px;margin:10px 0}\n.muted{color:var(--cinza);font-size:12px}\n.bar{height:26px;border-radius:6px;background:var(--verde);color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;padding:0 8px;white-space:nowrap;min-width:28px}\n.brow{display:flex;align-items:center;gap:10px;margin:6px 0}\n.brow .lb{width:170px;font-size:12px;font-weight:600;flex-shrink:0}\n.brow .rt{font-size:12px;color:var(--cinza);white-space:nowrap}\n.toast{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:var(--verde);color:#fff;padding:10px 16px;border-radius:10px;font-weight:600;z-index:90;display:none}\n.loading{position:fixed;inset:0;background:rgba(255,246,233,.7);display:flex;align-items:center;justify-content:center;z-index:80;font-weight:700;color:var(--verde)}\n.feed .it{background:#fff;border:1px solid var(--linha);border-radius:10px;padding:10px 12px;margin-bottom:8px}\n.mention{background:#FFE7DD;color:#b3360f;border-radius:4px;padding:0 3px;font-weight:700}\n.sug{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px}\n.hint{background:#EEF5DC;border-radius:10px;padding:10px 12px;font-size:12px;margin-bottom:12px}\n@media(max-width:700px){main{padding:12px}.col{min-width:220px;width:220px}.grid2{grid-template-columns:1fr}.grid2 .full{grid-column:1}.brow .lb{width:110px}}\n</style>\n</head>\n<body>\n<header>\n  <div class=\"logo\">Morphê <span>Comercial</span></div>\n  <nav id=\"nav\"></nav>\n  <div class=\"me\"><button class=\"btn lar sm\" onclick=\"quickAdd()\">+ Adicionar lead</button> 👤 Você: <select id=\"meSel\"></select></div>\n</header>\n<main id=\"main\"></main>\n<div class=\"ov\" id=\"ov\" onclick=\"closeDrawer()\"></div>\n<div class=\"drawer\" id=\"drawer\"></div>\n<div class=\"toast\" id=\"toast\"></div>\n<div class=\"loading\" id=\"loading\">Carregando…</div>\n\n<script>\nvar D = null, VIEW = 'hoje', ME = '', F = {q:'', resp:'', seg:'', orig:'', pvTodos:false}, OPEN = null, DRAG = null, PERIOD = 90;\nvar TABS = [['hoje','Hoje'],['funil','Funil'],['leads','Leads'],['sug','Sugeridos'],['buscar','Buscar no Maps'],['mural','Mural'],['dash','Dashboard'],['posVenda','Pós-venda 🔄']];\nvar COR = {'Sugerido':'#B9A6E0','Mapeado':'#9E9E9E','Contato feito':'#4A7BF7','Respondeu':'#2EA3C7','Amostra enviada':'#F2C94C','Negociação':'#F2572B','Ganho':'#2E9E5B','Perdido':'#E5484D','Frio':'#9E9E9E'};\nvar RESPONDEU = ['Respondeu','Interessado','Pediu valores / material','Pediu amostra','Encaminhou para outro contato/canal','Pediu para retornar depois','Sem interesse'];\n\nfunction $(id){return document.getElementById(id)}\nfunction esc(s){return String(s==null?'':s).replace(/[&<>\"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c]})}\nfunction toast(m){var t=$('toast');t.textContent=m;t.style.display='block';clearTimeout(t._h);t._h=setTimeout(function(){t.style.display='none'},2600)}\nfunction busy(b){$('loading').style.display=b?'flex':'none'}\nfunction call(fn){var args=[].slice.call(arguments,1);busy(true);return new Promise(function(ok,ko){google.script.run.withSuccessHandler(function(r){busy(false);ok(r)}).withFailureHandler(function(e){busy(false);alertErr(e);ko(e)})[fn].apply(null,args)})}\nfunction alertErr(e){toast('Erro: '+(e&&e.message||e))}\nfunction d(s){if(!s)return null;var x=new Date(String(s).length<=10?s+'T00:00:00':s);return isNaN(x)?null:x}\nfunction fmt(s){var x=d(s);return x?('0'+x.getDate()).slice(-2)+'/'+('0'+(x.getMonth()+1)).slice(-2):''}\nfunction fmtH(s){var x=d(s);return x?fmt(s)+' '+('0'+x.getHours()).slice(-2)+':'+('0'+x.getMinutes()).slice(-2):''}\nfunction today(){var t=new Date();t.setHours(0,0,0,0);return t}\nfunction days(s){var x=d(s);if(!x)return null;x.setHours(0,0,0,0);return Math.round((today()-x)/864e5)}\nfunction ini(n){return String(n||'?').trim().slice(0,2).toUpperCase()}\nfunction isoDay(dt){return dt.getFullYear()+'-'+('0'+(dt.getMonth()+1)).slice(-2)+'-'+('0'+dt.getDate()).slice(-2)}\nfunction lead(id){return D.leads.filter(function(l){return String(l.ID)===String(id)})[0]}\nfunction atvOf(id){return D.atividades.filter(function(a){return String(a['Lead ID'])===String(id)}).sort(function(a,b){return a.Data<b.Data?1:-1})}\nfunction ativo(l){return ['Sugerido','Ganho','Perdido','Frio'].indexOf(l.Etapa)<0}\nfunction waLink(n,msg){var p=String(n||'').replace(/\\D/g,'');if(!p)return '';if(p.length<=11)p='55'+p;return 'https://wa.me/'+p+(msg?'?text='+encodeURIComponent(msg):'')}\nfunction igLink(h){h=String(h||'').trim();if(!h)return '';if(/^http/.test(h))return h;return 'https://instagram.com/'+h.replace(/^@/,'')}\nfunction opts(list,sel,blank){return (blank!==undefined?'<option value=\"\">'+blank+'</option>':'')+list.map(function(x){return '<option'+(x===sel?' selected':'')+'>'+esc(x)+'</option>'}).join('')}\n\nfunction load(){busy(true);google.script.run.withSuccessHandler(function(r){busy(false);setData(r);render()}).withFailureHandler(function(e){busy(false);$('main').innerHTML='<div class=\"card\">Não foi possível carregar: '+esc(e.message||e)+'</div>'}).getAll()}\nfunction setData(r){D=r;var saved='';try{saved=localStorage.getItem('morphe_me')||''}catch(e){}\n  ME=(r.me&&r.me.nome)||saved||'';var s=$('meSel');s.innerHTML=opts(D.users.map(function(u){return u.nome}),ME,'👤 Escolha seu nome');s.className=ME?'':'vazio';s.onchange=function(){ME=s.value;s.className=ME?'':'vazio';try{localStorage.setItem('morphe_me',ME)}catch(e){};render()};}\nfunction after(r){setData(r);render();if(OPEN)openLead(OPEN)}\n\nfunction renderNav(){var nSug=D.leads.filter(function(l){return l.Etapa==='Sugerido'}).length;var nMe=pendentes(ME).length;\n  var nVenc=D.leads.filter(function(l){var x=days(l['Data próximo passo']);return ativo(l)&&x!==null&&x>0&&(!ME||l['Responsável']===ME)}).length;\n  var nPV=D.leads.filter(function(l){if(l.Etapa!=='Ganho')return false;var ds=diasSemContato(l);var vp=days(l['Data próximo passo']);return (vp!==null&&vp>0)||(ds===null||ds>=30);}).length;\n  $('nav').innerHTML=TABS.map(function(t){var b='';if(t[0]==='sug'&&nSug)b='<span class=\"badge\">'+nSug+'</span>';if(t[0]==='hoje'&&nVenc)b='<span class=\"badge\" style=\"background:var(--verm)\">'+nVenc+' venc.</span>';else if(t[0]==='hoje'&&nMe)b='<span class=\"badge\">'+nMe+'</span>';else if(t[0]==='posVenda'&&nPV)b='<span class=\"badge\" style=\"background:var(--laranja)\">'+nPV+'</span>';return '<button class=\"'+(VIEW===t[0]?'on':'')+'\" onclick=\"go(\\''+t[0]+'\\')\">'+t[1]+b+'</button>'}).join('')}\nfunction go(v){VIEW=v;render();window.scrollTo(0,0)}\nfunction render(){if(!D)return;renderNav();({hoje:vHoje,funil:vFunil,leads:vLeads,sug:vSug,buscar:vBuscar,mural:vMural,dash:vDash,posVenda:vPosVenda})[VIEW]()}\n\nfunction filtered(){var q=F.q.toLowerCase();return D.leads.filter(function(l){return (!q||(l.Estabelecimento+' '+l.Cidade+' '+l.Instagram+' '+l.Observações).toLowerCase().indexOf(q)>=0)&&(!F.resp||(F.resp==='__sem'?!l['Responsável']:l['Responsável']===F.resp))&&(!F.seg||l.Segmento===F.seg)&&(!F.orig||l.Origem===F.orig)})}\nfunction pessoaChips(){var base=D.leads.filter(function(l){return l.Etapa!=='Sugerido'});function n(p){return base.filter(function(l){return p==='__sem'?!l['Responsável']:(!p||l['Responsável']===p)}).length}\n  var ps=[['','Todos']].concat(D.users.map(function(u){return [u.nome,u.nome]})).concat([['__sem','Sem dono']]);\n  return '<div class=\"acts\" style=\"margin:0 0 10px\">'+ps.map(function(p){var on=(F.resp||'')===p[0];return '<button class=\"btn sm '+(on?'':'sec')+'\" onclick=\"F.resp=\\''+p[0]+'\\';render()\">'+(p[0]&&p[0]!=='__sem'?'<span class=\"av\" style=\"width:18px;height:18px;font-size:9px;margin-right:4px\">'+ini(p[1])+'</span>':'')+esc(p[1])+' <b>'+n(p[0])+'</b></button>'}).join('')+(semDono().length?'<button class=\"btn sm lar\" onclick=\"distribuirSemDono()\">Distribuir os '+semDono().length+' sem dono</button>':'')+'</div>'}\nfunction filterBar(extra){return pessoaChips()+'<div class=\"filters\"><input placeholder=\"🔎 Buscar…\" value=\"'+esc(F.q)+'\" oninput=\"F.q=this.value;clearTimeout(window._q);window._q=setTimeout(render,250)\" style=\"min-width:220px\">'+\n '<select onchange=\"F.resp=this.value;render()\">'+opts(D.users.map(function(u){return u.nome}),F.resp,'Todos os responsáveis')+'</select>'+\n '<select onchange=\"F.seg=this.value;render()\">'+opts(D.segmentos,F.seg,'Todos os segmentos')+'</select>'+\n '<select onchange=\"F.orig=this.value;render()\">'+opts(D.origens,F.orig,'Todas as origens')+'</select>'+(extra||'')+'</div>'}\n\nfunction parado(l){var lim=7;var u=days(l['Último contato']||l['Etapa desde']);return ativo(l)&&u!==null&&u>=lim}\nfunction diasSemContato(l){var u=D.atividades.filter(function(a){return String(a['Lead ID'])===String(l.ID)&&a.Tipo==='Contato'}).sort(function(a,b){return a.Data<b.Data?1:-1});return u.length?days(u[0].Data):null}\nfunction frioInfo(l){var atv=D.atividades.filter(function(a){return String(a['Lead ID'])===String(l.ID)&&a.Tipo==='Contato'});return {n:atv.length,dias:atv.length?days(atv[0].Data):null}}\nfunction card(l){\n  // Card especial para leads Frios\n  if(l.Etapa==='Frio'){var fi=frioInfo(l);var dsc2=fi.dias!==null?fi.dias+'d sem resposta':'';\n    return '<div class=\"lc\" style=\"background:#f5f5f5;border-color:#ccc;opacity:.85\" onclick=\"openLead(\\''+l.ID+'\\')\"><div class=\"n\" style=\"color:#555\">❄️ '+esc(l.Estabelecimento)+'</div><div class=\"m\">'+esc([l.Segmento,l.Cidade].filter(String).join(' · '))+'</div>'+\n    '<div class=\"f\"><span class=\"av\" style=\"background:#ccc;color:#555\" title=\"'+esc(l['Responsável'])+'\">'+ini(l['Responsável'])+'</span><span class=\"tag gray\">'+fi.n+' tentativas · '+dsc2+'</span></div></div>'}\n  var dp=days(l['Data próximo passo']);var venc=dp!==null&&dp>0&&ativo(l);var tag='';\n  if(dp!==null&&ativo(l)){if(dp>0)tag='<span class=\"tag red\">atrasado '+dp+'d</span>';else if(dp===0)tag='<span class=\"tag yel\">hoje</span>';else tag='<span class=\"tag gray\">'+fmt(l['Data próximo passo'])+'</span>'}\n  else if(parado(l))tag='<span class=\"tag yel\">parado</span>';\n  var sh=slaSite(l);if(sh!==null)tag='<span class=\"tag '+(sh>=24?'red':'yel')+'\">site · '+(sh>=24?Math.floor(sh/24)+'d sem resposta':sh+'h')+'</span>';\n  var dsc=ativo(l)?diasSemContato(l):null;var dscTag=(dsc!==null&&dsc>14)?'<span class=\"tag gray\" style=\"margin-left:4px\">'+dsc+'d s/contato</span>':'';\n  return '<div class=\"lc'+(venc?' venc':'')+'\" draggable=\"true\" ondragstart=\"DRAG=\\''+l.ID+'\\'\" onclick=\"openLead(\\''+l.ID+'\\')\"><div class=\"n\">'+esc(l.Estabelecimento)+'</div><div class=\"m\">'+esc([l.Segmento,l.Cidade].filter(String).join(' · '))+'</div>'+\n  (l['Próximo passo']&&ativo(l)?'<div class=\"m\">➡ '+esc(l['Próximo passo'])+'</div>':'')+\n  '<div class=\"f\"><span class=\"av\" title=\"'+esc(l['Responsável'])+'\">'+ini(l['Responsável'])+'</span>'+tag+dscTag+'</div></div>'}\n\n// ---------------- HOJE\nfunction pendentes(nome){var t=today();return D.leads.filter(function(l){var x=days(l['Data próximo passo']);return ativo(l)&&x!==null&&x>=0&&(!nome||l['Responsável']===nome)})}\nfunction vHoje(){var so=F.hojeTodos?'':ME;\n  var pend=pendentes(so).sort(function(a,b){return a['Data próximo passo']<b['Data próximo passo']?-1:1});\n  var semana=D.leads.filter(function(l){var x=days(l['Data próximo passo']);return ativo(l)&&x!==null&&x<0&&x>=-7&&(!so||l['Responsável']===so)});\n  var par=D.leads.filter(function(l){return parado(l)&&!l['Data próximo passo']&&(!so||l['Responsável']===so)});\n  var men=D.atividades.filter(function(a){return a.Tipo!=='Tarefa'&&ME&&String(a.Menções).split(', ').indexOf(ME)>=0}).sort(function(a,b){return a.Data<b.Data?1:-1}).slice(0,8);\n  function lista(arr,vazio){return arr.length?arr.map(card).join(''):'<div class=\"muted\">'+vazio+'</div>'}\n  $('main').innerHTML='<div class=\"filters\"><h2 style=\"margin:0\">Bom dia'+(ME?', '+esc(ME):'')+' 👋</h2><label class=\"muted\" style=\"margin-left:auto\"><input type=\"checkbox\" '+(F.hojeTodos?'checked':'')+' onchange=\"F.hojeTodos=this.checked;render()\"> ver o time todo</label><button class=\"btn\" onclick=\"newLead()\">+ Novo lead</button></div>'+\n   '<div class=\"row\"><div class=\"card kpi\"><div class=\"l\">Para fazer hoje</div><div class=\"v\">'+pend.length+'</div><div class=\"s\">follow-ups vencidos ou de hoje</div></div>'+\n   '<div class=\"card kpi\"><div class=\"l\">Próximos 7 dias</div><div class=\"v\">'+semana.length+'</div></div>'+\n   '<div class=\"card kpi\"><div class=\"l\">Leads parados</div><div class=\"v\">'+par.length+'</div><div class=\"s\">sem contato há 7+ dias e sem próximo passo</div></div>'+\n   '<div class=\"card kpi\"><div class=\"l\">Sugeridos p/ triar</div><div class=\"v\">'+D.leads.filter(function(l){return l.Etapa==='Sugerido'}).length+'</div></div></div>'+\n   '<div class=\"row\" style=\"margin-top:14px;align-items:flex-start\"><div class=\"card\" style=\"flex:2;min-width:300px\"><h3>🔥 Fazer hoje</h3><div class=\"sug\">'+lista(pend,'Nada vencido. 🎉')+'</div></div>'+\n   '<div style=\"flex:1;min-width:280px;display:flex;flex-direction:column;gap:14px\">'+cardTarefas(so)+cardSite(so)+'<div class=\"card\"><h3>💬 Menções para você</h3><div class=\"feed\">'+(men.length?men.map(feedItem).join(''):'<div class=\"muted\">Sem menções.</div>')+'</div></div>'+\n   '<div class=\"card\"><h3>📅 Próximos 7 dias</h3>'+lista(semana,'Nada agendado.')+'</div><div class=\"card\"><h3>💤 Parados</h3>'+lista(par.slice(0,10),'Nenhum.')+'</div></div></div>'}\n\n// ---------------- FUNIL (kanban)\nfunction vFunil(){var ls=filtered();var et=D.etapas.filter(function(e){return e!=='Sugerido'});\n  $('main').innerHTML=filterBar('<button class=\"btn\" onclick=\"newLead()\" style=\"margin-left:auto\">+ Novo lead</button>')+'<div class=\"board\">'+et.map(function(e){var arr=ls.filter(function(l){return l.Etapa===e});\n    if(e==='Perdido'||e==='Ganho')arr=arr.sort(function(a,b){return a['Atualizado em']<b['Atualizado em']?1:-1});\n    return '<div class=\"col\" style=\"--c:'+COR[e]+'\" ondragover=\"event.preventDefault();this.classList.add(\\'over\\')\" ondragleave=\"this.classList.remove(\\'over\\')\" ondrop=\"this.classList.remove(\\'over\\');drop(\\''+e+'\\')\"><div class=\"hd\">'+e+' <small>'+arr.length+'</small></div><div class=\"bd\">'+arr.slice(0,80).map(card).join('')+(arr.length>80?'<div class=\"muted\">+'+(arr.length-80)+' (use a busca)</div>':'')+'</div></div>'}).join('')+'</div>'}\nfunction drop(etapa){var id=DRAG;DRAG=null;if(!id)return;var l=lead(id);if(!l||l.Etapa===etapa)return;\n  if(etapa==='Perdido'){var m=prompt('Motivo da perda:\\n'+D.motivos.map(function(x,i){return (i+1)+'. '+x}).join('\\n'),'1');if(m===null)return;var mm=D.motivos[(+m||1)-1]||m;call('moveLead',id,etapa,mm).then(after);return}\n  call('moveLead',id,etapa,'').then(function(r){after(r);toast(l.Estabelecimento+' → '+etapa+(etapa==='Ganho'?' 🎉 (enviado para o dashboard de clientes)':''))})}\n\n// ---------------- LEADS (tabela)\nfunction vLeads(){var ls=filtered().sort(function(a,b){return a.Estabelecimento.localeCompare(b.Estabelecimento)});\n  $('main').innerHTML=filterBar('<select onchange=\"F.et=this.value;render()\">'+opts(D.etapas,F.et,'Todas as etapas')+'</select><button class=\"btn\" onclick=\"newLead()\" style=\"margin-left:auto\">+ Novo lead</button>')+\n  '<div class=\"muted\" style=\"margin-bottom:6px\">'+ls.filter(function(l){return !F.et||l.Etapa===F.et}).length+' leads</div><div style=\"overflow:auto\"><table><tr><th>Estabelecimento</th><th>Etapa</th><th>Segmento</th><th>Cidade</th><th>Resp.</th><th>Último contato</th><th>Próximo passo</th><th>Contato</th></tr>'+\n  ls.filter(function(l){return !F.et||l.Etapa===F.et}).map(function(l){return '<tr class=\"click\" onclick=\"openLead(\\''+l.ID+'\\')\"><td><b>'+esc(l.Estabelecimento)+'</b></td><td><span class=\"tag\" style=\"background:'+COR[l.Etapa]+'22;color:'+COR[l.Etapa]+'\">'+esc(l.Etapa)+'</span></td><td>'+esc(l.Segmento)+'</td><td>'+esc(l.Cidade)+'</td><td>'+esc(l['Responsável'])+'</td><td>'+fmt(l['Último contato'])+'</td><td>'+esc(l['Próximo passo'])+(l['Data próximo passo']?' · '+fmt(l['Data próximo passo']):'')+'</td><td>'+(l.WhatsApp?'📱 ':'')+(l.Instagram?'📷 ':'')+(l['E-mail']?'✉️':'')+'</td></tr>'}).join('')+'</table></div>'}\n\n// ---------------- SUGERIDOS\nfunction vSug(){var ls=D.leads.filter(function(l){return l.Etapa==='Sugerido'});\n  $('main').innerHTML='<h2>Sugeridos para prospectar ('+ls.length+')</h2><div class=\"hint\">Leads encontrados pela pesquisa semanal e pelo Google Maps. <b>Aceitar</b> manda para \"Mapeado\" com você como responsável; <b>Descartar</b> tira da lista. Quando a sugestão parece com um lead que você já tem, aparece um aviso amarelo. Telefone/WhatsApp só aparecem quando divulgados publicamente.</div>'+\n  (ls.length?'<div class=\"sug\">'+ls.map(function(l){return '<div class=\"card\"><div style=\"font-weight:700;font-size:15px\">'+esc(l.Estabelecimento)+'</div><div class=\"muted\">'+esc([l.Segmento,l.Cidade].filter(String).join(' · '))+'</div><div class=\"muted\" style=\"margin:4px 0\">'+esc(l.Bairro||'')+'</div><div class=\"muted\">'+esc(l.Observações||'')+'</div>'+\n   '<div class=\"acts\">'+(l.Instagram?'<a class=\"btn sec sm\" target=\"_blank\" href=\"'+igLink(l.Instagram)+'\">📷 '+esc(l.Instagram)+'</a>':'')+(l.WhatsApp?'<a class=\"btn sec sm\" target=\"_blank\" href=\"'+waLink(l.WhatsApp)+'\">📱 WhatsApp</a>':'')+(l.Telefone?'<span class=\"tag gray\">☎ '+esc(l.Telefone)+'</span>':'')+(l.Site?'<a class=\"btn sec sm\" target=\"_blank\" href=\"'+esc(l.Site)+'\">🌐 site</a>':'')+(l['Google Maps']?'<a class=\"btn sec sm\" target=\"_blank\" href=\"'+esc(l['Google Maps'])+'\">📍 Maps</a>':'')+'</div>'+\n   (function(){var p=parecidos(l);return p.length?'<div class=\"hint\" style=\"background:#FFF3CD;margin:8px 0\">⚠️ Parecido com o que você já tem: <b>'+p.map(function(x){return esc(x.Estabelecimento)+' ('+esc(x.Etapa)+')'}).join(', ')+'</b><br><button class=\"btn sm\" style=\"margin-top:6px;background:#7a5b00\" onclick=\"removerSug(\\''+l.ID+'\\')\">É o mesmo — remover sugestão</button></div>':''})()+\n   '<div class=\"acts\"><button class=\"btn sm\" onclick=\"aceitar(\\''+l.ID+'\\')\">✓ Aceitar</button><button class=\"btn sec sm\" onclick=\"openLead(\\''+l.ID+'\\')\">Editar</button><button class=\"btn sec sm\" style=\"color:var(--verm);border-color:var(--verm)\" onclick=\"descartar(\\''+l.ID+'\\')\">Descartar</button></div></div>'}).join('')+'</div>':'<div class=\"card muted\">Nenhuma sugestão pendente. Use a aba \"Buscar no Maps\" ou aguarde a busca de segunda-feira.</div>')}\nfunction aceitar(id){if(!ME){toast('Escolha seu nome no topo');return}call('bulkMove',[id],'Mapeado',ME).then(function(r){after(r);toast('Aceito ✓')})}\nfunction descartar(id){call('moveLead',id,'Perdido','Descartado na triagem').then(after)}\n\n// ---------------- BUSCAR NO MAPS\nvar RES=[];\nfunction vBuscar(){var chips=['empório de produtos naturais','restaurante vegano','empório vegano','poke','buffet vegano','cafeteria vegana','hortifruti','mercado natural','loja de produtos naturais','restaurante plant-based','bistrô vegano','food truck vegano','padaria vegana'];\n  $('main').innerHTML='<h2>Buscar estabelecimentos no Google Maps</h2>'+(D.temChaveMaps?'':'<div class=\"hint\">⚠️ Para usar a busca, cole a chave do Google Maps (Places API) na aba <b>Config</b> da planilha, célula J2. O passo a passo está na célula I10.</div>')+\n  '<div class=\"card\"><div class=\"filters\"><input id=\"bq\" placeholder=\"O que buscar (ex.: empório natural)\" style=\"min-width:260px\" value=\"'+esc(F.bq||'empório de produtos naturais')+'\"><input id=\"bc\" placeholder=\"Cidade (ex.: Campinas, SP)\" value=\"'+esc(F.bc||'São Paulo, SP')+'\"><select id=\"bs\">'+opts(D.segmentos,F.bs||'Empório / loja natural')+'</select>'+\n  '<select id=\"brad\" title=\"Raio de busca\"><option value=\"1000\">1 km</option><option value=\"2000\">2 km</option><option value=\"5000\" selected>5 km</option><option value=\"10000\">10 km</option><option value=\"20000\">20 km</option><option value=\"50000\">50 km</option></select>'+\n  '<button class=\"btn\" onclick=\"buscar()\">Buscar</button></div>'+\n  '<div class=\"acts\">'+chips.map(function(c){return '<button class=\"btn sec sm\" onclick=\"$(\\'bq\\').value=\\''+c+'\\';$(\\'bq\\').focus()\">'+c+'</button>'}).join('')+'</div>'+\n  '<div class=\"muted\" style=\"margin-top:6px\">💡 Dica: busque bairro por bairro para encontrar mais resultados. Ex.: \"Vila Madalena, SP\" depois \"Pinheiros, SP\"</div></div>'+\n  '<div id=\"bres\" style=\"margin-top:14px\">'+resTable()+'</div>'}\nfunction buscar(){F.bq=$('bq').value;F.bc=$('bc').value;F.bs=$('bs').value;call('searchPlaces',F.bq,F.bc).then(function(r){RES=r;$('bres').innerHTML=resTable()})}\nfunction resTable(){if(!RES.length)return '';var novos=RES.filter(function(p){return !p.jaExiste}).length;\n  return '<div class=\"filters\"><b>'+RES.length+' resultados · '+novos+' novos</b><button class=\"btn lar\" style=\"margin-left:auto\" onclick=\"importar()\">Adicionar selecionados aos Sugeridos</button></div><div style=\"overflow:auto\"><table><tr><th><input type=\"checkbox\" onchange=\"document.querySelectorAll(\\'.pk\\').forEach(function(c){if(!c.disabled)c.checked=event.target.checked})\"></th><th>Nome</th><th>Endereço</th><th>Telefone</th><th>Site</th><th>Nota</th></tr>'+\n  RES.map(function(p,i){return '<tr><td><input type=\"checkbox\" class=\"pk\" data-i=\"'+i+'\" '+(p.jaExiste?'disabled':'checked')+'></td><td><b>'+esc(p.nome)+'</b>'+(p.jaExiste?' <span class=\"tag gray\">já está no CRM</span>':'')+'<div class=\"muted\">'+esc(p.tipo)+'</div></td><td>'+esc(p.endereco)+'</td><td>'+esc(p.telefone||'—')+'</td><td>'+(p.site?'<a target=\"_blank\" href=\"'+esc(p.site)+'\">site</a>':'—')+' · <a target=\"_blank\" href=\"'+esc(p.maps)+'\">maps</a></td><td>'+(p.nota?p.nota+' ⭐ ('+p.avaliacoes+')':'—')+'</td></tr>'}).join('')+'</table></div>'}\nfunction importar(){var sel=[].slice.call(document.querySelectorAll('.pk:checked')).map(function(c){return RES[+c.dataset.i]});if(!sel.length){toast('Nada selecionado');return}\n  call('importPlaces',sel,F.bs,F.bc).then(function(r){RES=RES.map(function(p){if(sel.indexOf(p)>=0)p.jaExiste=true;return p});after(r);toast(sel.length+' adicionados aos Sugeridos')})}\n\n// ---------------- MURAL\nfunction feedItem(a){var l=a['Lead ID']?lead(a['Lead ID']):null;var txt=esc(a.Texto).replace(/@([A-Za-zÀ-ú]+)/g,'<span class=\"mention\">@$1</span>');\n  var ic={'Contato':'📞','Comentário':'💬','Etapa':'➡️','Sistema':'⚙️','Tarefa':'✅'}[a.Tipo]||'•';\n  return '<div class=\"it\"><div class=\"h muted\">'+ic+' <b>'+esc(a.Autor)+'</b> · '+fmtH(a.Data)+(l?' · <a href=\"#\" onclick=\"openLead(\\''+l.ID+'\\');return false\">'+esc(l.Estabelecimento)+'</a>':'')+(a.Canal?' · '+esc(a.Canal):'')+(a.Resultado?' · <b>'+esc(a.Resultado)+'</b>':'')+'</div><div class=\"t\" style=\"margin-top:4px\">'+txt+'</div></div>'}\nfunction vMural(){\n  var filtMural=window._filtMural||'',filtTipo=window._filtTipo||'';\n  var arr=D.atividades.slice().sort(function(a,b){return a.Data<b.Data?1:-1});\n  if(filtMural)arr=arr.filter(function(a){return a.Autor===filtMural});\n  if(filtTipo)arr=arr.filter(function(a){return a.Tipo===filtTipo});\n  arr=arr.slice(0,200);\n  $('main').innerHTML='<h2>Mural do time</h2><div class=\"card\" style=\"margin-bottom:14px\"><textarea id=\"mtxt\" placeholder=\"Escreva para o time… use @Nome para avisar alguém (ex.: @Maria)\"></textarea><div class=\"filters\" style=\"margin:8px 0 0\"><span class=\"muted\">'+D.users.map(function(u){return '@'+u.nome}).join(' ')+'</span><button class=\"btn\" style=\"margin-left:auto\" onclick=\"postMural()\">Publicar</button></div></div>'+\n  '<div class=\"filters\" style=\"margin-bottom:12px\"><select onchange=\"window._filtMural=this.value;vMural()\"><option value=\"\">👥 Toda a equipe</option>'+D.users.map(function(u){return '<option'+(filtMural===u.nome?' selected':'')+' value=\"'+esc(u.nome)+'\">'+esc(u.nome)+'</option>'}).join('')+'</select>'+\n  '<select onchange=\"window._filtTipo=this.value;vMural()\"><option value=\"\">Todos os tipos</option>'+['Contato','Comentário','Etapa','Tarefa','Sistema'].map(function(t){return '<option'+(filtTipo===t?' selected':'')+' value=\"'+esc(t)+'\">'+esc(t)+'</option>'}).join('')+'</select>'+\n  '<span class=\"muted\" style=\"margin-left:auto\">'+arr.length+' registros</span></div>'+\n  '<div class=\"feed\">'+arr.map(feedItem).join('')+'</div>'}\nfunction mentionsIn(t){return D.users.map(function(u){return u.nome}).filter(function(n){return new RegExp('@'+n+'\\\\b','i').test(t)})}\nfunction postMural(){var t=$('mtxt').value.trim();if(!t)return;if(!ME){toast('Escolha seu nome no topo');return}call('addActivity',{tipo:'Comentário',texto:t,mencoes:mentionsIn(t)}).then(function(r){after(r);toast('Publicado')})}\n\n// ---------------- DUPLICADOS (comparação)\nvar GEN=['restaurante','emporio','emporium','cozinha','vegana','vegano','vegan','veg','vegetariano','vegetariana','vegetal','natural','naturais','produtos','cafe','cafeteria','bistro','bar','loja','mercado','supermercado','supermercados','hortifruti','comida','culinaria','plant','based','delicias','casa','espaco','store','food','foods','market','saudavel','saudaveis','granel','atelie','organicos','organico','sao','paulo','campinas','maria','santa','verde','vida','grao','graos','terra','mundo','sabor','sabores'];\nfunction nrm(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'')}\nfunction toks(s){return nrm(s).replace(/[^a-z0-9 ]/g,' ').split(/\\s+/).map(function(w){return w.replace(/(.)\\1+/g,'$1')}).filter(function(w){return w.length>=4&&GEN.indexOf(w)<0})}\nfunction cmp(s){return nrm(s).replace(/[^a-z0-9]/g,'').replace(/(.)\\1+/g,'$1')}\nfunction igH(h){h=nrm(h).trim();var m=h.match(/instagram\\.com\\/([^/?#\\s]+)/);if(m)h=m[1];return h.replace(/^@/,'')}\nfunction telD(t){var x=String(t||'').replace(/\\D/g,'');if(x.length>11&&x.indexOf('55')===0)x=x.slice(2);return x.length>=10?x:''}\nfunction parecidos(o,ignId){var n=cmp(o.Estabelecimento),tk=toks(o.Estabelecimento),g=igH(o.Instagram),f=telD(o.WhatsApp||o.Telefone);\n  return D.leads.filter(function(l){if(l.ID===(ignId||o.ID)||l.Etapa==='Sugerido'&&o.Etapa==='Sugerido')return false;var m=cmp(l.Estabelecimento);\n    if(n&&m&&(n===m||(n.length>=6&&m.indexOf(n)>=0)||(m.length>=6&&n.indexOf(m)>=0)))return true;\n    if(g&&igH(l.Instagram)===g)return true;if(f&&(telD(l.WhatsApp)===f||telD(l.Telefone)===f))return true;\n    var t2=toks(l.Estabelecimento);return tk.some(function(w){return t2.indexOf(w)>=0})}).slice(0,4)}\nfunction removerSug(id){call('removerSugerido',id).then(function(r){after(r);toast('Sugestão removida')})}\n\n// ---------------- ADIÇÃO RÁPIDA\nvar QA_MODE='um';\n/* entende o que foi colado: @insta, link do insta, telefone, e-mail, site */\nfunction parseContato(s){var o={};s=String(s||'');var tm=s.match(/(\\+?55\\s?)?\\(?\\d{2}\\)?[\\s.-]?9?\\d{4}[\\s.-]?\\d{4}/);if(tm&&telD(tm[0])){o.WhatsApp=tm[0].trim();s=s.replace(tm[0],' ')}s.split(/[\\s,;|]+/).forEach(function(p){if(!p)return;\n  if(/instagram\\.com\\//i.test(p)||/^@[\\w.]+$/.test(p))o.Instagram='@'+igH(p);\n  else if(/^[\\w.+-]+@[\\w-]+\\.[\\w.]+$/.test(p))o['E-mail']=p;\n  else if(/^https?:\\/\\//i.test(p)||/^www\\./i.test(p))o.Site=p;\n  else if(telD(p))o.WhatsApp=p});return o}\nfunction parseLinha(line){var parts=line.split(/\\t|;|,(?![^(]*\\))| - | – /).map(function(x){return x.trim()}).filter(String);if(!parts.length)return null;\n  var o={Estabelecimento:''},resto=[];parts.forEach(function(p){var c=parseContato(p);if(Object.keys(c).length&&p.split(/\\s+/).length<=3){Object.keys(c).forEach(function(k){o[k]=c[k]})}else resto.push(p)});\n  if(!resto.length&&o.Instagram)resto.push(o.Instagram.slice(1));o.Estabelecimento=resto.shift()||'';if(resto.length)o.Cidade=resto.shift();if(resto.length)o['Observações']=resto.join(' · ');return o.Estabelecimento?o:null}\nfunction quickAdd(){OPEN=null;var dr=$('drawer');var et=['Mapeado','Contato feito','Respondeu','Amostra enviada','Negociação'];\n  dr.innerHTML='<div class=\"top\"><h2>Adicionar lead</h2><div style=\"font-size:12px;opacity:.85\">Rápido: só o nome já basta. O sistema avisa se o lead já existe.</div></div><div class=\"body\">'+\n  '<div class=\"acts\"><button class=\"btn sm '+(QA_MODE==='um'?'':'sec')+'\" onclick=\"QA_MODE=\\'um\\';quickAdd()\">Um lead</button><button class=\"btn sm '+(QA_MODE==='lista'?'':'sec')+'\" onclick=\"QA_MODE=\\'lista\\';quickAdd()\">Colar uma lista</button></div>'+\n  (QA_MODE==='um'?\n   '<div class=\"grid2\"><label class=\"full\">Nome do estabelecimento *<input id=\"qNome\" oninput=\"qaCheck()\" placeholder=\"ex.: Empório Verde Vida\"></label>'+\n   '<label class=\"full\">Instagram, WhatsApp, e-mail ou link (cole qualquer um)<input id=\"qCont\" oninput=\"qaCheck()\" placeholder=\"@emporioverde  ou  11 99999-9999  ou  link do Instagram\"></label>'+\n   '<div class=\"full muted\" id=\"qParse\"></div><div class=\"full\" id=\"qDup\"></div>'+\n   '<label>Cidade<input id=\"qCid\" placeholder=\"São Paulo, SP\"></label><label>Segmento<select id=\"qSeg\">'+opts(D.segmentos,'','(adivinhar pelo nome)')+'</select></label>'+\n   '<label>Situação<select id=\"qEt\">'+opts(et,'Mapeado')+'</select></label><label>Origem<select id=\"qOri\">'+opts(D.origens,'Prospecção ativa')+'</select></label>'+\n   '<label>Responsável<select id=\"qResp\">'+opts(D.users.map(function(u){return u.nome}),ME,'—')+'</select></label><label>Canal do 1º contato<select id=\"qCan\">'+opts(D.canais,'Instagram')+'</select></label>'+\n   '<label class=\"full\">Observação<input id=\"qObs\"></label></div>'+\n   '<div class=\"acts\"><button class=\"btn\" onclick=\"qaSalvar(false)\">Salvar</button><button class=\"btn sec\" onclick=\"qaSalvar(true)\">Salvar e adicionar outro</button><a href=\"#\" class=\"muted\" style=\"align-self:center\" onclick=\"newLeadFull();return false\">formulário completo</a></div>'\n  :\n   '<div class=\"hint\">Cole <b>um lead por linha</b>. Pode ser só o nome, ou nome + @instagram/telefone + cidade, separados por vírgula ou ponto e vírgula. Dá para copiar direto de uma planilha.<br><span class=\"muted\">Ex.: <code>Empório Verde, @emporioverde, Campinas</code> · <code>Casa Vegana; 11 98888-7777</code> · <code>instagram.com/raizesveg</code></span></div>'+\n   '<textarea id=\"qLista\" style=\"min-height:180px\" oninput=\"qaPrev()\" placeholder=\"Empório Verde, @emporioverde, Campinas&#10;Casa Vegana; 11 98888-7777\"></textarea>'+\n   '<div class=\"grid2\" style=\"margin-top:8px\"><label>Situação de todos<select id=\"qEt\">'+opts(et,'Mapeado')+'</select></label><label>Origem<select id=\"qOri\">'+opts(D.origens,'Prospecção ativa')+'</select></label>'+\n   '<label>Responsável<select id=\"qResp\">'+opts(D.users.map(function(u){return u.nome}),ME,'—')+'</select></label><label>Canal do 1º contato<select id=\"qCan\">'+opts(D.canais,'Instagram')+'</select></label></div>'+\n   '<div id=\"qPrev\" style=\"margin-top:10px\"></div><div class=\"acts\"><button class=\"btn\" onclick=\"qaLista()\">Adicionar todos</button></div>')+\n  '<div class=\"acts\"><button class=\"btn sec\" onclick=\"closeDrawer()\">Fechar</button></div></div>';\n  dr.classList.add('open');$('ov').classList.add('open');setTimeout(function(){var f=$('qNome')||$('qLista');if(f)f.focus()},250)}\nfunction qaCheck(){var c=parseContato($('qCont').value);var lab={Instagram:'📷',WhatsApp:'📱','E-mail':'✉️',Site:'🌐'};\n  $('qParse').innerHTML=Object.keys(c).map(function(k){return lab[k]+' '+esc(c[k])}).join(' · ');\n  if(c.Instagram&&!$('qNome').value.trim())$('qNome').placeholder=c.Instagram.slice(1);\n  var nome=$('qNome').value.trim()||(c.Instagram?c.Instagram.slice(1):'');var p=nome.length>=3||c.Instagram||c.WhatsApp?parecidos({Estabelecimento:nome,Instagram:c.Instagram,WhatsApp:c.WhatsApp}):[];\n  $('qDup').innerHTML=p.length?'<div class=\"hint\" style=\"background:#FFF3CD;margin:0\">⚠️ Já existe algo parecido: '+p.map(function(l){return '<a href=\"#\" onclick=\"openLead(\\''+l.ID+'\\');return false\"><b>'+esc(l.Estabelecimento)+'</b></a> ('+esc(l.Etapa)+(l['Responsável']?', '+esc(l['Responsável']):'')+')'}).join(', ')+'</div>':''}\nfunction qaObj(base){return Object.assign(base,{Etapa:$('qEt').value,Origem:$('qOri').value,'Responsável':$('qResp').value||ME,canal:$('qCan').value})}\nfunction qaSalvar(outro){var c=parseContato($('qCont').value);var nome=$('qNome').value.trim()||(c.Instagram?c.Instagram.slice(1):'');if(!nome){toast('Coloque o nome');return}\n  var o=qaObj(Object.assign({Estabelecimento:nome,Cidade:$('qCid').value.trim(),Segmento:$('qSeg').value,'Observações':$('qObs').value.trim()},c));\n  call('addLeads',[o]).then(function(r){setData(r);render();if(r.ignorados.length){toast('Não adicionado: '+r.ignorados[0]);return}toast('✓ '+nome+' adicionado');if(outro)quickAdd();else closeDrawer()})}\nfunction qaPrev(){var ls=$('qLista').value.split(/\\n/).map(parseLinha).filter(Boolean);\n  $('qPrev').innerHTML=ls.length?'<div class=\"muted\" style=\"margin-bottom:4px\">'+ls.length+' lead(s) encontrados:</div><table><tr><th>Nome</th><th>Contato</th><th>Cidade</th><th></th></tr>'+ls.map(function(o){var p=parecidos(o);return '<tr><td>'+esc(o.Estabelecimento)+'</td><td>'+esc([o.Instagram,o.WhatsApp,o['E-mail']].filter(Boolean).join(' '))+'</td><td>'+esc(o.Cidade||'')+'</td><td>'+(p.length?'<span class=\"tag yel\">parecido: '+esc(p[0].Estabelecimento)+'</span>':'<span class=\"tag\">novo</span>')+'</td></tr>'}).join('')+'</table><div class=\"muted\">Os iguais (mesmo nome, @ ou telefone) são pulados automaticamente. Os \"parecidos\" entram, mas confira.</div>':''}\nfunction qaLista(){var ls=$('qLista').value.split(/\\n/).map(parseLinha).filter(Boolean);if(!ls.length){toast('Cole pelo menos uma linha');return}\n  call('addLeads',ls.map(function(o){return qaObj(o)})).then(function(r){setData(r);render();closeDrawer();toast(r.criados.length+' adicionado(s)'+(r.ignorados.length?' · '+r.ignorados.length+' já existiam':''));if(r.ignorados.length)setTimeout(function(){alert('Já existiam (não duplicados):\\n'+r.ignorados.join('\\n'))},300)})}\n\n// ---------------- MODELOS DE MENSAGEM\nvar TPL_DEF=[{nome:'1º contato',texto:'Oi{contato}! Aqui é {eu}, da Morphê 🌱 Somos o primeiro análogo vegetal de salmão em posta do Brasil. Posso mandar nossa tabela e uma amostra para o {estabelecimento}?'}];\nfunction modelos(){var ms=(D.modelos&&D.modelos.length)?D.modelos:TPL_DEF;var dono=function(t){var m=String(t.nome).match(/\\(([^)]+)\\)\\s*$/);return m?m[1].trim():''};var meus=ms.filter(function(t){return dono(t)&&dono(t)===ME});return meus.concat(ms.filter(function(t){return !dono(t)}))}\nfunction tplTexto(l){var i=+(($('tplSel')||{}).value||0),t=modelos().at(i)||modelos().at(0);var pessoa=String(l['Contato (pessoa)']||'').replace(/\\(.*\\)/,'').trim().split(/\\s+/).at(0)||'';\n  var hr=new Date().getHours(),sau=hr<12?'Bom dia':hr<18?'Boa tarde':'Boa noite';return t.texto.split('{saudacao}').join(sau).split('{contato}').join(pessoa?' '+pessoa:'').split('{estabelecimento}').join(l.Estabelecimento||'').split('{eu}').join(ME||'')}\nfunction tplBar(l){return '<div class=\"box\" style=\"margin-top:0\"><h3>Mensagem pronta</h3><div class=\"filters\" style=\"margin:0\"><select id=\"tplSel\" onchange=\"tplPrev()\">'+modelos().map(function(t,i){return '<option value=\"'+i+'\">'+esc(t.nome)+'</option>'}).join('')+'</select>'+\n  (l.WhatsApp||l.Telefone?'<button class=\"btn sm\" onclick=\"tplWa()\">📱 WhatsApp</button>':'')+'<button class=\"btn sm sec\" onclick=\"tplCopy()\">📋 Copiar'+(l.Instagram?' p/ Instagram':'')+'</button>'+(l['E-mail']?'<button class=\"btn sm sec\" onclick=\"tplMail()\">✉️ E-mail</button>':'')+'</div><div class=\"muted\" id=\"tplPrev\" style=\"margin-top:8px;white-space:pre-wrap\">'+esc(tplTexto(l))+'</div><div class=\"muted\" style=\"margin-top:4px;font-size:11px\">Os modelos ficam na aba Config da planilha (colunas P e Q).</div></div>'}\nfunction tplPrev(){var l=lead(OPEN);if(l)$('tplPrev').textContent=tplTexto(l)}\nfunction tplWa(){var l=lead(OPEN);window.open(waLink(l.WhatsApp||l.Telefone,tplTexto(l)),'_blank')}\nfunction tplCopy(){var l=lead(OPEN),t=tplTexto(l);var ok=function(){toast('Mensagem copiada'+(l.Instagram?' — abrindo o Instagram':''));if(l.Instagram)window.open(igLink(l.Instagram),'_blank')};\n  try{navigator.clipboard.writeText(t).then(ok,function(){fallbackCopy(t);ok()})}catch(e){fallbackCopy(t);ok()}}\nfunction fallbackCopy(t){var a=document.createElement('textarea');a.value=t;document.body.appendChild(a);a.select();try{document.execCommand('copy')}catch(e){}a.remove()}\nfunction tplMail(){var l=lead(OPEN),t=modelos().at(+$('tplSel').value)||modelos().at(0);window.open('mailto:'+encodeURIComponent(l['E-mail'])+'?subject='+encodeURIComponent('Morphê · '+t.nome)+'&body='+encodeURIComponent(tplTexto(l)),'_blank')}\n\n// ---------------- TAREFAS\nfunction tarefas(filtro){return D.atividades.filter(function(a){return a.Tipo==='Tarefa'&&(!filtro||filtro(a))})}\nfunction tarefaLinha(a,mostraLead){var dp=days(a.Prazo),feita=a.Resultado==='Feita',l=a['Lead ID']?lead(a['Lead ID']):null;\n  var tag=feita?'':(dp===null?'':dp>0?'<span class=\"tag red\">atrasada '+dp+'d</span>':dp===0?'<span class=\"tag yel\">hoje</span>':'<span class=\"tag gray\">'+fmt(a.Prazo)+'</span>');\n  return '<div class=\"brow\" style=\"align-items:flex-start\"><input type=\"checkbox\" '+(feita?'checked':'')+' onchange=\"marcarTarefa(\\''+a.ID+'\\',this.checked)\" style=\"margin-top:3px\"><div style=\"flex:1'+(feita?';text-decoration:line-through;opacity:.6':'')+'\"><div>'+esc(a.Texto)+' '+tag+'</div><div class=\"muted\">'+esc(a['Menções']||'')+(mostraLead&&l?' · <a href=\"#\" onclick=\"openLead(\\''+l.ID+'\\');return false\">'+esc(l.Estabelecimento)+'</a>':'')+'</div></div></div>'}\nfunction tarefasBox(l){var ts=tarefas(function(a){return a['Lead ID']===l.ID}).sort(function(a,b){return (a.Resultado==='Feita')-(b.Resultado==='Feita')||String(a.Prazo||'9').localeCompare(String(b.Prazo||'9'))});\n  return '<div class=\"box\"><h3>Tarefas</h3>'+(ts.length?ts.map(function(a){return tarefaLinha(a,false)}).join(''):'<div class=\"muted\">Nenhuma tarefa.</div>')+\n  '<div class=\"grid2\" style=\"margin-top:8px\"><label class=\"full\">Nova tarefa<input id=\"tTxt\" placeholder=\"ex.: mandar tabela de preços\"></label><label>Prazo<input id=\"tPrazo\" type=\"date\" value=\"'+isoDay(new Date(Date.now()+864e5))+'\"></label><label>Para quem<select id=\"tPes\">'+opts(D.users.map(function(u){return u.nome}),ME)+'</select></label></div>'+\n  '<div class=\"acts\" style=\"margin-bottom:0\"><button class=\"btn sm\" onclick=\"novaTarefa()\">+ Adicionar tarefa</button></div></div>'}\nfunction novaTarefa(){var t=$('tTxt').value.trim();if(!t){toast('Escreva a tarefa');return}if(!ME){toast('Escolha seu nome no topo');return}\n  call('addTask',{leadId:OPEN,texto:t,prazo:$('tPrazo').value,pessoa:$('tPes').value}).then(function(r){after(r);toast('Tarefa criada')})}\nfunction marcarTarefa(id,feita){call('doneTask',id,feita).then(function(r){after(r);toast(feita?'Feita ✓':'Reaberta')})}\nfunction cardTarefas(so){var ts=tarefas(function(a){return a.Resultado!=='Feita'&&(!so||a['Menções']===so)}).sort(function(a,b){return String(a.Prazo||'9').localeCompare(String(b.Prazo||'9'))});\n  return '<div class=\"card\"><h3>✅ '+(so?'Minhas tarefas':'Tarefas do time')+' ('+ts.length+')</h3>'+(ts.length?ts.slice(0,12).map(function(a){return tarefaLinha(a,true)}).join(''):'<div class=\"muted\">Nenhuma tarefa aberta. Crie na ficha do lead.</div>')+'</div>'}\n\n// ---------------- LEAD DO SITE: RESPOSTA EM 24 H\nfunction slaSite(l){if(!/site|formul/i.test(l.Origem||'')||l['Último contato']||['Sugerido','Mapeado'].indexOf(l.Etapa)<0)return null;var c=d(l['Criado em']);if(!c)return null;return Math.max(0,Math.floor((Date.now()-c)/36e5))}\nfunction cardSite(so){var ls=D.leads.filter(function(l){return slaSite(l)!==null&&(!so||l['Responsável']===so||!l['Responsável'])});if(!ls.length)return '';\n  return '<div class=\"card\" style=\"border-color:var(--verm)\"><h3 style=\"color:var(--verm)\">🌐 Site sem resposta ('+ls.length+')</h3><div class=\"muted\" style=\"margin-bottom:8px\">Quem preencheu o formulário do site deve ser respondido em até 24 h.</div>'+ls.map(card).join('')+'</div>'}\n\n// ---------------- RODÍZIO\nfunction semDono(){return D.leads.filter(function(l){return !l['Responsável']&&['Sugerido','Ganho','Perdido'].indexOf(l.Etapa)<0})}\nfunction distribuirSemDono(){var ids=semDono().map(function(l){return l.ID});if(!ids.length)return;\n  if(!confirm('Distribuir '+ids.length+' lead(s) sem dono em rodízio entre: '+(D.rodizio||[]).join(', ')+'?\\n(A lista do rodízio fica na aba Config.)'))return;\n  call('distribuir',ids).then(function(r){after(r);toast(ids.length+' lead(s) distribuídos')})}\n\n// ---------------- RELATÓRIO\nfunction enviarResumo(){call('relatorioParaMim').then(function(e){toast('Resumo enviado para '+e)})}\n\n// ---------------- DRAWER (ficha do lead)\nfunction closeDrawer(){OPEN=null;$('drawer').classList.remove('open');$('ov').classList.remove('open')}\nfunction newLead(){quickAdd()}\nfunction newLeadFull(){OPEN=null;drawerForm({Etapa:'Mapeado','Responsável':ME,Origem:'Prospecção ativa'},true)}\nfunction openLead(id){OPEN=id;var l=lead(id);if(!l)return closeDrawer();drawerForm(l,false)}\nfunction field(k,l,type,list){var v=l[k]||'';if(type==='date')v=String(v).slice(0,10);\n  if(list)return '<label>'+k+'<select data-k=\"'+k+'\">'+opts(list,v,'—')+'</select></label>';\n  return '<label>'+k+'<input data-k=\"'+k+'\" type=\"'+(type||'text')+'\" value=\"'+esc(v)+'\"></label>'}\nfunction drawerForm(l,novo){var ah=novo?[]:atvOf(l.ID);var msg='Olá! Aqui é '+(ME||'')+' da Morphê 🌱 — somos o primeiro análogo de salmão em posta do Brasil. Posso te mandar nossa tabela e uma amostra?';\n  var dr=$('drawer');dr.innerHTML='<div class=\"top\"><h2>'+(novo?'Novo lead':esc(l.Estabelecimento))+'</h2><div style=\"font-size:12px;opacity:.85\">'+(novo?'':esc(l.Etapa)+' desde '+fmt(l['Etapa desde'])+' · resp.: '+esc(l['Responsável']||'—'))+'</div></div><div class=\"body\">'+\n  (novo?'':tplBar(l)+tarefasBox(l)+'<div class=\"acts\">'+(l.Etapa==='Frio'?'<button class=\"btn sm\" style=\"background:#555\" onclick=\"reativarFrio(\\''+l.ID+'\\')\">🔥 Reativar lead</button>':'')+' '+(l.Instagram?'<a class=\"btn sm sec\" target=\"_blank\" href=\"'+igLink(l.Instagram)+'\">📷 Instagram</a>':'')+(l.Site?'<a class=\"btn sm sec\" target=\"_blank\" href=\"'+esc(l.Site)+'\">🌐 Site</a>':'')+(l['Google Maps']?'<a class=\"btn sm sec\" target=\"_blank\" href=\"'+esc(l['Google Maps'])+'\">📍 Maps</a>':'')+'</div>'+\n  '<div class=\"box\"><h3>Registrar contato / comentário</h3><div class=\"grid2\"><label>Tipo<select id=\"aTipo\"><option>Contato</option><option>Comentário</option></select></label><label>Canal<select id=\"aCanal\">'+opts(D.canais,'WhatsApp')+'</select></label>'+\n  '<label class=\"full\">Resultado<select id=\"aRes\">'+opts(D.resultados,'','—')+'</select></label><label class=\"full\">O que aconteceu (use @Nome para avisar alguém)<textarea id=\"aTxt\"></textarea></label>'+\n  '<label>Próximo passo<input id=\"aProx\" value=\"'+esc(l['Próximo passo']||'')+'\"></label><label>Data do próximo passo<input id=\"aData\" type=\"date\" value=\"'+String(l['Data próximo passo']||'').slice(0,10)+'\"></label>'+\n  '<label>Mover para etapa<select id=\"aEt\">'+opts(D.etapas,'','(manter '+esc(l.Etapa)+')')+'</select></label><label>Motivo (se perdido)<select id=\"aMot\">'+opts(D.motivos,'','—')+'</select></label></div>'+\n  '<div class=\"acts\"><button class=\"btn\" onclick=\"salvarAtv()\">Salvar registro</button><button class=\"btn sec sm\" onclick=\"$(\\'aData\\').value=isoDay(new Date(Date.now()+3*864e5))\">+3 dias</button><button class=\"btn sec sm\" onclick=\"$(\\'aData\\').value=isoDay(new Date(Date.now()+7*864e5))\">+7 dias</button></div></div>')+\n  '<h3 style=\"margin-top:14px\">Dados</h3><div class=\"grid2\">'+field('Estabelecimento',l)+field('Etapa',l,'',D.etapas)+field('Segmento',l,'',D.segmentos)+field('Origem',l,'',D.origens)+field('Cidade',l)+field('Bairro',l)+\n  field('Contato (pessoa)',l)+field('Responsável',l,'',D.users.map(function(u){return u.nome}))+field('WhatsApp',l)+field('Telefone',l)+field('Instagram',l)+field('E-mail',l)+field('Site',l)+field('Potencial (R$/mês)',l,'number')+\n  field('Próximo passo',l)+field('Data próximo passo',l,'date')+field('Motivo de perda',l,'',D.motivos)+'<label class=\"full\">Observações<textarea data-k=\"Observações\">'+esc(l['Observações']||'')+'</textarea></label></div>'+\n  '<div class=\"acts\"><button class=\"btn\" onclick=\"salvarLead('+(novo?'true':'false')+')\">'+(novo?'Criar lead':'Salvar dados')+'</button><button class=\"btn sec\" onclick=\"closeDrawer()\">Fechar</button></div>'+\n  (novo?'':'<h3 style=\"margin-top:14px\">Histórico</h3><div class=\"tl\">'+(ah.length?ah.map(function(a){var c={'Contato':'#4A7BF7','Comentário':'#F2572B','Etapa':'#2E9E5B'}[a.Tipo]||'#A5C93A';\n    return '<div class=\"it\" style=\"--c:'+c+'\"><div class=\"h\">'+fmtH(a.Data)+' · <b>'+esc(a.Autor)+'</b> · '+esc(a.Tipo)+(a.Canal?' · '+esc(a.Canal):'')+(a.Resultado?' · <b>'+esc(a.Resultado)+'</b>':'')+'</div><div class=\"t\">'+esc(a.Texto).replace(/@([A-Za-zÀ-ú]+)/g,'<span class=\"mention\">@$1</span>')+'</div></div>'}).join(''):'<div class=\"muted\">Sem registros ainda.</div>')+'</div>')+'</div>';\n  dr.classList.add('open');$('ov').classList.add('open')}\nfunction reativarFrio(id){if(!confirm('Reativar este lead? Ele volta para \"Contato feito\" e entra no funil ativo novamente.'))return;call('reativarLead',id).then(function(r){after(r);toast('Lead reativado 🔥');openLead(id)})}\nfunction salvarLead(novo){var o={};if(!novo)o.ID=OPEN;document.querySelectorAll('#drawer [data-k]').forEach(function(el){o[el.dataset.k]=el.value});\n  if(!o.Estabelecimento){toast('Informe o nome');return}call('saveLead',o).then(function(r){if(novo){var n=r.leads.filter(function(x){return x.Estabelecimento===o.Estabelecimento}).pop();OPEN=n?n.ID:null}after(r);toast('Salvo ✓')})}\nfunction salvarAtv(){if(!ME){toast('Escolha seu nome no topo');return}var t=$('aTxt').value.trim(),tipo=$('aTipo').value;\n  if(!t&&!$('aRes').value&&tipo==='Comentário'){toast('Escreva algo');return}\n  call('addActivity',{leadId:OPEN,tipo:tipo,canal:tipo==='Contato'?$('aCanal').value:'',resultado:$('aRes').value,texto:t,proximoPasso:$('aProx').value,dataProx:$('aData').value,novaEtapa:$('aEt').value,motivo:$('aMot').value,mencoes:mentionsIn(t)}).then(function(r){$('aTxt').value='';$('aRes').value='';after(r);toast('Registrado ✓')})}\n\n// ---------------- DASHBOARD\nfunction inPeriod(s){if(!PERIOD)return true;var x=days(s);return x!==null&&x<=PERIOD}\nfunction pct(a,b){return b?Math.round(a/b*100)+'%':'—'}\nfunction bars(rows,max,cor){max=max||Math.max.apply(null,rows.map(function(r){return r[1]}).concat([1]));return rows.map(function(r){return '<div class=\"brow\"><div class=\"lb\">'+esc(r[0])+'</div><div style=\"flex:1\"><div class=\"bar\" style=\"width:'+Math.max(r[1]/max*100,4)+'%;background:'+(r[3]||cor||'var(--verde)')+'\">'+r[1]+'</div></div><div class=\"rt\">'+(r[2]||'')+'</div></div>'}).join('')}\nfunction kpi(l,v,s,cor){return '<div class=\"card kpi\"><div class=\"l\">'+l+'</div><div class=\"v\" style=\"'+(cor?'color:'+cor:'')+'\">'+v+'</div><div class=\"s\">'+(s||'')+'</div></div>'}\nfunction vDash(){\n  var L=D.leads.filter(function(l){return l.Etapa!=='Sugerido'&&inPeriod(l['Criado em'])});\n  var A=D.atividades.filter(function(a){return inPeriod(a.Data)});\n  var C=A.filter(function(a){return a.Tipo==='Contato'});\n  var ordem=['Mapeado','Contato feito','Respondeu','Amostra enviada','Negociação','Ganho'];\n  // alcançou etapa >= X\n  var maxIdx={};L.forEach(function(l){maxIdx[l.ID]=ordem.indexOf(l.Etapa)});\n  D.atividades.forEach(function(a){if(a.Tipo==='Etapa'&&maxIdx[a['Lead ID']]!==undefined){var dest=String(a.Texto).split('→').pop().trim().replace(/\\s*\\(.*$/,'');var i=ordem.indexOf(dest);if(i>maxIdx[a['Lead ID']])maxIdx[a['Lead ID']]=i}});\n  L.forEach(function(l){if(l.Etapa==='Perdido'&&maxIdx[l.ID]<1)maxIdx[l.ID]=1});\n  var fun=ordem.map(function(e,i){var n=L.filter(function(l){return maxIdx[l.ID]>=i}).length;return [e,n]});\n  fun=fun.map(function(r,i){return [r[0],r[1],i?pct(r[1],fun[i-1][1])+' da etapa anterior':'',COR[r[0]]]});\n  var ganhos=L.filter(function(l){return l.Etapa==='Ganho'}).length,perd=L.filter(function(l){return l.Etapa==='Perdido'}).length;\n  // canal\n  var porCanal={};C.forEach(function(a){var k=a.Canal||'Outro';porCanal[k]=porCanal[k]||{};var key=a['Lead ID'];porCanal[k][key]=porCanal[k][key]||RESPONDEU.indexOf(a.Resultado)>=0});\n  var canalRows=Object.keys(porCanal).map(function(k){var v=porCanal[k],t=Object.keys(v).length,r=Object.keys(v).filter(function(x){return v[x]}).length;return [k,t,r]}).sort(function(a,b){return b[1]-a[1]});\n  // migração de canal\n  var byLead={};D.atividades.filter(function(a){return a.Tipo==='Contato'&&a['Lead ID']}).sort(function(a,b){return a.Data<b.Data?-1:1}).forEach(function(a){(byLead[a['Lead ID']]=byLead[a['Lead ID']]||[]).push(a)});\n  var mig={};Object.keys(byLead).forEach(function(id){var arr=byLead[id];for(var i=1;i<arr.length;i++){if(arr[i].Canal&&arr[i-1].Canal&&arr[i].Canal!==arr[i-1].Canal){if(!inPeriod(arr[i].Data))continue;var k=arr[i-1].Canal+' → '+arr[i].Canal;mig[k]=mig[k]||{t:0,r:0};mig[k].t++;var resp=arr.slice(i).some(function(x){return x.Canal===arr[i].Canal&&RESPONDEU.indexOf(x.Resultado)>=0&&x.Resultado!=='Encaminhou para outro contato/canal'});if(resp)mig[k].r++;break}}});\n  var migRows=Object.keys(mig).map(function(k){return [k,mig[k].t,mig[k].r]}).sort(function(a,b){return b[1]-a[1]});\n  function grp(key){var g={};L.forEach(function(l){var k=l[key]||'(vazio)';g[k]=g[k]||{t:0,g:0,a:0};g[k].t++;if(l.Etapa==='Ganho')g[k].g++;if(maxIdx[l.ID]>=3)g[k].a++});return Object.keys(g).map(function(k){return [k,g[k].t,g[k].a+' amostras · '+g[k].g+' ganhos ('+pct(g[k].g,g[k].t)+')']}).sort(function(a,b){return b[1]-a[1]})}\n  var pessoa=D.users.map(function(u){var c=C.filter(function(a){return a.Autor===u.nome}).length;var g=L.filter(function(l){return l['Responsável']===u.nome&&l.Etapa==='Ganho'}).length;return [u.nome,c,g+' ganhos']}).filter(function(r){return r[1]||r[2]!=='0 ganhos'});\n  var mot={};L.filter(function(l){return l.Etapa==='Perdido'}).forEach(function(l){var k=l['Motivo de perda']||'(sem motivo)';mot[k]=(mot[k]||0)+1});\n  var motRows=Object.keys(mot).map(function(k){return [k,mot[k],'',' #E5484D']}).sort(function(a,b){return b[1]-a[1]});\n  var sem=D.atividades.filter(function(a){return a.Tipo==='Contato'&&days(a.Data)!==null&&days(a.Data)<7}).length;\n  var amostras=L.filter(function(l){return maxIdx[l.ID]>=3}).length;\n  var tempo=L.filter(function(l){return l.Etapa==='Ganho'}).map(function(l){var a=d(l['Criado em']),b=d(l['Etapa desde']);return a&&b?(b-a)/864e5:null}).filter(function(x){return x!==null});\n  var respTot=canalRows.reduce(function(s,r){return s+r[2]},0),contTot=canalRows.reduce(function(s,r){return s+r[1]},0);\n  var ativos=D.leads.filter(ativo).length;\n  // === NOVAS MÉTRICAS ===\n  // #2 Tempo médio por etapa (via histórico de mudanças)\n  var etapaHist={};D.atividades.filter(function(a){return a.Tipo==='Etapa'&&a['Lead ID']}).sort(function(a,b){return a.Data<b.Data?-1:1}).forEach(function(a){var id=a['Lead ID'];etapaHist[id]=etapaHist[id]||[];etapaHist[id].push({dt:d(a.Data),etapa:String(a.Texto).split('→')[0].trim()})});\n  var etapaDias={};Object.keys(etapaHist).forEach(function(id){var arr=etapaHist[id];for(var i=0;i<arr.length-1;i++){var e=arr[i].etapa,diff=arr[i+1].dt-arr[i].dt;if(diff>0){etapaDias[e]=etapaDias[e]||{s:0,n:0};etapaDias[e].s+=diff/864e5;etapaDias[e].n++}}});\n  var tempoRows=ordem.slice(0,-1).map(function(e){var t=etapaDias[e];return t&&t.n>0?[e,Math.round(t.s/t.n),Math.round(t.s/t.n)+'d médio ('+t.n+' obs.)']:null}).filter(Boolean);\n  // #4 Dia da semana com mais respostas\n  var nDias=['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];var porDia=[0,0,0,0,0,0,0];\n  D.atividades.filter(function(a){return a.Tipo==='Contato'&&RESPONDEU.indexOf(a.Resultado)>=0&&inPeriod(a.Data)}).forEach(function(a){var dt=d(a.Data);if(dt)porDia[dt.getDay()]++});\n  var diaRows=nDias.map(function(n,i){return [n,porDia[i],porDia[i]+' respostas']});\n  // #6 Top cidades\n  var cidades={};D.leads.filter(function(l){return ativo(l)}).forEach(function(l){var c=(l.Cidade||'(sem cidade)').trim().split(',')[0].trim();cidades[c]=(cidades[c]||0)+1});\n  var cidadeRows=Object.keys(cidades).map(function(c){return [c,cidades[c],cidades[c]+' leads ativos']}).sort(function(a,b){return b[1]-a[1]}).slice(0,10);\n  // #9 Leads sem contato há X dias\n  var xSC=+(window._xSC||30);\n  var semCont=D.leads.filter(ativo).map(function(l){var ds=diasSemContato(l);return {l:l,d:ds}}).filter(function(x){return x.d===null||x.d>=xSC}).sort(function(a,b){return (b.d===null?9999:b.d)-(a.d===null?9999:a.d)});\n  // #11 Taxa de resposta por semana (últimas 8 semanas)\n  var porSem={};D.atividades.filter(function(a){return a.Tipo==='Contato'&&inPeriod(a.Data)}).forEach(function(a){var dt=d(a.Data);if(!dt)return;var dow=dt.getDay();var seg=new Date(dt-dow*864e5+(dow===0?-6:1)*864e5);var sk=isoDay(seg);porSem[sk]=porSem[sk]||{c:0,r:0};porSem[sk].c++;if(RESPONDEU.indexOf(a.Resultado)>=0)porSem[sk].r++});\n  var semRows=Object.keys(porSem).sort().slice(-8).map(function(s){var v=porSem[s];return [s.slice(5),v.c,pct(v.r,v.c)+' resp.']});\n  // segmento por conversão (#5)\n  var segConv={};L.forEach(function(l){var k=l.Segmento||'(vazio)';segConv[k]=segConv[k]||{t:0,g:0};segConv[k].t++;if(l.Etapa==='Ganho')segConv[k].g++});\n  var segConvRows=Object.keys(segConv).filter(function(k){return segConv[k].t>=2}).map(function(k){var v=segConv[k];return [k,v.g,pct(v.g,v.t)+' ('+v.g+'/'+v.t+')']}).sort(function(a,b){return (b[0]===0?0:b[1]/segConv[b[0]].t)-(a[0]===0?0:a[1]/segConv[a[0]].t)});\n  var maxSegConv=Math.max.apply(null,segConvRows.map(function(r){return r[1]}).concat([1]));\n\n  $('main').innerHTML='<div class=\"filters\"><h2 style=\"margin:0\">Dashboard comercial</h2><button class=\"btn sec sm\" onclick=\"enviarResumo()\">📧 Receber o resumo semanal agora</button><select style=\"margin-left:auto\" onchange=\"PERIOD=+this.value;render()\">'+[[30,'Últimos 30 dias'],[90,'Últimos 90 dias'],[0,'Tudo']].map(function(p){return '<option value=\"'+p[0]+'\"'+(PERIOD===p[0]?' selected':'')+'>'+p[1]+'</option>'}).join('')+'</select></div>'+\n  // KPIs\n  '<div class=\"row\">'+\n   kpi('Leads no funil',ativos,'em andamento agora')+\n   kpi('Contatos na semana',sem,'meta: '+D.metaContatosSemana)+\n   kpi('Taxa de resposta geral',pct(respTot,contTot),respTot+' de '+contTot+' leads contatados')+\n   kpi('Amostras enviadas',amostras,pct(amostras,L.length)+' dos leads')+\n   kpi('Novos clientes',ganhos,'conversão '+pct(ganhos,L.length)+' · meta/mês: '+D.metaGanhosMes)+\n   kpi('Tempo até fechar',tempo.length?Math.round(tempo.reduce(function(a,b){return a+b},0)/tempo.length)+' dias':'—','média, da entrada ao ganho')+\n   kpi('Sem contato há 30d+',semCont.length,'leads ativos sem contato recente',semCont.length>5?'var(--verm)':'')+'</div>'+\n  // Funil + Canal\n  '<div class=\"row\" style=\"margin-top:14px;align-items:flex-start\">'+\n   '<div class=\"card\" style=\"flex:1;min-width:340px\"><h3>Funil de conversão — taxa por etapa</h3>'+bars(fun)+'<div class=\"muted\">Perdidos no período: '+perd+'</div></div>'+\n   '<div class=\"card\" style=\"flex:1;min-width:340px\"><h3>Taxa de resposta por canal</h3>'+(canalRows.length?canalRows.map(function(r){return '<div class=\"brow\"><div class=\"lb\">'+esc(r[0])+'</div><div style=\"flex:1;background:#F4EBDA;border-radius:6px\"><div class=\"bar\" style=\"width:'+Math.max(r[2]/Math.max(r[1],1)*100,6)+'%;background:var(--verde2)\">'+pct(r[2],r[1])+'</div></div><div class=\"rt\">'+r[2]+' de '+r[1]+' leads</div></div>'}).join(''):'<div class=\"muted\">Sem contatos no período.</div>')+\n   '<h3 style=\"margin-top:16px\">Quando muda de canal</h3>'+(migRows.length?migRows.map(function(r){return '<div class=\"brow\"><div class=\"lb\">'+esc(r[0])+'</div><div style=\"flex:1;background:#F4EBDA;border-radius:6px\"><div class=\"bar\" style=\"width:'+Math.max(r[2]/Math.max(r[1],1)*100,6)+'%;background:var(--laranja)\">'+pct(r[2],r[1])+'</div></div><div class=\"rt\">'+r[2]+' de '+r[1]+' responderam</div></div>'}).join(''):'<div class=\"muted\">Nenhuma troca de canal.</div>')+'</div></div>'+\n  // Tempo médio por etapa + Dia da semana\n  '<div class=\"row\" style=\"margin-top:14px;align-items:flex-start\">'+\n   '<div class=\"card\" style=\"flex:1;min-width:320px\"><h3>⏱ Tempo médio em cada etapa</h3>'+(tempoRows.length?bars(tempoRows,Math.max.apply(null,tempoRows.map(function(r){return r[1]})),'var(--azul)'):'<div class=\"muted\">Sem histórico suficiente de mudanças de etapa ainda.</div>')+'</div>'+\n   '<div class=\"card\" style=\"flex:1;min-width:320px\"><h3>📅 Quando recebemos respostas (dia da semana)</h3>'+bars(diaRows,Math.max.apply(null,porDia.concat([1])),'var(--lima)')+'</div></div>'+\n  // Segmento conversão + Cidades\n  '<div class=\"row\" style=\"margin-top:14px;align-items:flex-start\">'+\n   '<div class=\"card\" style=\"flex:1;min-width:320px\"><h3>🏆 Segmentos com melhor conversão</h3>'+(segConvRows.length?segConvRows.map(function(r){return '<div class=\"brow\"><div class=\"lb\">'+esc(r[0])+'</div><div style=\"flex:1\"><div class=\"bar\" style=\"width:'+Math.max(r[1]/maxSegConv*100,4)+'%;background:var(--lima)\">'+r[1]+'</div></div><div class=\"rt\">'+r[2]+'</div></div>'}).join(''):'<div class=\"muted\">Sem dados.</div>')+'</div>'+\n   '<div class=\"card\" style=\"flex:1;min-width:320px\"><h3>📍 Top cidades com leads ativos</h3>'+bars(cidadeRows,0,'var(--azul)')+'</div></div>'+\n  // Origem + Semana\n  '<div class=\"row\" style=\"margin-top:14px;align-items:flex-start\">'+\n   '<div class=\"card\" style=\"flex:1;min-width:320px\"><h3>Por origem</h3>'+bars(grp('Origem'),0,'var(--azul)')+'</div>'+\n   '<div class=\"card\" style=\"flex:1;min-width:320px\"><h3>📈 Taxa de resposta por semana</h3>'+(semRows.length?bars(semRows,Math.max.apply(null,semRows.map(function(r){return r[1]}).concat([1])),'var(--verde2)'):'<div class=\"muted\">Sem dados.</div>')+'</div></div>'+\n  // Pessoas + Motivos\n  '<div class=\"row\" style=\"margin-top:14px;align-items:flex-start\">'+\n   '<div class=\"card\" style=\"flex:1;min-width:320px\"><h3>Contatos por pessoa</h3>'+bars(pessoa,0,'var(--verde)')+'</div>'+\n   '<div class=\"card\" style=\"flex:1;min-width:320px\"><h3>Motivos de perda</h3>'+(motRows.length?bars(motRows,0,'var(--verm)'):'<div class=\"muted\">Nenhum.</div>')+'</div></div>'+\n  // Leads sem contato há X dias (#9)\n  '<div class=\"card\" style=\"margin-top:14px\"><h3>🔔 Leads sem contato recente</h3>'+\n  '<div class=\"filters\" style=\"margin-bottom:8px\"><span class=\"muted\">Sem contato há mais de </span><select style=\"width:80px\" onchange=\"window._xSC=+this.value;vDash()\">'+[14,21,30,45,60].map(function(n){return '<option'+(xSC===n?' selected':'')+' value=\"'+n+'\">'+n+'d</option>'}).join('')+'</select><span class=\"muted\"> ('+semCont.length+' leads)</span></div>'+\n  (semCont.length?'<div style=\"overflow:auto\"><table><tr><th>Estabelecimento</th><th>Responsável</th><th>Etapa</th><th>Dias s/contato</th><th></th></tr>'+semCont.slice(0,20).map(function(x){return '<tr class=\"click\" onclick=\"openLead(\\''+x.l.ID+'\\')\"><td><b>'+esc(x.l.Estabelecimento)+'</b></td><td>'+esc(x.l['Responsável']||'—')+'</td><td>'+esc(x.l.Etapa)+'</td><td><span class=\"tag '+(x.d>30?'red':x.d>14?'yel':'gray')+'\">'+( x.d===null?'nunca contactado':x.d+'d')+'</span></td><td><a href=\"#\" class=\"muted\" onclick=\"openLead(\\''+x.l.ID+'\\');return false\">abrir</a></td></tr>'}).join('')+'</table></div>'+'<div class=\"muted\" style=\"margin-top:6px\">'+(semCont.length>20?'Mostrando 20 de '+semCont.length+'. Filtre por etapa na aba Leads para ver todos.':'')+'</div>':'<div class=\"muted\">Todos os leads ativos tiveram contato recente. 🎉</div>')+'</div>'}\n\n// ------------------------------------------------------------------ PÓS-VENDA\nfunction followUpPV(id){\n  var l=lead(id);if(!l)return;\n  var dr=$('drawer');\n  var canais=['WhatsApp','Instagram','E-mail','Ligação','Visita','Outro'];\n  var vendas=['Ótimas — vendeu muito bem','Boas — saindo bem','Regulares — saindo devagar','Fracas — quase não vendeu','Não vendeu nada','Não souberam dizer'];\n  var estoques=['Sim, tem bastante','Sim, mas está acabando','Não, zerou','Não soube informar'];\n  var interesses=['Sim — quer pedir agora','Sim — mas ainda tem estoque','Talvez — vai retornar','Não agora','Não quer mais'];\n  dr.innerHTML='<div class=\"top\"><h2>📋 Follow-up pós-venda</h2><div style=\"font-size:12px;opacity:.85\">'+esc(l.Estabelecimento)+' · '+esc(l.Segmento||'')+'</div></div>'+\n  '<div class=\"body\"><p class=\"muted\" style=\"margin:0 0 12px\">Preencha após o contato com o cliente. Essas respostas ficam salvas no histórico e ajudam a decidir próximos passos.</p>'+\n  '<label>Canal do contato<select id=\"pvCanal\">'+canais.map(function(c){return'<option>'+c+'</option>'}).join('')+'</select></label>'+\n  '<h3 style=\"margin:16px 0 8px;color:var(--verde)\">Perguntas guiadas</h3>'+\n  '<label>📊 Como foram as vendas do produto desde o último pedido?<select id=\"pvVendas\"><option value=\"\">Selecione</option>'+vendas.map(function(v){return'<option>'+v+'</option>'}).join('')+'</select></label>'+\n  '<label>📦 Ainda tem produto em estoque?<select id=\"pvEstoque\"><option value=\"\">Selecione</option>'+estoques.map(function(v){return'<option>'+v+'</option>'}).join('')+'</select></label>'+\n  '<label>🔄 Interesse em novo pedido?<select id=\"pvInteresse\" onchange=\"pvInteresseMudou()\"><option value=\"\">Selecione</option>'+interesses.map(function(v){return'<option>'+v+'</option>'}).join('')+'</select></label>'+\n  '<div id=\"pvMotDiv\" style=\"display:none\"><label>⚠️ Motivo da queda / não recompra (opcional)<input id=\"pvMotivo\" placeholder=\"ex.: temporada baixa, caixa curto, produto encalhado...\"></label></div>'+\n  '<label>💬 Observações livres / feedback do cliente<textarea id=\"pvObs\" rows=\"3\" placeholder=\"ex.: adoraram o sabor, mas querem formato em cubos; pediram desconto; esperam retorno do chef...\"></textarea></label>'+\n  '<h3 style=\"margin:16px 0 8px;color:var(--verde)\">Próximo passo</h3>'+\n  '<div class=\"grid2\"><label class=\"full\">O que fazer<input id=\"pvProx\" placeholder=\"ex.: enviar proposta de pedido, ligar em 2 semanas, aguardar resposta do chef...\"></label>'+\n  '<label>Quando<input id=\"pvData\" type=\"date\" value=\"'+isoDay(new Date(Date.now()+7*864e5))+'\"></label></div>'+\n  '<div style=\"display:flex;gap:8px;margin-top:14px\"><button class=\"btn\" onclick=\"salvarFollowUpPV(\\''+id+'\\')\">Salvar follow-up</button><button class=\"btn sec\" onclick=\"closeDrawer()\">Cancelar</button></div></div>';\n  dr.classList.add('open');$('ov').classList.add('open');\n}\nfunction pvInteresseMudou(){\n  var v=$('pvInteresse').value;\n  $('pvMotDiv').style.display=(!v||v.indexOf('Sim')===0)?'none':'block';\n}\nfunction salvarFollowUpPV(id){\n  var canal=$('pvCanal').value;\n  var vendas=$('pvVendas').value;\n  var estoque=$('pvEstoque').value;\n  var interesse=$('pvInteresse').value;\n  var motivo=$('pvMotivo')?$('pvMotivo').value:'';\n  var obs=$('pvObs').value.trim();\n  var prox=$('pvProx').value.trim();\n  var data=$('pvData').value;\n  var linhas=[];\n  if(vendas)linhas.push('📊 Vendas: '+vendas);\n  if(estoque)linhas.push('📦 Estoque: '+estoque);\n  if(interesse)linhas.push('🔄 Interesse em repor: '+interesse);\n  if(motivo)linhas.push('⚠️ Motivo da queda/não recompra: '+motivo);\n  if(obs)linhas.push('💬 '+obs);\n  var texto='[Follow-up pós-venda]\\n'+linhas.join('\\n');\n  var resultado=interesse.indexOf('Sim — quer pedir')===0?'Interessado':interesse.indexOf('Sim — mas')===0?'Pediu para retornar depois':interesse.indexOf('Talvez')===0?'Pediu para retornar depois':interesse.indexOf('Não')===0?'Não tem interesse':'Não respondeu';\n  if(!canal){toast('Escolha o canal');return;}\n  OPEN=null;\n  call('addActivity',{leadId:id,tipo:'Contato',canal:canal,resultado:resultado,texto:texto,proximoPasso:prox||undefined,dataProx:data||undefined}).then(function(r){after(r);closeDrawer();toast('Follow-up registrado ✓');go('posVenda')});\n}\n\nfunction vPosVenda(){\n  var so=F.pvTodos?'':ME;\n  var clientes=D.leads.filter(function(l){return l.Etapa==='Ganho'&&(!so||l['Responsável']===so)});\n  clientes=clientes.map(function(l){\n    var ds=diasSemContato(l);var vp=days(l['Data próximo passo']);\n    var urg=0;\n    if(vp!==null&&vp>0)urg=100+vp;\n    else if(ds===null)urg=90;\n    else if(ds>=30)urg=80+Math.min(ds,50);\n    else if(ds>=14)urg=60+ds;\n    else urg=ds||0;\n    return{l:l,ds:ds,vp:vp,urg:urg};\n  }).sort(function(a,b){return b.urg-a.urg});\n  var urgentes=clientes.filter(function(x){return x.urg>=60});\n  var semCont30=clientes.filter(function(x){return x.ds===null||x.ds>=30});\n  var vencidos=clientes.filter(function(x){return x.vp!==null&&x.vp>0});\n  function tagDs(x){if(x.ds===null)return'<span class=\"tag red\">nunca contactado</span>';if(x.ds>30)return'<span class=\"tag red\">'+x.ds+'d</span>';if(x.ds>14)return'<span class=\"tag yel\">'+x.ds+'d</span>';return'<span class=\"tag gray\">'+x.ds+'d</span>';}\n  function tagVp(x){if(x.vp===null)return'<span class=\"tag gray\">—</span>';if(x.vp>0)return'<span class=\"tag red\">'+x.vp+'d de atraso</span>';if(x.vp===0)return'<span class=\"tag yel\">hoje</span>';return'<span class=\"tag\">em '+(-x.vp)+'d</span>';}\n  function row(x){var l=x.l;return'<tr><td class=\"click\" onclick=\"openLead(\\''+l.ID+'\\')\"><b>'+esc(l.Estabelecimento)+'</b><div class=\"muted\">'+esc([l.Segmento,l.Cidade].filter(String).join(' · '))+'</div></td><td>'+esc(l['Responsável']||'—')+'</td><td>'+tagDs(x)+'</td><td style=\"max-width:180px\">'+esc(l['Próximo passo']||'—')+'</td><td>'+tagVp(x)+'</td><td><button class=\"btn sm\" onclick=\"followUpPV(\\''+l.ID+'\\')\">📋 Follow-up</button>'+(l.WhatsApp?'<a href=\"'+waLink(l.WhatsApp)+'\" target=\"_blank\" class=\"btn sm sec\" style=\"margin-left:4px\">📱</a>':'')+'</td></tr>';}\n  $('main').innerHTML='<div class=\"filters\"><h2 style=\"margin:0\">Pós-venda — acompanhamento de clientes</h2><label class=\"muted\" style=\"margin-left:auto\"><input type=\"checkbox\" '+(F.pvTodos?'checked':'')+' onchange=\"F.pvTodos=this.checked;render()\"> ver o time todo</label></div>'+\n  '<div class=\"row\">'+\n  '<div class=\"card kpi\"><div class=\"l\">Clientes ativos</div><div class=\"v\">'+clientes.length+'</div><div class=\"s\">status Ganho</div></div>'+\n  '<div class=\"card kpi\"><div class=\"l\">Atenção urgente</div><div class=\"v\" style=\"color:var(--verm)\">'+urgentes.length+'</div><div class=\"s\">14d+ sem contato ou atrasado</div></div>'+\n  '<div class=\"card kpi\"><div class=\"l\">Sem contato 30d+</div><div class=\"v\" style=\"color:var(--verm)\">'+semCont30.length+'</div><div class=\"s\">risco de churn</div></div>'+\n  '<div class=\"card kpi\"><div class=\"l\">Follow-ups vencidos</div><div class=\"v\" style=\"color:var(--laranja)\">'+vencidos.length+'</div><div class=\"s\">próximo passo atrasado</div></div>'+\n  '</div>'+\n  '<div class=\"card\" style=\"margin-top:14px\"><h3>🔥 Prioritários — contate agora</h3>'+\n  (urgentes.length?'<div class=\"sug\">'+urgentes.slice(0,8).map(function(x){var l=x.l;return'<div class=\"card\"><div style=\"font-weight:700;font-size:15px\">'+esc(l.Estabelecimento)+'</div><div class=\"muted\">'+esc([l.Segmento,l.Cidade].filter(String).join(' · '))+'</div><div style=\"margin:6px 0;display:flex;gap:6px;flex-wrap:wrap\">'+tagDs(x)+tagVp(x)+'</div><div class=\"muted\" style=\"margin-bottom:8px\">'+esc(l['Próximo passo']||'Sem próximo passo definido')+'</div><div class=\"acts\"><button class=\"btn sm\" onclick=\"followUpPV(\\''+l.ID+'\\')\">📋 Registrar follow-up</button>'+(l.WhatsApp?'<a href=\"'+waLink(l.WhatsApp)+'\" target=\"_blank\" class=\"btn sm sec\">📱 WhatsApp</a>':'')+(l.Instagram?'<a href=\"'+igLink(l.Instagram)+'\" target=\"_blank\" class=\"btn sm sec\">📷 Instagram</a>':'')+'</div></div>'}).join('')+'</div>':'<div class=\"muted\">Nenhum urgente no momento. 🎉</div>')+'</div>'+\n  '<div class=\"card\" style=\"margin-top:14px\"><h3>Lista completa · ordenado por urgência</h3>'+\n  (clientes.length?'<div style=\"overflow:auto\"><table><tr><th>Cliente</th><th>Resp.</th><th>Último contato</th><th>Próximo passo</th><th>Data/prazo</th><th></th></tr>'+clientes.map(row).join('')+'</table></div>':'<div class=\"muted\">Nenhum cliente ainda. Quando um lead chega em Ganho, aparece aqui automaticamente.</div>')+'</div>';\n}\n\ndocument.addEventListener('keydown',function(e){if(e.key==='Escape')closeDrawer()});\nload();\n</script>\n</body>\n</html>\n";
