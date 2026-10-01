// ==================================================================
// Inteligência comercial: Carteira/Churn, Sell-in/Sell-out, Previsão.
// calc*_ são funções puras (recebem dados simples + data de hoje em ISO), sem tocar em planilha.
// Colunas do Monitoramento (posição fixa, conforme as abas):
//   Clientes A:H  | Pedidos B2B A:X | Follow-ups A:G | Sell-out A:E | Ficha Cliente A:K
// ==================================================================

var INT_ = {
  CICLO_PADRAO: 21,
  PROB_ETAPA: { 'Mapeado': 0.02, 'Contato feito': 0.05, 'Respondeu': 0.15, 'Amostra enviada': 0.30, 'Negociação': 0.55 },
  SELL_HEADER: ['Data', 'Cliente', 'Unidades vendidas', 'Observação', 'Registrado por'],
  FICHA_HEADER: ['Cliente', 'Ticket médio do cardápio (R$)', 'Público', 'Item no cardápio', 'No cardápio desde', 'Posição no cardápio',
    'Pedidos do item/semana', 'Motivo da baixa adesão', 'Ações em andamento', 'Notas', 'Atualizado em']
};

function iDia_(s) { if (!s) return null; var p = String(s).slice(0, 10).split('-'); if (p.length < 3) return null; return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 864e5; }
function fromDia_(n) { if (n === null || n === undefined) return ''; var d = new Date(n * 864e5); return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2); }
function nrm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, ''); }
function mediana_(a) { if (!a.length) return 0; var s = a.slice().sort(function (x, y) { return x - y; }), m = Math.floor(s.length / 2); return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }
function soma_(a, f) { var t = 0; a.forEach(function (x) { t += f(x) || 0; }); return t; }
function r1_(n) { return Math.round(n * 10) / 10; }

// ------------------------------------------------------------------ leitura do Monitoramento
function cacheGet_(k) {
  try { var v = CacheService.getScriptCache().get(k); if (!v) return null; return JSON.parse(Utilities.newBlob(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(v), 'application/x-gzip'))).getDataAsString()); } catch (e) { return null; }
}
function cachePut_(k, obj, seg) {
  try { var b = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(obj))).getBytes()); if (b.length < 95000) CacheService.getScriptCache().put(k, b, seg || 120); } catch (e) {}
}
function limparCacheMonitor_() { try { CacheService.getScriptCache().remove('mon_v1'); } catch (e) {} }

function lerMonitor_() {
  var c = cacheGet_('mon_v1'); if (c) return c;
  var ss = SpreadsheetApp.openById(CFG.MONITOR_ID);
  function rng(nome, cols) { var s = ss.getSheetByName(nome); if (!s) return []; var n = s.getLastRow() - 1; if (n < 1) return []; return s.getRange(2, 1, n, cols).getValues(); }
  function dia(v) { return (v instanceof Date) ? Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd') : (v ? String(v).slice(0, 10) : ''); }
  function num(v) { if (typeof v === 'number') return v; var s = String(v || '').replace(/[R$\s%]/g, '').replace(/\./g, '').replace(',', '.'); var n = parseFloat(s); return isNaN(n) ? 0 : n; }
  var mon = { clientes: [], pedidos: [], fups: [], sellout: [], fichas: {}, listas: {} };
  rng('Clientes', 8).forEach(function (r) {
    if (!r[0]) return;
    mon.clientes.push({ nome: String(r[0]).trim(), tipo: String(r[1] || '').trim(), cidade: cidadeCanon_(r[2]) || String(r[2] || ''), ciclo: num(r[3]), antec: num(r[4]), obs: String(r[5] || ''), motivo: String(r[6] || ''), registro: dia(r[7]) });
  });
  rng('Pedidos B2B', 25).forEach(function (r) {
    if (!r[1]) return;
    mon.pedidos.push({ data: dia(r[0]), cliente: String(r[1]).trim(), tipo: String(r[2] || '').trim(), postas: num(r[3]), c200: num(r[4]), c300: num(r[5]), d200: num(r[6]), d300: num(r[7]),
      un: num(r[8]), kg: num(r[9]), entrega: dia(r[10]), transporte: String(r[12] || ''), origem: String(r[13] || ''), obs: String(r[14] || ''), tabela: String(r[15] || ''), valor: num(r[20]), custo: num(r[21]), margemR: num(r[22]), margemPct: (typeof r[23] === 'number') ? Math.round(r[23] * 1000) / 10 : num(r[23]), frete: num(r[24]) });
  });
  rng('Follow-ups', 7).forEach(function (r) {
    if (!r[1]) return;
    mon.fups.push({ data: dia(r[0]), cliente: String(r[1]).trim(), canal: String(r[2] || ''), resultado: String(r[3] || ''), prox: String(r[4] || ''), obs: String(r[5] || ''), tipo: String(r[6] || '') });
  });
  rng('Sell-out', 5).forEach(function (r) {
    if (!r[1]) return;
    mon.sellout.push({ data: dia(r[0]), cliente: String(r[1]).trim(), un: num(r[2]), obs: String(r[3] || ''), autor: String(r[4] || '') });
  });
  rng('Ficha Cliente', 11).forEach(function (r) {
    if (!r[0]) return;
    mon.fichas[nrm_(r[0])] = { cliente: String(r[0]).trim(), ticketCardapio: num(r[1]), publico: String(r[2] || ''), noCardapio: String(r[3] || ''), desde: dia(r[4]), posicao: String(r[5] || ''),
      pedidosSemana: num(r[6]), motivo: String(r[7] || ''), acoes: String(r[8] || ''), notas: String(r[9] || ''), atualizado: dia(r[10]) };
  });
  mon.listas = lerListas_(ss, mon);
  cachePut_('mon_v1', mon, 120);
  return mon;
}

function listaValidacao_(sheet, a1) {
  try {
    var dv = sheet.getRange(a1).getDataValidation(); if (!dv) return null;
    var t = dv.getCriteriaType(), v = dv.getCriteriaValues();
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return v[0].map(String).filter(String);
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return v[0].getValues().map(function (x) { return String(x[0]).trim(); }).filter(String);
  } catch (e) {}
  return null;
}
function distintos_(arr, f) { var o = {}, r = []; arr.forEach(function (x) { var v = String(f(x) || '').trim(); if (v && !o[v]) { o[v] = 1; r.push(v); } }); return r; }
// Os menus do app espelham os menus suspensos da própria planilha (assim o Calc continua contando certo)
function lerListas_(ss, mon) {
  function L(aba, a1) { var sh = ss.getSheetByName(aba); return sh ? listaValidacao_(sh, a1) : null; }
  return {
    resultadosFup: L('Follow-ups', 'D2') || ['✅ Comprou', '📅 Vai comprar (agendado)', '⏸️ Adiou / pediu p/ voltar depois', '🔇 Não respondeu', '❌ Recusou'],
    canais: L('Follow-ups', 'C2') || ['WhatsApp', 'Instagram', 'Ligação', 'E-mail', 'Presencial'],
    tiposContato: L('Follow-ups', 'G2') || ['Recompra', 'Amostra', 'Reativação'],
    transportes: L('Pedidos B2B', 'M2') || distintos_(mon.pedidos, function (p) { return p.transporte; }),
    origens: L('Pedidos B2B', 'N2') || distintos_(mon.pedidos, function (p) { return p.origem; }),
    tabelas: L('Pedidos B2B', 'P2') || ['Antiga', 'Nova'],
    faixas: L('Pedidos B2B', 'Q2') || ['Auto', 'Varejo', 'Faixa 2', 'Faixa 3'],
    tiposCliente: ['Ativo', 'Prospecção', 'Inativo'],
    ultimaTabela: (function () { var u = mon.pedidos.filter(function (p) { return p.tipo === 'Pedido' && p.tabela; }).sort(function (a, b) { return a.data < b.data ? 1 : -1; })[0]; return u ? u.tabela : ''; })(),
    motivosQueda: L('Clientes', 'G2') || ['Preço', 'Não girou / vendeu pouco', 'Concorrente', 'Frete / logística', 'Produto / qualidade', 'Fechou / mudou de foco', 'Sazonalidade / estoque alto', 'Sem resposta', 'Outro']
  };
}

function hojeISO_() { return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd'); }

// ------------------------------------------------------------------ Carteira e churn
function skusTxt_(a) { var o = []; [['postas', 'Posta'], ['c200', 'Cubos 200g'], ['c300', 'Cubos 300g'], ['d200', 'Desfiado 200g'], ['d300', 'Desfiado 300g']].forEach(function (k) { if (a[k[0]] > 0) o.push(a[k[0]] + '× ' + k[1]); }); return o.join(' · '); }
function ehProspect_(t) { return /prospe/i.test(String(t || '')); }
function ehInativo_(t) { return /inativ/i.test(String(t || '')); }
var BAIXA_VENDA_ = /baixo em vendas|baixa venda|n[aã]o girou|vendeu pouco|ningu[eé]m pediu|sem giro/i;
var MIN_CID_ = { de: 1, da: 1, do: 1, das: 1, dos: 1, e: 1 };
/** "são paulo", "São Paulo, SP", "SÃO PAULO/SP" -> "São Paulo/SP". Vazio ou "(confirmar cidade)" -> "". */
function cidadeCanon_(c) {
  var t = String(c || '').replace(/\(.*?\)/g, ' ').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  var uf = '', m = t.match(/[\/,\-]\s*([A-Za-z]{2})\s*$/);
  if (m) { uf = m[1].toUpperCase(); t = t.slice(0, m.index); }
  t = t.replace(/[\/,]+\s*[A-Za-z]{2}\s*$/, '').replace(/[,\/\-]+\s*$/, '').trim();
  if (!t) return '';
  var nome = t.toLowerCase().split(' ').map(function (w, i) { return (i > 0 && MIN_CID_[w]) ? w : w.charAt(0).toUpperCase() + w.slice(1); }).join(' ');
  if (!uf && /^s[aã]o paulo$/i.test(nome)) uf = 'SP';
  return nome + (uf ? '/' + uf : '');
}
function regiao_(cidade) {
  var c = cidadeCanon_(cidade) || String(cidade || '');
  if (/\/SP\s*$/i.test(c)) return /^s[ãa]o paulo/i.test(c) ? 'SP Capital' : 'Interior SP';
  if (/\/RS\s*$/i.test(c)) return 'Rio Grande do Sul';
  return 'Outros estados';
}
// Situação equivalente à da aba Calc: janela do FUP começa "antecedência" dias antes da próxima compra
function situacao_(n, dUlt, ciclo, antec, H, fu, ultimoPedido) {
  if (!n || dUlt === null) return '⚪ Sem pedidos';
  var prox = dUlt + ciclo, ini = prox - antec;
  if (H < ini) return '🟢 Em dia';
  var d0 = Math.max(ini, dUlt);
  var feito = fu.filter(function (f) { var d = iDia_(f.data); return d !== null && d >= d0; });
  if (feito.length) return nrm_(feito[0].resultado).indexOf('naorespondeu') >= 0 ? '🟠 Sem resposta' : '✅ FUP feito';
  return H > prox ? '🔴 Atrasado' : '🟡 Fazer FUP';
}
function agruparPor_(arr, f) { var o = {}; arr.forEach(function (x) { var k = nrm_(f(x)); (o[k] = o[k] || []).push(x); }); return o; }

function tokensNome_(n) {
  var stop = { emporio: 1, restaurante: 1, bar: 1, cafe: 1, de: 1, do: 1, da: 1, dos: 1, das: 1, e: 1, the: 1, loja: 1, mercadinho: 1 };
  var num = { '1': 'um', '2': 'dois', '3': 'tres', '4': 'quatro', '5': 'cinco' };
  return String(n || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(function (t) { return t && !stop[t]; }).map(function (t) { return num[t] || t; });
}
function simNome_(a, b) {
  var A = tokensNome_(a), B = tokensNome_(b); if (!A.length || !B.length) return 0;
  function lev(x, y) { var m = [], i, j; for (i = 0; i <= x.length; i++) { m[i] = [i]; } for (j = 1; j <= y.length; j++) m[0][j] = j;
    for (i = 1; i <= x.length; i++) for (j = 1; j <= y.length; j++) m[i][j] = Math.min(m[i - 1][j] + 1, m[i][j - 1] + 1, m[i - 1][j - 1] + (x[i - 1] === y[j - 1] ? 0 : 1)); return m[x.length][y.length]; }
  var hit = 0; A.forEach(function (t) { if (B.some(function (u) { return t === u || (Math.min(t.length, u.length) >= 5 && 1 - lev(t, u) / Math.max(t.length, u.length) >= 0.7); })) hit++; });
  return hit / Math.max(A.length, B.length);
}
function duplicadosProvaveis_(mon, ped) {
  var out = [], cl = mon.clientes;
  for (var i = 0; i < cl.length; i++) for (var j = i + 1; j < cl.length; j++) {
    var veio = /veio do morph/i.test(cl[i].obs) || /veio do morph/i.test(cl[j].obs);
    var comum = tokensNome_(cl[i].nome).filter(function (t) { return t.length >= 5 && tokensNome_(cl[j].nome).indexOf(t) >= 0; }).length;
    var umSemPedido = !(ped[nrm_(cl[i].nome)] || []).length || !(ped[nrm_(cl[j].nome)] || []).length;
    if (simNome_(cl[i].nome, cl[j].nome) >= 0.6 || nrm_(cl[i].nome) === nrm_(cl[j].nome) || (veio && umSemPedido && comum >= 1))
      out.push({ a: cl[i].nome, b: cl[j].nome, pedidosA: (ped[nrm_(cl[i].nome)] || []).length, pedidosB: (ped[nrm_(cl[j].nome)] || []).length });
  }
  return out;
}

function calcCarteira_(mon, hoje) {
  var H = iDia_(hoje);
  var ped = agruparPor_(mon.pedidos, function (p) { return p.cliente; });
  var fup = agruparPor_(mon.fups, function (f) { return f.cliente; });
  var ativos = [], prospects = [], inativos = [];
  var conv = { comAmostra: 0, converteram: 0, dias: [] }, amostrasLista = [], pendentes = 0;

  mon.clientes.forEach(function (c) {
    var k = nrm_(c.nome), todos = ped[k] || [];
    var pedidos = todos.filter(function (p) { return p.tipo === 'Pedido' && p.data; }).sort(function (a, b) { return a.data < b.data ? -1 : 1; });
    var amostras = todos.filter(function (p) { return p.tipo === 'Amostra'; }).sort(function (a, b) { return (a.data || '') < (b.data || '') ? -1 : 1; });
    var fu = (fup[k] || []).slice().sort(function (a, b) { return a.data < b.data ? 1 : -1; });
    if (ehInativo_(c.tipo)) { inativos.push({ nome: c.nome, cidade: c.cidade, motivo: c.motivo }); return; }
    var ehAtivo = pedidos.length > 0 || (c.tipo === 'Ativo' && !amostras.length); // 'Ativo' sem pedido real e com amostra é prospect

    var amEnt = amostras.filter(function (a) { return a.entrega; }); // só amostra com entrega confirmada conta
    pendentes += amostras.length - amEnt.length;
    var dEnt0 = amEnt.length ? Math.min.apply(null, amEnt.map(function (a) { return iDia_(a.entrega); })) : null, convDia = null;
    if (amEnt.length) {
      conv.comAmostra++;
      var conv1 = pedidos.filter(function (p) { return iDia_(p.data) >= dEnt0; })[0];
      if (conv1) { conv.converteram++; convDia = iDia_(conv1.data) - dEnt0; conv.dias.push(convDia); }
    }
    amostras.forEach(function (a) { amostrasLista.push({ cliente: c.nome, data: a.data, entrega: a.entrega, skus: skusTxt_(a), un: a.un, kg: a.kg, custo: Math.round(a.custo * 100) / 100, frete: a.frete || 0, entregue: !!a.entrega, converteu: !!(a.entrega && dEnt0 !== null && convDia !== null), diasConv: a.entrega ? convDia : null, tipoCliente: ehAtivo ? 'cliente' : 'prospect' }); });

    if (!ehAtivo) {
      var ult = amostras[amostras.length - 1] || null;
      var refDia = ult ? iDia_(ult.entrega || ult.data) : null;
      var resp = ult ? fu.some(function (f) { var d = iDia_(f.data); return d !== null && refDia !== null && d >= refDia && nrm_(f.resultado).indexOf('naorespondeu') < 0; }) : false;
      var dias = refDia !== null ? H - refDia : null;
      var temEntrega = amostras.some(function (a) { return a.entrega; });
      var status = !ult ? 'sem amostra' : (!temEntrega ? 'aguardando entrega' : (resp ? 'respondeu' : (dias >= 7 ? 'sem resposta 7d+' : 'aguardando retorno')));
      prospects.push({ nome: c.nome, cidade: c.cidade, pedidoAmostra: ult ? ult.data : '', entrega: ult ? ult.entrega : '', diasDesde: dias, status: status,
        skus: ult ? skusTxt_(ult) : '', custo: Math.round(soma_(amostras, function (a) { return a.custo; }) * 100) / 100, nAmostras: amostras.length, diasParaEntregar: ult && ult.entrega && ult.data ? iDia_(ult.entrega) - iDia_(ult.data) : null, ultimoFup: fu[0] ? fu[0].resultado : '', motivo: c.motivo });
      return;
    }

    var n = pedidos.length, ultimo = n ? pedidos[n - 1] : null;
    var dUlt = ultimo ? iDia_(ultimo.data) : null, diasDesde = dUlt !== null ? H - dUlt : null;
    var ints = []; for (var i = 1; i < n; i++) ints.push(iDia_(pedidos[i].data) - iDia_(pedidos[i - 1].data));
    var cicloCad = c.tipo === 'Ativo' ? c.ciclo : 0;
    var ciclo = n >= 3 ? Math.max(1, Math.round(mediana_(ints))) : (cicloCad > 0 ? cicloCad : INT_.CICLO_PADRAO);
    var ratio = diasDesde !== null ? diasDesde / ciclo : null;
    function kgJanela(de, ate) { return soma_(pedidos, function (p) { var d = iDia_(p.data); return (d > H - ate && d <= H - de) ? p.kg : 0; }); }
    var kg90 = kgJanela(0, 90), kgPrev = kgJanela(90, 180), variacao = kgPrev > 0 ? (kg90 - kgPrev) / kgPrev : null;
    var ult3 = pedidos.slice(-3), ticket = ult3.length ? soma_(ult3, function (p) { return p.valor; }) / ult3.length : 0, kgMedio = ult3.length ? soma_(ult3, function (p) { return p.kg; }) / ult3.length : 0;
    var mix = { postas: soma_(pedidos, function (p) { return p.postas; }), c200: soma_(pedidos, function (p) { return p.c200; }), c300: soma_(pedidos, function (p) { return p.c300; }),
      d200: soma_(pedidos, function (p) { return p.d200; }), d300: soma_(pedidos, function (p) { return p.d300; }) };
    var vTot = soma_(pedidos, function (p) { return p.valor; }), mTot = soma_(pedidos, function (p) { return p.margemR; });
    var semResp = 0, recusou = false;
    for (var j = 0; j < fu.length; j++) {
      var rr = nrm_(fu[j].resultado);
      if (rr.indexOf('recusou') >= 0 && j === 0) recusou = true;
      if (rr.indexOf('naorespondeu') >= 0 || rr.indexOf('adiou') >= 0) semResp++; else break;
    }
    var pts = { atraso: 0, queda: 0, fup: 0, motivo: 0 }, flags = [];
    if (ratio !== null) {
      pts.atraso = Math.min(50, Math.max(0, (ratio - 0.9) * 40));
      if (ratio > 1.05) flags.push(diasDesde + ' dias sem pedir (ciclo de ' + ciclo + ')');
    }
    if (variacao !== null && variacao < 0) { pts.queda = Math.min(25, -variacao * 50); if (variacao <= -0.2) flags.push('volume ' + Math.round(variacao * 100) + '% vs. 90 dias anteriores'); }
    if (semResp) { pts.fup = Math.min(15, semResp * 5); flags.push(semResp + ' follow-up(s) seguidos sem resposta/adiados'); }
    if (recusou) { pts.fup = 15; flags.push('último follow-up: recusou'); }
    var fupBaixa = fu.filter(function (f) { return BAIXA_VENDA_.test(f.obs + ' ' + f.prox + ' ' + f.resultado); })[0];
    if (c.motivo) { pts.motivo = 10; flags.push('motivo de queda registrado: ' + c.motivo); }
    else if (fupBaixa) { pts.motivo = 8; flags.push('follow-up: “' + (fupBaixa.obs || fupBaixa.prox) + '”'); }
    var score = n ? Math.round(Math.min(100, pts.atraso + pts.queda + pts.fup + pts.motivo)) : null;
    var nivel = score === null ? 'sem-pedidos' : (score >= 60 ? 'critico' : (score >= 40 ? 'alto' : (score >= 20 ? 'medio' : 'baixo')));
    var antecDias = c.antec > 0 ? c.antec : 3;
    var sit = situacao_(n, dUlt, ciclo, antecDias, H, fu, ultimo ? ultimo.data : '');
    ativos.push({ situacao: sit, nome: c.nome, cidade: c.cidade, tipoCadastro: c.tipo, nPedidos: n, ultimoPedido: ultimo ? ultimo.data : '', diasDesde: diasDesde, ciclo: ciclo,
      proxRecompra: dUlt !== null ? fromDia_(dUlt + ciclo) : '', atrasoDias: diasDesde !== null ? diasDesde - ciclo : null, kg90: r1_(kg90), kgPrev90: r1_(kgPrev), variacao: variacao,
      ticketMedio: Math.round(ticket), kgMedio: r1_(kgMedio), receita: Math.round(vTot), custo: Math.round(vTot - mTot), antec: c.antec > 0 ? c.antec : 3, regiao: regiao_(c.cidade), obs: c.obs, mix: mix, margemPct: vTot ? Math.round(mTot / vTot * 1000) / 10 : null, score: score, nivel: nivel, pts: pts, flags: flags,
      ultimoFup: fu[0] ? { data: fu[0].data, resultado: fu[0].resultado, prox: fu[0].prox, obs: fu[0].obs } : null, proxPasso: (fu.filter(function (f) { return f.prox; })[0] || {}).prox || '',
      notas: fu.filter(function (f) { return f.obs || f.prox; }).slice(0, 4).map(function (f) { return { data: f.data, resultado: f.resultado, obs: f.obs, prox: f.prox }; }), motivo: c.motivo, ficha: (mon.fichas || {})[k] || null,
      historico: pedidos.slice(-6).reverse().map(function (p) { return { data: p.data, un: p.un, kg: p.kg, valor: Math.round(p.valor), margemPct: p.margemPct }; }) });
  });

  var duplicados = duplicadosProvaveis_(mon, ped);
  ativos.sort(function (a, b) { return (b.score === null ? -1 : b.score) - (a.score === null ? -1 : a.score); });
  var risco = ativos.filter(function (a) { return a.nivel === 'critico' || a.nivel === 'alto'; });
  var sem7 = prospects.filter(function (p) { return p.status === 'sem resposta 7d+'; });
  return { hoje: hoje, ativos: ativos, prospects: prospects, amostras: amostrasLista.sort(function (a, b) { return (b.data || '') < (a.data || '') ? -1 : 1; }), inativos: inativos, duplicados: duplicados, resumo: {
    nAtivos: ativos.length, nRisco: risco.length, valorEmRisco: Math.round(soma_(risco, function (a) { return a.ticketMedio; })),
    nProspects: prospects.length, prospectsSem7d: sem7.length, amostrasTotal: conv.comAmostra, amostrasPendentes: pendentes, convertidos: conv.converteram, custoAmostras: Math.round(soma_(amostrasLista, function (a) { return a.custo; })),
    custoPorConvertido: conv.converteram ? Math.round(soma_(amostrasLista, function (a) { return a.custo; }) / conv.converteram) : null,
    taxaConversao: conv.comAmostra ? conv.converteram / conv.comAmostra : null, diasMedioConversao: conv.dias.length ? Math.round(soma_(conv.dias, function (d) { return d; }) / conv.dias.length) : null } };
}

// ------------------------------------------------------------------ Sell-in / Sell-out
function calcSell_(mon, hoje) {
  var H = iDia_(hoje);
  var ped = agruparPor_(mon.pedidos.filter(function (p) { return p.tipo === 'Pedido'; }), function (p) { return p.cliente; });
  var so = agruparPor_(mon.sellout, function (s) { return s.cliente; });
  var carteira = null;
  var linhas = [];
  mon.clientes.forEach(function (c) {
    var k = nrm_(c.nome), pedidos = ped[k] || [];
    var regs = so[k] || [];
    if (!pedidos.length && !regs.length) return;
    if (ehProspect_(c.tipo) && !pedidos.length) return;
    function j(arr, f, dias) { return soma_(arr, function (x) { var d = iDia_(x.data); return (d !== null && d > H - dias && d <= H) ? f(x) : 0; }); }
    var si = { d30: j(pedidos, function (p) { return p.un; }, 30), d60: j(pedidos, function (p) { return p.un; }, 60), d90: j(pedidos, function (p) { return p.un; }, 90) };
    var sa = { d30: j(regs, function (s) { return s.un; }, 30), d60: j(regs, function (s) { return s.un; }, 60), d90: j(regs, function (s) { return s.un; }, 90) };
    var temSell = regs.length > 0;
    var primeiro = temSell ? Math.min.apply(null, regs.map(function (s) { return iDia_(s.data); })) : null;
    var giro = temSell && si.d90 > 0 ? sa.d90 / si.d90 : null;
    var estoque = Math.max(0, si.d90 - sa.d90);
    var semanas = temSell ? Math.max(1, Math.min(90, H - primeiro + 1)) / 7 : 0;
    var porSemana = temSell && sa.d90 > 0 ? sa.d90 / semanas : 0;
    var cobertura = porSemana > 0 ? estoque / porSemana : null;
    var ultimo = pedidos.length ? pedidos.map(function (p) { return p.data; }).sort().pop() : '';
    var alerta = !temSell ? 'sem-dado' : (giro !== null && giro < 0.4 ? 'baixa-adesao' : (cobertura !== null && cobertura < 1.5 ? 'repor' : 'ok'));
    linhas.push({ nome: c.nome, cidade: c.cidade, sellIn: si, sellOut: sa, giro: giro, estoque: estoque, coberturaSemanas: cobertura === null ? null : r1_(cobertura), ultimoPedido: ultimo,
      ultimoSellOut: regs.length ? regs.map(function (s) { return s.data; }).sort().pop() : '', alerta: alerta, ficha: (mon.fichas || {})[k] || null });
  });
  var ordem = { 'baixa-adesao': 0, 'repor': 1, 'sem-dado': 2, 'ok': 3 };
  linhas.sort(function (a, b) { return ordem[a.alerta] - ordem[b.alerta]; });
  return { hoje: hoje, linhas: linhas, resumo: { baixaAdesao: linhas.filter(function (l) { return l.alerta === 'baixa-adesao'; }).length, repor: linhas.filter(function (l) { return l.alerta === 'repor'; }).length,
    semDado: linhas.filter(function (l) { return l.alerta === 'sem-dado'; }).length } };
}

// ------------------------------------------------------------------ Base de previsão (a parte de leads é calculada no navegador)
function calcPrevisaoBase_(mon, hoje, carteira) {
  var ped = agruparPor_(mon.pedidos.filter(function (p) { return p.tipo === 'Pedido' && p.data; }), function (p) { return p.cliente; });
  var primeiros = [], primeirosKg = [];
  Object.keys(ped).forEach(function (k) { var a = ped[k].slice().sort(function (x, y) { return x.data < y.data ? -1 : 1; }); if (a[0].valor > 0) primeiros.push(a[0].valor); if (a[0].kg > 0) primeirosKg.push(a[0].kg); });
  var H = iDia_(hoje), lista = [];
  carteira.ativos.forEach(function (a) {
    if (!a.proxRecompra || a.nivel === 'sem-pedidos') return;
    var dias = iDia_(a.proxRecompra) - H;
    if (dias > 30) return;
    var prob = Math.max(0.2, 0.9 - (a.score || 0) / 100);
    lista.push({ nome: a.nome, prox: a.proxRecompra, prob: Math.round(prob * 100) / 100, valor: a.ticketMedio, kg: a.kgMedio, nivel: a.nivel });
  });
  return { ticketNovo: Math.round(mediana_(primeiros)), kgNovo: r1_(mediana_(primeirosKg)), probEtapa: INT_.PROB_ETAPA,
    recompra30: { valor: Math.round(soma_(lista, function (x) { return x.valor * x.prob; })), kg: r1_(soma_(lista, function (x) { return x.kg * x.prob; })), clientes: lista },
    ticketMensalPorCliente: carteira.ativos.length ? Math.round(soma_(carteira.ativos, function (a) { return a.kg90 > 0 && a.kgMedio > 0 ? a.ticketMedio * (a.kg90 / a.kgMedio) : 0; }) / carteira.ativos.length / 3) : 0 };
}

// ------------------------------------------------------------------ Painel geral (substitui Painel + Dashboard da planilha)
var SKU_ = [['postas', 'Posta (emb. 2 un)', 0.2], ['c200', 'Cubos 200g', 0.2], ['c300', 'Cubos 300g', 0.3], ['d200', 'Desfiado 200g', 0.2], ['d300', 'Desfiado 300g', 0.3]];

function calcPainel_(mon, hoje, carteira) {
  var H = iDia_(hoje);
  var ped = mon.pedidos.filter(function (p) { return p.tipo === 'Pedido' && p.data; });
  function jan(arr, de, ate, f) { return soma_(arr, function (p) { var d = iDia_(p.data); return (d > H - ate && d <= H - de) ? f(p) : 0; }); }
  function cont(arr, de, ate) { return arr.filter(function (p) { var d = iDia_(p.data); return d > H - ate && d <= H - de; }).length; }
  var rec30 = jan(ped, 0, 30, function (p) { return p.valor; }), recPrev = jan(ped, 30, 60, function (p) { return p.valor; });
  var kg30 = jan(ped, 0, 30, function (p) { return p.kg; }), kgPrev = jan(ped, 30, 60, function (p) { return p.kg; });
  var n90 = cont(ped, 0, 90), rec90 = jan(ped, 0, 90, function (p) { return p.valor; }), mar90 = jan(ped, 0, 90, function (p) { return p.margemR; });
  var com1 = carteira.ativos.filter(function (a) { return a.nPedidos >= 1; }), com2 = carteira.ativos.filter(function (a) { return a.nPedidos >= 2; });
  var cont_ = {}; carteira.ativos.forEach(function (a) { cont_[a.situacao] = (cont_[a.situacao] || 0) + 1; });
  var ordem = { '🔴 Atrasado': 0, '🟠 Sem resposta': 1, '🟡 Fazer FUP': 2 }, acao = [];
  carteira.ativos.forEach(function (a) {
    if (ordem[a.situacao] === undefined) return;
    acao.push({ tipo: 'cliente', nome: a.nome, situacao: a.situacao, dias: a.atrasoDias, prox: a.proxRecompra, ultimoFup: a.ultimoFup, nivel: a.nivel, ticket: a.ticketMedio });
  });
  acao.sort(function (x, y) { return (ordem[x.situacao] - ordem[y.situacao]) || ((y.dias || 0) - (x.dias || 0)); });
  carteira.prospects.forEach(function (pr) {
    if (pr.status === 'sem resposta 7d+' || pr.status === 'aguardando retorno') acao.push({ tipo: 'amostra', nome: pr.nome, situacao: '🔵 FUP da amostra', dias: pr.diasDesde, prox: '', ultimoFup: null, nivel: '', ticket: 0 });
  });
  // semanas (segunda a domingo), últimas 12
  var semanal = [];
  for (var w = 11; w >= 0; w--) {
    var fim = H - w * 7, ini = fim - 6;
    var dow = (new Date(fim * 864e5).getUTCDay() + 6) % 7; // 0 = segunda
    var segunda = fim - dow, domingo = segunda + 6;
    var sub = ped.filter(function (p) { var d = iDia_(p.data); return d >= segunda && d <= domingo; });
    semanal.push({ semana: fromDia_(segunda), receita: Math.round(soma_(sub, function (p) { return p.valor; })), kg: r1_(soma_(sub, function (p) { return p.kg; })), pedidos: sub.length });
  }
  var seen = {}; semanal = semanal.filter(function (x) { if (seen[x.semana]) return false; seen[x.semana] = 1; return true; });
  var produtos = SKU_.map(function (k) { var u = soma_(ped, function (p) { return p[k[0]]; }); return { sku: k[1], un: u, kg: r1_(u * k[2]), un90: jan(ped, 0, 90, function (p) { return p[k[0]]; }) }; });
  var reg = {}, regDe = {}; mon.clientes.forEach(function (c) { regDe[nrm_(c.nome)] = regiao_(c.cidade); });
  ped.forEach(function (p) { var r = regDe[nrm_(p.cliente)] || 'Outros estados'; reg[r] = reg[r] || { regiao: r, kg: 0, receita: 0 }; reg[r].kg += p.kg; reg[r].receita += p.valor; });
  var transp = {}; ped.forEach(function (p) { var t = p.transporte || '(não informado)'; transp[t] = transp[t] || { transporte: t, pedidos: 0 }; transp[t].pedidos++; });
  var motivos = {}; mon.clientes.forEach(function (c) { if (c.motivo) motivos[c.motivo] = (motivos[c.motivo] || 0) + 1; });
  var recTot = soma_(carteira.ativos, function (a) { return a.receita; });
  var porRec = carteira.ativos.slice().sort(function (x, y) { return y.receita - x.receita; });
  // comparativo 30d vs 30d anteriores (como no Dashboard da planilha)
  var comp = {
    receita: { prev: Math.round(recPrev), now: Math.round(rec30) },
    kg: { prev: r1_(kgPrev), now: r1_(kg30) },
    pedidos: { prev: cont(ped, 30, 60), now: cont(ped, 0, 30) },
    margem: { prev: Math.round(jan(ped, 30, 60, function (p) { return p.margemR; })), now: Math.round(jan(ped, 0, 30, function (p) { return p.margemR; })) }
  };
  // previsão de demanda (produção): clientes com próxima compra em até 14 dias × mix médio por pedido
  var prev14 = { postas: 0, c200: 0, c300: 0, d200: 0, d300: 0 }, nPrev14 = 0, valor14 = 0;
  carteira.ativos.forEach(function (a) {
    if (!a.nPedidos || !a.proxRecompra) return;
    if (iDia_(a.proxRecompra) - H > 14) return;
    nPrev14++; valor14 += a.ticketMedio * Math.max(0.2, 0.9 - (a.score || 0) / 100);
    ['postas', 'c200', 'c300', 'd200', 'd300'].forEach(function (k) { prev14[k] += a.mix[k] / a.nPedidos; });
  });
  var previstoSKU = SKU_.map(function (k) { return { sku: k[1], un: Math.round(prev14[k[0]]), kg: r1_(prev14[k[0]] * k[2]) }; });
  // saúde da carteira
  var recAcum = soma_(ped, function (p) { return p.valor; }), kgAcum = soma_(ped, function (p) { return p.kg; });
  var intervalos = [], leads = [];
  carteira.ativos.forEach(function (a) { if (a.nPedidos >= 2) intervalos.push(a.ciclo); });
  ped.forEach(function (p) { if (p.entrega && p.data) leads.push(iDia_(p.entrega) - iDia_(p.data)); });
  var fupsPos = mon.fups.filter(function (f) { return /comprou|vaicomprar/.test(nrm_(f.resultado)); }).length, fupsTot = mon.fups.filter(function (f) { return f.resultado; }).length;
  var segs = carteira.ativos.filter(function (a) { return a.nPedidos >= 2; }).map(function (a) { var ps = ped.filter(function (p) { return nrm_(p.cliente) === nrm_(a.nome); }).sort(function (x, y) { return x.data < y.data ? -1 : 1; }); return iDia_(ps[1].data) - iDia_(ps[0].data); });
  var compraram30 = {}; ped.forEach(function (p) { var d = iDia_(p.data); if (d > H - 30 && d <= H) compraram30[nrm_(p.cliente)] = 1; });
  var saude = { precoKg: kgAcum ? Math.round(recAcum / kgAcum * 100) / 100 : null, cicloMedio: intervalos.length ? Math.round(soma_(intervalos, function (x) { return x; }) / intervalos.length) : null,
    receitaAcum: Math.round(recAcum), ticketHistorico: ped.length ? Math.round(recAcum / ped.length) : 0, leadTime: leads.length ? r1_(soma_(leads, function (x) { return x; }) / leads.length) : null,
    compraram30: Object.keys(compraram30).length, sucessoFup: fupsTot ? fupsPos / fupsTot : null, diasSegundaCompra: segs.length ? Math.round(soma_(segs, function (x) { return x; }) / segs.length) : null,
    clientesQueda: carteira.ativos.filter(function (a) { return (a.variacao !== null && a.variacao <= -0.3) || a.nivel === 'critico'; }).length, margemAcum: Math.round(soma_(ped, function (p) { return p.margemR; })) };
  var ranking = porRec.map(function (a) { return { nome: a.nome, pedidos: a.nPedidos, kg: r1_(soma_(ped.filter(function (p) { return nrm_(p.cliente) === nrm_(a.nome); }), function (p) { return p.kg; })), receita: a.receita,
    ticket: a.nPedidos ? Math.round(a.receita / a.nPedidos) : 0, variacao: a.variacao, parte: recTot ? a.receita / recTot : 0, margemPct: a.margemPct, ritmo: a.variacao === null ? '' : (a.variacao <= -0.2 ? 'desacelerando' : (a.variacao >= 0.2 ? 'acelerando' : 'estável')) }; });
  return {
    kpis: { clientesAtivos: carteira.ativos.length, atrasados: cont_['🔴 Atrasado'] || 0, fupAgora: (cont_['🟡 Fazer FUP'] || 0) + (cont_['🟠 Sem resposta'] || 0) + carteira.prospects.filter(function (p) { return p.status === 'sem resposta 7d+'; }).length,
      taxaRecompra: com1.length ? com2.length / com1.length : null, receita30: Math.round(rec30), receitaPrev30: Math.round(recPrev), kg30: r1_(kg30), kgPrev30: r1_(kgPrev), pedidos30: cont(ped, 0, 30),
      ticketMedio90: n90 ? Math.round(rec90 / n90) : 0, margem90: rec90 ? Math.round(mar90 / rec90 * 1000) / 10 : null },
    comparativo: comp, previstoSKU: previstoSKU, previsto14: { clientes: nPrev14, valor: Math.round(valor14) }, saude: saude, ranking: ranking,
    situacoes: cont_, acaoImediata: acao, semanal: semanal, produtos: produtos,
    regioes: Object.keys(reg).map(function (k) { return { regiao: k, kg: r1_(reg[k].kg), receita: Math.round(reg[k].receita) }; }).sort(function (a, b) { return b.receita - a.receita; }),
    transportes: Object.keys(transp).map(function (k) { return transp[k]; }).sort(function (a, b) { return b.pedidos - a.pedidos; }),
    motivosQueda: Object.keys(motivos).map(function (k) { return { motivo: k, clientes: motivos[k] }; }).sort(function (a, b) { return b.clientes - a.clientes; }),
    margemPorCliente: porRec.map(function (a) { return { nome: a.nome, receita: a.receita, margemPct: a.margemPct }; }),
    concentracao: { top1: recTot ? (porRec[0] ? porRec[0].receita / recTot : 0) : null, top3: recTot ? soma_(porRec.slice(0, 3), function (a) { return a.receita; }) / recTot : null, top1Nome: porRec[0] ? porRec[0].nome : '' }
  };
}


// ------------------------------------------------------------------ Meta: quantos clientes para bater a meta (base: últimos pedidos)
function periodoMeta_(hoje, tipo) {
  var d = new Date(iDia_(hoje) * 864e5), y = d.getUTCFullYear(), m = d.getUTCMonth(), ini, fim, meses;
  if (tipo === 'anual') { ini = Date.UTC(y, 0, 1); fim = Date.UTC(y, 11, 31); meses = 12; }
  else if (tipo === 'trimestral') { var q = Math.floor(m / 3) * 3; ini = Date.UTC(y, q, 1); fim = Date.UTC(y, q + 3, 0); meses = 3; }
  else { ini = Date.UTC(y, m, 1); fim = Date.UTC(y, m + 1, 0); meses = 1; tipo = 'mensal'; }
  return { tipo: tipo, ini: ini / 864e5, fim: fim / 864e5, meses: meses };
}
function calcMeta_(mon, hoje, cfg, carteira) {
  cfg = cfg || {}; var metaKg = +cfg.kg > 0 ? +cfg.kg : 300, P = periodoMeta_(hoje, cfg.periodo), H = iDia_(hoje);
  var ped = mon.pedidos.filter(function (p) { return p.tipo === 'Pedido' && p.data; });
  var atingido = soma_(ped, function (p) { var d = iDia_(p.data); return d >= P.ini && d <= P.fim ? p.kg : 0; });
  var d30 = soma_(ped, function (p) { var d = iDia_(p.data); return d > H - 30 && d <= H ? p.kg : 0; });
  var com = carteira.ativos.filter(function (a) { return a.nPedidos >= 1 && a.kgMedio > 0; });
  // perfil médio do cliente, pelos ÚLTIMOS pedidos de cada um: kg/pedido × pedidos/mês (ciclo)
  var perfis = com.map(function (a) { return { kgPed: a.kgMedio, pedMes: 30.4 / Math.max(7, a.ciclo), kgMes: a.kgMedio * 30.4 / Math.max(7, a.ciclo) }; });
  var kgMes = perfis.length ? soma_(perfis, function (x) { return x.kgMes; }) / perfis.length : 0;
  var kgPed = perfis.length ? soma_(perfis, function (x) { return x.kgPed; }) / perfis.length : 0, pedMes = perfis.length ? soma_(perfis, function (x) { return x.pedMes; }) / perfis.length : 0;
  var metaMes = metaKg / P.meses, necess = kgMes > 0 ? Math.ceil(metaMes / kgMes) : null;
  var emDia = com.length;
  // projeção até o fim do período: já vendido + recompras esperadas (prob. por risco) dentro do período
  var esperado = 0; com.forEach(function (a) { if (!a.proxRecompra) return; var dp = iDia_(a.proxRecompra); if (dp <= P.fim) esperado += a.kgMedio * Math.max(0.2, 0.9 - (a.score || 0) / 100); });
  var proj = atingido + esperado, diasRest = Math.max(0, P.fim - H);
  return { metaKg: metaKg, periodo: P.tipo, inicio: fromDia_(P.ini), fim: fromDia_(P.fim), diasRestantes: diasRest, kgAtingido: r1_(atingido), progresso: metaKg ? atingido / metaKg : 0,
    kg30d: r1_(d30), projecaoFim: r1_(proj), projecaoPct: metaKg ? proj / metaKg : 0, kgMesPorCliente: r1_(kgMes), kgPorPedido: r1_(kgPed), pedidosMesPorCliente: Math.round(pedMes * 10) / 10,
    clientesNecessarios: necess, clientesAtivos: emDia, clientesFaltam: necess === null ? null : Math.max(0, necess - emDia),
    pedidosMesNecessarios: kgPed > 0 ? Math.ceil(metaMes / kgPed) : null, baseClientes: com.length };
}
function lerConfigMeta_() {
  var kg = 300, periodo = 'mensal';
  try { var a = setting_('Meta v2: kg'), b = setting_('Meta v2: período'); if (a !== '' && +a > 0) kg = +a; if (b) periodo = String(b); } catch (e) {}
  return { kg: kg, periodo: periodo };
}
function setConfig_(key, val) {
  var C = sh_('Config'), v = C.getRange('I2:J40').getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === key) { C.getRange(i + 2, 10).setValue(val); invalidarCacheConfig_(); return; }
  for (var j = 0; j < v.length; j++) if (!v[j][0]) { C.getRange(j + 2, 9, 1, 2).setValues([[key, val]]); invalidarCacheConfig_(); return; }
}
function salvarMetaKg(kg, periodo) {
  if (!(+kg > 0)) throw new Error('Informe a meta em kg (maior que zero).');
  if (['mensal', 'trimestral', 'anual'].indexOf(periodo) < 0) periodo = 'mensal';
  setConfig_('Meta v2: kg', +kg); setConfig_('Meta v2: período', periodo);
  return getInteligencia();
}

// ------------------------------------------------------------------ Entradas chamadas pelo front
function montarInteligencia_(mon, hoje, metaCfg) {
  var carteira = calcCarteira_(mon, hoje), meta = calcMeta_(mon, hoje, metaCfg, carteira);
  var pedidos = mon.pedidos.slice().sort(function (a, b) { return a.data < b.data ? 1 : -1; }).slice(0, 300);
  var fups = mon.fups.slice().sort(function (a, b) { return a.data < b.data ? 1 : -1; }).slice(0, 300);
  return { hoje: hoje, carteira: carteira, sell: calcSell_(mon, hoje), previsao: calcPrevisaoBase_(mon, hoje, carteira), painel: calcPainel_(mon, hoje, carteira), meta: meta,
    listas: mon.listas, clientes: mon.clientes, pedidos: pedidos, fups: fups };
}
function getInteligencia() {
  var out = montarInteligencia_(lerMonitor_(), hojeISO_(), lerConfigMeta_());
  out.geradoEm = iso_(new Date());
  return out;
}

function garantirAba_(ss, nome, header) {
  var s = ss.getSheetByName(nome);
  if (!s) {
    s = ss.insertSheet(nome);
    s.getRange(1, 1, 1, header.length).setValues([header]).setBackground('#1E5E4A').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
    s.setFrozenRows(1);
  }
  return s;
}

function registrarSellout(o) {
  var un = +o.un; if (!(un > 0)) throw new Error('Informe as unidades vendidas (maior que zero).');
  if (!o.cliente) throw new Error('Escolha o cliente.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), s = garantirAba_(ss, 'Sell-out', INT_.SELL_HEADER);
    var d = o.data ? parseDay_(o.data) : new Date();
    s.appendRow([d, String(o.cliente).trim(), un, o.obs || '', (me_().nome || '')]);
    s.getRange(s.getLastRow(), 1).setNumberFormat('dd/MM/yyyy');
    limparCacheMonitor_();
  } finally { lock.releaseLock(); }
  return getInteligencia();
}

function salvarFicha(o) {
  if (!o.cliente) throw new Error('Cliente não informado.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), s = garantirAba_(ss, 'Ficha Cliente', INT_.FICHA_HEADER);
    var linha = [String(o.cliente).trim(), +o.ticketCardapio || '', o.publico || '', o.noCardapio || '', o.desde ? parseDay_(o.desde) : '', o.posicao || '', +o.pedidosSemana || '',
      o.motivo || '', o.acoes || '', o.notas || '', new Date()];
    var nomes = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues() : [], r = -1;
    for (var i = 0; i < nomes.length; i++) if (nrm_(nomes[i][0]) === nrm_(o.cliente)) { r = i + 2; break; }
    if (r < 0) r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, linha.length).setValues([linha]);
    s.getRange(r, 5).setNumberFormat('dd/MM/yyyy'); s.getRange(r, 11).setNumberFormat('dd/MM/yyyy');
    limparCacheMonitor_();
  } finally { lock.releaseLock(); }
  return getInteligencia();
}

function primeiraLinhaVazia_(sheet, col, desde) {
  var n = Math.max(sheet.getMaxRows() - desde + 1, 1), v = sheet.getRange(desde, col, n, 1).getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === '' || v[i][0] === null) return desde + i;
  sheet.insertRowAfter(sheet.getMaxRows()); return sheet.getMaxRows();
}
function promoverParaAtivo_(ss, nome) {
  var cl = ss.getSheetByName('Clientes'); if (!cl) return;
  var v = cl.getRange(2, 1, Math.max(cl.getLastRow() - 1, 1), 2).getValues();
  for (var i = 0; i < v.length; i++) if (nrm_(v[i][0]) === nrm_(nome)) { if (v[i][1] !== 'Ativo') cl.getRange(i + 2, 2).setValue('Ativo'); return; }
}

/** Registra o ENVIO de amostra (cliente novo ou existente, SKUs, entrega e frete opcionais). Prospect não vira cliente: só no 1º pedido. */
function registrarAmostra(o) {
  if (!o.cliente) throw new Error('Escolha ou digite o cliente.');
  var q = [+o.postas || 0, +o.c200 || 0, +o.c300 || 0, +o.d200 || 0, +o.d300 || 0];
  if (!q.some(function (x) { return x > 0; })) throw new Error('Informe os SKUs enviados (quantidade de pelo menos um produto).');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  var salvo = null;
  try {
    var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), cl = ss.getSheetByName('Clientes'), nome = String(o.cliente).trim();
    var v = cl.getRange(2, 1, Math.max(cl.getLastRow() - 1, 1), 1).getValues(), existe = false;
    for (var i = 0; i < v.length; i++) if (nrm_(v[i][0]) === nrm_(nome)) { nome = String(v[i][0]).trim(); existe = true; break; }
    if (!existe) {
      var rc = primeiraLinhaVazia_(cl, 1, 2);
      cl.getRange(rc, 1, 1, 6).setValues([[nome, 'Prospecção', cidadeCanon_(o.cidade) || '', 7, 0, o.obsCliente || '']]);
    }
    var pb = ss.getSheetByName('Pedidos B2B'), row = primeiraLinhaVazia_(pb, 2, 2);
    pb.getRange(row, 1, 1, 8).setValues([[o.data ? parseDay_(o.data) : new Date(), nome, 'Amostra', q[0], q[1], q[2], q[3], q[4]]]);
    if (o.entrega) pb.getRange(row, 11).setValue(parseDay_(o.entrega));
    if (o.origem) pb.getRange(row, 14).setValue(o.origem);
    if (o.obs) pb.getRange(row, 15).setValue(o.obs);
    if (+o.frete > 0) pb.getRange(row, 25).setValue(+o.frete);
    SpreadsheetApp.flush();
    salvo = { linha: row, custo: pb.getRange(row, 22).getValue() };
    limparCacheMonitor_();
  } finally { lock.releaseLock(); }
  var out = getInteligencia(); out.salvo = salvo; return out;
}

/** o = {data, cliente, tipo:'Pedido', postas,c200,c300,d200,d300, entrega, transporte, origem, obs, tabela, valorAjustado} */
function registrarPedido(o) {
  if (!o.cliente) throw new Error('Escolha o cliente.');
  var q = [+o.postas || 0, +o.c200 || 0, +o.c300 || 0, +o.d200 || 0, +o.d300 || 0];
  if (!q.some(function (x) { return x > 0; })) throw new Error('Informe a quantidade de pelo menos um produto.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  var salvo = null;
  try {
    var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), pb = ss.getSheetByName('Pedidos B2B');
    var row = primeiraLinhaVazia_(pb, 2, 2);
    pb.getRange(row, 1, 1, 8).setValues([[o.data ? parseDay_(o.data) : new Date(), String(o.cliente).trim(), 'Pedido', q[0], q[1], q[2], q[3], q[4]]]);
    if (o.entrega) pb.getRange(row, 11).setValue(parseDay_(o.entrega));
    if (o.transporte) pb.getRange(row, 13).setValue(o.transporte);
    if (o.origem) pb.getRange(row, 14).setValue(o.origem);
    if (o.obs) pb.getRange(row, 15).setValue(o.obs);
    if (+o.frete > 0) pb.getRange(row, 25).setValue(+o.frete); // coluna Y (Frete R$)
    // Colunas I, J, L, S, V são fórmulas de matriz e R, U, W, X fórmulas por linha: nunca escrever nelas
    if (o.tabela) pb.getRange(row, 16).setValue(o.tabela);
    pb.getRange(row, 17).setValue(o.faixa || 'Auto');
    if (+o.valorAjustado > 0) pb.getRange(row, 20).setValue(+o.valorAjustado);
    SpreadsheetApp.flush();
    var r = pb.getRange(row, 9, 1, 16).getValues()[0];
    salvo = { linha: row, un: r[0], kg: r[1], valorFinal: r[12], margemR: r[14], margemPct: r[15] };
    promoverParaAtivo_(ss, o.cliente);
    limparCacheMonitor_();
  } finally { lock.releaseLock(); }
  var out = getInteligencia(); out.salvo = salvo; return out;
}

/** o = {cliente, entrega, data} → marca a entrega da amostra (a criação já é automática pelo CRM) */
function registrarEntregaAmostra(o) {
  registrarAmostrasMonitoramento_(o.cliente, o.dataPedido ? parseDay_(o.dataPedido) : null, o.entrega ? parseDay_(o.entrega) : new Date());
  limparCacheMonitor_();
  return getInteligencia();
}

/** o = {data, cliente, canal, resultado, prox, obs, tipo} */
function registrarFollowup(o) {
  if (!o.cliente) throw new Error('Escolha o cliente.');
  if (!o.resultado) throw new Error('Escolha o resultado.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), fu = ss.getSheetByName('Follow-ups');
    if (!fu.getRange('G1').getValue()) {
      fu.getRange('G1').setValue('Tipo de contato');
      fu.getRange('F1').copyTo(fu.getRange('G1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    var row = primeiraLinhaVazia_(fu, 2, 2);
    if (row > 2) fu.getRange(row - 1, 1, 1, 7).copyTo(fu.getRange(row, 1, 1, 7), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    fu.getRange(row, 1, 1, 7).setValues([[o.data ? parseDay_(o.data) : new Date(), String(o.cliente).trim(), o.canal || 'WhatsApp', o.resultado, o.prox || '', o.obs || '', o.tipo || 'Recompra']]);
    fu.getRange(row, 1).setNumberFormat('dd/MM/yyyy');
    limparCacheMonitor_();
  } finally { lock.releaseLock(); }
  return getInteligencia();
}

/** o = {nome, tipo, cidade, ciclo, antec, obs, motivo} — cria ou atualiza a linha do cliente na aba Clientes */
function salvarCliente(o) {
  if (!o.nome) throw new Error('Informe o nome do cliente.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), cl = ss.getSheetByName('Clientes');
    var v = cl.getRange(2, 1, Math.max(cl.getLastRow() - 1, 1), 8).getValues(), row = -1, antes = null;
    for (var i = 0; i < v.length; i++) if (nrm_(v[i][0]) === nrm_(o.nome)) { row = i + 2; antes = v[i]; break; }
    if (row < 0) row = primeiraLinhaVazia_(cl, 1, 2);
    var motivoMudou = (o.motivo || '') !== (antes ? String(antes[6] || '') : '');
    cl.getRange(row, 1, 1, 7).setValues([[String(o.nome).trim(), o.tipo || 'Ativo', o.cidade || '', +o.ciclo || '', o.antec === '' || o.antec === undefined ? '' : +o.antec, o.obs || '', o.motivo || '']]);
    if (o.motivo && motivoMudou) cl.getRange(row, 8).setValue(new Date()).setNumberFormat('dd/MM/yyyy');
    if (!o.motivo) cl.getRange(row, 8).clearContent();
    limparCacheMonitor_();
  } finally { lock.releaseLock(); }
  return getInteligencia();
}

// ------------------------------------------------------------------ Cadastros duplicados: mescla e apaga o duplicado (com backup)
function backupAba_(ss, nome) {
  var s = ss.getSheetByName(nome); if (!s) return;
  var dia = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd'), alvo = nome + '_backup_' + dia;
  if (!ss.getSheetByName(alvo)) { var c = s.copyTo(ss); c.setName(alvo); c.hideSheet(); }
}
/** Núcleo (sem trava, para poder rodar dentro das migrações): mantém "manter", move follow-ups/pedidos do "remover" e APAGA a linha dele em Clientes. */
function mesclarClientesCore_(manter, remover) {
  if (!manter || !remover || nrm_(manter) === nrm_(remover)) throw new Error('Escolha dois cadastros diferentes.');
  var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), cl = ss.getSheetByName('Clientes');
  var v = cl.getRange(2, 1, Math.max(cl.getLastRow() - 1, 1), 1).getValues(), rm = -1, mt = -1, nomeManter = '';
  for (var i = 0; i < v.length; i++) { if (nrm_(v[i][0]) === nrm_(remover)) rm = i + 2; if (nrm_(v[i][0]) === nrm_(manter)) { mt = i + 2; nomeManter = String(v[i][0]).trim(); } }
  if (rm < 0 || mt < 0) throw new Error('Não encontrei um dos cadastros na aba Clientes.');
  backupAba_(ss, 'Clientes'); backupAba_(ss, 'Follow-ups'); backupAba_(ss, 'Pedidos B2B');
  var movidos = 0;
  ['Follow-ups', 'Pedidos B2B'].forEach(function (aba) {
    var sh = ss.getSheetByName(aba); if (!sh || sh.getLastRow() < 2) return;
    var rng = sh.getRange(2, 2, sh.getLastRow() - 1, 1), vals = rng.getValues(), mudou = false;
    vals.forEach(function (r, k) { if (r[0] && nrm_(r[0]) === nrm_(remover)) { vals[k][0] = nomeManter; mudou = true; movidos++; } });
    if (mudou) rng.setValues(vals);
  });
  cl.deleteRow(rm);
  limparCacheMonitor_();
  Logger.log('mesclarClientes: "' + remover + '" -> "' + nomeManter + '" (' + movidos + ' linha(s) movidas)');
}
function mesclarClientes(manter, remover) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { mesclarClientesCore_(manter, remover); } finally { lock.releaseLock(); }
  return getInteligencia();
}
// Migração única: os 3 duplicados já conhecidos (a descrição escrita pelos colaboradores é preservada: só o nome do cliente é ajustado)
function migrarDuplicadosV1_() {
  var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), cl = ss.getSheetByName('Clientes'), log = [];
  var nomes = cl.getRange(2, 1, Math.max(cl.getLastRow() - 1, 1), 1).getValues().map(function (r) { return String(r[0]).trim(); });
  function existe(n) { return nomes.filter(function (x) { return nrm_(x) === nrm_(n); })[0] || ''; }
  [['Vegsim - Mooca', 'Vegsim Mocca'], ['Empório Quatro Estrelas', '4 Estrelas'], ['Manjericão do Quintal', 'Manjerição Pircicaba']].forEach(function (par) {
    var a = existe(par[0]), b = existe(par[1]);
    if (a && b) { mesclarClientesCore_(a, b); log.push(b + ' → ' + a); nomes = nomes.filter(function (x) { return nrm_(x) !== nrm_(b); }); }
  });
  return log.join('; ') || 'nenhum duplicado conhecido encontrado';
}

// ------------------------------------------------------------------ Dono do lead: ninguém começa um lead que é de outra pessoa; só o dono transfere
function quemSou_(quem) {
  var n = ''; try { n = String(me_().nome || '').trim(); } catch (e) {}
  n = n || String(quem || '').trim();
  if (!n) throw new Error('Escolha seu nome no canto superior direito antes de agir.');
  return n;
}
function leadDono_(id) {
  var L = sh_('Leads'), r = rowOf_(L, id); if (r < 0) return '';
  var h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0];
  return String(L.getRange(r, h.indexOf('Responsável') + 1).getValue() || '').trim();
}
/** Agir em lead de outra pessoa é permitido, mas o dono é AVISADO (pop-up) e a ação aparece no Mural e na linha do tempo do lead.
 *  Lead sem dono fica com quem agir primeiro. Retorna {eu, donoAntes}. */
function guardLead_(id, quem, acao) {
  var eu = quemSou_(quem), dono = leadDono_(id);
  if (!dono) writeLead_({ ID: id, 'Responsável': eu });
  else if (dono !== eu) avisarDono_(id, dono, eu, acao);
  return { eu: eu, donoAntes: dono };
}
function avisarDono_(id, dono, eu, acao) {
  var L = sh_('Leads'), r = rowOf_(L, id); if (r < 0) return;
  var h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0], nome = L.getRange(r, h.indexOf('Estabelecimento') + 1).getValue();
  logAtv_({ leadId: id, nome: nome, tipo: 'Aviso', autor: eu, texto: eu + ' ' + (acao || 'fez uma ação') + ' no lead “' + nome + '” (de ' + dono + ')', mencoes: dono });
}
/** Avisos ainda não vistos para uma pessoa (a tela mostra como pop-up). desde = ISO. */
function getAvisos(quem, desde) {
  var A = sh_('Atividades'), n = A.getLastRow(); if (n < 2 || !quem) return [];
  var h = A.getRange(1, 1, 1, A.getLastColumn()).getValues()[0], ini = Math.max(2, n - 399), v = A.getRange(ini, 1, n - ini + 1, h.length).getValues();
  var iT = h.indexOf('Tipo'), iM = h.indexOf('Menções'), iD = h.indexOf('Data'), d0 = desde ? new Date(desde) : new Date(Date.now() - 864e5);
  return v.filter(function (r) {
    if (r[iT] !== 'Aviso' || !(r[iD] instanceof Date) || r[iD] <= d0) return false;
    return String(r[iM] || '').split(',').map(function (x) { return x.trim(); }).indexOf(quem) >= 0;
  }).reverse().slice(0, 20).map(function (r) { return serializeRow_(h, r); });
}
function leadDuplicado_(nome, insta, fone) {
  var d = dupDe_(nome, insta, fone, ''); if (!d) return null;
  var L = sh_('Leads'), v = L.getDataRange().getValues(), h = v.shift(), iN = h.indexOf('Estabelecimento'), iR = h.indexOf('Responsável');
  for (var i = 0; i < v.length; i++) if (v[i][iN] === d) return { nome: d, dono: String(v[i][iR] || '').trim() };
  return { nome: d, dono: '' };
}
/** Só o dono transfere o lead para outro colaborador. */
function transferirLead(id, para, quem) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var L = sh_('Leads'), r = rowOf_(L, id); if (r < 0) throw new Error('Lead não encontrado');
    var eu = quemSou_(quem), dono = leadDono_(id);
    if (!dono) throw new Error('Este lead ainda não tem responsável.');
    if (dono !== eu) throw new Error('🔒 Só ' + dono + ' pode transferir este lead.');
    var nomes = users_().map(function (u) { return u.nome; });
    if (nomes.indexOf(para) < 0) throw new Error('Escolha um colaborador da lista.');
    if (para === eu) throw new Error('O lead já é seu.');
    writeLead_({ ID: id, 'Responsável': para });
    var h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0], nome = L.getRange(r, h.indexOf('Estabelecimento') + 1).getValue();
    var atv = logAtv_({ leadId: id, nome: nome, tipo: 'Sistema', texto: 'Lead transferido de ' + eu + ' para ' + para });
    notificar_([para], nome, eu + ' transferiu este lead para você', id);
    return { leads: [leadLeve_(id)], atividades: [atv] };
  } finally { lock.releaseLock(); }
}

// Migração única (ajustes de regra e de dados). Idempotente: pode rodar de novo sem duplicar nada.
function migrarAjustesV2_() {
  var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), log = [];
  var cl = ss.getSheetByName('Clientes'), pb = ss.getSheetByName('Pedidos B2B');
  // 1) coluna Frete (R$) nos pedidos
  if (pb && !pb.getRange('Y1').getValue()) { pb.getRange('X1').copyTo(pb.getRange('Y1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); pb.getRange('Y1').setValue('Frete (R$)'); pb.setColumnWidth(25, 90); log.push('coluna Frete'); }
  // 2) origem "Pesquisa Claude" -> "Lista pesquisada"
  try {
    var crm = ss_(), C = crm.getSheetByName('Config'), L = crm.getSheetByName('Leads'), n = 0;
    var lo = C.getRange(2, 2, 60, 1).getValues(); lo.forEach(function (r, i) { if (/^pesquisa claude/i.test(String(r[0]))) { C.getRange(2 + i, 2).setValue('Lista pesquisada'); n++; } });
    var h = L.getRange(1, 1, 1, L.getLastColumn()).getValues()[0], iO = h.indexOf('Origem') + 1;
    if (iO > 0 && L.getLastRow() > 1) { var ov = L.getRange(2, iO, L.getLastRow() - 1, 1).getValues(); ov.forEach(function (r, i) { if (/^pesquisa claude/i.test(String(r[0]))) { ov[i][0] = 'Lista pesquisada'; n++; } }); L.getRange(2, iO, ov.length, 1).setValues(ov); }
    invalidarCacheConfig_(); log.push(n + ' origem(ns) renomeada(s)');
  } catch (e) { log.push('origem: ' + e); }
  var mon = lerMonitorSemCache_();
  // 3) cidades padronizadas em Clientes
  if (cl && cl.getLastRow() > 1) {
    var cv = cl.getRange(2, 3, cl.getLastRow() - 1, 1).getValues(), m3 = 0;
    cv.forEach(function (r, i) { var c = cidadeCanon_(r[0]); if (c && c !== String(r[0]).trim()) { cv[i][0] = c; m3++; } });
    if (m3) cl.getRange(2, 3, cv.length, 1).setValues(cv); log.push(m3 + ' cidade(s) padronizada(s)');
  }
  // 4) Bruna era a Raissa: mescla o cadastro e corrige o lead no CRM
  var raissa = (mon.clientes.filter(function (c) { return /^raissa/i.test(c.nome); })[0] || {}).nome, bruna = (mon.clientes.filter(function (c) { return /^bruna/i.test(c.nome); })[0] || {}).nome;
  if (raissa && bruna) { mesclarClientesCore_(raissa, bruna); log.push('Bruna mesclada em ' + raissa); }
  try {
    var L2 = sh_('Leads'), h2 = L2.getRange(1, 1, 1, L2.getLastColumn()).getValues()[0], iN = h2.indexOf('Estabelecimento') + 1;
    if (L2.getLastRow() > 1) { var nv = L2.getRange(2, iN, L2.getLastRow() - 1, 1).getValues(), c4 = 0; nv.forEach(function (r, i) { if (/^bruna/i.test(String(r[0]))) { nv[i][0] = raissa || 'Raissa (Amostras para clientes)'; c4++; } }); if (c4) { L2.getRange(2, iN, nv.length, 1).setValues(nv); log.push(c4 + ' lead(s) Bruna -> Raissa'); } }
  } catch (e) { log.push('lead Bruna: ' + e); }
  // 5) linhas de amostra
  if (pb && pb.getLastRow() > 1) {
    var rows = pb.getRange(2, 1, pb.getLastRow() - 1, 15).getValues(), vistos = {};
    rows.forEach(function (r, i) { if (r[1] && r[2] === 'Amostra') { var k = nrm_(r[1]); vistos[k] = (vistos[k] || 0) + 1; } });
    rows.forEach(function (r, i) {
      var linha = 2 + i, nm = String(r[1] || ''), vazio = !(+r[3] || +r[4] || +r[5] || +r[6] || +r[7]);
      if (!nm || r[2] !== 'Amostra' || !vazio) return;
      if (/raissa/i.test(nm)) { pb.getRange(linha, 4, 1, 5).setValues([[5, 3, 0, 5, 0]]); log.push('Raissa: SKUs 5 postas / 3 cubos 200g / 5 desfiado 200g'); }
      else if (/damodara|personal chef/i.test(nm)) {
        if (vistos[nrm_(nm)] > 1) { pb.getRange(linha, 1, 1, 3).clearContent(); pb.getRange(linha, 4, 1, 12).clearContent(); vistos[nrm_(nm)]--; log.push('amostra duplicada vazia removida: ' + nm); }
        else { pb.getRange(linha, 1).setValue(new Date(2026, 8, 29)); pb.getRange(linha, 4, 1, 5).setValues([[0, 0, 0, 1, 0]]); log.push(nm + ': 1 desfiado 200g, pedido 29/09, entrega a confirmar'); }
      }
    });
  }
  // 6) Taverna Medieval (amostra da Maria, 30/09, 1 desfiado 200g): cria só se ainda não existir
  var tem = (mon.clientes.concat([]).filter(function (c) { return /taverna medieval/i.test(c.nome); })[0]);
  if (!tem && cl && pb) {
    var rc = primeiraLinhaVazia_(cl, 1, 2); cl.getRange(rc, 1, 1, 6).setValues([['Taverna Medieval', 'Prospecção', '', 7, 0, 'Amostra enviada pela Maria']]);
    var rp = primeiraLinhaVazia_(pb, 2, 2); pb.getRange(rp, 1, 1, 8).setValues([[new Date(2026, 8, 30), 'Taverna Medieval', 'Amostra', 0, 0, 0, 1, 0]]); pb.getRange(rp, 11).setValue(new Date(2026, 8, 30)); pb.getRange(rp, 14).setValue('Prospecção direta');
    log.push('Taverna Medieval registrada');
  }
  // 7) "Ativo" sem pedido real e com amostra vira Prospecção
  mon = lerMonitorSemCache_();
  var comPedido = {}, comAmostra = {};
  mon.pedidos.forEach(function (p) { if (p.tipo === 'Pedido') comPedido[nrm_(p.cliente)] = 1; if (p.tipo === 'Amostra') comAmostra[nrm_(p.cliente)] = 1; });
  if (cl && cl.getLastRow() > 1) {
    var tv = cl.getRange(2, 1, cl.getLastRow() - 1, 2).getValues(), m7 = 0;
    tv.forEach(function (r, i) { if (r[0] && r[1] === 'Ativo' && !comPedido[nrm_(r[0])] && comAmostra[nrm_(r[0])]) { cl.getRange(2 + i, 2).setValue('Prospecção'); m7++; } });
    if (m7) log.push(m7 + ' cadastro(s) "Ativo" sem pedido real -> Prospecção');
  }
  limparCacheMonitor_();
  return log.join('; ') || 'nada a ajustar';
}
function lerMonitorSemCache_() { limparCacheMonitor_(); return lerMonitor_(); }

// Migração única (roda sozinha pelo CRM): prepara a planilha de Monitoramento para o app unificado.
function migrarMonitoramentoV1_() {
  var ss = SpreadsheetApp.openById(CFG.MONITOR_ID), log = [];
  var cl = ss.getSheetByName('Clientes');
  if (cl) {
    // o menu de Tipo passa a aceitar "Prospecção" (antes era "Prospect (amostra)")
    cl.getRange('B2:B60').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Ativo', 'Prospecção', 'Inativo'], true).setAllowInvalid(false).build());
    var n = Math.max(cl.getLastRow() - 1, 1), v = cl.getRange(2, 2, n, 1).getValues(), c = 0;
    v.forEach(function (r, i) { if (/prospect\s*\(amostra\)/i.test(String(r[0]))) { cl.getRange(2 + i, 2).setValue('Prospecção'); c++; } });
    log.push(c + ' cliente(s) "Prospect (amostra)" renomeado(s) para "Prospecção"');
  }
  var fu = ss.getSheetByName('Follow-ups');
  if (fu) {
    if (!fu.getRange('G1').getValue()) { fu.getRange('F1').copyTo(fu.getRange('G1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); fu.getRange('G1').setValue('Tipo de contato'); }
    fu.getRange('G2:G1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Recompra', 'Amostra', 'Reativação'], true).setAllowInvalid(true).build());
    log.push('Follow-ups: coluna Tipo de contato');
  }
  garantirAba_(ss, 'Sell-out', INT_.SELL_HEADER);
  garantirAba_(ss, 'Ficha Cliente', INT_.FICHA_HEADER);
  log.push('abas Sell-out e Ficha Cliente');
  limparCacheMonitor_();
  return log.join('; ');
}

if (typeof module !== 'undefined') module.exports = { cidadeCanon_: cidadeCanon_, calcMeta_: calcMeta_, montarInteligencia_: montarInteligencia_, calcPainel_: calcPainel_, calcCarteira_: calcCarteira_, calcSell_: calcSell_, calcPrevisaoBase_: calcPrevisaoBase_, iDia_: iDia_, fromDia_: fromDia_, nrm_: nrm_, INT_: INT_ };
