/**
 * Inventário Moto Caldas — backend (Google Apps Script, preso à planilha
 * "Inventário - Moto Caldas"). Recebe as contagens enviadas pelos apps das
 * coletoras (botão ☁️ Enviar) e serve os dados pro dashboard.
 *
 * Mesma receita do backend da Agenda Pós-Venda (Codigo.gs + /exec).
 *
 * Segredos ficam em Script Properties (Configurações do projeto → Propriedades):
 *   TOKEN_ENVIO  — token que os apps mandam junto da contagem
 *   SENHA        — senha do dashboard (leitura)
 *
 * Ações:
 *   POST {token, contagem}                    → grava contagem (idempotente por id)
 *   GET  ?action=ping                         → {ok:true}
 *   GET  ?action=getData&senha=...            → lista de contagens
 *   GET  ?action=getItens&senha=...&id=...    → itens de uma contagem
 */

var ABA_CONTAGENS = 'contagens';
var ABA_ITENS = 'itens';
// Bases de contagem (planilha sobe pelo dashboard; coletora só baixa)
var BASES = {
  diaria:  {aba: 'base_diaria',  cab: ['codigo', 'desc', 'loc', 'saldo', 'qtdMov']},
  locacao: {aba: 'base_locacao', cab: ['codigo', 'loc', 'saldo', 'curva', 'marca',
                                       'ultEntrada', 'ultSaida']},
  tipo:    {aba: 'base_tipo',    cab: ['codigo', 'desc', 'tipo', 'loc', 'saldo']},
  // capacetes: TODOS os capacetes (tipos CAPACETES+BOUTIQUE, sem acessórios),
  // INCLUSIVE saldo 0 — o modo 🪖 flagra venda por código trocado
  capacetes: {aba: 'base_capacetes', cab: ['codigo', 'desc', 'tipo', 'loc', 'saldo']}
};
var ABA_BASES_META = 'bases_meta';
var CAB_BASES_META = ['tipo', 'enviado_em', 'por', 'n_pecas'];
var CAB_CONTAGENS = ['id', 'recebido_em', 'origem', 'modo', 'nome',
                     'num_contagem', 'conferente', 'data_app', 'status',
                     'locacao', 'total_itens', 'total_bipes', 'divergencias',
                     'total_fab'];
var CAB_ITENS = ['contagem_id', 'codigo', 'descricao', 'local', 'saldo',
                 'contado', 'extra', 'tamanho', 'cor'];
// Códigos do fabricante (fora do catálogo carregado) — só auditoria, não contam
var ABA_FAB = 'fabricante';
var CAB_FAB = ['contagem_id', 'codigo', 'bipes', 'primeiro', 'ultimo', 'modo'];

// Histórico de bases (guarda TODAS as versões; apaga manual no dashboard).
// A base MAIS RECENTE segue também na aba per-tipo + bases_meta (compat).
var ABA_HIST = 'bases_hist';
var CAB_HIST = ['versao_id', 'tipo', 'enviado_em', 'por', 'n_pecas'];
var ABA_BLOBS = 'base_blobs';          // peças da versão em JSON, fatiado em chunks
var CAB_BLOBS = ['versao_id', 'chunk_idx', 'chunk'];
var BLOB_CHUNK = 45000;                // limite ~50k chars por célula

function _escreverBlob(versaoId, pecas) {
  var sh = _aba(ABA_BLOBS, CAB_BLOBS);
  var json = JSON.stringify(pecas), linhas = [];
  for (var i = 0, idx = 0; i < json.length; i += BLOB_CHUNK, idx++) {
    linhas.push([versaoId, idx, json.substring(i, i + BLOB_CHUNK)]);
  }
  if (linhas.length) {
    sh.getRange(sh.getLastRow() + 1, 1, linhas.length, CAB_BLOBS.length).setValues(linhas);
  }
}
function _lerBlob(versaoId) {
  var sh = _aba(ABA_BLOBS, CAB_BLOBS), vals = sh.getDataRange().getValues(), chunks = [];
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(versaoId)) chunks.push([Number(vals[i][1]), vals[i][2]]);
  }
  if (!chunks.length) return null;
  chunks.sort(function (a, b) { return a[0] - b[0]; });
  try { return JSON.parse(chunks.map(function (c) { return c[1]; }).join('')); }
  catch (e) { return null; }
}
function _apagarBlob(versaoId) {
  var sh = _aba(ABA_BLOBS, CAB_BLOBS), vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === String(versaoId)) sh.deleteRow(i + 1);
  }
}
// Versões de um tipo (do histórico), mais recente primeiro.
function _versoesDe(tipo) {
  var sh = _aba(ABA_HIST, CAB_HIST), vals = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < vals.length; i++) {
    if (!tipo || String(vals[i][1]) === String(tipo)) {
      var o = {};
      CAB_HIST.forEach(function (c, k) { o[c] = vals[i][k]; });
      out.push(o);
    }
  }
  out.sort(function (a, b) { return new Date(b.enviado_em) - new Date(a.enviado_em); });
  return out;
}

function _props() { return PropertiesService.getScriptProperties(); }

// Segredo: Script Property vence; senão cai no secrets.gs (fora do git)
function _segredo(nome) {
  var v = _props().getProperty(nome);
  if (v) return v;
  try { return SEGREDOS[nome]; } catch (e) { return null; }
}

function _aba(nome, cabecalho) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(cabecalho);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'upload_base') return _uploadBase(body);
    if (body.action === 'deleteBase') return _deleteBase(body);
    if (!body.token || body.token !== _segredo('TOKEN_ENVIO')) {
      return _json({ok: false, erro: 'token inválido'});
    }
    var c = body.contagem || {};
    if (!c.id || !c.modo || !c.itens || !c.itens.length) {
      return _json({ok: false, erro: 'contagem incompleta'});
    }
    var shC = _aba(ABA_CONTAGENS, CAB_CONTAGENS);
    var shI = _aba(ABA_ITENS, CAB_ITENS);

    // idempotência: reenvio da mesma contagem substitui a anterior
    _apagarContagem(shC, shI, c.id);

    var divergencias = 0;
    var linhas = c.itens.map(function (it) {
      var saldo = (it.saldo === undefined || it.saldo === null) ? '' : it.saldo;
      var contado = (it.qtd === undefined || it.qtd === null) ? '' : it.qtd;
      if (saldo !== '' && contado !== '' && Number(saldo) !== Number(contado)) {
        divergencias++;
      }
      return [c.id, it.codigo || '', it.desc || '', it.loc || '', saldo,
              contado, it.extra ? 'EXT' : '', it.tamanho || '', it.cor || ''];
    });
    if (linhas.length) {
      shI.getRange(shI.getLastRow() + 1, 1, linhas.length, CAB_ITENS.length)
         .setValues(linhas);
    }
    // Códigos do fabricante (auditoria) — fora do catálogo, não contam no inventário
    var fab = c.fab || [];
    if (fab.length) {
      var shF = _aba(ABA_FAB, CAB_FAB);
      var linhasFab = fab.map(function (f) {
        return [c.id, f.codigo || '', f.qtd || '', f.primeiro || '',
                f.ultimo || '', f.modo || c.modo];
      });
      shF.getRange(shF.getLastRow() + 1, 1, linhasFab.length, CAB_FAB.length)
         .setValues(linhasFab);
    }
    shC.appendRow([c.id, new Date(), c.origem || '', c.modo, c.nome || '',
                   c.numContagem || '', c.conferente || '', c.data || '',
                   c.status || '', c.locacao || '', c.itens.length,
                   c.totalBipes || '', divergencias, fab.length]);
    return _json({ok: true, itens: linhas.length, divergencias: divergencias,
                  fab: fab.length});
  } catch (err) {
    return _json({ok: false, erro: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function _apagarContagem(shC, shI, id) {
  var vals = shC.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === String(id)) shC.deleteRow(i + 1);
  }
  var itens = shI.getDataRange().getValues();
  // apaga de baixo pra cima, em blocos contíguos quando possível
  for (var j = itens.length - 1; j >= 1; j--) {
    if (String(itens[j][0]) === String(id)) shI.deleteRow(j + 1);
  }
  // fabricante (auditoria) da mesma contagem
  var shF = _aba(ABA_FAB, CAB_FAB);
  var fabs = shF.getDataRange().getValues();
  for (var k = fabs.length - 1; k >= 1; k--) {
    if (String(fabs[k][0]) === String(id)) shF.deleteRow(k + 1);
  }
}

function _autorizado(p) {
  return (p.senha && p.senha === _segredo('SENHA')) ||
         (p.token && p.token === _segredo('TOKEN_ENVIO'));
}

function _uploadBase(body) {
  if (!_autorizado(body)) return _json({ok: false, erro: 'não autorizado'});
  var cfg = BASES[body.tipo];
  if (!cfg || !body.pecas || !body.pecas.length) {
    return _json({ok: false, erro: 'base inválida'});
  }
  var sh = _aba(cfg.aba, cfg.cab);
  sh.clearContents();
  sh.appendRow(cfg.cab);
  var linhas = body.pecas.map(function (pc) {
    return cfg.cab.map(function (c) {
      var v = pc[c];
      return (v === undefined || v === null) ? '' : v;
    });
  });
  sh.getRange(2, 1, linhas.length, cfg.cab.length).setValues(linhas);
  var shM = _aba(ABA_BASES_META, CAB_BASES_META);
  var vals = shM.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === body.tipo) shM.deleteRow(i + 1);
  }
  var agora = new Date();
  shM.appendRow([body.tipo, agora, body.por || '', body.pecas.length]);
  // guarda a versão no histórico (mantém TODAS)
  var versaoId = body.tipo + '-' + agora.getTime();
  _aba(ABA_HIST, CAB_HIST).appendRow([versaoId, body.tipo, agora,
                                      body.por || '', body.pecas.length]);
  _escreverBlob(versaoId, body.pecas);
  return _json({ok: true, n: body.pecas.length, versao_id: versaoId});
}

// grava um conjunto de peças na aba per-tipo + bases_meta (o "mais recente")
function _gravarPerTipo(tipo, pecas, ver) {
  var cfg = BASES[tipo];
  if (!cfg) return;
  var sh = _aba(cfg.aba, cfg.cab);
  sh.clearContents();
  sh.appendRow(cfg.cab);
  if (pecas.length) {
    var linhas = pecas.map(function (pc) {
      return cfg.cab.map(function (c) {
        var v = pc[c]; return (v === undefined || v === null) ? '' : v; });
    });
    sh.getRange(2, 1, linhas.length, cfg.cab.length).setValues(linhas);
  }
  var shM = _aba(ABA_BASES_META, CAB_BASES_META), mv = shM.getDataRange().getValues();
  for (var i = mv.length - 1; i >= 1; i--) {
    if (String(mv[i][0]) === tipo) shM.deleteRow(i + 1);
  }
  shM.appendRow([tipo, (ver && ver.enviado_em) || new Date(),
                 (ver && ver.por) || '', pecas.length]);
}
function _limparPerTipo(tipo) {
  var cfg = BASES[tipo];
  if (cfg) { var sh = _aba(cfg.aba, cfg.cab); sh.clearContents(); sh.appendRow(cfg.cab); }
  var shM = _aba(ABA_BASES_META, CAB_BASES_META), mv = shM.getDataRange().getValues();
  for (var i = mv.length - 1; i >= 1; i--) {
    if (String(mv[i][0]) === tipo) shM.deleteRow(i + 1);
  }
}
// Migra as bases atuais (só na aba per-tipo) para o histórico (hist + blob),
// preservando o timestamp original. Idempotente — pula tipo que já tem versão.
function _migrarHist() {
  var tipos = ['diaria', 'locacao', 'tipo', 'capacetes'], feitos = [];
  tipos.forEach(function (t) {
    if (_versoesDe(t).length) return;
    var m = _metaBase(t);
    if (!m) return;
    var cfg = BASES[t], sh = _aba(cfg.aba, cfg.cab), vals = sh.getDataRange().getValues(), pecas = [];
    for (var b = 1; b < vals.length; b++) {
      var pb = {}; cfg.cab.forEach(function (c, k) { pb[c] = vals[b][k]; }); pecas.push(pb);
    }
    if (!pecas.length) return;
    var ts = m.enviado_em ? new Date(m.enviado_em) : new Date();
    var vid = t + '-' + ts.getTime();
    _aba(ABA_HIST, CAB_HIST).appendRow([vid, t, ts, m.por || '', pecas.length]);
    _escreverBlob(vid, pecas);
    feitos.push({tipo: t, versao_id: vid, n: pecas.length});
  });
  return feitos;
}
function _deleteBase(body) {
  if (!_autorizado(body)) return _json({ok: false, erro: 'não autorizado'});
  var vid = body.versao_id;
  if (!vid) return _json({ok: false, erro: 'versao_id ausente'});
  // base legado (pré-histórico): limpa a aba per-tipo + meta
  if (String(vid).indexOf('legado-') === 0) {
    _limparPerTipo(String(vid).substring(7));
    return _json({ok: true, legado: true});
  }
  var shH = _aba(ABA_HIST, CAB_HIST), vals = shH.getDataRange().getValues(), tipo = null;
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === String(vid)) { tipo = vals[i][1]; shH.deleteRow(i + 1); }
  }
  _apagarBlob(vid);
  if (tipo) {
    // reflete o novo "mais recente" na aba per-tipo (usada pelo getBase sem versão)
    var rest = _versoesDe(tipo);
    if (rest.length) {
      var pc = _lerBlob(rest[0].versao_id);
      if (pc) _gravarPerTipo(tipo, pc, rest[0]);
    } else {
      _limparPerTipo(tipo);
    }
  }
  return _json({ok: true});
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') return _json({ok: true, servico: 'inventario-mc'});
  if (p.action === 'getBase') {
    if (!_autorizado(p)) return _json({ok: false, erro: 'não autorizado'});
    var cfgB = BASES[p.tipo];
    if (!cfgB) return _json({ok: false, erro: 'tipo inválido'});
    // versão específica escolhida pelo contador → lê do blob
    // ('legado-*' = base pré-histórico, cai no read da aba per-tipo abaixo)
    if (p.versao && p.versao.indexOf('legado-') !== 0) {
      var pecasV = _lerBlob(p.versao);
      if (!pecasV) return _json({ok: false, erro: 'versão não encontrada'});
      var mv = _versoesDe(p.tipo).filter(function (v) {
        return String(v.versao_id) === String(p.versao); })[0] || null;
      return _json({ok: true, meta: mv ? {enviado_em: mv.enviado_em, por: mv.por,
                    n: mv.n_pecas, versao_id: mv.versao_id} : _metaBase(p.tipo),
                    pecas: pecasV});
    }
    var shB = _aba(cfgB.aba, cfgB.cab);
    var valsB = shB.getDataRange().getValues();
    var pecas = [];
    for (var b = 1; b < valsB.length; b++) {
      var pb = {};
      cfgB.cab.forEach(function (c, k) { pb[c] = valsB[b][k]; });
      pecas.push(pb);
    }
    return _json({ok: true, meta: _metaBase(p.tipo), pecas: pecas});
  }
  if (p.action === 'getBases') {
    if (!_autorizado(p)) return _json({ok: false, erro: 'não autorizado'});
    return _json({ok: true, bases: {diaria: _metaBase('diaria'),
                                    locacao: _metaBase('locacao'),
                                    tipo: _metaBase('tipo'),
                                    capacetes: _metaBase('capacetes')}});
  }
  if (p.action === 'migrarHist') {
    if (!p.senha || p.senha !== _segredo('SENHA')) {
      return _json({ok: false, erro: 'senha inválida'});
    }
    return _json({ok: true, migrados: _migrarHist()});
  }
  if (p.action === 'getBasesHist') {
    if (!_autorizado(p)) return _json({ok: false, erro: 'não autorizado'});
    var versoes = _versoesDe(p.tipo || null);
    // compat: bases pré-histórico (só na aba per-tipo) aparecem como 'legado'
    var tiposComHist = {};
    versoes.forEach(function (v) { tiposComHist[v.tipo] = true; });
    var tiposAlvo = p.tipo ? [p.tipo] : ['diaria', 'locacao', 'tipo', 'capacetes'];
    tiposAlvo.forEach(function (t) {
      if (tiposComHist[t]) return;
      var m = _metaBase(t);
      if (m) versoes.push({versao_id: 'legado-' + t, tipo: t,
                           enviado_em: m.enviado_em, por: m.por, n_pecas: m.n});
    });
    versoes.sort(function (a, b) { return new Date(b.enviado_em) - new Date(a.enviado_em); });
    return _json({ok: true, versoes: versoes});
  }
  if (!p.senha || p.senha !== _segredo('SENHA')) {
    return _json({ok: false, erro: 'senha inválida'});
  }
  if (p.action === 'getData') {
    var shC = _aba(ABA_CONTAGENS, CAB_CONTAGENS);
    var vals = shC.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < vals.length; i++) {
      var o = {};
      CAB_CONTAGENS.forEach(function (col, k) { o[col] = vals[i][k]; });
      out.push(o);
    }
    return _json({ok: true, contagens: out});
  }
  if (p.action === 'getItens') {
    var shI = _aba(ABA_ITENS, CAB_ITENS);
    var tudo = shI.getDataRange().getValues();
    var itens = [];
    for (var j = 1; j < tudo.length; j++) {
      if (String(tudo[j][0]) !== String(p.id)) continue;
      var it = {};
      CAB_ITENS.forEach(function (col, k) { it[col] = tudo[j][k]; });
      itens.push(it);
    }
    var shF = _aba(ABA_FAB, CAB_FAB);
    var tudoF = shF.getDataRange().getValues();
    var fab = [];
    for (var f = 1; f < tudoF.length; f++) {
      if (String(tudoF[f][0]) !== String(p.id)) continue;
      var ff = {};
      CAB_FAB.forEach(function (col, k) { ff[col] = tudoF[f][k]; });
      fab.push(ff);
    }
    return _json({ok: true, itens: itens, fab: fab});
  }
  return _json({ok: false, erro: 'ação desconhecida'});
}

function _metaBase(tipo) {
  var shM = _aba(ABA_BASES_META, CAB_BASES_META);
  var mv = shM.getDataRange().getValues();
  for (var j = mv.length - 1; j >= 1; j--) {
    if (String(mv[j][0]) === tipo) {
      return {enviado_em: mv[j][1], por: mv[j][2], n: mv[j][3]};
    }
  }
  return null;
}

/** Rode UMA vez no editor pra conceder as permissões do script. */
function autorizar() {
  var ss = SpreadsheetApp.getActive();
  Logger.log('OK — planilha: ' + ss.getName());
}
