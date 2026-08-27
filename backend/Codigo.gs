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
                                       'ultEntrada', 'ultSaida']}
};
var ABA_BASES_META = 'bases_meta';
var CAB_BASES_META = ['tipo', 'enviado_em', 'por', 'n_pecas'];
var CAB_CONTAGENS = ['id', 'recebido_em', 'origem', 'modo', 'nome',
                     'num_contagem', 'conferente', 'data_app', 'status',
                     'locacao', 'total_itens', 'total_bipes', 'divergencias'];
var CAB_ITENS = ['contagem_id', 'codigo', 'descricao', 'local', 'saldo',
                 'contado', 'extra', 'tamanho', 'cor'];

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
    shC.appendRow([c.id, new Date(), c.origem || '', c.modo, c.nome || '',
                   c.numContagem || '', c.conferente || '', c.data || '',
                   c.status || '', c.locacao || '', c.itens.length,
                   c.totalBipes || '', divergencias]);
    return _json({ok: true, itens: linhas.length, divergencias: divergencias});
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
  shM.appendRow([body.tipo, new Date(), body.por || '', body.pecas.length]);
  return _json({ok: true, n: body.pecas.length});
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') return _json({ok: true, servico: 'inventario-mc'});
  if (p.action === 'getBase') {
    if (!_autorizado(p)) return _json({ok: false, erro: 'não autorizado'});
    var cfgB = BASES[p.tipo];
    if (!cfgB) return _json({ok: false, erro: 'tipo inválido'});
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
                                    locacao: _metaBase('locacao')}});
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
    return _json({ok: true, itens: itens});
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
