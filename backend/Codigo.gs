/**
 * Inventário Moto Caldas — backend (Google Apps Script, preso à planilha
 * "Inventário - Moto Caldas"). Recebe as contagens enviadas pelos apps das
 * coletoras (botão ☁️ Enviar) e serve os dados pro dashboard.
 *
 * Mesma receita do backend da Agenda Pós-Venda (Codigo.gs + /exec).
 *
 * Segredos ficam em Script Properties (Configurações do projeto → Propriedades):
 *   TOKEN_ENVIO  — token das coletoras. É PÚBLICO (está no app do GitHub Pages),
 *                  então só pode: enviar contagem, ler base e ler o mapa.
 *   SENHA        — senha do dashboard (tudo).
 *   TOKEN_ADMIN  — token do bot (estoque_diario): mesmo poder da senha, sem
 *                  depender da senha da equipe.
 *
 * Ações (leitura aceita GET ?action=... ou POST {action,...}; o dashboard usa
 * POST pra senha não ir na URL):
 *   POST {token, contagem}                    → grava contagem (idempotente por id)
 *   GET  ?action=ping                         → {ok:true}
 *   getData (admin)                           → lista de contagens (com categoria)
 *   getItens / getItensLote (admin)           → itens de uma / várias contagens
 *   getMapa (coletora)                        → mapa persistente das prateleiras
 *   arquivar_limpar {categorias?} (admin)     → arquiva e zera (só as categorias)
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
                     'total_fab', 'categoria'];
// Categoria = o que foi contado. A coletora manda desde a v30 (Netum) / v37
// (Eyoyo); as linhas antigas são deduzidas em _categoriaDe.
var CATEGORIAS = ['capacetes', 'prateleira', 'diaria', 'locacao', 'tipo', 'roupa', 'livre'];
// Mapa das prateleiras PERSISTENTE: sobrevive ao "Salvar e limpar" (a página de
// vendas depende dele). Atualizado a cada contagem de prateleira recebida.
var ABA_MAPA = 'mapa_prateleiras';
var CAB_MAPA = ['loc', 'contagem_id', 'recebido_em', 'conferente', 'codigo', 'descricao', 'qtd'];
// Ações só de leitura: rodam sem o lock (não disputam com o envio das coletoras).
var LEITURAS = ['ping', 'getMapa', 'getBase', 'getBases', 'getBasesHist', 'getEstoque',
                'getData', 'getItens', 'getItensLote', 'migrarHist'];
var CAB_ITENS = ['contagem_id', 'codigo', 'descricao', 'local', 'saldo',
                 'contado', 'extra', 'tamanho', 'cor'];
// Códigos do fabricante (fora do catálogo carregado) — só auditoria, não contam
var ABA_FAB = 'fabricante';
var CAB_FAB = ['contagem_id', 'codigo', 'bipes', 'primeiro', 'ultimo', 'modo'];
// Pasta do Drive onde o "Salvar e limpar" arquiva a cópia da planilha
var PASTA_ARQUIVO = '1Ba5GXsPF9UgUUFZlejqpBWeGl2r_W_QZ';

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

// Datas SEMPRE como texto ISO (UTC). Gravar `new Date()` na célula passava pelo
// fuso da planilha e voltava 7h atrasado no getValues (bug achado 22/09/26).
function _agoraISO() { return new Date().toISOString(); }
function _iso(v) { return (v instanceof Date) ? v.toISOString() : (v || _agoraISO()); }

// Migração 22/09/26: converte as células de data gravadas como Date (deslocadas
// pelo fuso da planilha) em texto ISO correto. Mede o desvio REAL com um
// ida-e-volta numa aba de rascunho e aplica a todas. Idempotente: só toca célula
// que ainda é Date. body.teste=true só mede e devolve o desvio.
function _corrigirDatas(body) {
  if (!_ehAdmin(body)) return _json({ok: false, erro: 'senha inválida'});
  var ss = SpreadsheetApp.getActive();
  var tmp = ss.insertSheet('_tz_tmp_' + Date.now());
  var ref = new Date();
  tmp.getRange(1, 1).setValue(ref);
  SpreadsheetApp.flush();
  var lido = tmp.getRange(1, 1).getValue();
  ss.deleteSheet(tmp);
  var desvioMs = (lido instanceof Date) ? ref.getTime() - lido.getTime() : 0;
  var alvos = [[ABA_CONTAGENS, CAB_CONTAGENS, 'recebido_em'],
               [ABA_BASES_META, CAB_BASES_META, 'enviado_em'],
               [ABA_HIST, CAB_HIST, 'enviado_em'],
               [ABA_EST_META, CAB_EST_META, 'gerado_em']];
  var resumo = {ok: true, desvio_h: Math.round(desvioMs / 36e4) / 10, teste: !!body.teste, abas: {}};
  if (body.teste) return _json(resumo);
  alvos.forEach(function (a) {
    var sh = _aba(a[0], a[1]), col = a[1].indexOf(a[2]) + 1, n = sh.getLastRow() - 1, feitas = 0;
    // coluna inteira como texto: linhas novas (appendRow) herdam e não viram Date
    sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
    if (n < 1) { resumo.abas[a[0]] = 0; return; }
    var rng = sh.getRange(2, col, n, 1), vals = rng.getValues();
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (v instanceof Date) { vals[i][0] = new Date(v.getTime() + desvioMs).toISOString(); feitas++; }
    }
    if (feitas) rng.setValues(vals);
    resumo.abas[a[0]] = feitas;
  });
  return _json(resumo);
}

// Apaga UMA contagem (e seus itens/fabricante) — uso administrativo (senha),
// p.ex. pra tirar contagem de teste da planilha.
function _apagarContagemAdmin(body) {
  if (!_ehAdmin(body)) return _json({ok: false, erro: 'senha inválida'});
  if (!body.id) return _json({ok: false, erro: 'id obrigatório'});
  var shC = _aba(ABA_CONTAGENS, CAB_CONTAGENS), shI = _aba(ABA_ITENS, CAB_ITENS);
  _apagarContagem(shC, shI, String(body.id));   // já limpa itens + fabricante
  // a contagem apagada pode ser a que o mapa mostra: refaz aquela prateleira
  // com o que sobrou (sem nenhuma, a prateleira sai do mapa)
  var locs = _mapaTirarContagem(String(body.id));
  locs.forEach(function (l) { _mapaRefazerLoc(l); });
  try { CacheService.getScriptCache().remove('mapa_prateleiras'); } catch (err) {}
  return _json({ok: true, id: body.id, mapa_refeito: locs});
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return _json({ok: false, erro: 'corpo inválido'}); }
  if (LEITURAS.indexOf(body.action) >= 0) return _ler(body);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (body.action === 'upload_base') return _uploadBase(body);
    if (body.action === 'deleteBase') return _deleteBase(body);
    if (body.action === 'upload_estoque') return _uploadEstoque(body);
    if (body.action === 'arquivar_limpar') return _arquivarLimpar(body);
    if (body.action === 'corrigir_datas') return _corrigirDatas(body);
    if (body.action === 'apagar_contagem') return _apagarContagemAdmin(body);
    if (body.action === 'reconstruir_mapa') return _reconstruirMapaAdmin(body);
    if (!body.token || body.token !== _segredo('TOKEN_ENVIO')) {
      return _json({ok: false, erro: 'token inválido'});
    }
    var c = body.contagem || {};
    if (!c.id || !c.modo || !c.itens || !c.itens.length) {
      return _json({ok: false, erro: 'contagem incompleta'});
    }
    var shC = _aba(ABA_CONTAGENS, CAB_CONTAGENS);
    _garantirCab(shC, CAB_CONTAGENS);
    var shI = _aba(ABA_ITENS, CAB_ITENS);
    var categoria = CATEGORIAS.indexOf(c.categoria) >= 0 ? c.categoria
                    : _categoriaDe(c.modo, c.locacao, '', _tiposDaBase());

    // idempotência: reenvio da mesma contagem substitui a anterior
    _apagarContagem(shC, shI, c.id);

    // Na prateleira o saldo que vem é o TOTAL da peça na loja e o contado é só o
    // daquela prateleira: comparar os dois dava divergência falsa em toda peça
    // espalhada. Lá a divergência não existe (o confronto é na aba Capacetes).
    var compara = categoria !== 'prateleira';
    var divergencias = 0;
    var linhas = c.itens.map(function (it) {
      var saldo = (it.saldo === undefined || it.saldo === null) ? '' : it.saldo;
      var contado = (it.qtd === undefined || it.qtd === null) ? '' : it.qtd;
      if (compara && saldo !== '' && contado !== '' && Number(saldo) !== Number(contado)) {
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
    var recebido = _agoraISO();
    shC.appendRow([c.id, recebido, c.origem || '', c.modo, c.nome || '',
                   c.numContagem || '', c.conferente || '', c.data || '',
                   c.status || '', c.locacao || '', c.itens.length,
                   c.totalBipes || '', divergencias, fab.length, categoria]);
    if (categoria === 'prateleira' && _ehPrateleira(c.locacao)) {
      _mapaGravar(String(c.locacao), String(c.id), recebido, c.conferente || '',
                  c.itens.map(function (it) {
                    return {codigo: it.codigo || '', descricao: it.desc || '', qtd: Number(it.qtd) || 0};
                  }));
    }
    try { CacheService.getScriptCache().remove('mapa_prateleiras'); } catch (err) {}
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

// Dois níveis. O TOKEN_ENVIO é público (vai no app das coletoras, repo aberto):
// com ele só se envia contagem e se lê base/mapa. Trocar ou apagar base, subir o
// estoque e ler o estoque (tem custo e preço) exige a senha ou o TOKEN_ADMIN.
function _ehAdmin(p) {
  var adm = _segredo('TOKEN_ADMIN');
  return !!((p.senha && p.senha === _segredo('SENHA')) ||
            (adm && p.token_admin && p.token_admin === adm));
}
function _ehColetora(p) {
  return _ehAdmin(p) || !!(p.token && p.token === _segredo('TOKEN_ENVIO'));
}

// Garante o cabeçalho completo numa aba antiga (coluna nova no fim).
function _garantirCab(sh, cab) {
  if (sh.getLastColumn() >= cab.length) return;
  sh.getRange(1, 1, 1, cab.length).setValues([cab]);
}

// Tipos da base por Tipo vigente (pra reconhecer contagem antiga do modo Tipo,
// cujo "local" é o nome do tipo).
function _tiposDaBase() {
  var cfg = BASES.tipo, sh = _aba(cfg.aba, cfg.cab), v = sh.getDataRange().getValues(), out = {};
  var k = cfg.cab.indexOf('tipo');
  for (var i = 1; i < v.length; i++) { var t = String(v[i][k] || '').trim().toUpperCase(); if (t) out[t] = true; }
  out['SEM TIPO'] = true;
  return out;
}
// Categoria da contagem. Gravada (v30+) vence; nas antigas: modo próprio
// (diaria/roupa/livre), "Prateleira N"/"Exposição" = prateleira, o grupo único
// "TODOS OS CAPACETES" = capacetes, nome de tipo da base = tipo, o resto = locação.
function _categoriaDe(modo, locacao, gravada, tipos) {
  if (CATEGORIAS.indexOf(String(gravada || '')) >= 0) return String(gravada);
  modo = String(modo || '');
  if (modo && modo !== 'locacao') return CATEGORIAS.indexOf(modo) >= 0 ? modo : 'livre';
  var loc = String(locacao || '').trim();
  if (_ehPrateleira(loc)) return 'prateleira';
  if (loc.toUpperCase() === 'TODOS OS CAPACETES') return 'capacetes';
  if (tipos && tipos[loc.toUpperCase()]) return 'tipo';
  return 'locacao';
}

function _uploadBase(body) {
  if (!_ehAdmin(body)) return _json({ok: false, erro: 'não autorizado'});
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
  shM.appendRow([body.tipo, agora.toISOString(), body.por || '', body.pecas.length]);
  // guarda a versão no histórico (mantém TODAS)
  var versaoId = body.tipo + '-' + agora.getTime();
  _aba(ABA_HIST, CAB_HIST).appendRow([versaoId, body.tipo, agora.toISOString(),
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
  shM.appendRow([tipo, _iso(ver && ver.enviado_em),
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
    _aba(ABA_HIST, CAB_HIST).appendRow([vid, t, ts.toISOString(), m.por || '', pecas.length]);
    _escreverBlob(vid, pecas);
    feitos.push({tipo: t, versao_id: vid, n: pecas.length});
  });
  return feitos;
}
// ===== BASE DO DIA (estoque enriquecido p/ o módulo de Análises) =====
var ABA_EST_META = 'estoque_meta';
var CAB_EST_META = ['gerado_em', 'n_pecas', 'por'];
function _uploadEstoque(body) {
  if (!_ehAdmin(body)) return _json({ok: false, erro: 'não autorizado'});
  if (!body.pecas || !body.pecas.length) return _json({ok: false, erro: 'sem peças'});
  _apagarBlob('estoque-atual');
  _escreverBlob('estoque-atual', body.pecas);
  var shM = _aba(ABA_EST_META, CAB_EST_META);
  shM.clearContents();
  shM.appendRow(CAB_EST_META);
  shM.appendRow([_agoraISO(), body.pecas.length, body.por || '']);
  return _json({ok: true, n: body.pecas.length});
}
function _metaEstoque() {
  var shM = _aba(ABA_EST_META, CAB_EST_META), v = shM.getDataRange().getValues();
  return v.length > 1 ? {gerado_em: v[1][0], n: v[1][1], por: v[1][2]} : null;
}
// "Salvar e limpar": copia a planilha INTEIRA pra PASTA_ARQUIVO (Drive) e zera
// SÓ as contagens (contagens + itens + fabricante). Bases e estoque ficam.
// body.teste = arquiva sem limpar (validação em produção sem destruir dados).
// body.categorias = ['capacetes','prateleira'...] limpa SÓ as contagens dessas
// categorias (as outras seguem em andamento). Sem o campo = limpa tudo (compat).
// O mapa das prateleiras NUNCA é limpo aqui.
function _arquivarLimpar(body) {
  if (!_ehAdmin(body)) return _json({ok: false, erro: 'senha inválida'});
  var cats = (body.categorias || []).filter(function (c) { return CATEGORIAS.indexOf(c) >= 0; });
  var ss = SpreadsheetApp.getActive();
  var nome = 'Inventário MC — arquivo ' +
             Utilities.formatDate(new Date(), 'America/Sao_Paulo', "dd-MM-yyyy HH'h'mm") +
             (cats.length ? ' (' + cats.join(', ') + ')' : '');
  var copia = ss.copy(nome);
  DriveApp.getFileById(copia.getId()).moveTo(DriveApp.getFolderById(PASTA_ARQUIVO));
  var url = copia.getUrl();
  if (body.teste) return _json({ok: true, teste: true, url: url, nome: nome});
  if (!cats.length) {
    [[ABA_CONTAGENS, CAB_CONTAGENS], [ABA_ITENS, CAB_ITENS], [ABA_FAB, CAB_FAB]]
      .forEach(function (par) {
        var sh = _aba(par[0], par[1]);
        sh.clearContents();
        sh.appendRow(par[1]);
      });
    return _json({ok: true, url: url, nome: nome,
                  limpo: ['contagens', 'itens', 'fabricante']});
  }
  var shC = _aba(ABA_CONTAGENS, CAB_CONTAGENS), vals = shC.getDataRange().getValues();
  var iMod = CAB_CONTAGENS.indexOf('modo'), iLoc = CAB_CONTAGENS.indexOf('locacao'),
      iCat = CAB_CONTAGENS.indexOf('categoria'), tipos = _tiposDaBase(), sai = {}, n = 0;
  for (var i = 1; i < vals.length; i++) {
    var cat = _categoriaDe(vals[i][iMod], vals[i][iLoc], vals[i][iCat], tipos);
    if (cats.indexOf(cat) >= 0) { sai[String(vals[i][0])] = true; n++; }
  }
  [[ABA_CONTAGENS, CAB_CONTAGENS], [ABA_ITENS, CAB_ITENS], [ABA_FAB, CAB_FAB]]
    .forEach(function (par) { _filtrarAba(par[0], par[1], sai); });
  return _json({ok: true, url: url, nome: nome, categorias: cats, contagens_limpas: n});
}
// Reescreve a aba sem as linhas cujo 1º campo (id/contagem_id) está em `sai`.
// Uma escrita só (deleteRow linha a linha estoura o tempo com milhares de itens).
function _filtrarAba(nome, cab, sai) {
  var sh = _aba(nome, cab), v = sh.getDataRange().getValues();
  if (v.length < 2) return;
  var larg = Math.max(cab.length, v[0].length);
  var fica = v.slice(1).filter(function (r) { return !sai[String(r[0])]; })
              .map(function (r) { while (r.length < larg) r.push(''); return r; });
  sh.getRange(2, 1, v.length - 1, v[0].length).clearContent();
  if (fica.length) sh.getRange(2, 1, fica.length, larg).setValues(fica);
}

function _deleteBase(body) {
  if (!_ehAdmin(body)) return _json({ok: false, erro: 'não autorizado'});
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
  return _ler((e && e.parameter) || {});
}

function _ler(p) {
  if (p.action === 'ping') return _json({ok: true, servico: 'inventario-mc',
    tz_script: Session.getScriptTimeZone(),
    tz_planilha: SpreadsheetApp.getActive().getSpreadsheetTimeZone()});
  if (p.action === 'getMapa') {
    // Mapa das prateleiras (página pública das vendas): contagem mais recente
    // de cada prateleira + itens, numa chamada só. Cache 3 min (invalidado ao
    // receber contagem). &nocache=1 força recalcular.
    if (!_ehColetora(p)) return _json({ok: false, erro: 'não autorizado'});
    var cache = CacheService.getScriptCache(), hit = p.nocache ? null : cache.get('mapa_prateleiras');
    if (!hit) {
      hit = JSON.stringify(_mapaPrateleiras());
      try { cache.put('mapa_prateleiras', hit, 180); } catch (err) {}
    }
    return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  }
  if (p.action === 'getBase') {
    if (!_ehColetora(p)) return _json({ok: false, erro: 'não autorizado'});
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
    if (!_ehColetora(p)) return _json({ok: false, erro: 'não autorizado'});
    return _json({ok: true, bases: {diaria: _metaBase('diaria'),
                                    locacao: _metaBase('locacao'),
                                    tipo: _metaBase('tipo'),
                                    capacetes: _metaBase('capacetes')}});
  }
  if (p.action === 'migrarHist') {
    if (!_ehAdmin(p)) return _json({ok: false, erro: 'senha inválida'});
    return _json({ok: true, migrados: _migrarHist()});
  }
  if (p.action === 'getEstoque') {
    if (!_ehAdmin(p)) return _json({ok: false, erro: 'não autorizado'});
    return _json({ok: true, meta: _metaEstoque(), pecas: _lerBlob('estoque-atual') || []});
  }
  if (p.action === 'getBasesHist') {
    if (!_ehColetora(p)) return _json({ok: false, erro: 'não autorizado'});
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
  if (!_ehAdmin(p)) return _json({ok: false, erro: 'senha inválida'});
  if (p.action === 'getData') {
    var shC = _aba(ABA_CONTAGENS, CAB_CONTAGENS);
    var vals = shC.getDataRange().getValues();
    var out = [], tipos = null;
    for (var i = 1; i < vals.length; i++) {
      var o = {};
      CAB_CONTAGENS.forEach(function (col, k) { o[col] = vals[i][k] === undefined ? '' : vals[i][k]; });
      if (CATEGORIAS.indexOf(String(o.categoria)) < 0) {
        if (o.modo === 'locacao' && !tipos) tipos = _tiposDaBase();
        o.categoria = _categoriaDe(o.modo, o.locacao, o.categoria, tipos);
      }
      // contagens de prateleira antigas gravaram divergência contra o saldo total
      if (o.categoria === 'prateleira') o.divergencias = 0;
      o.recebido_em = _iso(o.recebido_em);
      out.push(o);
    }
    return _json({ok: true, contagens: out});
  }
  if (p.action === 'getItensLote') {
    // várias contagens numa passada só (o dashboard abria uma chamada por contagem)
    var ids = {}, lista = String(p.ids || '').split(',').filter(String);
    lista.forEach(function (x) { ids[x] = true; });
    var resI = {}, resF = {};
    lista.forEach(function (x) { resI[x] = []; resF[x] = []; });
    var tI = _aba(ABA_ITENS, CAB_ITENS).getDataRange().getValues();
    for (var a = 1; a < tI.length; a++) {
      var idI = String(tI[a][0]); if (!ids[idI]) continue;
      var oi = {}; CAB_ITENS.forEach(function (col, k) { oi[col] = tI[a][k]; }); resI[idI].push(oi);
    }
    var tF = _aba(ABA_FAB, CAB_FAB).getDataRange().getValues();
    for (var b2 = 1; b2 < tF.length; b2++) {
      var idF = String(tF[b2][0]); if (!ids[idF]) continue;
      var of = {}; CAB_FAB.forEach(function (col, k) { of[col] = tF[b2][k]; }); resF[idF].push(of);
    }
    return _json({ok: true, itens: resI, fab: resF});
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

function _ehPrateleira(loc) {
  loc = String(loc || '');
  return /^Prateleira \d+$/.test(loc) || loc === 'Exposição';
}
// ===== MAPA DAS PRATELEIRAS (persistente) =====
// Lê a aba mapa_prateleiras. Na 1ª vez (aba vazia) semeia com as contagens que
// ainda existem na planilha — depois disso o "Salvar e limpar" não mexe mais nele.
function _mapaPrateleiras() {
  var sh = _aba(ABA_MAPA, CAB_MAPA), v = sh.getDataRange().getValues();
  if (v.length < 2) { _mapaSemear(); v = sh.getDataRange().getValues(); }
  var porLoc = {};
  for (var i = 1; i < v.length; i++) {
    var loc = String(v[i][0]);
    if (!porLoc[loc]) porLoc[loc] = {loc: loc, contagem_id: String(v[i][1]), recebido_em: _iso(v[i][2]),
                                    conferente: String(v[i][3] || ''), itens: []};
    var q = Number(v[i][6]) || 0;
    if (v[i][4] !== '' && q > 0) porLoc[loc].itens.push({codigo: String(v[i][4]), descricao: String(v[i][5] || ''), qtd: q});
  }
  return {ok: true, gerado_em: _agoraISO(), persistente: true,
          locais: Object.keys(porLoc).map(function (l) { return porLoc[l]; })};
}
// Grava o retrato de UMA prateleira, se a contagem for mais nova que a do mapa
// (ou a mesma, no reenvio). Prateleira contada vazia vira uma linha-marcador.
function _mapaGravar(loc, contagemId, recebido, conferente, itens) {
  var sh = _aba(ABA_MAPA, CAB_MAPA), v = sh.getDataRange().getValues(), atual = null;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === loc) { atual = v[i]; break; }
  if (atual && String(atual[1]) !== contagemId && _iso(atual[2]) > recebido) return false;
  _filtrarMapa(function (r) { return String(r[0]) !== loc; });
  var linhas = itens.filter(function (it) { return it.qtd > 0; }).map(function (it) {
    return [loc, contagemId, recebido, conferente, it.codigo, it.descricao, it.qtd];
  });
  if (!linhas.length) linhas = [[loc, contagemId, recebido, conferente, '', '', 0]];
  var r0 = sh.getLastRow() + 1;
  // data e código como TEXTO: o Sheets converteria o ISO em Date (fuso, bug 22/09)
  // e o código numérico perderia zero à esquerda
  sh.getRange(r0, 1, linhas.length, CAB_MAPA.length).setNumberFormat('@');
  sh.getRange(r0, 1, linhas.length, CAB_MAPA.length).setValues(linhas);
  return true;
}
function _filtrarMapa(manter) {
  var sh = _aba(ABA_MAPA, CAB_MAPA), v = sh.getDataRange().getValues();
  if (v.length < 2) return;
  var fica = v.slice(1).filter(manter);
  sh.getRange(2, 1, v.length - 1, CAB_MAPA.length).clearContent();
  if (fica.length) sh.getRange(2, 1, fica.length, CAB_MAPA.length).setValues(fica);
}
// Tira do mapa as linhas de uma contagem; devolve as prateleiras afetadas.
function _mapaTirarContagem(id) {
  var v = _aba(ABA_MAPA, CAB_MAPA).getDataRange().getValues(), locs = {};
  for (var i = 1; i < v.length; i++) if (String(v[i][1]) === id) locs[String(v[i][0])] = true;
  if (Object.keys(locs).length) _filtrarMapa(function (r) { return String(r[1]) !== id; });
  return Object.keys(locs);
}
// Refaz uma prateleira com a contagem mais recente que ainda existe dela.
function _mapaRefazerLoc(loc) {
  var m = _mapaDeContagens().locais.filter(function (x) { return x.loc === loc; })[0];
  if (m) _mapaGravar(loc, m.id, m.recebido_em, m.conferente, m.itens);
}
function _mapaSemear() {
  _mapaDeContagens().locais.forEach(function (m) {
    _mapaGravar(m.loc, m.id, m.recebido_em, m.conferente, m.itens);
  });
}
// Admin: funde no mapa as contagens de prateleira que estão na planilha (só
// substitui prateleira por contagem MAIS NOVA; nunca apaga prateleira do mapa).
function _reconstruirMapaAdmin(body) {
  if (!_ehAdmin(body)) return _json({ok: false, erro: 'senha inválida'});
  var feitas = [];
  _mapaDeContagens().locais.forEach(function (m) {
    if (_mapaGravar(m.loc, m.id, m.recebido_em, m.conferente, m.itens)) feitas.push(m.loc);
  });
  try { CacheService.getScriptCache().remove('mapa_prateleiras'); } catch (err) {}
  return _json({ok: true, atualizadas: feitas});
}

// Contagem MAIS RECENTE de cada prateleira (modo locacao, locacao 'Prateleira N'
// ou 'Exposição') + itens com contado>0, direto das contagens. Uma passada em
// cada aba. Serve pra semear/refazer o mapa persistente.
function _mapaDeContagens() {
  var shC = _aba(ABA_CONTAGENS, CAB_CONTAGENS), vals = shC.getDataRange().getValues();
  var iLoc = CAB_CONTAGENS.indexOf('locacao'), iRec = CAB_CONTAGENS.indexOf('recebido_em'),
      iMod = CAB_CONTAGENS.indexOf('modo'), iConf = CAB_CONTAGENS.indexOf('conferente');
  var recente = {};
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (String(r[iMod]) !== 'locacao' || !_ehPrateleira(r[iLoc])) continue;
    var t = new Date(r[iRec]).getTime() || 0, loc = String(r[iLoc]);
    if (!recente[loc] || t > recente[loc].t) {
      recente[loc] = {t: t, id: String(r[0]), recebido_em: _iso(r[iRec]), conferente: String(r[iConf] || '')};
    }
  }
  var porId = {};
  Object.keys(recente).forEach(function (l) { porId[recente[l].id] = l; });
  var shI = _aba(ABA_ITENS, CAB_ITENS), tudo = shI.getDataRange().getValues(), itens = {};
  for (var j = 1; j < tudo.length; j++) {
    var loc2 = porId[String(tudo[j][0])];
    if (!loc2) continue;
    var q = Number(tudo[j][5]) || 0;
    if (q <= 0) continue;
    (itens[loc2] = itens[loc2] || []).push({codigo: String(tudo[j][1]), descricao: String(tudo[j][2] || ''), qtd: q});
  }
  var locais = Object.keys(recente).map(function (l) {
    return {loc: l, id: recente[l].id, recebido_em: recente[l].recebido_em, conferente: recente[l].conferente, itens: itens[l] || []};
  });
  return {ok: true, gerado_em: _agoraISO(), locais: locais};
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

/** Rode UMA vez no editor pra conceder o escopo do Drive (Salvar e limpar). */
function autorizarDrive() {
  var pasta = DriveApp.getFolderById(PASTA_ARQUIVO);
  Logger.log('OK — pasta de arquivo: ' + pasta.getName());
}

