/**
 * Portal do Recanto da Siriema
 * Planilha Google + Apps Script. Cole este arquivo como "Codigo.gs".
 *
 * Primeira vez: rode a função setup() e anote a senha provisória
 * que aparece no Registro de execução.
 */

const TZ = 'America/Sao_Paulo';

const ABAS = {
  Config:      ['Chave', 'Valor', 'Descricao'],
  Usuarios:    ['ID', 'Usuario', 'Nome', 'Perfil', 'Salt', 'Hash', 'Ativo', 'TrocarSenha', 'CriadoEm', 'UltimoAcesso'],
  Associados:  ['ID', 'Chacara', 'Nome', 'Celular', 'Mensalidade', 'CobrarDesde', 'Ativo', 'Obs', 'CriadoEm', 'Regularizacao'],
  Lancamentos: ['ID', 'Data', 'Tipo', 'Categoria', 'Campanha', 'Chacara', 'PagoPor', 'MesReferencia', 'Descricao', 'Valor',
                'Status', 'Motivo', 'CriadoPor', 'CriadoEm', 'CanceladoPor', 'CanceladoEm', 'Comprovante'],
  Campanhas:   ['ID', 'Nome', 'Descricao', 'Meta', 'Inicio', 'Status', 'CriadoEm', 'Oculta'],
  Reservas:    ['ID', 'Dia', 'Periodo', 'Chacara', 'Nome', 'Celular', 'Tipo', 'Taxa', 'TaxaPaga', 'Pix', 'Status', 'Motivo',
                'PedidoEm', 'Lancamento'],
  Avisos:      ['ID', 'CriadoEm', 'Para', 'Celular', 'Mensagem', 'Reserva', 'Status'],
  Terrenos:    ['ID', 'Numero', 'Associado', 'Situacao', 'Cobrado', 'Motivo', 'AtualizadoEm'],
  Historico:   ['Data', 'Usuario', 'Acao', 'Detalhe']
};

// Colunas guardadas como texto (evita o Sheets transformar "17/18" ou "2026-10" em data)
const COLUNAS_TEXTO = {
  Config: ['Valor'],
  Usuarios: ['CriadoEm', 'UltimoAcesso'],
  Associados: ['Chacara', 'Celular', 'CobrarDesde', 'CriadoEm'],
  Lancamentos: ['Data', 'Chacara', 'MesReferencia', 'CriadoEm', 'CanceladoEm'],
  Campanhas: ['Inicio', 'CriadoEm'],
  Reservas: ['Dia', 'Chacara', 'Celular', 'PedidoEm'],
  Avisos: ['CriadoEm', 'Celular'],
  Terrenos: ['Numero', 'AtualizadoEm']
};

const CONFIG_PADRAO = [
  ['nomeAssociacao', 'Recanto da Siriema', 'Nome que aparece no topo do portal'],
  ['saldoInicial', '0', 'Saldo do caixa no dia da abertura (R$)'],
  ['dataAbertura', '', 'Dia em que o caixa começou a ser controlado aqui (AAAA-MM-DD)'],
  ['mensalidadePadrao', '60', 'Valor padrão da mensalidade (R$)'],
  ['diaVencimento', '10', 'Dia do mês em que a mensalidade vence'],
  ['mesesAtraso', '2', 'A partir de quantos meses em aberto a chácara fica "Atrasada"'],
  ['mesesInadimplente', '4', 'A partir de quantos meses em aberto a chácara fica "Inadimplente"'],
  ['taxaReserva', '50', 'Taxa de manutenção da reserva (R$) para não associados, atrasados e inadimplentes'],
  ['limiteReservasMes', '', 'Máximo de reservas por chácara no mês (em branco = sem limite)'],
  ['periodos', 'Manhã,Tarde,Noite,Dia inteiro', 'Períodos da reserva, separados por vírgula'],
  ['chavePix', '', 'Chave Pix da associação para pagamento da taxa'],
  ['avisoInadimplente', 'Sua chácara está com mensalidades em aberto há 4 meses ou mais. Por isso, a reserva da área de lazer está suspensa. Procure a diretoria para regularizar. (Texto provisório: a diretoria vai escrever o texto final.)', 'Aviso que aparece quando um inadimplente tenta reservar'],
  ['conferirCelular', 'sim', 'Na reserva, conferir o celular com o cadastro da chácara (sim/não)']
];

/* ===================================================================
   SETUP E UTILIDADES
   =================================================================== */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss && props.getProperty('PLANILHA_ID')) ss = SpreadsheetApp.openById(props.getProperty('PLANILHA_ID'));
  if (!ss) throw new Error('Abra o Apps Script pela planilha (Extensões > Apps Script) e rode de novo.');
  props.setProperty('PLANILHA_ID', ss.getId());

  Object.keys(ABAS).forEach(function (nome) {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    const cab = ABAS[nome];
    const atual = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    if (!atual.length || !atual[0]) {
      sh.getRange(1, 1, 1, cab.length).setValues([cab]);
    } else {
      // acrescenta colunas novas que ainda não existem, sem mexer nas antigas
      cab.forEach(function (c) {
        if (atual.indexOf(c) < 0) sh.getRange(1, sh.getLastColumn() + 1).setValue(c);
      });
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#003912').setFontColor('#ffffff');
    (COLUNAS_TEXTO[nome] || []).forEach(function (c) {
      const i = cabecalho_(sh).indexOf(c);
      if (i >= 0) sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  });
  const padrao = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1') || ss.getSheetByName('Planilha1');
  if (padrao && padrao.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(padrao);

  // Config padrão (só as chaves que faltam)
  const cfg = tabela_('Config');
  const existentes = cfg.rows.map(function (r) { return r.Chave; });
  CONFIG_PADRAO.forEach(function (c) {
    if (existentes.indexOf(c[0]) < 0) {
      const valor = c[0] === 'dataAbertura' ? hoje_() : c[1];
      inserir_('Config', { Chave: c[0], Valor: valor, Descricao: c[2] });
    }
  });

  // Primeiro usuário (administrador)
  const us = tabela_('Usuarios');
  if (!us.rows.length) {
    const senha = senhaProvisoria_();
    const salt = Utilities.getUuid();
    inserir_('Usuarios', {
      ID: proximoId_('U'), Usuario: 'diretoria', Nome: 'Diretoria', Perfil: 'admin',
      Salt: salt, Hash: hash_(senha, salt), Ativo: 'sim', TrocarSenha: 'sim', CriadoEm: agora_(), UltimoAcesso: ''
    });
    Logger.log('==============================================');
    Logger.log('Usuário: diretoria');
    Logger.log('Senha provisória: ' + senha);
    Logger.log('No primeiro acesso o sistema pede uma senha nova.');
    Logger.log('==============================================');
  } else {
    Logger.log('Setup conferido. Nada foi apagado. Usuários já existem.');
  }
  const criados = criarTerrenosDoCadastro_();
  if (criados) Logger.log('Terrenos criados a partir do cadastro: ' + criados);
  registrar_('sistema', 'setup', 'Estrutura conferida');
}

/** Quando alguém edita a planilha direto, o portal deixa de usar os dados guardados em cache. */
function onEdit() { mudouDados_(); }

/** Recupera o acesso do primeiro administrador (rode pelo editor). */
function redefinirSenhaAdmin() {
  const us = tabela_('Usuarios');
  const adm = us.rows.filter(function (r) { return r.Perfil === 'admin'; })[0];
  if (!adm) throw new Error('Nenhum administrador encontrado. Rode setup().');
  const senha = senhaProvisoria_();
  const salt = Utilities.getUuid();
  atualizar_('Usuarios', adm._linha, { Salt: salt, Hash: hash_(senha, salt), Ativo: 'sim', TrocarSenha: 'sim' });
  CacheService.getScriptCache().put('bloq_usuario_' + adm.ID, '1', 21600);
  CacheService.getScriptCache().remove('tent_' + String(adm.Usuario).toLowerCase());
  Logger.log('Usuário: ' + adm.Usuario + '  Senha provisória: ' + senha);
}

function planilha_() {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILHA_ID');
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function aba_(nome) {
  const sh = planilha_().getSheetByName(nome);
  if (!sh) throw new Error('A aba "' + nome + '" não existe. Rode setup() no editor.');
  return sh;
}

function cabecalho_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

function normaliza_(nomeAba, coluna, v) {
  if (v instanceof Date) {
    if (coluna === 'MesReferencia' || coluna === 'CobrarDesde') return Utilities.formatDate(v, TZ, 'yyyy-MM');
    if (coluna === 'CriadoEm' || coluna === 'PedidoEm' || coluna === 'CanceladoEm' || coluna === 'UltimoAcesso' || coluna === 'Data' && nomeAba === 'Historico')
      return Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss");
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  return v;
}

// Memória por execução: cada aba é lida uma vez por chamada; qualquer gravação limpa a memória.
var MEMO_ = {};
function limpaMemo_() { MEMO_ = {}; }
function mudouDados_() {
  limpaMemo_();
  try { PropertiesService.getScriptProperties().setProperty('versao_dados', String(Date.now())); } catch (e) { }
}
function versaoDados_() { return PropertiesService.getScriptProperties().getProperty('versao_dados') || '0'; }
function emCache_(chave, fn) {
  const cache = CacheService.getScriptCache();
  const k = chave + '_' + versaoDados_();
  try { const c = cache.get(k); if (c) return JSON.parse(c); } catch (e) { }
  const r = fn();
  try { const t = JSON.stringify(r); if (t.length < 95000) cache.put(k, t, 600); } catch (e) { }
  return r;
}

function tabela_(nome) {
  if (MEMO_[nome]) return { sh: MEMO_[nome].sh, cab: MEMO_[nome].cab, rows: MEMO_[nome].rows.map(function (r) { return Object.assign({}, r); }) };
  const sh = aba_(nome);
  const dados = sh.getDataRange().getValues();
  const cab = (dados[0] || []).map(String);
  const rows = [];
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    if (linha.every(function (c) { return c === '' || c === null; })) continue;
    const o = { _linha: i + 1 };
    cab.forEach(function (c, j) { o[c] = normaliza_(nome, c, linha[j]); });
    rows.push(o);
  }
  MEMO_[nome] = { sh: sh, cab: cab, rows: rows };
  return { sh: sh, cab: cab, rows: rows.map(function (r) { return Object.assign({}, r); }) };
}

function seguro_(nomeAba, coluna, v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  let s = String(v);
  if ((COLUNAS_TEXTO[nomeAba] || []).indexOf(coluna) >= 0) return "'" + s;
  if (/^[=+\-@]/.test(s)) return "'" + s;
  return s;
}

function inserir_(nomeAba, obj) {
  const sh = aba_(nomeAba);
  const cab = cabecalho_(sh);
  sh.appendRow(cab.map(function (c) { return seguro_(nomeAba, c, obj[c]); }));
  mudouDados_();
}

function inserirLote_(nomeAba, objs) {
  if (!objs.length) return;
  const sh = aba_(nomeAba);
  const cab = cabecalho_(sh);
  const linhas = objs.map(function (o) { return cab.map(function (c) { return seguro_(nomeAba, c, o[c]); }); });
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, cab.length).setValues(linhas);
  mudouDados_();
}

function atualizar_(nomeAba, linha, parcial) {
  const sh = aba_(nomeAba);
  const cab = cabecalho_(sh);
  Object.keys(parcial).forEach(function (c) {
    const j = cab.indexOf(c);
    if (j >= 0) sh.getRange(linha, j + 1).setValue(seguro_(nomeAba, c, parcial[c]));
  });
  if (nomeAba !== 'Usuarios' && nomeAba !== 'Historico') mudouDados_(); else limpaMemo_();
}

function proximoId_(prefixo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const chave = 'cont_' + prefixo;
    const n = Number(props.getProperty(chave) || 0) + 1;
    props.setProperty(chave, String(n));
    return prefixo + ('00000' + n).slice(-5);
  } finally { lock.releaseLock(); }
}

function hoje_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function agora_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function mesAtual_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM'); }
function diaDoMes_() { return Number(Utilities.formatDate(new Date(), TZ, 'd')); }

function somaMes_(ym, n) {
  const p = ym.split('-');
  let y = Number(p[0]), m = Number(p[1]) - 1 + n;
  y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
  return y + '-' + ('0' + (m + 1)).slice(-2);
}

function numero_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/[R$\s]/g, '');
  if (!s) return 0;
  const n = s.indexOf(',') >= 0 ? Number(s.replace(/\./g, '').replace(',', '.')) : Number(s);
  return isNaN(n) ? 0 : n;
}

function texto_(v, max) {
  return String(v === null || v === undefined ? '' : v).trim().slice(0, max || 300);
}

function soDigitos_(v) { return String(v || '').replace(/\D/g, ''); }

function registrar_(usuario, acao, detalhe) {
  try {
    const sh = aba_('Historico');
    sh.appendRow([agora_(), seguro_('Historico', 'Usuario', usuario), seguro_('Historico', 'Acao', acao), seguro_('Historico', 'Detalhe', String(detalhe).slice(0, 500))]);
  } catch (e) { }
}

function config_() {
  const o = {};
  tabela_('Config').rows.forEach(function (r) { o[r.Chave] = String(r.Valor === null ? '' : r.Valor); });
  CONFIG_PADRAO.forEach(function (c) { if (!(c[0] in o)) o[c[0]] = c[1]; });
  return o;
}

function periodos_(cfg) {
  return String(cfg.periodos || '').split(',').map(function (s) { return s.trim(); }).filter(String);
}

/* ===================================================================
   LOGIN E SESSÃO
   =================================================================== */

function hash_(senha, salt) {
  let h = salt + '|' + senha;
  for (let i = 0; i < 200; i++) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8);
    h = bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  }
  return h;
}

function senhaProvisoria_() {
  const letras = 'abcdefghjkmnpqrstuvwxyz';
  const u = Utilities.getUuid().replace(/-/g, '');
  let s = '';
  for (let i = 0; i < 4; i++) s += letras[parseInt(u.substr(i * 2, 2), 16) % letras.length];
  return s + (1000 + parseInt(u.substr(10, 4), 16) % 9000);
}

function login(usuario, senha) {
  usuario = texto_(usuario, 60).toLowerCase();
  const cache = CacheService.getScriptCache();
  if (cache.get('bloqlogin_' + usuario)) throw new Error('Muitas tentativas erradas. Espere 10 minutos e tente de novo.');
  const u = tabela_('Usuarios').rows.filter(function (r) { return String(r.Usuario).toLowerCase() === usuario; })[0];
  if (!u || String(u.Ativo) !== 'sim' || hash_(String(senha || ''), u.Salt) !== u.Hash) {
    const n = Number(cache.get('tent_' + usuario) || 0) + 1;
    cache.put('tent_' + usuario, String(n), 600);
    if (n >= 5) { cache.put('bloqlogin_' + usuario, '1', 600); cache.remove('tent_' + usuario); }
    throw new Error('Usuário ou senha incorretos.');
  }
  cache.remove('tent_' + usuario);
  cache.remove('bloq_usuario_' + u.ID);
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('s_' + token, JSON.stringify({ id: u.ID, criado: Date.now() }), 21600);
  atualizar_('Usuarios', u._linha, { UltimoAcesso: agora_() });
  registrar_(u.Usuario, 'login', '');
  return { token: token, usuario: perfilPublico_(u) };
}

function perfilPublico_(u) {
  return { id: u.ID, usuario: u.Usuario, nome: u.Nome, perfil: u.Perfil, trocarSenha: String(u.TrocarSenha) === 'sim' };
}

function sessao_(token, opcoes) {
  opcoes = opcoes || {};
  const cache = CacheService.getScriptCache();
  const bruto = token ? cache.get('s_' + token) : null;
  if (!bruto) throw new Error('SESSAO_EXPIRADA: Sua sessão terminou. Entre de novo.');
  const s = JSON.parse(bruto);
  if (cache.get('bloq_usuario_' + s.id)) { cache.remove('s_' + token); throw new Error('SESSAO_EXPIRADA: Entre de novo.'); }
  const u = tabela_('Usuarios').rows.filter(function (r) { return r.ID === s.id; })[0];
  if (!u || String(u.Ativo) !== 'sim') { cache.remove('s_' + token); throw new Error('SESSAO_EXPIRADA: Usuário desativado.'); }
  if (String(u.TrocarSenha) === 'sim' && !opcoes.permitirTroca) throw new Error('SESSAO_TROCAR: Crie uma senha nova para continuar.');
  if (opcoes.admin && u.Perfil !== 'admin') throw new Error('Somente o administrador pode fazer isso.');
  cache.put('s_' + token, bruto, 21600);
  return u;
}

function sair(token) {
  if (token) CacheService.getScriptCache().remove('s_' + token);
  return true;
}

function trocarSenha(token, atual, nova) {
  const u = sessao_(token, { permitirTroca: true });
  if (hash_(String(atual || ''), u.Salt) !== u.Hash) throw new Error('A senha atual está errada.');
  nova = String(nova || '');
  if (nova.length < 6) throw new Error('A senha nova precisa ter pelo menos 6 letras ou números.');
  const salt = Utilities.getUuid();
  atualizar_('Usuarios', u._linha, { Salt: salt, Hash: hash_(nova, salt), TrocarSenha: 'não' });
  registrar_(u.Usuario, 'trocou a senha', '');
  return perfilPublico_(Object.assign({}, u, { TrocarSenha: 'não' }));
}

function quemSou(token) {
  return perfilPublico_(sessao_(token, { permitirTroca: true }));
}

/* ===================================================================
   CÁLCULOS: SALDO E ADIMPLÊNCIA
   =================================================================== */

function lancamentosAtivos_() {
  return tabela_('Lancamentos').rows.filter(function (l) { return l.Status !== 'Cancelado'; })
    .map(function (l) { l.Valor = numero_(l.Valor); l.Chacara = String(l.Chacara || ''); return l; });
}

function saldoAtual_(lans, cfg) {
  const abertura = cfg.dataAbertura || '0000-00-00';
  let s = numero_(cfg.saldoInicial);
  lans.forEach(function (l) {
    if (String(l.Data) < abertura) return;
    s += l.Tipo === 'Entrada' ? l.Valor : -l.Valor;
  });
  return Math.round(s * 100) / 100;
}

function mesesDevidos_(desde, cfg) {
  const atual = mesAtual_();
  const ultimo = diaDoMes_() > Number(cfg.diaVencimento || 10) ? atual : somaMes_(atual, -1);
  const meses = [];
  if (!desde) return meses;
  let m = String(desde).slice(0, 7);
  let guarda = 0;
  while (m <= ultimo && guarda++ < 240) { meses.push(m); m = somaMes_(m, 1); }
  return meses;
}

function situacao_(abertos, cfg) {
  if (abertos >= Number(cfg.mesesInadimplente || 4)) return 'inadimplente';
  if (abertos >= Number(cfg.mesesAtraso || 2)) return 'atrasado';
  return 'em dia';
}

/** Lista de associados com meses pagos, em aberto e situação. */
function adimplencia_(cfg, lans) {
  const assoc = tabela_('Associados').rows.filter(function (a) { return String(a.Ativo) !== 'não'; });
  const pagosPorChacara = {};
  lans.forEach(function (l) {
    if (l.Categoria !== 'Mensalidade' || !l.MesReferencia) return;
    const k = String(l.Chacara);
    (pagosPorChacara[k] = pagosPorChacara[k] || {})[String(l.MesReferencia).slice(0, 7)] = true;
  });
  return assoc.map(function (a) {
    const ch = String(a.Chacara);
    const devidos = mesesDevidos_(a.CobrarDesde, cfg);
    const pagos = pagosPorChacara[ch] || {};
    const abertos = devidos.filter(function (m) { return !pagos[m]; });
    return {
      id: a.ID, chacara: ch, nome: String(a.Nome), celular: String(a.Celular || ''),
      mensalidade: numero_(a.Mensalidade) || numero_(cfg.mensalidadePadrao),
      cobrarDesde: String(a.CobrarDesde || ''), obs: String(a.Obs || ''),
      abertos: abertos, situacao: situacao_(abertos.length, cfg)
    };
  }).sort(ordemChacara_);
}

function ordemChacara_(a, b) {
  const na = parseInt(a.chacara, 10), nb = parseInt(b.chacara, 10);
  if (na !== nb) return (isNaN(na) ? 1e9 : na) - (isNaN(nb) ? 1e9 : nb);
  return String(a.chacara).localeCompare(String(b.chacara));
}

/* ===================================================================
   PÚBLICO (sem login, sem nomes)
   =================================================================== */

function resumoPublico(ym) {
  ym = /^\d{4}-\d{2}$/.test(String(ym || '')) ? ym : mesAtual_();
  return emCache_('pub_' + ym + '_' + mesAtual_() + '_' + diaDoMes_(), function () { return resumoPublicoCalc_(ym); });
}

function resumoPublicoCalc_(ym) {
  const cfg = config_();
  const lans = lancamentosAtivos_();
  const doMes = lans.filter(function (l) { return String(l.Data).slice(0, 7) === ym; });
  const entradas = doMes.filter(function (l) { return l.Tipo === 'Entrada'; });
  const saidas = doMes.filter(function (l) { return l.Tipo === 'Saída'; });
  const soma = function (a) { return Math.round(a.reduce(function (s, l) { return s + l.Valor; }, 0) * 100) / 100; };
  const porCategoria = {};
  entradas.forEach(function (l) { porCategoria[l.Categoria] = (porCategoria[l.Categoria] || 0) + l.Valor; });
  const camps = campanhasComTotais_(lans).filter(function (c) { return !c.oculta; }).map(function (c) {
    return { id: c.id, nome: c.nome, descricao: c.descricao, meta: c.meta, arrecadado: c.arrecadado, gasto: c.gasto, status: c.status };
  });
  const nomesCamp = {};
  camps.forEach(function (c) { nomesCamp[c.id] = c.nome; });
  const linhaPublica = function (l) {
    return {
      id: l.ID, comprovante: l.Tipo === 'Saída' && !!l.Comprovante,
      data: l.Data, tipo: l.Tipo, categoria: l.Categoria, campanha: nomesCamp[l.Campanha] || '',
      chacara: '', mes: '', descricao: l.Categoria === 'Mensalidade' ? '' : String(l.Descricao || ''), valor: l.Valor
    };
  };
  return {
    nome: cfg.nomeAssociacao,
    mes: ym,
    mesAtual: mesAtual_(),
    saldoAtual: saldoAtual_(lans, cfg),
    entradasMes: soma(entradas),
    saidasMes: soma(saidas),
    entradasPorCategoria: porCategoria,
    entradas: entradas.sort(porData_).map(linhaPublica),
    saidas: saidas.sort(porData_).map(linhaPublica),
    campanhas: camps,
    atualizadoEm: agora_()
  };
}

function porData_(a, b) { return String(a.Data).localeCompare(String(b.Data)); }

function infoReserva() {
  return emCache_('info_' + hoje_(), infoReservaCalc_);
}
function infoReservaCalc_() {
  const cfg = config_();
  const chacaras = tabela_('Associados').rows.filter(function (a) { return String(a.Ativo) !== 'não'; })
    .map(function (a) { return { chacara: String(a.Chacara) }; }).sort(ordemChacara_).map(function (a) { return a.chacara; });
  return {
    periodos: periodos_(cfg),
    taxa: numero_(cfg.taxaReserva),
    limite: cfg.limiteReservasMes ? Number(cfg.limiteReservasMes) : null,
    chacaras: chacaras,
    hoje: hoje_()
  };
}

/** Dias ocupados do mês, sem nomes. */
function calendario(ym) {
  ym = /^\d{4}-\d{2}$/.test(String(ym || '')) ? ym : mesAtual_();
  return emCache_('cal_' + ym, function () { return calendarioCalc_(ym); });
}
function calendarioCalc_(ym) {
  return tabela_('Reservas').rows
    .filter(function (r) { return r.Status === 'Ativa' && String(r.Dia).slice(0, 7) === ym; })
    .map(function (r) { return { dia: r.Dia, periodo: r.Periodo, tipo: r.Tipo }; });
}

function sobrepoe_(p1, p2) { return p1 === p2 || p1 === 'Dia inteiro' || p2 === 'Dia inteiro'; }

/**
 * Pedido de reserva (público).
 * dados: { dia, periodo, chacara ('' = não associado), nome, celular }
 */
function reservar(dados) {
  dados = dados || {};
  const cfg = config_();
  const dia = texto_(dados.dia, 10);
  const periodo = texto_(dados.periodo, 30);
  const chacara = texto_(dados.chacara, 20);
  const nome = texto_(dados.nome, 80);
  const celular = texto_(dados.celular, 30);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dia)) throw new Error('Escolha o dia no calendário.');
  if (dia < hoje_()) throw new Error('Esse dia já passou. Escolha outro dia.');
  if (periodos_(cfg).indexOf(periodo) < 0) throw new Error('Escolha o período.');
  if (!nome) throw new Error('Escreva o seu nome.');
  if (soDigitos_(celular).length < 10) throw new Error('Escreva o celular com DDD. Exemplo: (31) 99999-0000.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lans = lancamentosAtivos_();
    let associado = null, situacao = 'não associado';
    if (chacara) {
      associado = adimplencia_(cfg, lans).filter(function (a) { return a.chacara === chacara; })[0];
      if (!associado) throw new Error('Não encontramos a chácara ' + chacara + ' no cadastro. Se não é associado, escolha "Não sou associado".');
      if (String(cfg.conferirCelular).toLowerCase() !== 'não' && associado.celular) {
        const a = soDigitos_(associado.celular).slice(-8), b = soDigitos_(celular).slice(-8);
        if (a !== b) throw new Error('O celular não confere com o cadastro da chácara ' + chacara + '. Use o celular cadastrado ou procure a diretoria.');
      }
      situacao = associado.situacao;
    }
    if (situacao === 'inadimplente') {
      return { ok: false, bloqueado: true, titulo: 'Reserva não permitida', mensagem: cfg.avisoInadimplente };
    }

    const reservas = tabela_('Reservas').rows;
    if (chacara && cfg.limiteReservasMes) {
      const noMes = reservas.filter(function (r) {
        return r.Status === 'Ativa' && String(r.Chacara) === chacara && String(r.Dia).slice(0, 7) === dia.slice(0, 7);
      }).length;
      if (noMes >= Number(cfg.limiteReservasMes)) throw new Error('A chácara ' + chacara + ' já chegou ao limite de ' + cfg.limiteReservasMes + ' reserva(s) neste mês.');
    }

    const tipo = situacao === 'em dia' ? 'Associado' : 'Não associado';
    const taxa = tipo === 'Associado' ? 0 : numero_(cfg.taxaReserva);
    const conflitos = reservas.filter(function (r) { return r.Status === 'Ativa' && r.Dia === dia && sobrepoe_(r.Periodo, periodo); });
    const base = {
      Dia: dia, Periodo: periodo, Chacara: chacara, Nome: nome, Celular: celular, Tipo: tipo,
      Taxa: taxa, TaxaPaga: taxa ? 'não' : '', Pix: cfg.chavePix, Motivo: '', PedidoEm: agora_(), Lancamento: ''
    };

    if (tipo === 'Associado') {
      const primeiro = conflitos.filter(function (r) { return r.Tipo === 'Associado'; })
        .sort(function (a, b) { return String(a.PedidoEm).localeCompare(String(b.PedidoEm)); })[0];
      if (primeiro) {
        const id = proximoId_('F');
        inserir_('Reservas', Object.assign({ ID: id, Status: 'Fila', Motivo: 'Aguardando a chácara ' + primeiro.Chacara }, base));
        aviso_(primeiro.Nome, primeiro.Celular, 'Olá, ' + primeiro.Nome + '. A chácara ' + chacara + ' também pediu a área de lazer no dia ' + dataBR_(dia) +
          ' (' + periodo + '). Você pediu primeiro e tem prioridade. Se puderem, conversem para combinar. Contato: ' + nome + ', ' + celular + '.', primeiro.ID);
        aviso_(nome, celular, 'Olá, ' + nome + '. O dia ' + dataBR_(dia) + ' (' + periodo + ') já foi pedido antes pela chácara ' + primeiro.Chacara +
          ', que tem prioridade. Você está na fila de espera. Se quiserem combinar, o contato é ' + primeiro.Nome + ', ' + primeiro.Celular + '.', id);
        registrar_('público', 'reserva na fila', id + ' ' + dia + ' ' + periodo + ' ch ' + chacara);
        return {
          ok: false, fila: true, titulo: 'Este dia já foi pedido',
          mensagem: 'A chácara ' + primeiro.Chacara + ' pediu este dia antes e tem prioridade. Vocês dois serão avisados para conversar. Você ficou na fila de espera: se a outra reserva for cancelada, a vaga passa para você.'
        };
      }
      conflitos.filter(function (r) { return r.Tipo !== 'Associado'; }).forEach(function (r) {
        atualizar_('Reservas', r._linha, { Status: 'Substituída', Motivo: 'Um associado reservou o mesmo dia' });
        aviso_(r.Nome, r.Celular, 'Olá, ' + r.Nome + '. A sua reserva da área de lazer do dia ' + dataBR_(r.Dia) + ' (' + r.Periodo +
          ') foi substituída, porque um associado reservou a mesma data. ' +
          (String(r.TaxaPaga) === 'sim' ? 'A diretoria vai combinar com você a devolução da taxa.' : 'Fale com a diretoria se quiser outra data.'), r.ID);
      });
    } else if (conflitos.length) {
      throw new Error('Este dia e período já estão reservados. Escolha outro dia ou outro período.');
    }

    const id = proximoId_('V');
    inserir_('Reservas', Object.assign({ ID: id, Status: 'Ativa' }, base));
    registrar_('público', 'reserva', id + ' ' + dia + ' ' + periodo + ' ' + tipo + (chacara ? ' ch ' + chacara : ''));
    return {
      ok: true, id: id, titulo: taxa ? 'Reserva registrada' : 'Reserva confirmada',
      taxa: taxa, pix: taxa ? cfg.chavePix : '', dia: dia, periodo: periodo, situacao: situacao,
      mensagem: taxa
        ? 'Para garantir a data, pague a taxa de manutenção pelo Pix. A diretoria confirma quando o pagamento entrar.'
        : 'Tudo certo. A área de lazer está reservada para você.'
    };
  } finally { lock.releaseLock(); }
}

function aviso_(para, celular, mensagem, reserva) {
  inserir_('Avisos', { ID: proximoId_('A'), CriadoEm: agora_(), Para: para, Celular: celular, Mensagem: mensagem, Reserva: reserva || '', Status: 'Pendente' });
}

function dataBR_(iso) { const p = String(iso).split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }

/* ===================================================================
   DIRETORIA
   =================================================================== */

/** Tudo que a tela da diretoria precisa para abrir. */
function painel(token) {
  const u = sessao_(token);
  const cfg = config_();
  const lans = lancamentosAtivos_();
  const adim = adimplencia_(cfg, lans);
  const avisos = tabela_('Avisos').rows.filter(function (a) { return a.Status === 'Pendente'; }).length;
  const reservasFuturas = tabela_('Reservas').rows.filter(function (r) { return r.Status === 'Ativa' && r.Dia >= hoje_(); });
  return {
    usuario: perfilPublico_(u),
    saldoAtual: saldoAtual_(lans, cfg),
    associados: adim,
    campanhas: campanhasComTotais_(lans),
    config: u.Perfil === 'admin' ? configLista_() : null,
    periodos: periodos_(cfg),
    mensalidadePadrao: numero_(cfg.mensalidadePadrao),
    pagadores: pagadores_(lans, adim),
    avisosPendentes: avisos,
    taxasPendentes: reservasFuturas.filter(function (r) { return numero_(r.Taxa) > 0 && String(r.TaxaPaga) !== 'sim'; }).length,
    proximasReservas: reservasFuturas.sort(function (a, b) { return String(a.Dia).localeCompare(String(b.Dia)); }).slice(0, 5)
      .map(reservaTela_),
    hoje: hoje_(), mesAtual: mesAtual_()
  };
}

function pagadores_(lans, adim) {
  const vistos = {};
  lans.forEach(function (l) { if (l.PagoPor) vistos[String(l.PagoPor)] = true; });
  adim.forEach(function (a) { vistos[a.nome] = true; });
  return Object.keys(vistos).sort();
}

function configLista_() {
  return tabela_('Config').rows.map(function (r) { return { chave: r.Chave, valor: String(r.Valor), descricao: String(r.Descricao || '') }; });
}

/* ---------- Associados ---------- */

/* ---------- Terrenos ---------- */

const SITUACOES_TERRENO_ = ['Com casa', 'Sem construção', 'Em obra', 'Não informado'];

function terrenosPorAssociado_() {
  const mapa = {};
  tabela_('Terrenos').rows.forEach(function (t) {
    (mapa[t.Associado] = mapa[t.Associado] || []).push({ id: t.ID, numero: String(t.Numero), situacao: t.Situacao || 'Não informado',
      cobrado: String(t.Cobrado) !== 'não', motivo: String(t.Motivo || '') });
  });
  Object.keys(mapa).forEach(function (k) { mapa[k].sort(function (a, b) { return (parseInt(a.numero, 10) || 0) - (parseInt(b.numero, 10) || 0); }); });
  return mapa;
}

/** Cria um terreno para cada número do código da chácara ("65/66/67") de quem ainda não tem terrenos. */
function criarTerrenosDoCadastro_() {
  const ja = terrenosPorAssociado_();
  const novos = [];
  tabela_('Associados').rows.forEach(function (a) {
    if (ja[a.ID]) return;
    String(a.Chacara || '').split('/').map(function (n) { return n.trim(); }).filter(String).forEach(function (n) {
      novos.push({ ID: proximoId_('T'), Numero: n, Associado: a.ID, Situacao: 'Não informado', Cobrado: 'sim', Motivo: '', AtualizadoEm: agora_() });
    });
  });
  inserirLote_('Terrenos', novos);
  return novos.length;
}

function removerLinhas_(nomeAba, teste) {
  const t = tabela_(nomeAba);
  const linhas = t.rows.filter(teste).map(function (r) { return r._linha; }).sort(function (a, b) { return b - a; });
  apagarLinhas_(t.sh, linhas);
  if (linhas.length) mudouDados_();
  return linhas.length;
}

/** Apaga linhas agrupando as vizinhas num só comando (muito mais rápido que uma a uma). */
function apagarLinhas_(sh, linhas) {
  const ord = linhas.slice().sort(function (a, b) { return b - a; });
  if (!ord.length) return;
  // O Sheets não deixa apagar todas as linhas abaixo do cabeçalho: garante uma linha vazia sobrando.
  const congeladas = Math.max(1, sh.getFrozenRows ? sh.getFrozenRows() : 1);
  if (sh.getMaxRows() - ord.length <= congeladas) sh.insertRowsAfter(sh.getMaxRows(), 1);
  let i = 0;
  while (i < ord.length) {
    let fim = ord[i], ini = fim;
    while (i + 1 < ord.length && ord[i + 1] === ini - 1) { i++; ini = ord[i]; }
    if (fim === ini) sh.deleteRow(ini); else sh.deleteRows(ini, fim - ini + 1);
    i++;
  }
}

/** Quando o código da chácara muda, o histórico (lançamentos e reservas) acompanha. */
function renomearChacara_(antigo, novo) {
  if (!antigo || antigo === novo) return;
  ['Lancamentos', 'Reservas'].forEach(function (aba) {
    const sh = aba_(aba);
    const cab = cabecalho_(sh), j = cab.indexOf('Chacara');
    const n = sh.getLastRow() - 1;
    if (j < 0 || n < 1) return;
    const rg = sh.getRange(2, j + 1, n, 1);
    const vals = rg.getValues();
    let mudou = false;
    const novos = vals.map(function (v) { if (String(v[0]) === antigo) { mudou = true; return ["'" + novo]; } return [v[0] === '' ? '' : "'" + String(v[0])]; });
    if (mudou) rg.setValues(novos);
  });
  mudouDados_();
}

/**
 * dados: { id, nome, celular, mensalidade, cobrarDesde, obs,
 *          terrenos: [{ numero, situacao, cobrado, motivo }] }
 * O código da chácara é formado pelos números dos terrenos ("65/66").
 */
function salvarAssociado(token, dados) {
  const u = sessao_(token);
  dados = dados || {};
  const nome = texto_(dados.nome, 100);
  if (!nome) throw new Error('Informe o nome do associado.');
  let terrenos = Array.isArray(dados.terrenos) ? dados.terrenos : String(dados.chacara || '').split('/').map(function (n) { return { numero: n }; });
  terrenos = terrenos.map(function (t) {
    return { numero: texto_(t.numero, 12).replace(/\s+/g, ''), situacao: SITUACOES_TERRENO_.indexOf(t.situacao) >= 0 ? t.situacao : 'Não informado',
      cobrado: t.cobrado !== false, motivo: texto_(t.motivo, 200) };
  }).filter(function (t) { return t.numero; });
  if (!terrenos.length) throw new Error('Informe o número de pelo menos um terreno.');
  const nums = terrenos.map(function (t) { return t.numero; });
  const rep = nums.filter(function (n, i) { return nums.indexOf(n) !== i; });
  if (rep.length) throw new Error('O terreno ' + rep[0] + ' foi informado duas vezes.');
  if (nums.some(function (n) { return n.indexOf('/') >= 0; })) throw new Error('Informe um número por linha de terreno.');
  const semMotivo = terrenos.filter(function (t) { return !t.cobrado && !t.motivo; })[0];
  if (semMotivo) throw new Error('Escreva o motivo de não cobrar o terreno ' + semMotivo.numero + '.');
  nums.sort(function (a, b) { return (parseInt(a, 10) || 0) - (parseInt(b, 10) || 0) || a.localeCompare(b); });
  const chacara = nums.join('/');

  const t = tabela_('Associados');
  const ativos = {};
  t.rows.forEach(function (r) { if (String(r.Ativo) !== 'não') ativos[r.ID] = r; });
  const donos = terrenosPorAssociado_();
  Object.keys(donos).forEach(function (aid) {
    if (aid === dados.id || !ativos[aid]) return;
    donos[aid].forEach(function (x) { if (nums.indexOf(x.numero) >= 0) throw new Error('O terreno ' + x.numero + ' já está na conta de ' + ativos[aid].Nome + ' (chácara ' + ativos[aid].Chacara + ').'); });
  });
  const repetida = t.rows.filter(function (r) { return String(r.Chacara) === chacara && r.ID !== dados.id && String(r.Ativo) !== 'não'; })[0];
  if (repetida) throw new Error('A chácara ' + chacara + ' já está cadastrada para ' + repetida.Nome + '.');

  const cobrarDesde = /^\d{4}-\d{2}$/.test(String(dados.cobrarDesde || '')) ? dados.cobrarDesde : mesAtual_();
  const obj = {
    Chacara: chacara, Nome: nome, Celular: texto_(dados.celular, 30),
    Mensalidade: numero_(dados.mensalidade) || numero_(config_().mensalidadePadrao),
    CobrarDesde: cobrarDesde, Obs: texto_(dados.obs, 300), Ativo: dados.ativo === false ? 'não' : 'sim'
  };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let id = dados.id;
  try {
    if (id) {
      const r = t.rows.filter(function (x) { return x.ID === id; })[0];
      if (!r) throw new Error('Associado não encontrado.');
      atualizar_('Associados', r._linha, obj);
      renomearChacara_(String(r.Chacara), chacara);
    } else {
      id = proximoId_('C');
      inserir_('Associados', Object.assign({ ID: id, CriadoEm: agora_() }, obj));
    }
    removerLinhas_('Terrenos', function (x) { return x.Associado === id; });
    inserirLote_('Terrenos', terrenos.map(function (x) {
      return { ID: proximoId_('T'), Numero: x.numero, Associado: id, Situacao: x.situacao, Cobrado: x.cobrado ? 'sim' : 'não', Motivo: x.cobrado ? '' : x.motivo, AtualizadoEm: agora_() };
    }));
  } finally { lock.releaseLock(); }
  registrar_(u.Usuario, dados.id ? 'editou associado' : 'cadastrou associado', id + ' ch ' + chacara + ' · terrenos ' +
    terrenos.map(function (x) { return x.numero + (x.cobrado ? '' : ' (não cobrado: ' + x.motivo + ')'); }).join(', '));
  return id;
}

function historicoChacara(token, chacara) {
  sessao_(token);
  const cfg = config_();
  chacara = String(chacara);
  const todos = tabela_('Lancamentos').rows.map(function (l) { l.Valor = numero_(l.Valor); l.Chacara = String(l.Chacara || ''); return l; });
  const ativos = todos.filter(function (l) { return l.Status !== 'Cancelado'; });
  const a = adimplencia_(cfg, ativos).filter(function (x) { return x.chacara === chacara; })[0] || null;
  const nomesCamp = {};
  tabela_('Campanhas').rows.forEach(function (c) { nomesCamp[c.ID] = c.Nome; });
  const pagos = {};
  ativos.forEach(function (l) { if (l.Chacara === chacara && l.Categoria === 'Mensalidade') pagos[String(l.MesReferencia).slice(0, 7)] = l; });
  const devidos = a ? mesesDevidos_(a.cobrarDesde, cfg) : [];
  const meses = devidos.slice().reverse().map(function (m) {
    const l = pagos[m];
    return { mes: m, pago: !!l, data: l ? l.Data : '', valor: l ? l.Valor : (a ? a.mensalidade : 0), pagoPor: l ? String(l.PagoPor || '') : '' };
  });
  const movimentos = todos.filter(function (l) { return l.Chacara === chacara; })
    .sort(function (x, y) { return String(y.Data).localeCompare(String(x.Data)); })
    .map(function (l) { return lancTela_(l, nomesCamp); });
  const reservas = tabela_('Reservas').rows.filter(function (r) { return String(r.Chacara) === chacara; })
    .sort(function (x, y) { return String(y.Dia).localeCompare(String(x.Dia)); }).map(reservaTela_);
  return { associado: a, meses: meses, movimentos: movimentos, reservas: reservas };
}

/* ---------- Lançamentos ---------- */

function lancTela_(l, nomesCamp) {
  return {
    id: l.ID, data: l.Data, tipo: l.Tipo, categoria: l.Categoria, campanha: l.Campanha || '', campanhaNome: (nomesCamp || {})[l.Campanha] || '',
    chacara: String(l.Chacara || ''), pagoPor: String(l.PagoPor || ''), mes: String(l.MesReferencia || ''),
    descricao: String(l.Descricao || ''), valor: numero_(l.Valor), status: l.Status || 'Ativo', motivo: String(l.Motivo || ''),
    criadoPor: String(l.CriadoPor || ''), comprovante: !!l.Comprovante
  };
}

/**
 * Mensalidade: cria um lançamento por mês pago.
 * dados: { chacara, meses: ['2026-09','2026-10'], valorPorMes, data, pagoPor }
 */
function lancarMensalidade(token, dados) {
  const u = sessao_(token);
  dados = dados || {};
  const chacara = texto_(dados.chacara, 20);
  const meses = (dados.meses || []).filter(function (m) { return /^\d{4}-\d{2}$/.test(m); });
  const valor = numero_(dados.valorPorMes);
  const data = /^\d{4}-\d{2}-\d{2}$/.test(String(dados.data || '')) ? dados.data : hoje_();
  if (!chacara) throw new Error('Escolha a chácara.');
  if (!meses.length) throw new Error('Marque pelo menos um mês.');
  if (!(valor > 0)) throw new Error('Informe o valor da mensalidade.');
  const arq = dados.arquivo ? salvarArquivo_(dados.arquivo, { data: data, tipo: 'Entrada', categoria: 'Mensalidade ch ' + chacara, valor: valor * meses.length }) : '';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ja = {};
    lancamentosAtivos_().forEach(function (l) { if (l.Chacara === chacara && l.Categoria === 'Mensalidade') ja[String(l.MesReferencia).slice(0, 7)] = true; });
    const repetidos = meses.filter(function (m) { return ja[m]; });
    if (repetidos.length) throw new Error('Estes meses já estão pagos: ' + repetidos.map(mesBR_).join(', ') + '.');
    const ids = [];
    meses.sort().forEach(function (m) {
      const id = proximoId_('L');
      inserir_('Lancamentos', {
        ID: id, Data: data, Tipo: 'Entrada', Categoria: 'Mensalidade', Campanha: '', Chacara: chacara,
        PagoPor: texto_(dados.pagoPor, 100), MesReferencia: m, Descricao: 'Mensalidade ' + mesBR_(m), Valor: valor,
        Status: 'Ativo', Motivo: '', CriadoPor: u.Usuario, CriadoEm: agora_(), CanceladoPor: '', CanceladoEm: '', Comprovante: arq
      });
      ids.push(id);
    });
    registrar_(u.Usuario, 'mensalidade', 'ch ' + chacara + ' ' + meses.join(',') + ' ' + valor);
    return { ids: ids, total: Math.round(valor * meses.length * 100) / 100 };
  } finally { lock.releaseLock(); }
}

/**
 * Outra entrada ou despesa.
 * dados: { tipo: 'Entrada'|'Saída', categoria, campanha, chacara, pagoPor, data, descricao, valor }
 */
function lancar(token, dados) {
  const u = sessao_(token);
  dados = dados || {};
  const tipo = dados.tipo === 'Saída' ? 'Saída' : 'Entrada';
  const valor = numero_(dados.valor);
  const descricao = texto_(dados.descricao, 200);
  const data = /^\d{4}-\d{2}-\d{2}$/.test(String(dados.data || '')) ? dados.data : hoje_();
  const categoria = texto_(dados.categoria, 40) || (tipo === 'Saída' ? 'Despesa' : 'Outras entradas');
  if (!(valor > 0)) throw new Error('Informe um valor maior que zero.');
  if (!descricao) throw new Error('Escreva uma descrição. Exemplo: Conta de luz do poço.');
  const campanha = texto_(dados.campanha, 20);
  if (campanha && !tabela_('Campanhas').rows.some(function (c) { return c.ID === campanha; })) throw new Error('Campanha não encontrada.');
  const id = proximoId_('L');
  const arq = dados.arquivo ? salvarArquivo_(dados.arquivo, { data: data, tipo: tipo, categoria: campanha ? 'Campanha' : categoria, valor: valor, id: id }) : '';
  inserir_('Lancamentos', {
    ID: id, Data: data, Tipo: tipo, Categoria: campanha ? 'Campanha' : categoria, Campanha: campanha, Chacara: texto_(dados.chacara, 20),
    PagoPor: texto_(dados.pagoPor, 100), MesReferencia: '', Descricao: descricao, Valor: valor,
    Status: 'Ativo', Motivo: '', CriadoPor: u.Usuario, CriadoEm: agora_(), CanceladoPor: '', CanceladoEm: '', Comprovante: arq
  });
  registrar_(u.Usuario, 'lançamento', id + ' ' + tipo + ' ' + valor + ' ' + descricao);
  return id;
}

function listarLancamentos(token, ym) {
  sessao_(token);
  ym = /^\d{4}-\d{2}$/.test(String(ym || '')) ? ym : mesAtual_();
  const nomesCamp = {};
  tabela_('Campanhas').rows.forEach(function (c) { nomesCamp[c.ID] = c.Nome; });
  return tabela_('Lancamentos').rows.filter(function (l) { return String(l.Data).slice(0, 7) === ym; })
    .sort(function (a, b) { return String(b.Data).localeCompare(String(a.Data)) || String(b.ID).localeCompare(String(a.ID)); })
    .map(function (l) { return lancTela_(l, nomesCamp); });
}

function cancelarLancamento(token, id, motivo) {
  const u = sessao_(token);
  motivo = texto_(motivo, 200);
  if (!motivo) throw new Error('Escreva o motivo do cancelamento.');
  const l = tabela_('Lancamentos').rows.filter(function (x) { return x.ID === id; })[0];
  if (!l) throw new Error('Lançamento não encontrado.');
  if (l.Status === 'Cancelado') throw new Error('Este lançamento já estava cancelado.');
  atualizar_('Lancamentos', l._linha, { Status: 'Cancelado', Motivo: motivo, CanceladoPor: u.Usuario, CanceladoEm: agora_() });
  registrar_(u.Usuario, 'cancelou lançamento', id + ' ' + motivo);
  return true;
}

/* ---------- Regularização da contagem ---------- */

/**
 * dados: { id, modo: 'recomecar' | 'acordo', mes (AAAA-MM), motivo, valor, data, pagoPor, arquivo }
 * recomecar: a mensalidade passa a contar a partir de "mes" (meses anteriores não são cobrados).
 * acordo:    registra o valor pago no acordo e dá os meses até "mes" como quitados.
 */
function regularizarAssociado(token, dados) {
  const u = sessao_(token);
  dados = dados || {};
  const a = tabela_('Associados').rows.filter(function (x) { return x.ID === dados.id; })[0];
  if (!a) throw new Error('Associado não encontrado.');
  const mes = String(dados.mes || '');
  if (!/^\d{4}-\d{2}$/.test(mes)) throw new Error('Escolha o mês.');
  const motivo = texto_(dados.motivo, 200);
  const hoje = dataBR_(hoje_());
  let nota;
  if (dados.modo === 'acordo') {
    const valor = numero_(dados.valor);
    if (!(valor > 0)) throw new Error('Informe o valor pago no acordo.');
    const data = /^\d{4}-\d{2}-\d{2}$/.test(String(dados.data || '')) ? dados.data : hoje_();
    const id = proximoId_('L');
    const arq = dados.arquivo ? salvarArquivo_(dados.arquivo, { data: data, tipo: 'Entrada', categoria: 'Acordo ch ' + a.Chacara, valor: valor, id: id }) : '';
    inserir_('Lancamentos', {
      ID: id, Data: data, Tipo: 'Entrada', Categoria: 'Acordo de regularização', Campanha: '', Chacara: String(a.Chacara),
      PagoPor: texto_(dados.pagoPor, 100), MesReferencia: '', Descricao: 'Acordo: quitou as mensalidades até ' + mesBR_(mes) + (motivo ? ' (' + motivo + ')' : ''),
      Valor: valor, Status: 'Ativo', Motivo: '', CriadoPor: u.Usuario, CriadoEm: agora_(), CanceladoPor: '', CanceladoEm: '', Comprovante: arq
    });
    nota = hoje + ': acordo de R$ ' + valor.toFixed(2).replace('.', ',') + ' quitou até ' + mesBR_(mes) + (motivo ? '. ' + motivo : '');
    atualizar_('Associados', a._linha, { CobrarDesde: somaMes_(mes, 1), Regularizacao: (String(a.Regularizacao || '') + '\n' + nota).trim() });
  } else {
    nota = hoje + ': recomeçou a contar em ' + mesBR_(mes) + (motivo ? '. Motivo: ' + motivo : '');
    atualizar_('Associados', a._linha, { CobrarDesde: mes, Regularizacao: (String(a.Regularizacao || '') + '\n' + nota).trim() });
  }
  registrar_(u.Usuario, 'regularização', a.Chacara + ' ' + nota);
  return true;
}

/* ---------- Sincronização (o aplicativo guarda tudo no aparelho) ---------- */

// Lançamento em lista compacta: [id, data, tipo, categoria, campanha, chacara, mes, descricao, valor, comprovante, status, pagoPor, motivo, criadoPor]
/** Entradas que não vêm de uma pessoa: só essas mostram a descrição no portal público. */
const ENTRADAS_SEM_PESSOA_ = ['Rendimento', 'Ajuste de conferência'];

function lancCompacto_(l, completo) {
  // No portal público, entradas pagas por moradores não levam o texto livre (pode ter nome): a tela mostra a campanha e a chácara.
  const descPublica = l.Tipo !== 'Entrada' || ENTRADAS_SEM_PESSOA_.indexOf(String(l.Categoria)) >= 0;
  const base = [l.ID, String(l.Data), l.Tipo, l.Categoria, l.Campanha || '', String(l.Chacara || ''), String(l.MesReferencia || '').slice(0, 7),
    completo || descPublica ? String(l.Descricao || '') : '', numero_(l.Valor),
    l.Comprovante && (completo || l.Tipo === 'Saída') ? 1 : 0];
  if (completo) base.push(l.Status || 'Ativo', String(l.PagoPor || ''), String(l.Motivo || ''), String(l.CriadoPor || ''));
  return base;
}

function configPublica_(cfg) {
  return {
    nomeAssociacao: cfg.nomeAssociacao, saldoInicial: numero_(cfg.saldoInicial), dataAbertura: cfg.dataAbertura || '0000-00-00',
    diaVencimento: Number(cfg.diaVencimento || 10), mesesAtraso: Number(cfg.mesesAtraso || 2), mesesInadimplente: Number(cfg.mesesInadimplente || 4),
    taxaReserva: numero_(cfg.taxaReserva), limiteReservasMes: cfg.limiteReservasMes ? Number(cfg.limiteReservasMes) : null, periodos: periodos_(cfg)
  };
}

function sincronizarPublico() {
  return emCacheGrande_('sinc_pub_' + hoje_(), function () {
    const cfg = config_();
    const camps = tabela_('Campanhas').rows;
    const ocultas = {};
    camps.forEach(function (c) { if (String(c.Oculta) === 'sim') ocultas[c.ID] = true; });
    const lans = tabela_('Lancamentos').rows.filter(function (l) { return l.Status !== 'Cancelado'; }).map(function (l) {
      const c = lancCompacto_(l, false); if (ocultas[c[4]]) c[4] = '';
      c[5] = ''; c[6] = ''; // público: nenhum pagamento é ligado a chácara nem a mês de referência
      return c;
    });
    // Só quantidades (nunca quais chácaras): total, pagaram o mês atual, atrasadas e inadimplentes
    const ativos = lancamentosAtivos_();
    const adim = adimplencia_(cfg, ativos);
    const pagouMes = {};
    ativos.forEach(function (l) { if (l.Categoria === 'Mensalidade' && String(l.MesReferencia).slice(0, 7) === mesAtual_()) pagouMes[String(l.Chacara)] = true; });
    const contagem = {
      total: adim.length,
      pagaramMes: adim.filter(function (x) { return pagouMes[x.chacara]; }).length,
      atrasados: adim.filter(function (x) { return x.situacao === 'atrasado'; }).length,
      inadimplentes: adim.filter(function (x) { return x.situacao === 'inadimplente'; }).length
    };
    return {
      versao: 2, geradoEm: agora_(), hoje: hoje_(), config: configPublica_(cfg), contagem: contagem, lancamentos: lans,
      campanhas: camps.filter(function (c) { return !ocultas[c.ID]; }).map(function (c) {
        return { id: c.ID, nome: String(c.Nome), descricao: String(c.Descricao || ''), meta: numero_(c.Meta), inicio: String(c.Inicio || ''), status: c.Status || 'Ativa' };
      }),
      associados: tabela_('Associados').rows.filter(function (a) { return String(a.Ativo) !== 'não'; })
        .map(function (a) { return { chacara: String(a.Chacara) }; }), // só a lista para o formulário de reserva
      reservas: tabela_('Reservas').rows.filter(function (r) { return r.Status === 'Ativa' && r.Dia >= somaMes_(mesAtual_(), -1) + '-01'; })
        .map(function (r) { return { dia: r.Dia, periodo: r.Periodo, tipo: r.Tipo }; })
    };
  });
}

function sincronizarDiretoria(token) {
  const u = sessao_(token);
  const cfg = config_();
  const admin = u.Perfil === 'admin';
  const c = configPublica_(cfg);
  c.chavePix = cfg.chavePix; c.mensalidadePadrao = numero_(cfg.mensalidadePadrao);
  if (admin) { c.avisoInadimplente = cfg.avisoInadimplente; c.conferirCelular = cfg.conferirCelular; }
  return {
    versao: 2, geradoEm: agora_(), hoje: hoje_(), usuario: perfilPublico_(u), config: c,
    configLista: admin ? configLista_() : null,
    usuarios: admin ? listarUsuariosSemSessao_() : null,
    lancamentos: tabela_('Lancamentos').rows.map(function (l) { return lancCompacto_(l, true); }),
    campanhas: tabela_('Campanhas').rows.map(function (k) {
      return { id: k.ID, nome: String(k.Nome), descricao: String(k.Descricao || ''), meta: numero_(k.Meta), inicio: String(k.Inicio || ''), status: k.Status || 'Ativa', oculta: String(k.Oculta) === 'sim' };
    }),
    associados: (function () { const ter = terrenosPorAssociado_(); return tabela_('Associados').rows.map(function (a) {
      return { id: a.ID, terrenos: ter[a.ID] || [], chacara: String(a.Chacara), nome: String(a.Nome), celular: String(a.Celular || ''), mensalidade: numero_(a.Mensalidade) || numero_(cfg.mensalidadePadrao),
        desde: String(a.CobrarDesde || '').slice(0, 7), ativo: String(a.Ativo) !== 'não', obs: String(a.Obs || ''), regularizacao: String(a.Regularizacao || '') };
    }); })(),
    reservas: tabela_('Reservas').rows.filter(function (r) { return r.Dia >= somaMes_(mesAtual_(), -3) + '-01'; }).map(reservaTela_),
    avisos: tabela_('Avisos').rows.filter(function (a) { return a.Status === 'Pendente'; })
      .map(function (a) { return { id: a.ID, criadoEm: a.CriadoEm, para: String(a.Para), celular: String(a.Celular || ''), mensagem: String(a.Mensagem), reserva: a.Reserva }; })
  };
}

function listarUsuariosSemSessao_() {
  return tabela_('Usuarios').rows.map(function (u) {
    return { id: u.ID, usuario: u.Usuario, nome: u.Nome, perfil: u.Perfil, ativo: String(u.Ativo) === 'sim', ultimoAcesso: String(u.UltimoAcesso || '') };
  });
}

/** Cache para resultados grandes (divide em partes de 90 mil caracteres). */
function emCacheGrande_(chave, fn) {
  const cache = CacheService.getScriptCache();
  const base = chave + '_' + versaoDados_();
  try {
    const n = cache.get(base + '_n');
    if (n) {
      const ks = []; for (let i = 0; i < Number(n); i++) ks.push(base + '_' + i);
      const partes = cache.getAll(ks);
      if (ks.every(function (k) { return partes[k] != null; })) return JSON.parse(ks.map(function (k) { return partes[k]; }).join(''));
    }
  } catch (e) { }
  const r = fn();
  try {
    const t = JSON.stringify(r), mapa = {}, tam = 90000;
    let i = 0;
    for (; i * tam < t.length; i++) mapa[base + '_' + i] = t.slice(i * tam, (i + 1) * tam);
    mapa[base + '_n'] = String(i);
    if (i <= 20) cache.putAll(mapa, 600);
  } catch (e) { }
  return r;
}

/* ---------- Comprovantes (Google Drive) ---------- */

const TIPOS_ARQUIVO_ = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'application/pdf': 'pdf' };

function pastaComprovantes_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    let raiz = null;
    const id = props.getProperty('PASTA_ID');
    if (id) { try { raiz = DriveApp.getFolderById(id); } catch (e) { raiz = null; } }
    if (!raiz) { raiz = DriveApp.createFolder('Portal Siriema - Comprovantes'); props.setProperty('PASTA_ID', raiz.getId()); }
    const filha = function (pai, nome) { const it = pai.getFoldersByName(nome); return it.hasNext() ? it.next() : pai.createFolder(nome); };
    const meses = ['01 - Janeiro', '02 - Fevereiro', '03 - Março', '04 - Abril', '05 - Maio', '06 - Junho', '07 - Julho', '08 - Agosto', '09 - Setembro', '10 - Outubro', '11 - Novembro', '12 - Dezembro'];
    return filha(filha(raiz, String(data).slice(0, 4)), meses[Number(String(data).slice(5, 7)) - 1]);
  } finally { lock.releaseLock(); }
}

/** arquivo: { nome, tipo, base64 }  ->  ID do arquivo no Drive */
function salvarArquivo_(arquivo, info) {
  const tipo = String(arquivo.tipo || '');
  const ext = TIPOS_ARQUIVO_[tipo];
  if (!ext) throw new Error('O comprovante precisa ser uma foto (JPG ou PNG) ou um PDF.');
  const b64 = String(arquivo.base64 || '').replace(/^data:[^,]+,/, '');
  if (b64.length > 11000000) throw new Error('O arquivo é grande demais. O limite é 8 MB.');
  const bytes = Utilities.base64Decode(b64);
  const limpa = function (t) { return String(t).replace(/[\\\/:*?"<>|]/g, '-').slice(0, 40); };
  const nome = [info.data, info.tipo === 'Saída' ? 'Saida' : 'Entrada', limpa(info.categoria), 'R$' + Number(info.valor || 0).toFixed(2).replace('.', ','), info.id || ''].filter(String).join('_') + '.' + ext;
  const f = pastaComprovantes_(info.data).createFile(Utilities.newBlob(bytes, tipo, nome));
  return f.getId();
}

function arquivoParaTela_(fileId) {
  const f = DriveApp.getFileById(fileId);
  const blob = f.getBlob();
  return { nome: f.getName(), tipo: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) };
}

function anexarComprovante(token, id, arquivo) {
  const u = sessao_(token);
  const l = tabela_('Lancamentos').rows.filter(function (x) { return x.ID === id; })[0];
  if (!l) throw new Error('Lançamento não encontrado.');
  const fid = salvarArquivo_(arquivo, { data: l.Data, tipo: l.Tipo, categoria: l.Categoria, valor: numero_(l.Valor), id: l.ID });
  atualizar_('Lancamentos', l._linha, { Comprovante: fid });
  registrar_(u.Usuario, 'anexou comprovante', id);
  return true;
}

function verComprovante(token, id) {
  sessao_(token);
  const l = tabela_('Lancamentos').rows.filter(function (x) { return x.ID === id; })[0];
  if (!l || !l.Comprovante) throw new Error('Este lançamento não tem comprovante.');
  return arquivoParaTela_(l.Comprovante);
}

/** Qualquer pessoa pode ver o comprovante de uma DESPESA. */
function comprovantePublico(id) {
  const l = tabela_('Lancamentos').rows.filter(function (x) { return x.ID === id; })[0];
  if (!l || l.Tipo !== 'Saída' || l.Status === 'Cancelado' || !l.Comprovante) throw new Error('Comprovante não disponível.');
  return arquivoParaTela_(l.Comprovante);
}

/* ---------- Campanhas ---------- */

function campanhasComTotais_(lans) {
  return tabela_('Campanhas').rows.map(function (c) {
    const da = lans.filter(function (l) { return l.Campanha === c.ID; });
    const arr = da.filter(function (l) { return l.Tipo === 'Entrada'; });
    const contrib = {};
    arr.forEach(function (l) { if (l.Chacara) contrib[l.Chacara] = true; });
    return {
      id: c.ID, nome: String(c.Nome), descricao: String(c.Descricao || ''), meta: numero_(c.Meta), inicio: String(c.Inicio || ''),
      status: c.Status || 'Ativa', oculta: String(c.Oculta) === 'sim',
      arrecadado: Math.round(arr.reduce(function (s, l) { return s + l.Valor; }, 0) * 100) / 100,
      gasto: Math.round(da.filter(function (l) { return l.Tipo === 'Saída'; }).reduce(function (s, l) { return s + l.Valor; }, 0) * 100) / 100,
      contribuintes: Object.keys(contrib).length
    };
  });
}

function salvarCampanha(token, dados) {
  const u = sessao_(token);
  dados = dados || {};
  const nome = texto_(dados.nome, 80);
  if (!nome) throw new Error('Dê um nome para a campanha. Exemplo: Luz Forte.');
  const obj = {
    Nome: nome, Descricao: texto_(dados.descricao, 300), Meta: numero_(dados.meta),
    Inicio: /^\d{4}-\d{2}-\d{2}$/.test(String(dados.inicio || '')) ? dados.inicio : hoje_(),
    Status: dados.status === 'Concluída' ? 'Concluída' : 'Ativa',
    Oculta: dados.oculta ? 'sim' : 'não'
  };
  if (dados.id) {
    const c = tabela_('Campanhas').rows.filter(function (x) { return x.ID === dados.id; })[0];
    if (!c) throw new Error('Campanha não encontrada.');
    atualizar_('Campanhas', c._linha, obj);
    registrar_(u.Usuario, 'editou campanha', dados.id + ' ' + nome);
    return dados.id;
  }
  const id = proximoId_('K');
  inserir_('Campanhas', Object.assign({ ID: id, CriadoEm: agora_() }, obj));
  registrar_(u.Usuario, 'criou campanha', id + ' ' + nome);
  return id;
}

/* ---------- Reservas (diretoria) ---------- */

function reservaTela_(r) {
  return {
    id: r.ID, dia: r.Dia, periodo: r.Periodo, chacara: String(r.Chacara || ''), nome: String(r.Nome || ''), celular: String(r.Celular || ''),
    tipo: r.Tipo, taxa: numero_(r.Taxa), taxaPaga: String(r.TaxaPaga) === 'sim', pix: String(r.Pix || ''),
    status: r.Status, motivo: String(r.Motivo || ''), pedidoEm: String(r.PedidoEm || '')
  };
}

function reservasDiretoria(token) {
  sessao_(token);
  const rs = tabela_('Reservas').rows;
  const ord = function (a, b) { return String(a.Dia).localeCompare(String(b.Dia)) || String(a.PedidoEm).localeCompare(String(b.PedidoEm)); };
  const hoje = hoje_();
  return {
    ativas: rs.filter(function (r) { return r.Status === 'Ativa' && r.Dia >= hoje; }).sort(ord).map(reservaTela_),
    fila: rs.filter(function (r) { return r.Status === 'Fila' && r.Dia >= hoje; }).sort(ord).map(reservaTela_),
    encerradas: rs.filter(function (r) { return r.Status !== 'Ativa' && r.Status !== 'Fila' || r.Dia < hoje; })
      .sort(function (a, b) { return String(b.Dia).localeCompare(String(a.Dia)); }).slice(0, 40).map(reservaTela_),
    avisos: tabela_('Avisos').rows.filter(function (a) { return a.Status === 'Pendente'; })
      .sort(function (a, b) { return String(b.CriadoEm).localeCompare(String(a.CriadoEm)); })
      .map(function (a) { return { id: a.ID, criadoEm: a.CriadoEm, para: String(a.Para), celular: String(a.Celular || ''), mensagem: String(a.Mensagem), reserva: a.Reserva }; })
  };
}

function cancelarReserva(token, id, motivo) {
  const u = sessao_(token);
  motivo = texto_(motivo, 200) || 'Cancelada pela diretoria';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rs = tabela_('Reservas').rows;
    const r = rs.filter(function (x) { return x.ID === id; })[0];
    if (!r) throw new Error('Reserva não encontrada.');
    if (r.Status !== 'Ativa' && r.Status !== 'Fila') throw new Error('Esta reserva já não está ativa.');
    atualizar_('Reservas', r._linha, { Status: 'Cancelada', Motivo: motivo + ' (por ' + u.Usuario + ')' });
    aviso_(r.Nome, r.Celular, 'Olá, ' + r.Nome + '. A sua reserva da área de lazer do dia ' + dataBR_(r.Dia) + ' (' + r.Periodo + ') foi cancelada pela diretoria. Motivo: ' + motivo + '.', r.ID);
    let promovida = null;
    if (r.Status === 'Ativa') {
      const ativasRestantes = rs.filter(function (x) { return x.Status === 'Ativa' && x.ID !== r.ID && x.Dia === r.Dia; });
      const candidato = rs.filter(function (x) {
        return x.Status === 'Fila' && x.Dia === r.Dia && sobrepoe_(x.Periodo, r.Periodo) &&
          !ativasRestantes.some(function (y) { return sobrepoe_(y.Periodo, x.Periodo); });
      }).sort(function (a, b) { return String(a.PedidoEm).localeCompare(String(b.PedidoEm)); })[0];
      if (candidato) {
        atualizar_('Reservas', candidato._linha, { Status: 'Ativa', Motivo: 'Saiu da fila: reserva anterior cancelada' });
        aviso_(candidato.Nome, candidato.Celular, 'Olá, ' + candidato.Nome + '. A área de lazer ficou livre no dia ' + dataBR_(candidato.Dia) + ' (' + candidato.Periodo + ') e a sua reserva foi confirmada.', candidato.ID);
        promovida = candidato.ID;
      }
    }
    registrar_(u.Usuario, 'cancelou reserva', id + ' ' + motivo + (promovida ? ' / promovida ' + promovida : ''));
    return { promovida: promovida };
  } finally { lock.releaseLock(); }
}

/** Marca a taxa como paga e lança a entrada no caixa. */
function marcarTaxaPaga(token, id, data) {
  const u = sessao_(token);
  const r = tabela_('Reservas').rows.filter(function (x) { return x.ID === id; })[0];
  if (!r) throw new Error('Reserva não encontrada.');
  if (String(r.TaxaPaga) === 'sim') throw new Error('Esta taxa já foi marcada como paga.');
  const lid = proximoId_('L');
  inserir_('Lancamentos', {
    ID: lid, Data: /^\d{4}-\d{2}-\d{2}$/.test(String(data || '')) ? data : hoje_(), Tipo: 'Entrada', Categoria: 'Taxa de reserva', Campanha: '',
    Chacara: String(r.Chacara || ''), PagoPor: r.Nome, MesReferencia: '', Descricao: 'Taxa de reserva ' + dataBR_(r.Dia) + ' (' + r.Periodo + ')',
    Valor: numero_(r.Taxa), Status: 'Ativo', Motivo: '', CriadoPor: u.Usuario, CriadoEm: agora_(), CanceladoPor: '', CanceladoEm: ''
  });
  atualizar_('Reservas', r._linha, { TaxaPaga: 'sim', Lancamento: lid });
  registrar_(u.Usuario, 'taxa paga', id + ' ' + lid);
  return lid;
}

function marcarAvisoEnviado(token, id) {
  const u = sessao_(token);
  const a = tabela_('Avisos').rows.filter(function (x) { return x.ID === id; })[0];
  if (!a) throw new Error('Aviso não encontrado.');
  atualizar_('Avisos', a._linha, { Status: 'Enviado' });
  registrar_(u.Usuario, 'aviso enviado', id);
  return true;
}

/* ---------- Relatórios ---------- */

function relatorioMes(token, ym) {
  sessao_(token);
  const cfg = config_();
  ym = /^\d{4}-\d{2}$/.test(String(ym || '')) ? ym : mesAtual_();
  const lans = lancamentosAtivos_();
  const inicio = ym + '-01';
  const abertura = cfg.dataAbertura || '0000-00-00';
  let saldoAnterior = numero_(cfg.saldoInicial);
  lans.forEach(function (l) { if (String(l.Data) >= abertura && String(l.Data) < inicio) saldoAnterior += l.Tipo === 'Entrada' ? l.Valor : -l.Valor; });
  const nomesCamp = {};
  tabela_('Campanhas').rows.forEach(function (c) { nomesCamp[c.ID] = c.Nome; });
  const doMes = lans.filter(function (l) { return String(l.Data).slice(0, 7) === ym; }).sort(porData_);
  const agrupa = function (tipo) {
    const g = {};
    doMes.filter(function (l) { return l.Tipo === tipo; }).forEach(function (l) {
      const k = l.Categoria === 'Campanha' ? 'Campanha: ' + (nomesCamp[l.Campanha] || '') : l.Categoria;
      g[k] = Math.round(((g[k] || 0) + l.Valor) * 100) / 100;
    });
    return Object.keys(g).map(function (k) { return { categoria: k, valor: g[k] }; }).sort(function (a, b) { return b.valor - a.valor; });
  };
  const te = doMes.filter(function (l) { return l.Tipo === 'Entrada'; }).reduce(function (s, l) { return s + l.Valor; }, 0);
  const ts = doMes.filter(function (l) { return l.Tipo === 'Saída'; }).reduce(function (s, l) { return s + l.Valor; }, 0);
  const r2 = function (n) { return Math.round(n * 100) / 100; };
  return {
    mes: ym, saldoAnterior: r2(saldoAnterior), entradas: r2(te), saidas: r2(ts), saldoFinal: r2(saldoAnterior + te - ts),
    entradasPorCategoria: agrupa('Entrada'), saidasPorCategoria: agrupa('Saída'),
    lancamentos: doMes.map(function (l) { return lancTela_(l, nomesCamp); })
  };
}

function relatorioCampanha(token, id) {
  sessao_(token);
  const lans = lancamentosAtivos_();
  const c = campanhasComTotais_(lans).filter(function (x) { return x.id === id; })[0];
  if (!c) throw new Error('Campanha não encontrada.');
  const nomesCamp = {}; nomesCamp[c.id] = c.nome;
  const da = lans.filter(function (l) { return l.Campanha === id; }).sort(porData_);
  const porChacara = {};
  da.filter(function (l) { return l.Tipo === 'Entrada'; }).forEach(function (l) {
    const k = l.Chacara || 'Sem chácara';
    porChacara[k] = Math.round(((porChacara[k] || 0) + l.Valor) * 100) / 100;
  });
  return {
    campanha: c,
    porChacara: Object.keys(porChacara).map(function (k) { return { chacara: k, valor: porChacara[k] }; }).sort(ordemChacara_),
    lancamentos: da.map(function (l) { return lancTela_(l, nomesCamp); })
  };
}

function relatorioAdimplencia(token) {
  sessao_(token);
  const cfg = config_();
  const adim = adimplencia_(cfg, lancamentosAtivos_());
  return {
    regras: { atraso: Number(cfg.mesesAtraso), inadimplente: Number(cfg.mesesInadimplente), vencimento: Number(cfg.diaVencimento) },
    associados: adim.map(function (a) {
      return { chacara: a.chacara, nome: a.nome, celular: a.celular, mensalidade: a.mensalidade, abertos: a.abertos, situacao: a.situacao,
        devendo: Math.round(a.abertos.length * a.mensalidade * 100) / 100 };
    })
  };
}

/* ---------- Configurações e usuários (administrador) ---------- */

function salvarConfig(token, valores) {
  const u = sessao_(token, { admin: true });
  valores = valores || {};
  const t = tabela_('Config');
  const permitidas = CONFIG_PADRAO.map(function (c) { return c[0]; });
  Object.keys(valores).forEach(function (k) {
    if (permitidas.indexOf(k) < 0) return;
    let v = texto_(valores[k], 1500);
    if (['saldoInicial', 'mensalidadePadrao', 'taxaReserva'].indexOf(k) >= 0) v = String(numero_(v));
    if (['diaVencimento', 'mesesAtraso', 'mesesInadimplente'].indexOf(k) >= 0 && !(Number(v) > 0)) throw new Error('O campo "' + k + '" precisa ser um número maior que zero.');
    if (k === 'limiteReservasMes' && v && !(Number(v) > 0)) throw new Error('O limite de reservas precisa ser um número, ou ficar em branco.');
    if (k === 'dataAbertura' && v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error('Data de abertura inválida.');
    const r = t.rows.filter(function (x) { return x.Chave === k; })[0];
    if (r) atualizar_('Config', r._linha, { Valor: v });
    else inserir_('Config', { Chave: k, Valor: v, Descricao: '' });
  });
  const cfg = config_();
  if (Number(cfg.mesesInadimplente) <= Number(cfg.mesesAtraso)) throw new Error('"Inadimplente" precisa começar depois de "Atrasado".');
  registrar_(u.Usuario, 'configurações', Object.keys(valores).join(','));
  return configLista_();
}

function listarUsuarios(token) {
  sessao_(token, { admin: true });
  return tabela_('Usuarios').rows.map(function (u) {
    return { id: u.ID, usuario: u.Usuario, nome: u.Nome, perfil: u.Perfil, ativo: String(u.Ativo) === 'sim', ultimoAcesso: String(u.UltimoAcesso || '') };
  });
}

function criarUsuario(token, dados) {
  const adm = sessao_(token, { admin: true });
  dados = dados || {};
  const usuario = texto_(dados.usuario, 40).toLowerCase().replace(/\s+/g, '.');
  const nome = texto_(dados.nome, 80);
  if (!/^[a-z0-9._-]{3,40}$/.test(usuario)) throw new Error('O usuário deve ter pelo menos 3 letras, sem acentos. Exemplo: maria.silva');
  if (!nome) throw new Error('Informe o nome da pessoa.');
  if (tabela_('Usuarios').rows.some(function (u) { return String(u.Usuario).toLowerCase() === usuario; })) throw new Error('Este usuário já existe.');
  const senha = senhaProvisoria_();
  const salt = Utilities.getUuid();
  const id = proximoId_('U');
  inserir_('Usuarios', {
    ID: id, Usuario: usuario, Nome: nome, Perfil: dados.perfil === 'admin' ? 'admin' : 'operador',
    Salt: salt, Hash: hash_(senha, salt), Ativo: 'sim', TrocarSenha: 'sim', CriadoEm: agora_(), UltimoAcesso: ''
  });
  registrar_(adm.Usuario, 'criou usuário', usuario);
  return { usuario: usuario, senha: senha };
}

function novaSenhaUsuario(token, id) {
  const adm = sessao_(token, { admin: true });
  const u = tabela_('Usuarios').rows.filter(function (x) { return x.ID === id; })[0];
  if (!u) throw new Error('Usuário não encontrado.');
  const senha = senhaProvisoria_();
  const salt = Utilities.getUuid();
  atualizar_('Usuarios', u._linha, { Salt: salt, Hash: hash_(senha, salt), TrocarSenha: 'sim', Ativo: 'sim' });
  CacheService.getScriptCache().put('bloq_usuario_' + u.ID, '1', 21600);
  registrar_(adm.Usuario, 'nova senha', u.Usuario);
  return { usuario: u.Usuario, senha: senha };
}

function ativarUsuario(token, id, ativo) {
  const adm = sessao_(token, { admin: true });
  if (id === adm.ID && !ativo) throw new Error('Você não pode desativar a sua própria conta.');
  const u = tabela_('Usuarios').rows.filter(function (x) { return x.ID === id; })[0];
  if (!u) throw new Error('Usuário não encontrado.');
  atualizar_('Usuarios', u._linha, { Ativo: ativo ? 'sim' : 'não' });
  if (!ativo) CacheService.getScriptCache().put('bloq_usuario_' + u.ID, '1', 21600);
  registrar_(adm.Usuario, ativo ? 'ativou usuário' : 'desativou usuário', u.Usuario);
  return true;
}

function mesBR_(ym) {
  const nomes = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const p = String(ym).split('-');
  return nomes[Number(p[1]) - 1] + '/' + p[0];
}

/* ===================================================================
   WEB APP
   =================================================================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Recanto da Siriema')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lista explícita do que o site (GitHub Pages) pode chamar. */
function apiPublica_() {
  return {
    resumoPublico: resumoPublico, infoReserva: infoReserva, calendario: calendario, reservar: reservar, comprovantePublico: comprovantePublico,
    anexarComprovante: anexarComprovante, verComprovante: verComprovante,
    sincronizarPublico: sincronizarPublico, sincronizarDiretoria: sincronizarDiretoria, regularizarAssociado: regularizarAssociado,
    login: login, sair: sair, trocarSenha: trocarSenha, quemSou: quemSou, painel: painel,
    salvarAssociado: salvarAssociado, historicoChacara: historicoChacara,
    lancarMensalidade: lancarMensalidade, lancar: lancar, listarLancamentos: listarLancamentos, cancelarLancamento: cancelarLancamento,
    salvarCampanha: salvarCampanha,
    reservasDiretoria: reservasDiretoria, cancelarReserva: cancelarReserva, marcarTaxaPaga: marcarTaxaPaga, marcarAvisoEnviado: marcarAvisoEnviado,
    relatorioMes: relatorioMes, relatorioCampanha: relatorioCampanha, relatorioAdimplencia: relatorioAdimplencia,
    salvarConfig: salvarConfig, listarUsuarios: listarUsuarios, criarUsuario: criarUsuario, novaSenhaUsuario: novaSenhaUsuario, ativarUsuario: ativarUsuario
  };
}

function doPost(e) {
  let saida;
  try {
    const corpo = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const api = apiPublica_();
    const fn = api[corpo.fn];
    if (!fn || !Object.prototype.hasOwnProperty.call(api, corpo.fn)) throw new Error('Função não permitida.');
    saida = { ok: true, result: fn.apply(null, Array.isArray(corpo.args) ? corpo.args : []) };
  } catch (err) {
    saida = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}
