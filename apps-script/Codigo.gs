/*****************************************************************
 * LIVRO DE OCORRÊNCIAS DE TURNO – MANUTENÇÃO (relatório separado por oficina)
 * Google Planilhas + Apps Script (site para tablet + e-mail por turno)
 *
 * Abas criadas pelo setup():
 *   Cadastro    -> Crachá | Nome | Ciclo | Início do ciclo | Ativo
 *   Excecoes    -> trocas de turno, férias, atestados (sobrepõem o ciclo)
 *   Ocorrencias -> todos os lançamentos
 *   Pendencias  -> pendências/andamentos em aberto e finalizadas
 *   Config      -> horários dos turnos, e-mails dos líderes etc.
 *   Envios      -> registro dos relatórios já enviados
 *****************************************************************/

const ABA = { CAD: 'Cadastro', EXC: 'Excecoes', OCO: 'Ocorrencias', PEN: 'Pendencias', CFG: 'Config', ENV: 'Envios' };
const TZ = 'America/Sao_Paulo';
const UTC_OFFSET_H = 3;            // Brasil sem horário de verão: hora local = UTC-3
const TURNOS = ['A', 'B', 'C', 'D'];   // D = 6x1 (tem relatório e horário de envio próprios)
function rot_(t) { return t === 'D' ? '6X1' : 'TURNO ' + t; }
function letra_(t) { return t === 'D' ? '6x1' : t; }
const TOLERANCIA_MIN = 90;         // antecedência/atraso aceito para sugerir o turno da escala

const CAB = {
  Cadastro:    ['Crachá', 'Nome', 'Ciclo', '1º dia de serviço', 'Ativo', 'Oficina', 'Perfil'],
  Excecoes:    ['Data início', 'Data fim', 'Crachá', 'Turno (A/B/C/D=6x1/F=folga)', 'Motivo'],
  Ocorrencias: ['ID', 'Data turno', 'Turno', 'Crachá', 'Nome', 'Ocorrência', 'Início', 'Fim', 'Status',
                'Descrição da pendência', 'Nº LT', 'Ordem/Nota', 'Registrado em', 'Atualizado em',
                'Fora da escala', 'Após envio', 'Ref. pendência', 'Justificativa sem ordem', 'Equipamento/TAG', 'Oficina'],
  Pendencias:  ['ID', 'Aberta em', 'Turno', 'Crachá', 'Nome', 'Ocorrência', 'Descrição da pendência', 'Nº LT',
                'Ordem/Nota', 'Situação', 'Status original', 'ID ocorrência origem', 'Finalizada em',
                'Turno finalização', 'Crachá finalizador', 'Nome finalizador', 'O que foi feito', 'Registrado em', 'Equipamento/TAG', 'Oficina'],
  Config:      ['Parâmetro', 'Valor', 'Observação'],
  Envios:      ['Data turno', 'Turno', 'Enviado em', 'Destinatários', 'Ocorrências', 'Modo', 'Oficina']
};

const CONFIG_PADRAO = [
  ['TURNO_A_INICIO', '07:00', 'hh:mm'],
  ['TURNO_A_FIM', '15:35', 'hh:mm'],
  ['TURNO_B_INICIO', '15:20', 'hh:mm'],
  ['TURNO_B_FIM', '23:40', 'hh:mm'],
  ['TURNO_C_INICIO', '23:40', 'hh:mm'],
  ['TURNO_C_FIM', '07:00', 'hh:mm (termina no dia seguinte)'],
  ['TURNO_D_INICIO', '07:00', 'hh:mm - 6x1'],
  ['TURNO_D_FIM', '16:48', 'hh:mm - 6x1'],
  ['ENVIO_A', '15:50', 'Horário do e-mail do turno A'],
  ['ENVIO_D', '17:00', 'Horário do e-mail do 6x1'],
  ['ENVIO_B', '23:50', 'Horário do e-mail do turno B'],
  ['ENVIO_C', '07:10', 'Horário do e-mail do turno C (manhã seguinte)'],
  ['ATRASO_ENVIO_MIN', '15', 'Reserva: usado só se o ENVIO do turno estiver vazio (minutos após o fim)'],
  ['EMAILS_LIDERES', '', 'Separe por vírgula: lider1@empresa.com, lider2@empresa.com'],
  ['EMAIL_ADMIN', '', 'Recebe aviso se o envio automático falhar (vazio = dono da planilha)'],
  ['ENVIO_ATIVO', 'SIM', 'SIM ou NAO'],
  ['LOGOUT_MIN', '5', 'Minutos sem uso para o tablet voltar à tela do crachá'],
  ['AMBIENTE', 'PRODUCAO', 'PRODUCAO ou TESTE (TESTE mostra faixa amarela no site e [TESTE] no assunto do e-mail)']
];

/* =================== SITE =================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(tituloSite_())
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Nome da aba do navegador: "[TESTE]" na frente quando AMBIENTE = TESTE
function tituloSite_() {
  let teste = false;
  try { teste = ehTeste_(); } catch (e) {}
  return (teste ? '[TESTE] ' : '') + 'Livro de Ocorrências';
}

/* =================== MENU NA PLANILHA =================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Livro de Ocorrências')
    .addItem('1. Configuração inicial', 'setup')
    .addItem('Enviar relatório de um turno agora...', 'enviarManual')
    .addItem('Reinstalar envio automático', 'instalarGatilho')
    .addItem('Definir senha do administrador', 'definirSenhaAdmin')
    .addSeparator()
    .addItem('Limpar histórico de lançamentos...', 'limparHistorico')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  ss.setSpreadsheetTimeZone(TZ);

  Object.keys(CAB).forEach(function (nome) {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, CAB[nome].length).setValues([CAB[nome]])
      .setFontWeight('bold').setBackground('#0032A0').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), CAB[nome].length).setNumberFormat('@'); // tudo como texto
  });

  const cfg = ss.getSheetByName(ABA.CFG);
  // Acrescenta parâmetros que ainda não existem (sem mexer nos que já foram ajustados)
  const chavesCfg = cfg.getLastRow() > 1 ? cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  CONFIG_PADRAO.forEach(function (c) { if (chavesCfg.indexOf(c[0]) < 0) cfg.appendRow(c); });

  const cad = ss.getSheetByName(ABA.CAD);
  if (cad.getLastRow() < 2) {
    cad.getRange(2, 1, 4, 5).setValues([
      ['1001', 'ELETRICISTA EXEMPLO TURNO A', 'A,A,A,A,A,F', '09/11/2026', 'SIM'],
      ['1002', 'ELETRICISTA EXEMPLO TURNO B', 'B,B,B,B,B,F', '09/11/2026', 'SIM'],
      ['1003', 'ELETRICISTA EXEMPLO TURNO C', 'C,C,C,C,C,F', '09/11/2026', 'SIM'],
      ['1004', 'FOLGUISTA EXEMPLO', 'B,B,C,C,F,F', '09/11/2026', 'SIM']
    ]);
    cad.getRange('H1').setValue('Ciclo: um item por dia (A, B, C, D=6x1 ou F=folga). "1º dia de serviço" = qualquer data em que a pessoa voltou de folga (folguista: 1º dia de B). O sistema calcula as folgas a partir dela.');
  }
  cad.getRange('E2:E500').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NAO']).build());
  const exc = ss.getSheetByName(ABA.EXC);
  exc.getRange('D2:D1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['A', 'B', 'C', 'D', 'F']).build());

  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1) ss.deleteSheet(padrao);

  instalarGatilho();
  const ui = ui_();
  if (!ui) {   // rodado pelo editor do Apps Script: sem janelas na planilha
    console.log('Configuração concluída. Agora, na PLANILHA, use o menu Livro de Ocorrências > Definir senha do administrador.');
    return;
  }
  if (!PropertiesService.getScriptProperties().getProperty('ADMIN_HASH')) definirSenhaAdmin();
  ui.alert('Configuração concluída.\n\n1) Preencha a aba Cadastro com os eletricistas.\n' +
    '2) Coloque os e-mails dos líderes na aba Config.\n3) Publique o site em Implantar > Nova implantação > App da Web.');
}

function instalarGatilho() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'verificarEnvios') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('verificarEnvios').timeBased().everyMinutes(10).create();
}

/* =================== UTILITÁRIOS DE DATA/HORA =================== */

function pad2(n) { return ('0' + n).slice(-2); }
function hmToMin(s) {
  const m = String(s || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? (+m[1]) * 60 + (+m[2]) : null;
}
function dkFromParts(y, m, d) { return y + '-' + pad2(m) + '-' + pad2(d); }
function dkAdd(dk, n) {
  const p = dk.split('-').map(Number);
  const t = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return dkFromParts(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
}
function dkDiff(a, b) { // b - a em dias
  const pa = a.split('-').map(Number), pb = b.split('-').map(Number);
  return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000);
}
function dkToBr(dk) { const p = dk.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
function paraDk(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return dkFromParts(+m[3], +m[2], +m[1]);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? s.slice(0, 10) : '';
}
function paraHm(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  const m = hmToMin(v);
  return m === null ? '' : pad2(Math.floor(m / 60)) + ':' + pad2(m % 60);
}
function localMs(dk, min) {
  const p = dk.split('-').map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2], 0, min) + UTC_OFFSET_H * 3600000;
}
function agora() {
  const now = new Date();
  const s = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm');
  return { ms: now.getTime(), dk: s.slice(0, 10), hm: s.slice(11), carimbo: Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm:ss') };
}
function normCracha(v) {
  const s = String(v || '').trim();
  return /^\d+$/.test(s) ? String(Number(s)) : s.toUpperCase();
}

/* =================== OFICINAS =================== */

const SEM_OFICINA = 'SEM OFICINA';
function normOf_(s) { return String(s || '').trim().toUpperCase() || SEM_OFICINA; }
// Chave usada na Config para os e-mails da oficina: INSTRUMENTAÇÃO -> EMAILS_INSTRUMENTACAO
function chaveOf_(of) {
  return 'EMAILS_' + normOf_(of).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');
}
function mapaOficinas_() {
  const m = {};
  pessoas_().forEach(function (p) { m[p.cracha] = normOf_(p.oficina); });
  return m;
}
// Oficina de uma ocorrência/pendência: a gravada na linha ou, em linhas antigas, a do cadastro atual
function ofDe_(gravada, cracha, mapa) { return gravada ? normOf_(gravada) : (mapa[normCracha(cracha)] || SEM_OFICINA); }

function emailsOficina_(cfg, of) {
  const proprios = String(cfg[chaveOf_(of)] || '').split(/[,;\s]+/).filter(function (e) { return e.indexOf('@') > 0; });
  if (proprios.length) return proprios;
  return String(cfg.EMAILS_LIDERES || '').split(/[,;\s]+/).filter(function (e) { return e.indexOf('@') > 0; });
}

// Oficinas que têm alguém escalado ou algum lançamento no turno
function oficinasDoTurno_(dk, turno) {
  const mapa = mapaOficinas_(), set = {};
  escalados_(dk, turno, pessoas_(), excecoes_()).forEach(function (p) { set[normOf_(p.oficina)] = 1; });
  ler_(ABA.OCO).rows.map(ocoParaObj_).forEach(function (o) {
    if (o.dk === dk && o.turno === turno) set[ofDe_(o.oficina, o.cracha, mapa)] = 1;
  });
  return Object.keys(set).sort();
}

/* =================== LEITURA DAS ABAS =================== */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function aba_(nome) {
  const sh = ss_().getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não encontrada. Rode a Configuração inicial.');
  return sh;
}
function ler_(nome) {
  const sh = aba_(nome);
  const n = sh.getLastRow() - 1;
  const cols = CAB[nome].length;
  const vals = n > 0 ? sh.getRange(2, 1, n, cols).getValues() : [];
  return { sh: sh, rows: vals };
}
function col_(nome, titulo) { return CAB[nome].indexOf(titulo); }

function config_() {
  const cfg = {};
  ler_(ABA.CFG).rows.forEach(function (r) { if (r[0]) cfg[String(r[0]).trim()] = r[1]; });
  CONFIG_PADRAO.forEach(function (c) { if (cfg[c[0]] === undefined || (cfg[c[0]] === '' && /^TURNO_/.test(c[0]))) cfg[c[0]] = c[1]; });
  cfg.turnos = {};
  TURNOS.forEach(function (t) {
    cfg.turnos[t] = { ini: hmToMin(paraHm(cfg['TURNO_' + t + '_INICIO'])), fim: hmToMin(paraHm(cfg['TURNO_' + t + '_FIM'])) };
  });
  cfg.atraso = Number(cfg.ATRASO_ENVIO_MIN || 15);
  cfg.envio = {};
  TURNOS.forEach(function (t) { cfg.envio[t] = hmToMin(paraHm(cfg['ENVIO_' + t])); });
  cfg.logout = Number(cfg.LOGOUT_MIN || 5);
  return cfg;
}

function pessoas_() {
  return ler_(ABA.CAD).rows.filter(function (r) { return String(r[0]).trim() !== ''; }).map(function (r) {
    return {
      cracha: normCracha(r[0]),
      nome: String(r[1]).trim().toUpperCase(),
      ciclo: String(r[2]).toUpperCase().split(/[,;\s]+/).filter(String),
      inicio: paraDk(r[3]),
      ativo: String(r[4]).trim().toUpperCase() !== 'NAO', oficina: String(r[5] || '').trim().toUpperCase(),
      lider: ehLider_(r[6])
    };
  });
}
function ehLider_(v) { return /^L[IÍ]DER$/.test(String(v || '').trim().toUpperCase()); }
function excecoes_() {
  return ler_(ABA.EXC).rows.filter(function (r) { return r[0] && r[2]; }).map(function (r) {
    const ini = paraDk(r[0]);
    return { ini: ini, fim: paraDk(r[1]) || ini, cracha: normCracha(r[2]), turno: String(r[3]).trim().toUpperCase(), motivo: r[4] };
  });
}

/* =================== ESCALA =================== */

// Retorna 'A', 'B', 'C' ou 'F' (folga) para a pessoa na data.
function turnoEscala(p, dk, excs) {
  for (let i = 0; i < excs.length; i++) {
    const e = excs[i];
    if (e.cracha === p.cracha && dk >= e.ini && dk <= e.fim) return e.turno;
  }
  if (!p.ciclo.length || !p.inicio) return '';
  const n = p.ciclo.length;
  const idx = ((dkDiff(p.inicio, dk) % n) + n) % n;
  return p.ciclo[idx];
}

function janela(dk, t, cfg) {
  const j = cfg.turnos[t];
  const fim = j.fim <= j.ini ? j.fim + 1440 : j.fim;
  return { ini: localMs(dk, j.ini), fim: localMs(dk, fim) };
}
function distancia(ms, w) { return ms < w.ini ? w.ini - ms : (ms >= w.fim ? ms - w.fim : 0); }

// Para cada turno, qual data de turno está mais próxima de agora
function opcoesTurno_(now, cfg) {
  const op = {};
  TURNOS.forEach(function (t) {
    let melhor = null;
    [-1, 0, 1].forEach(function (k) {
      const dk = dkAdd(now.dk, k);
      const d = distancia(now.ms, janela(dk, t, cfg));
      if (!melhor || d < melhor.d) melhor = { dk: dk, d: d };
    });
    op[t] = { dk: melhor.dk, br: dkToBr(melhor.dk), atual: melhor.d === 0 };
  });
  return op;
}

function escalados_(dk, t, pess, excs) {
  return pess.filter(function (p) { return p.ativo && turnoEscala(p, dk, excs) === t; });
}

/* =================== API CHAMADA PELO TABLET =================== */

function identificar(crachaDigitado) {
  const cracha = normCracha(crachaDigitado);
  const p = pessoas_().filter(function (x) { return x.cracha === cracha; })[0];
  if (!p) throw new Error('Crachá ' + crachaDigitado + ' não cadastrado.');
  if (!p.ativo) throw new Error('Crachá ' + crachaDigitado + ' está inativo no cadastro.');

  const cfg = config_(), now = agora(), excs = excecoes_();
  const opcoes = opcoesTurno_(now, cfg);

  // Turno da escala mais próximo de agora (ontem, hoje ou amanhã)
  let sug = null;
  [-1, 0, 1].forEach(function (k) {
    const dk = dkAdd(now.dk, k);
    const t = turnoEscala(p, dk, excs);
    if (TURNOS.indexOf(t) < 0) return;
    const d = distancia(now.ms, janela(dk, t, cfg));
    if (!sug || d < sug.d) sug = { turno: t, dk: dk, d: d };
  });
  let foraEscala = false;
  if (!sug || sug.d > TOLERANCIA_MIN * 60000) {
    const atual = TURNOS.filter(function (t) { return opcoes[t].atual; })[0] || 'A';
    sug = { turno: atual, dk: opcoes[atual].dk };
    foraEscala = true;
  }
  return {
    cracha: p.cracha, nome: p.nome, oficina: p.oficina, lider: p.lider,
    turno: sug.turno, dk: sug.dk, br: dkToBr(sug.dk),
    escalaHoje: turnoEscala(p, now.dk, excs),
    escalaPorDia: [-1, 0, 1].reduce(function (m, k) { const d = dkAdd(now.dk, k); m[d] = turnoEscala(p, d, excs); return m; }, {}),
    foraEscala: foraEscala, opcoes: opcoes, logoutMin: cfg.logout
  };
}

// cracha vazio (administrador) = todas as oficinas; senão, só a oficina de quem entrou
function carregarTurno(cracha, turno, dk) {
  const oco = ler_(ABA.OCO).rows;
  const C = function (t) { return col_(ABA.OCO, t); };
  const cfg = config_(), mapa = mapaOficinas_();
  const minhaOf = cracha ? (mapa[normCracha(cracha)] || SEM_OFICINA) : null;
  const lista = oco.filter(function (r) { return paraDk(r[C('Data turno')]) === dk && r[C('Turno')] === turno; })
    .map(ocoParaObj_)
    .map(function (o) { o.oficina = ofDe_(o.oficina, o.cracha, mapa); return o; })
    .filter(function (o) { return !minhaOf || o.oficina === minhaOf; })
    .sort(ordenarOco_(cfg.turnos[turno].ini));
  const now = agora();
  const pend = ler_(ABA.PEN).rows.filter(function (r) {
      return r[col_(ABA.PEN, 'Situação')] === 'ABERTA' &&
        (!minhaOf || ofDe_(r[col_(ABA.PEN, 'Oficina')], r[col_(ABA.PEN, 'Crachá')], mapa) === minhaOf);
    })
    .map(function (r) {
      const P = function (t) { return r[col_(ABA.PEN, t)]; };
      const ab = paraDk(P('Aberta em'));
      return {
        id: P('ID'), abertaEm: dkToBr(ab), dias: dkDiff(ab, now.dk), turno: P('Turno'), nome: P('Nome'),
        ocorrencia: P('Ocorrência'), descricao: P('Descrição da pendência'), lt: P('Nº LT'),
        ordem: P('Ordem/Nota'), statusOriginal: P('Status original'), equipamento: P('Equipamento/TAG') || '',
        oficina: ofDe_(P('Oficina'), P('Crachá'), mapa)
      };
    }).sort(function (a, b) { return b.dias - a.dias; });
  return { ocorrencias: lista, pendencias: pend, enviado: minhaOf ? jaEnviado_(dk, turno, minhaOf) : jaEnviado_(dk, turno), cracha: normCracha(cracha), equipamentos: listaEquip_(oco) };
}

// Equipamentos já usados (mais frequentes primeiro), para sugerir no tablet
function listaEquip_(rows) {
  const c = col_(ABA.OCO, 'Equipamento/TAG'), cont = {};
  rows.forEach(function (r) { const e = String(r[c] || '').trim(); if (e) cont[e] = (cont[e] || 0) + 1; });
  return Object.keys(cont).sort(function (a, b) { return cont[b] - cont[a]; }).slice(0, 300);
}

function salvarOcorrencia(cracha, d) { return salvarOcorrencia_(cracha, d, false); }

function travado_(row, porAdmin) {
  if (porAdmin === true) return;
  const of = ofDe_(row[col_(ABA.OCO, 'Oficina')], row[col_(ABA.OCO, 'Crachá')], mapaOficinas_());
  if (jaEnviado_(paraDk(row[col_(ABA.OCO, 'Data turno')]), row[col_(ABA.OCO, 'Turno')], of) && row[col_(ABA.OCO, 'Após envio')] !== 'SIM')
    throw new Error('O relatório deste turno já foi enviado aos líderes. Só o administrador pode alterar.');
}

function salvarOcorrencia_(cracha, d, porAdmin) {
  cracha = normCracha(cracha);
  const p = pessoas_().filter(function (x) { return x.cracha === cracha; })[0];
  if (!p) throw new Error('Crachá não cadastrado.');
  validar_(d);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = aba_(ABA.OCO), now = agora();
    const fora = turnoEscala(p, d.dk, excecoes_()) !== d.turno ? 'SIM' : '';
    const apos = jaEnviado_(d.dk, d.turno, normOf_(p.oficina)) ? 'SIM' : '';
    let obj;
    if (d.id) {
      const lin = acharLinha_(ABA.OCO, 'ID', d.id);
      if (!lin) throw new Error('Lançamento não encontrado.');
      const row = sh.getRange(lin, 1, 1, CAB.Ocorrencias.length).getValues()[0];
      if (normCracha(row[col_(ABA.OCO, 'Crachá')]) !== cracha) throw new Error('Só quem lançou pode editar.');
      travado_(row, porAdmin);
      row[col_(ABA.OCO, 'Ocorrência')] = d.ocorrencia;
      row[col_(ABA.OCO, 'Início')] = d.inicio;
      row[col_(ABA.OCO, 'Fim')] = d.fim;
      row[col_(ABA.OCO, 'Status')] = d.status;
      row[col_(ABA.OCO, 'Descrição da pendência')] = d.descricao;
      row[col_(ABA.OCO, 'Nº LT')] = d.lt;
      row[col_(ABA.OCO, 'Ordem/Nota')] = d.ordem;
      row[col_(ABA.OCO, 'Justificativa sem ordem')] = d.justificativa;
      row[col_(ABA.OCO, 'Equipamento/TAG')] = d.equipamento;
      if (!row[col_(ABA.OCO, 'Oficina')]) row[col_(ABA.OCO, 'Oficina')] = normOf_(p.oficina);
      row[col_(ABA.OCO, 'Atualizado em')] = now.carimbo + (porAdmin === true ? ' (ADMIN)' : '');
      sh.getRange(lin, 1, 1, row.length).setValues([row]);
      obj = ocoParaObj_(row);
    } else {
      const row = [Utilities.getUuid().slice(0, 8), dkToBr(d.dk), d.turno, p.cracha, p.nome, d.ocorrencia,
        d.inicio, d.fim, d.status, d.descricao, d.lt, d.ordem, now.carimbo + (porAdmin === true ? ' (ADMIN)' : ''), '', fora, apos, d.ref || '', d.justificativa, d.equipamento, normOf_(p.oficina)];
      sh.appendRow(row);
      obj = ocoParaObj_(row);
    }
    if (!d.ref) sincronizarPendencia_(obj);
  } finally { lock.releaseLock(); }
  return carregarTurno(cracha, d.turno, d.dk);
}

function excluirOcorrencia(cracha, id, turno, dk) { return excluirOcorrencia_(cracha, id, turno, dk, false); }

function excluirOcorrencia_(cracha, id, turno, dk, porAdmin) {
  cracha = normCracha(cracha);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const lin = acharLinha_(ABA.OCO, 'ID', id);
    if (!lin) throw new Error('Lançamento não encontrado.');
    const sh = aba_(ABA.OCO);
    const row = sh.getRange(lin, 1, 1, CAB.Ocorrencias.length).getValues()[0];
    if (normCracha(row[col_(ABA.OCO, 'Crachá')]) !== cracha) throw new Error('Só quem lançou pode excluir.');
    travado_(row, porAdmin);
    const linP = acharLinha_(ABA.PEN, 'ID ocorrência origem', id);
    if (linP) {
      const sit = aba_(ABA.PEN).getRange(linP, col_(ABA.PEN, 'Situação') + 1).getValue();
      if (sit === 'ABERTA') aba_(ABA.PEN).deleteRow(linP);
    }
    sh.deleteRow(lin);
  } finally { lock.releaseLock(); }
  return carregarTurno(cracha, turno, dk);
}

function finalizarPendencia(cracha, pendId, turno, dk, d) {
  cracha = normCracha(cracha);
  const p = pessoas_().filter(function (x) { return x.cracha === cracha; })[0];
  if (!p) throw new Error('Crachá não cadastrado.');
  if (!d || String(d.feito || '').trim().length < 5) throw new Error('Descreva o que foi feito para finalizar.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let pend;
  try {
    const shP = aba_(ABA.PEN);
    const lin = acharLinha_(ABA.PEN, 'ID', pendId);
    if (!lin) throw new Error('Pendência não encontrada.');
    const row = shP.getRange(lin, 1, 1, CAB.Pendencias.length).getValues()[0];
    if (row[col_(ABA.PEN, 'Situação')] !== 'ABERTA') throw new Error('Essa pendência já foi finalizada.');
    const now = agora();
    row[col_(ABA.PEN, 'Situação')] = 'FINALIZADA';
    row[col_(ABA.PEN, 'Finalizada em')] = dkToBr(dk);
    row[col_(ABA.PEN, 'Turno finalização')] = turno;
    row[col_(ABA.PEN, 'Crachá finalizador')] = p.cracha;
    row[col_(ABA.PEN, 'Nome finalizador')] = p.nome;
    row[col_(ABA.PEN, 'O que foi feito')] = String(d.feito).trim().toUpperCase();
    row[col_(ABA.PEN, 'Registrado em')] = now.carimbo;
    shP.getRange(lin, 1, 1, row.length).setValues([row]);
    pend = row;
  } finally { lock.releaseLock(); }

  // Lança também como ocorrência de quem finalizou
  const P = function (t) { return pend[col_(ABA.PEN, t)]; };
  const texto = 'FINALIZAÇÃO DE PENDÊNCIA (' + P('Aberta em') + ' ' + rot_(P('Turno')) + ' - ' + P('Nome') + '): ' +
    String(d.feito).trim().toUpperCase();
  return salvarOcorrencia(cracha, {
    turno: turno, dk: dk, ocorrencia: texto, inicio: d.inicio, fim: d.fim, status: 'CONCLUIDO',
    descricao: '', lt: d.lt || P('Nº LT'), ordem: d.ordem || P('Ordem/Nota'), justificativa: d.justificativa, equipamento: P('Equipamento/TAG') || '', ref: pendId
  });
}

/* =================== AUXILIARES =================== */

function validar_(d) {
  if (TURNOS.indexOf(d.turno) < 0) throw new Error('Turno inválido.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.dk || '')) throw new Error('Data inválida.');
  d.ocorrencia = String(d.ocorrencia || '').trim().toUpperCase();
  d.descricao = String(d.descricao || '').trim().toUpperCase();
  d.lt = String(d.lt || '').trim().toUpperCase();
  d.equipamento = String(d.equipamento || '').trim().toUpperCase();
  if (!d.ref && d.equipamento.length < 2) throw new Error('Informe o equipamento/TAG.');
  d.ordem = String(d.ordem || '').replace(/\D/g, '');
  d.inicio = paraHm(d.inicio); d.fim = paraHm(d.fim);
  d.status = String(d.status || '').toUpperCase();
  if (d.ocorrencia.length < 5) throw new Error('Descreva a ocorrência.');
  if (!d.inicio) throw new Error('Informe a hora de início.');
  // PENDENTE saiu das opções; só continua aceito para lançamentos antigos que já estavam assim
  if (['CONCLUIDO', 'ANDAMENTO'].indexOf(d.status) < 0 && !(d.status === 'PENDENTE' && d.id)) throw new Error('Escolha o status: Concluído ou Andamento.');
  if (d.status === 'CONCLUIDO' && !d.fim) throw new Error('Informe a hora de fim.');
  if (d.status !== 'CONCLUIDO' && !d.ref && d.descricao.length < 5) throw new Error('Descreva a pendência.');
  d.justificativa = d.ordem ? '' : String(d.justificativa || '').trim().toUpperCase();
  if (!d.ordem && d.justificativa.length < 5) throw new Error('Informe a ordem/nota ou justifique por que está sem.');
}

function ocoParaObj_(r) {
  const C = function (t) { return r[col_(ABA.OCO, t)]; };
  return {
    id: C('ID'), dk: paraDk(C('Data turno')), turno: C('Turno'), cracha: normCracha(C('Crachá')), nome: C('Nome'),
    ocorrencia: C('Ocorrência'), inicio: paraHm(C('Início')), fim: paraHm(C('Fim')), status: C('Status'),
    descricao: C('Descrição da pendência'), lt: C('Nº LT'), ordem: String(C('Ordem/Nota') || ''),
    foraEscala: C('Fora da escala') === 'SIM', aposEnvio: C('Após envio') === 'SIM', ref: C('Ref. pendência'),
    justificativa: C('Justificativa sem ordem') || '', equipamento: C('Equipamento/TAG') || '',
    oficina: C('Oficina') || ''
  };
}

function ordenarOco_(iniTurno) {
  const k = function (hm) { const m = hmToMin(hm); return m === null ? 9999 : (m < iniTurno - 120 ? m + 1440 : m); };
  return function (a, b) { return a.nome < b.nome ? -1 : a.nome > b.nome ? 1 : k(a.inicio) - k(b.inicio); };
}

function acharLinha_(nome, coluna, valor) {
  const sh = aba_(nome);
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const vals = sh.getRange(2, col_(nome, coluna) + 1, n, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(valor)) return i + 2;
  return 0;
}

function sincronizarPendencia_(o) {
  const sh = aba_(ABA.PEN);
  const lin = acharLinha_(ABA.PEN, 'ID ocorrência origem', o.id);
  if (o.status === 'CONCLUIDO') {
    if (lin && sh.getRange(lin, col_(ABA.PEN, 'Situação') + 1).getValue() === 'ABERTA') sh.deleteRow(lin);
    return;
  }
  if (lin) {
    const row = sh.getRange(lin, 1, 1, CAB.Pendencias.length).getValues()[0];
    if (row[col_(ABA.PEN, 'Situação')] !== 'ABERTA') return;
    row[col_(ABA.PEN, 'Ocorrência')] = o.ocorrencia;
    row[col_(ABA.PEN, 'Descrição da pendência')] = o.descricao;
    row[col_(ABA.PEN, 'Nº LT')] = o.lt;
    row[col_(ABA.PEN, 'Ordem/Nota')] = o.ordem;
    row[col_(ABA.PEN, 'Status original')] = o.status;
    row[col_(ABA.PEN, 'Equipamento/TAG')] = o.equipamento;
    if (!row[col_(ABA.PEN, 'Oficina')]) row[col_(ABA.PEN, 'Oficina')] = o.oficina || '';
    sh.getRange(lin, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow([Utilities.getUuid().slice(0, 8), dkToBr(o.dk), o.turno, o.cracha, o.nome, o.ocorrencia,
      o.descricao, o.lt, o.ordem, 'ABERTA', o.status, o.id, '', '', '', '', '', '', o.equipamento, o.oficina || '']);
  }
}

// Sem oficina: algum envio do turno. Com oficina: envio daquela oficina (envios antigos, sem oficina, valem para todas)
function jaEnviado_(dk, turno, oficina) {
  return ler_(ABA.ENV).rows.some(function (r) {
    if (paraDk(r[0]) !== dk || r[1] !== turno) return false;
    return oficina === undefined || !r[6] || normOf_(r[6]) === normOf_(oficina);
  });
}

/* =================== ENVIO DO RELATÓRIO =================== */

// Rodado pelo gatilho a cada 10 minutos
function verificarEnvios() {
  const cfg = config_();
  if (String(cfg.ENVIO_ATIVO || 'SIM').toUpperCase() !== 'SIM') return;
  const now = agora();
  [-1, 0].forEach(function (k) {
    const dk = dkAdd(now.dk, k);
    TURNOS.forEach(function (t) {
      const prazo = momentoEnvio_(dk, t, cfg);
      if (now.ms < prazo || now.ms >= prazo + 6 * 3600000) return;
      const faltam = oficinasDoTurno_(dk, t).filter(function (of) { return !jaEnviado_(dk, t, of); });
      if (faltam.length) {
        try { enviarRelatorio(dk, t, 'AUTOMÁTICO', faltam); }
        catch (e) { avisarFalha_(dk, t, e); }   // tenta de novo na próxima rodada (10 min), por até 6 h
      }
    });
  });
}

// Momento (ms) em que sai o e-mail do turno t da data dk
function momentoEnvio_(dk, t, cfg) {
  const w = janela(dk, t, cfg);
  const e = cfg.envio[t];
  if (e === null || e === undefined) return w.fim + cfg.atraso * 60000;
  let ms = localMs(dk, e);
  while (ms < w.fim) ms += 86400000;   // ex.: turno C (23:40) -> envio 07:15 do dia seguinte
  return ms;
}

// Registra a falha e avisa o administrador por e-mail (uma vez por turno)
function avisarFalha_(dk, t, erro) {
  const props = PropertiesService.getScriptProperties();
  const msg = String(erro && erro.message ? erro.message : erro);
  const chave = dk + '|' + t;
  props.setProperty('ULTIMA_FALHA', JSON.stringify({ dk: dk, turno: t, em: agora().carimbo, erro: msg }));
  if (props.getProperty('FALHA_AVISADA') === chave) return;
  try {
    let para = '';
    try { para = String(config_().EMAIL_ADMIN || '').trim(); } catch (e2) {}
    if (!para) para = Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(para, 'FALHA no envio do Livro de Ocorrências - ' + rot_(t) + ' - ' + dkToBr(dk),
      'O relatório do ' + rot_(t) + ' de ' + dkToBr(dk) + ' NÃO foi enviado aos líderes.\n\nMotivo: ' + msg +
      '\n\nO sistema continua tentando a cada 10 minutos, por até 6 horas. ' +
      'Você também pode enviar pela área do Administrador > Relatórios > ENVIAR AGORA.');
    props.setProperty('FALHA_AVISADA', chave);
  } catch (e3) { console.error('Não consegui avisar o administrador: ' + e3); }
}

function enviarManual() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Enviar relatório', 'Data do turno (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('Enviar relatório', 'Turno (A, B, C ou 6X1):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const dk = paraDk(r1.getResponseText());
  let t = r2.getResponseText().trim().toUpperCase();
  if (t === '6X1') t = 'D';
  if (!dk || TURNOS.indexOf(t) < 0) { ui.alert('Data ou turno inválido.'); return; }
  enviarRelatorio(dk, t, 'MANUAL');
  ui.alert('Relatório do ' + rot_(t) + ' de ' + dkToBr(dk) + ' enviado.');
}

// Envia um e-mail por oficina. oficinas: lista a enviar (vazio = todas as do turno)
function enviarRelatorio(dk, turno, modo, oficinas) {
  const cfg = config_();
  const lista = (oficinas && oficinas.length) ? oficinas.map(normOf_) : oficinasDoTurno_(dk, turno);
  if (!lista.length) throw new Error('Ninguém escalado e nenhum lançamento neste turno.');
  const erros = [];
  lista.forEach(function (of) {
    try { enviarOficina_(dk, turno, of, modo, cfg); }
    catch (e) { erros.push(of + ': ' + (e && e.message ? e.message : e)); }
  });
  if (erros.length) throw new Error(erros.join(' | '));
  const props = PropertiesService.getScriptProperties();
  try {
    const f = JSON.parse(props.getProperty('ULTIMA_FALHA') || 'null');
    if (f && f.dk === dk && f.turno === turno) props.deleteProperty('ULTIMA_FALHA');
  } catch (e) {}
  return lista;
}

function enviarOficina_(dk, turno, of, modo, cfg) {
  const emails = emailsOficina_(cfg, of);
  if (!emails.length) throw new Error('nenhum e-mail de líder cadastrado em Configurações.');
  if (MailApp.getRemainingDailyQuota() < emails.length) throw new Error('cota diária de e-mails do Google esgotada.');
  const dados = montarDados_(dk, turno, cfg, of);
  const html = montarHtml_(dados, cfg);
  const nomeTurno = turno === 'D' ? '6x1' : 'Turno ' + turno;
  const titulo = (ehTeste_(cfg) ? '[TESTE] ' : '') + 'Livro de Ocorrências - ' + of + ' - ' + nomeTurno + ' - ' + dkToBr(dk);
  const pdf = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF)
    .setName('Livro_Ocorrencias_' + of.replace(/[^A-Z0-9]+/gi, '_') + '_' + dk + '_' + nomeTurno.replace(' ', '_') + '.pdf');
  MailApp.sendEmail({ to: emails.join(','), subject: titulo, htmlBody: html, attachments: [pdf], name: 'Livro de Ocorrências' });
  aba_(ABA.ENV).appendRow([dkToBr(dk), turno, agora().carimbo, emails.join(', '), dados.ocorrencias.length, modo, of]);
}

function montarDados_(dk, turno, cfg, of) {
  of = normOf_(of);
  const pess = pessoas_(), excs = excecoes_(), mapa = mapaOficinas_();
  const ocos = ler_(ABA.OCO).rows.map(ocoParaObj_)
    .filter(function (o) { return o.dk === dk && o.turno === turno && ofDe_(o.oficina, o.cracha, mapa) === of; })
    .sort(ordenarOco_(cfg.turnos[turno].ini));
  const escal = escalados_(dk, turno, pess, excs).filter(function (p) { return normOf_(p.oficina) === of; });
  const lancaram = {};
  ocos.forEach(function (o) { lancaram[o.cracha] = true; });
  const naoLancaram = escal.filter(function (p) { return !lancaram[p.cracha]; }).map(function (p) { return p.nome; });
  const foraEscala = [];
  ocos.forEach(function (o) { if (o.foraEscala && foraEscala.indexOf(o.nome) < 0) foraEscala.push(o.nome); });

  const pendRows = ler_(ABA.PEN).rows;
  const P = function (r, t) { return r[col_(ABA.PEN, t)]; };
  const hoje = agora().dk;
  const daOf = function (r) { return ofDe_(P(r, 'Oficina'), P(r, 'Crachá'), mapa) === of; };
  const abertas = pendRows.filter(function (r) { return P(r, 'Situação') === 'ABERTA' && daOf(r); }).map(function (r) {
    return { abertaEm: P(r, 'Aberta em'), dias: dkDiff(paraDk(P(r, 'Aberta em')), hoje), turno: P(r, 'Turno'),
      nome: P(r, 'Nome'), ocorrencia: P(r, 'Ocorrência'), descricao: P(r, 'Descrição da pendência'),
      ordem: P(r, 'Ordem/Nota'), status: P(r, 'Status original'), equipamento: P(r, 'Equipamento/TAG') || '' };
  }).sort(function (a, b) { return b.dias - a.dias; });
  const finalizadas = pendRows.filter(function (r) {
    return P(r, 'Situação') === 'FINALIZADA' && paraDk(P(r, 'Finalizada em')) === dk && P(r, 'Turno finalização') === turno && daOf(r);
  }).map(function (r) {
    return { abertaEm: P(r, 'Aberta em'), turno: P(r, 'Turno'), nome: P(r, 'Nome'), ocorrencia: P(r, 'Ocorrência'),
      por: P(r, 'Nome finalizador'), feito: P(r, 'O que foi feito') };
  });

  const cont = { CONCLUIDO: 0, ANDAMENTO: 0, PENDENTE: 0 };
  ocos.forEach(function (o) { cont[o.status] = (cont[o.status] || 0) + 1; });
  return { dk: dk, turno: turno, oficina: of, ocorrencias: ocos, escalados: escal.length, lancaram: Object.keys(lancaram).length,
    naoLancaram: naoLancaram, foraEscala: foraEscala, abertas: abertas, finalizadas: finalizadas, cont: cont };
}

function esc_(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function montarHtml_(d, cfg) {
  const AZUL = '#0032A0', VERDE = '#00B5AD';
  const corStatus = { CONCLUIDO: '#00A651', ANDAMENTO: '#FFDE1A', PENDENTE: '#E53935' };
  const txtStatus = { CONCLUIDO: '#000', ANDAMENTO: '#000', PENDENTE: '#fff' };
  const td = 'border:1px solid #BFC7D5;padding:4px 6px;font-size:12px;vertical-align:middle;';
  const th = 'border:1px solid #BFC7D5;padding:6px;font-size:12px;background:' + AZUL + ';color:#fff;';
  const j = cfg.turnos[d.turno];
  const hm = function (m) { return pad2(Math.floor(m / 60)) + ':' + pad2(m % 60); };

  let h = '<div style="font-family:Verdana,Arial,sans-serif;color:#1a1a1a;max-width:1100px">';
  h += '<table style="width:100%;border-collapse:collapse"><tr>' +
    '<td style="background:' + AZUL + ';color:#fff;padding:12px 14px;font-size:18px;font-weight:bold">LIVRO DE OCORRÊNCIAS DE TURNO – MANUTENÇÃO<br><span style="font-size:14px">OFICINA: ' + esc_(d.oficina) + '</span></td>' +
    '<td style="background:' + VERDE + ';color:#fff;padding:12px 14px;font-size:16px;font-weight:bold;text-align:center;white-space:nowrap">' +
    rot_(d.turno) + '<br><span style="font-size:13px;font-weight:normal">' + dkToBr(d.dk) + ' · ' + hm(j.ini) + ' às ' + hm(j.fim) + '</span></td></tr></table>';

  // Resumo
  const card = function (rot, val, cor) {
    return '<td style="padding:10px;border:1px solid #DDE3EE;text-align:center;width:20%"><div style="font-size:22px;font-weight:bold;color:' + cor + '">' + val +
      '</div><div style="font-size:11px;color:#555">' + rot + '</div></td>';
  };
  h += '<table style="width:100%;border-collapse:collapse;margin:10px 0"><tr>' +
    card('OCORRÊNCIAS', d.ocorrencias.length, AZUL) + card('CONCLUÍDAS', d.cont.CONCLUIDO || 0, '#00A651') +
    card('EM ANDAMENTO', d.cont.ANDAMENTO || 0, '#B38F00') + card('PENDENTES', d.cont.PENDENTE || 0, '#E53935') +
    card('LANÇARAM / ESCALADOS', d.lancaram + ' / ' + d.escalados, AZUL) + '</tr></table>';

  if (d.naoLancaram.length) {
    h += '<div style="background:#FDECEA;border-left:5px solid #E53935;padding:8px 10px;font-size:12px;margin-bottom:8px">' +
      '<b>Escalados que não lançaram ocorrências:</b> ' + esc_(d.naoLancaram.join(', ')) + '</div>';
  }
  const semOrdem = d.ocorrencias.filter(function (o) { return !o.ordem; }).length;
  if (semOrdem) {
    h += '<div style="background:#FFF8D6;border-left:5px solid #FF9421;padding:8px 10px;font-size:12px;margin-bottom:8px">' +
      '<b>' + semOrdem + ' ocorrência(s) sem ordem/nota</b> — justificativas na última coluna.</div>';
  }
  if (d.foraEscala.length) {
    h += '<div style="background:#FFF8D6;border-left:5px solid #FFDE1A;padding:8px 10px;font-size:12px;margin-bottom:8px">' +
      '<b>Lançaram fora da própria escala:</b> ' + esc_(d.foraEscala.join(', ')) + '</div>';
  }

  // Tabela principal
  h += '<table style="width:100%;border-collapse:collapse"><tr>' +
    [['NOME', '10%'], ['EQUIPAMENTO', '11%'], ['OCORRÊNCIAS', '27%'], ['INÍCIO', '5%'], ['FIM', '5%'], ['STATUS', '9%'],
     ['DESCRIÇÃO DA PENDÊNCIA', '15%'], ['Nº LT', '6%'], ['ORDEM OU NOTA', '12%']]
      .map(function (c) { return '<th style="' + th + 'width:' + c[1] + '">' + c[0] + '</th>'; }).join('') + '</tr>';
  if (!d.ocorrencias.length) {
    h += '<tr><td colspan="9" style="' + td + 'text-align:center;color:#777;padding:14px">Nenhuma ocorrência lançada neste turno.</td></tr>';
  }
  let i = 0;
  while (i < d.ocorrencias.length) {
    const nome = d.ocorrencias[i].nome;
    let n = 0;
    while (i + n < d.ocorrencias.length && d.ocorrencias[i + n].nome === nome) n++;
    for (let k = 0; k < n; k++) {
      const o = d.ocorrencias[i + k];
      h += '<tr>';
      if (k === 0) h += '<td rowspan="' + n + '" style="' + td + 'font-weight:bold;text-align:center;background:#F3F6FB">' + esc_(nome) + '</td>';
      h += '<td style="' + td + 'font-weight:bold">' + esc_(o.equipamento) + '</td>' +
        '<td style="' + td + '">' + esc_(o.ocorrencia) + (o.aposEnvio ? ' <i style="color:#999">(lançado após envio)</i>' : '') + '</td>' +
        '<td style="' + td + 'text-align:center">' + esc_(o.inicio) + '</td>' +
        '<td style="' + td + 'text-align:center">' + esc_(o.fim) + '</td>' +
        '<td style="' + td + 'text-align:center;font-weight:bold;background:' + (corStatus[o.status] || '#fff') + ';color:' + (txtStatus[o.status] || '#000') + '">' + esc_(o.status) + '</td>' +
        '<td style="' + td + '">' + esc_(o.descricao) + '</td>' +
        '<td style="' + td + 'text-align:center;white-space:nowrap">' + esc_(o.lt) + '</td>' +
        (o.ordem ? '<td style="' + td + 'text-align:center">' + esc_(o.ordem) + '</td></tr>'
                 : '<td style="' + td + 'font-size:11px;background:#FFF6C2"><b>SEM ORDEM:</b> ' + esc_(o.justificativa) + '</td></tr>');
    }
    i += n;
  }
  h += '</table>';

  // Pendências finalizadas no turno
  if (d.finalizadas.length) {
    h += '<h3 style="color:' + AZUL + ';font-size:14px;margin:18px 0 6px">PENDÊNCIAS FINALIZADAS NESTE TURNO</h3>' +
      '<table style="width:100%;border-collapse:collapse"><tr>' +
      ['ABERTA EM', 'ABERTA POR', 'OCORRÊNCIA', 'FINALIZADA POR', 'O QUE FOI FEITO']
        .map(function (c) { return '<th style="' + th + 'background:' + VERDE + '">' + c + '</th>'; }).join('') + '</tr>';
    d.finalizadas.forEach(function (f) {
      h += '<tr><td style="' + td + '">' + esc_(f.abertaEm) + ' (' + esc_(letra_(f.turno)) + ')</td><td style="' + td + '">' + esc_(f.nome) +
        '</td><td style="' + td + '">' + esc_(f.ocorrencia) + '</td><td style="' + td + '">' + esc_(f.por) +
        '</td><td style="' + td + '">' + esc_(f.feito) + '</td></tr>';
    });
    h += '</table>';
  }

  // Pendências em aberto
  h += '<h3 style="color:' + AZUL + ';font-size:14px;margin:18px 0 6px">PENDÊNCIAS EM ABERTO (' + d.abertas.length + ')</h3>';
  if (!d.abertas.length) {
    h += '<div style="font-size:12px;color:#777">Nenhuma pendência em aberto.</div>';
  } else {
    h += '<table style="width:100%;border-collapse:collapse"><tr>' +
      ['DIAS', 'ABERTA EM', 'POR', 'OCORRÊNCIA', 'DESCRIÇÃO DA PENDÊNCIA', 'STATUS', 'ORDEM']
        .map(function (c) { return '<th style="' + th + '">' + c + '</th>'; }).join('') + '</tr>';
    d.abertas.forEach(function (a) {
      const cor = a.dias >= 7 ? '#E53935' : a.dias >= 3 ? '#FF9421' : '#1a1a1a';
      h += '<tr><td style="' + td + 'text-align:center;font-weight:bold;color:' + cor + '">' + a.dias + '</td><td style="' + td + '">' +
        esc_(a.abertaEm) + ' (' + esc_(letra_(a.turno)) + ')</td><td style="' + td + '">' + esc_(a.nome) + '</td><td style="' + td + '">' +
        (a.equipamento ? '<b>' + esc_(a.equipamento) + '</b> – ' : '') + esc_(a.ocorrencia) + '</td><td style="' + td + '">' + esc_(a.descricao) + '</td><td style="' + td + 'text-align:center;background:' +
        (corStatus[a.status] || '#fff') + ';color:' + (txtStatus[a.status] || '#000') + '">' + esc_(a.status) + '</td><td style="' + td + 'text-align:center">' +
        esc_(a.ordem) + '</td></tr>';
    });
    h += '</table>';
  }
  h += '<p style="font-size:10px;color:#888;margin-top:14px">Relatório gerado automaticamente em ' + agora().carimbo + '.</p></div>';
  return h;
}


/*****************************************************************
 * ADMINISTRADOR (protegido por senha)
 *****************************************************************/

const ADMIN_SESSAO_SEG = 7200;      // sessão de 2 h, renovada a cada uso
const ADMIN_MAX_FALHAS = 5;         // depois disso, bloqueia por 5 min

function hash_(s) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'livro-ocorrencias#' + s, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(b);
}

// Janelas da planilha só existem quando a função é chamada pelo menu da planilha
function ui_() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } }

function definirSenhaAdmin() {
  const ui = ui_();
  if (!ui) throw new Error('Abra a PLANILHA e use o menu Livro de Ocorrências > Definir senha do administrador.');
  const r1 = ui.prompt('Senha do administrador', 'Digite a nova senha (mínimo 6 caracteres):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('Senha do administrador', 'Repita a senha:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const s1 = r1.getResponseText(), s2 = r2.getResponseText();
  if (s1.length < 6) { ui.alert('A senha precisa ter pelo menos 6 caracteres.'); return; }
  if (s1 !== s2) { ui.alert('As senhas não conferem.'); return; }
  PropertiesService.getScriptProperties().setProperty('ADMIN_HASH', hash_(s1));
  ui.alert('Senha do administrador definida.');
}

function adminLogin(senha) {
  const cache = CacheService.getScriptCache();
  const falhas = Number(cache.get('adm_falhas') || 0);
  if (falhas >= ADMIN_MAX_FALHAS) throw new Error('Muitas tentativas erradas. Aguarde 5 minutos.');
  const h = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH');
  if (!h) throw new Error('Senha de administrador ainda não definida (menu da planilha).');
  if (hash_(String(senha || '')) !== h) {
    cache.put('adm_falhas', String(falhas + 1), 300);
    throw new Error('Senha incorreta.');
  }
  cache.remove('adm_falhas');
  const token = Utilities.getUuid();
  cache.put('adm_' + token, '1', ADMIN_SESSAO_SEG);
  const cfg = config_(), now = agora();
  return { token: token, hoje: now.dk, opcoes: opcoesTurno_(now, cfg), logoutMin: cfg.logout };
}

function exigirAdmin_(token) {
  const c = CacheService.getScriptCache();
  if (!token || !c.get('adm_' + token)) throw new Error('Sessão de administrador expirada. Entre de novo.');
  c.put('adm_' + token, '1', ADMIN_SESSAO_SEG);
}

function adminSair(token) { CacheService.getScriptCache().remove('adm_' + token); return true; }

/* ---------- Funcionários ---------- */

function adminDados(token) {
  exigirAdmin_(token);
  const cad = ler_(ABA.CAD).rows;
  const pessoas = cad.filter(function (r) { return String(r[0]).trim() !== ''; }).map(function (r) {
    return { cracha: normCracha(r[0]), nome: String(r[1]).toUpperCase(), ciclo: String(r[2]).toUpperCase().replace(/\s/g, ''),
      inicio: paraDk(r[3]), ativo: String(r[4]).trim().toUpperCase() !== 'NAO', oficina: String(r[5] || '').trim().toUpperCase(), lider: ehLider_(r[6]) };
  }).sort(function (a, b) { return a.nome < b.nome ? -1 : 1; });
  const exc = ler_(ABA.EXC).rows.map(function (r, i) {
    const ini = paraDk(r[0]);
    return { linha: i + 2, ini: ini, fim: paraDk(r[1]) || ini, cracha: normCracha(r[2]), turno: String(r[3]).toUpperCase(), motivo: String(r[4] || '') };
  }).filter(function (e) { return e.ini && e.cracha; });
  const nomes = {};
  pessoas.forEach(function (p) { nomes[p.cracha] = p.nome; });
  exc.forEach(function (e) { e.nome = nomes[e.cracha] || '(crachá ' + e.cracha + ')'; });
  const cfgRows = ler_(ABA.CFG).rows.filter(function (r) { return r[0]; }).map(function (r) {
    return { chave: String(r[0]), valor: /^(TURNO_|ENVIO_[ABCD]$)/.test(String(r[0])) ? paraHm(r[1]) : String(r[1]), obs: String(r[2] || '') };
  });
  const ofs = {};
  pessoas.forEach(function (p) { if (p.ativo) ofs[normOf_(p.oficina)] = 1; });
  const oficinas = Object.keys(ofs).sort().map(function (o) { return { nome: o, chave: chaveOf_(o) }; });
  return { pessoas: pessoas, excecoes: exc.sort(function (a, b) { return a.ini < b.ini ? 1 : -1; }), config: cfgRows, oficinas: oficinas };
}

function adminSalvarPessoa(token, p, crachaOriginal) {
  exigirAdmin_(token);
  const cracha = normCracha(p.cracha);
  if (!/^\d+$/.test(cracha)) throw new Error('Crachá deve ter só números.');
  const nome = String(p.nome || '').trim().toUpperCase();
  if (nome.length < 3) throw new Error('Informe o nome.');
  const lider = p.lider === true;
  const ciclo = String(p.ciclo || '').toUpperCase().split(/[,;\s]+/).filter(String);
  if ((!ciclo.length && !lider) || ciclo.some(function (c) { return ['A', 'B', 'C', 'D', 'F'].indexOf(c) < 0; }))
    throw new Error('Ciclo inválido. Use só A, B, C, D (6x1) ou F separados por vírgula.');
  const inicio = paraDk(p.inicio);
  if (!inicio && ciclo.length) throw new Error('Informe o 1º dia de serviço (primeiro dia de trabalho depois de uma folga).');
  const linhaNova = [cracha, nome, ciclo.join(','), inicio ? dkToBr(inicio) : '', p.ativo === false ? 'NAO' : 'SIM', String(p.oficina || '').trim().toUpperCase(),
    lider ? 'LÍDER' : 'COLABORADOR'];

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = aba_(ABA.CAD);
    const n = sh.getLastRow() - 1;
    const crachas = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return normCracha(r[0]); }) : [];
    const orig = crachaOriginal ? normCracha(crachaOriginal) : '';
    const idxOrig = orig ? crachas.indexOf(orig) : -1;
    const idxDup = crachas.indexOf(cracha);
    if (idxDup >= 0 && idxDup !== idxOrig) throw new Error('Já existe um funcionário com o crachá ' + cracha + '.');
    if (idxOrig >= 0) sh.getRange(idxOrig + 2, 1, 1, 7).setValues([linhaNova]);
    else sh.appendRow(linhaNova);
  } finally { lock.releaseLock(); }
  return adminDados(token);
}

/* ---------- Escala e exceções ---------- */

function adminEscala(token, dkIni, dias) {
  exigirAdmin_(token);
  dkIni = paraDk(dkIni) || agora().dk;
  dias = Math.min(Number(dias) || 14, 31);
  const datas = [];
  for (let i = 0; i < dias; i++) datas.push(dkAdd(dkIni, i));
  const excs = excecoes_();
  const linhas = pessoas_().filter(function (p) { return p.ativo; })
    .sort(function (a, b) { return a.nome < b.nome ? -1 : 1; })
    .map(function (p) {
      return { cracha: p.cracha, nome: p.nome,
        dias: datas.map(function (dk) {
          const exc = excs.some(function (e) { return e.cracha === p.cracha && dk >= e.ini && dk <= e.fim; });
          return { t: turnoEscala(p, dk, excs), exc: exc };
        }) };
    });
  const totais = datas.map(function (dk, i) {
    const c = { A: 0, B: 0, C: 0, D: 0 };
    linhas.forEach(function (l) { if (c[l.dias[i].t] !== undefined) c[l.dias[i].t]++; });
    return c;
  });
  return { datas: datas, linhas: linhas, totais: totais };
}

function adminSalvarExcecao(token, e) {
  exigirAdmin_(token);
  const ini = paraDk(e.ini), fim = paraDk(e.fim) || ini;
  const cracha = normCracha(e.cracha), turno = String(e.turno || '').toUpperCase();
  if (!ini) throw new Error('Informe a data de início.');
  if (fim < ini) throw new Error('A data fim é antes da data início.');
  if (['A', 'B', 'C', 'D', 'F'].indexOf(turno) < 0) throw new Error('Escolha o turno (ou folga).');
  if (!pessoas_().some(function (p) { return p.cracha === cracha; })) throw new Error('Funcionário não encontrado.');
  aba_(ABA.EXC).appendRow([dkToBr(ini), dkToBr(fim), cracha, turno, String(e.motivo || '').trim().toUpperCase()]);
  return adminDados(token);
}

function adminExcluirExcecao(token, linha, cracha) {
  exigirAdmin_(token);
  const sh = aba_(ABA.EXC);
  if (linha < 2 || linha > sh.getLastRow() || normCracha(sh.getRange(linha, 3).getValue()) !== normCracha(cracha))
    throw new Error('A exceção mudou na planilha. Recarregue e tente de novo.');
  sh.deleteRow(linha);
  return adminDados(token);
}

/* ---------- Lançamentos e pendências ---------- */

function adminLancamentos(token, dk, turno) {
  exigirAdmin_(token);
  if (TURNOS.indexOf(turno) < 0 || !paraDk(dk)) throw new Error('Escolha data e turno.');
  const r = carregarTurno('', turno, paraDk(dk));
  const excs = excecoes_();
  r.escalados = escalados_(paraDk(dk), turno, pessoas_(), excs).map(function (p) { return { cracha: p.cracha, nome: p.nome, oficina: normOf_(p.oficina) }; });
  r.enviadoOf = {};
  oficinasDoTurno_(paraDk(dk), turno).forEach(function (of) { r.enviadoOf[of] = jaEnviado_(paraDk(dk), turno, of); });
  return r;
}

function donoOcorrencia_(id) {
  const lin = acharLinha_(ABA.OCO, 'ID', id);
  if (!lin) throw new Error('Lançamento não encontrado.');
  return normCracha(aba_(ABA.OCO).getRange(lin, col_(ABA.OCO, 'Crachá') + 1).getValue());
}

function adminSalvarOcorrencia(token, d) {
  exigirAdmin_(token);
  const dono = d.id ? donoOcorrencia_(d.id) : normCracha(d.cracha);
  if (!dono) throw new Error('Escolha o eletricista.');
  salvarOcorrencia_(dono, d, true);
  return adminLancamentos(token, d.dk, d.turno);
}

function adminExcluirOcorrencia(token, id, turno, dk) {
  exigirAdmin_(token);
  excluirOcorrencia_(donoOcorrencia_(id), id, turno, dk, true);
  return adminLancamentos(token, dk, turno);
}

function adminCancelarPendencia(token, pendId, motivo, turno, dk) {
  exigirAdmin_(token);
  motivo = String(motivo || '').trim().toUpperCase();
  if (motivo.length < 5) throw new Error('Informe o motivo do cancelamento.');
  const sh = aba_(ABA.PEN);
  const lin = acharLinha_(ABA.PEN, 'ID', pendId);
  if (!lin) throw new Error('Pendência não encontrada.');
  const row = sh.getRange(lin, 1, 1, CAB.Pendencias.length).getValues()[0];
  if (row[col_(ABA.PEN, 'Situação')] !== 'ABERTA') throw new Error('Essa pendência não está mais aberta.');
  const now = agora();
  row[col_(ABA.PEN, 'Situação')] = 'CANCELADA';
  row[col_(ABA.PEN, 'Finalizada em')] = dkToBr(now.dk);
  row[col_(ABA.PEN, 'Nome finalizador')] = 'ADMINISTRADOR';
  row[col_(ABA.PEN, 'O que foi feito')] = 'CANCELADA: ' + motivo;
  row[col_(ABA.PEN, 'Registrado em')] = now.carimbo;
  sh.getRange(lin, 1, 1, row.length).setValues([row]);
  return adminLancamentos(token, dk, turno);
}

/* ---------- Relatórios ---------- */

// of vazio = todas as oficinas do turno (uma depois da outra)
function adminPreviewRelatorio(token, dk, turno, of) {
  exigirAdmin_(token);
  dk = paraDk(dk);
  if (!dk || TURNOS.indexOf(turno) < 0) throw new Error('Escolha data e turno.');
  const cfg = config_();
  const lista = of ? [normOf_(of)] : oficinasDoTurno_(dk, turno);
  if (!lista.length) return { html: '<p style="font-family:Verdana;padding:20px">Ninguém escalado e nenhum lançamento neste turno.</p>', itens: [] };
  const htmls = [], itens = [];
  lista.forEach(function (o) {
    htmls.push(montarHtml_(montarDados_(dk, turno, cfg, o), cfg));
    itens.push({ oficina: o, enviado: jaEnviado_(dk, turno, o), emails: emailsOficina_(cfg, o).join(', ') });
  });
  return { html: htmls.join('<hr style="border:0;border-top:4px dashed #BFC7D5;margin:28px 0">'), itens: itens };
}

function adminEnviarRelatorio(token, dk, turno, of) {
  exigirAdmin_(token);
  dk = paraDk(dk);
  if (!dk || TURNOS.indexOf(turno) < 0) throw new Error('Escolha data e turno.');
  return { ok: true, oficinas: enviarRelatorio(dk, turno, 'ADMIN', of ? [of] : null) };
}

function adminHistoricoEnvios(token) {
  exigirAdmin_(token);
  const envios = ler_(ABA.ENV).rows.slice(-30).reverse().map(function (r) {
    return { data: String(r[0]), turno: r[1], em: String(r[2]), para: String(r[3]), qtd: r[4], modo: r[5], oficina: String(r[6] || 'TODAS') };
  });
  let falha = null;
  try { falha = JSON.parse(PropertiesService.getScriptProperties().getProperty('ULTIMA_FALHA') || 'null'); } catch (e) {}
  if (falha) falha.data = dkToBr(falha.dk);
  return { envios: envios, falha: falha };
}

/* ---------- Configurações ---------- */

function adminSalvarConfig(token, valores) {
  exigirAdmin_(token);
  const permitidas = CONFIG_PADRAO.map(function (c) { return c[0]; });
  Object.keys(valores).forEach(function (k) {
    if (permitidas.indexOf(k) < 0 && !/^EMAILS_[A-Z0-9_]+$/.test(k)) throw new Error('Parâmetro desconhecido: ' + k);
    let v = String(valores[k] || '').trim();
    if (/^(TURNO_|ENVIO_[ABCD]$)/.test(k)) { v = paraHm(v); if (!v) throw new Error('Horário inválido em ' + k + '.'); }
    if (k === 'ATRASO_ENVIO_MIN' || k === 'LOGOUT_MIN') { if (!/^\d+$/.test(v)) throw new Error(k + ' deve ser um número.'); }
    if (k === 'EMAIL_ADMIN' && v && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) throw new Error('E-mail do administrador inválido.');
    if (k === 'ENVIO_ATIVO') v = v.toUpperCase() === 'NAO' ? 'NAO' : 'SIM';
    if (k.indexOf('EMAILS_') === 0) {
      const lista = v.split(/[,;\s]+/).filter(String);
      lista.forEach(function (e) { if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('E-mail inválido: ' + e); });
      v = lista.join(', ');
    }
    valores[k] = v;
  });
  const sh = aba_(ABA.CFG);
  const n = sh.getLastRow() - 1;
  const chaves = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  Object.keys(valores).forEach(function (k) {
    const i = chaves.indexOf(k);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(valores[k]);
    else {
      const padrao = CONFIG_PADRAO.filter(function (c) { return c[0] === k; })[0];
      sh.appendRow([k, valores[k], padrao ? padrao[2] : 'E-mails dos líderes da oficina (vazio = usa EMAILS_LIDERES)']);
    }
  });
  return adminDados(token);
}

function adminTrocarSenha(token, atual, nova) {
  exigirAdmin_(token);
  const props = PropertiesService.getScriptProperties();
  if (hash_(String(atual || '')) !== props.getProperty('ADMIN_HASH')) throw new Error('Senha atual incorreta.');
  if (String(nova || '').length < 6) throw new Error('A nova senha precisa ter pelo menos 6 caracteres.');
  props.setProperty('ADMIN_HASH', hash_(String(nova)));
  return true;
}


/*****************************************************************
 * TELA DO LÍDER (perfil LÍDER no cadastro)
 *****************************************************************/

// Dados de um dia (todos os turnos) da oficina do líder. Líder sem oficina vê todas.
function liderPainel(cracha, dkPedido) {
  cracha = normCracha(cracha);
  const p = pessoas_().filter(function (x) { return x.cracha === cracha; })[0];
  if (!p || !p.ativo || !p.lider) throw new Error('Acesso restrito aos líderes.');
  const now = agora(), cfg = config_(), mapa = mapaOficinas_(), excs = excecoes_(), pess = pessoas_();
  const opcoes = opcoesTurno_(now, cfg);
  const dk = paraDk(dkPedido) || (opcoes[TURNOS.filter(function (t) { return opcoes[t].atual; })[0] || 'A'] || {}).dk || now.dk;
  const of = p.oficina ? normOf_(p.oficina) : null;
  const daOf = function (o) { return !of || o === of; };

  const ocos = ler_(ABA.OCO).rows.map(ocoParaObj_)
    .map(function (o) { o.oficina = ofDe_(o.oficina, o.cracha, mapa); return o; })
    .filter(function (o) { return o.dk === dk && daOf(o.oficina); });
  const escalados = [], enviado = {}, janelas = {};
  TURNOS.forEach(function (t) {
    escalados_(dk, t, pess, excs).forEach(function (x) {
      if (daOf(normOf_(x.oficina))) escalados.push({ cracha: x.cracha, nome: x.nome, turno: t, oficina: normOf_(x.oficina) });
    });
    enviado[t] = of ? jaEnviado_(dk, t, of) : jaEnviado_(dk, t);
    const j = cfg.turnos[t];
    janelas[t] = paraHm(pad2(Math.floor(j.ini / 60)) + ':' + pad2(j.ini % 60)) + ' às ' + paraHm(pad2(Math.floor(j.fim / 60)) + ':' + pad2(j.fim % 60));
  });
  const P = function (r, t) { return r[col_(ABA.PEN, t)]; };
  const pend = ler_(ABA.PEN).rows.filter(function (r) {
    return P(r, 'Situação') === 'ABERTA' && daOf(ofDe_(P(r, 'Oficina'), P(r, 'Crachá'), mapa));
  }).map(function (r) {
    const ab = paraDk(P(r, 'Aberta em'));
    return { id: P(r, 'ID'), abertaEm: dkToBr(ab), dias: dkDiff(ab, now.dk), turno: P(r, 'Turno'), nome: P(r, 'Nome'),
      ocorrencia: P(r, 'Ocorrência'), descricao: P(r, 'Descrição da pendência'), lt: P(r, 'Nº LT'), ordem: P(r, 'Ordem/Nota'),
      statusOriginal: P(r, 'Status original'), equipamento: P(r, 'Equipamento/TAG') || '', oficina: ofDe_(P(r, 'Oficina'), P(r, 'Crachá'), mapa) };
  }).sort(function (a, b) { return b.dias - a.dias; });

  return { nome: p.nome, oficina: of || 'TODAS', dk: dk, br: dkToBr(dk), hoje: now.dk,
    turnoAtual: TURNOS.filter(function (t) { return opcoes[t].atual; })[0] || 'A',
    ocorrencias: ocos, escalados: escalados, pendencias: pend, enviado: enviado, janelas: janelas, logoutMin: cfg.logout };
}


/*****************************************************************
 * AMBIENTE DE TESTE E LIMPEZA DO HISTÓRICO
 *****************************************************************/

function ehTeste_(cfg) { return String((cfg || config_()).AMBIENTE || '').trim().toUpperCase() === 'TESTE'; }

// Chamado pela tela do crachá para saber se mostra a faixa de teste
function infoSite() { return { teste: ehTeste_() }; }

// Apaga lançamentos, pendências e registro de envios. Mantém Cadastro, Excecoes e Config.
// Antes de apagar, guarda uma cópia numa planilha nova no seu Drive.
function limparHistorico() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Limpar histórico de lançamentos',
    'Isso apaga TODOS os lançamentos, pendências e registros de envio.\n' +
    'Funcionários, escala e configurações continuam.\n' +
    'Uma cópia de segurança será salva no seu Google Drive antes.\n\n' +
    'Para confirmar, digite LIMPAR:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  if (r.getResponseText().trim().toUpperCase() !== 'LIMPAR') { ui.alert('Nada foi apagado.'); return; }

  const ss = ss_();
  const backup = SpreadsheetApp.create('Backup Livro de Ocorrências ' + agora().carimbo.replace(/[\/:]/g, '-'));
  [ABA.OCO, ABA.PEN, ABA.ENV].forEach(function (nome) {
    ss.getSheetByName(nome).copyTo(backup).setName(nome);
  });
  const vazia = backup.getSheets()[0];
  if (backup.getSheets().length > 1 && vazia.getLastRow() === 0) backup.deleteSheet(vazia);

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    [ABA.OCO, ABA.PEN, ABA.ENV].forEach(function (nome) {
      const sh = aba_(nome);
      if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    });
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('ULTIMA_FALHA'); props.deleteProperty('FALHA_AVISADA');
  } finally { lock.releaseLock(); }
  ui.alert('Histórico apagado.\n\nCópia de segurança: "' + backup.getName() + '" no seu Google Drive.');
}
