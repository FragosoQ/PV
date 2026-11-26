/**
 * Objetivos Writer — acrescenta linhas à folha de Objetivos a partir de adicionar.html
 *
 * Colunas (linha 1 = cabeçalho):
 *   A ID | B OBJETIVO | C DATA PEDIDO | D DATA PREVISTA | E PEDIDO POR: | F PAÍS | G STATUS | H MODELO | I EVOLUÇÃO
 *
 * INSTALAÇÃO (uma vez, com uma conta que tenha permissão de edição na Sheet):
 *   1. https://script.google.com → Novo projeto → colar este ficheiro em Code.gs
 *   2. Implementar → Nova implementação → Tipo: Aplicação Web
 *        Executar como:      Eu
 *        Quem tem acesso:    Qualquer pessoa
 *   3. Autorizar o acesso às Sheets quando pedido
 *   4. Copiar o URL que termina em /exec e colar em WRITER_URL no adicionar.html
 *
 * Teste: <URL>/exec  → devolve { ok: true, nextId: "00xx", rows: [...] }
 *
 * Depois de alterar este ficheiro: Implementar → Gerir implementações → editar (lápis)
 * → Versão: "Nova versão" → Implementar. Assim o URL /exec mantém-se.
 *
 * POST { action: "add" | "update", id, newId (opcional), objetivo, pedido, prevista,
 *        quem, pais, status, modelo, evolucao }
 *   add:    id opcional — se vier, é usado (tem de estar livre); senão usa o próximo livre
 *   update: id = registo a alterar; newId opcional para mudar o ID (tem de estar livre)
 */

const SHEET_ID = '1VAbosk2nuuzkSqgP5DTS0M3PBnV-Oep0XRGY-aOxjMo';
const SHEET_GID = 0;
// Grafia usada na Sheet. O pedido é comparado sem maiúsculas/acentos ("em curso" → "Em Curso").
const STATUS_OK = ['Em Análise', 'Agendado', 'Em Curso', 'Finalizado'];
const MAX_LEN = 300;

// Devolve todos os registos (lidos diretamente, sem a cache do gviz)
function doGet() {
  const sh = sheet_();
  const last = sh.getLastRow();
  const rows = [];
  if (last >= 2) {
    const range = sh.getRange(2, 1, last - 1, 9);
    const vals = range.getValues();
    const disp = range.getDisplayValues();
    vals.forEach((r, i) => {
      const d = disp[i];
      if (!String(d[1]).trim()) return;
      rows.push({
        id: String(d[0]).trim(),
        objetivo: d[1],
        pedido: iso_(r[2], d[2]),
        prevista: iso_(r[3], d[3]),
        quem: d[4], pais: d[5], status: d[6], modelo: d[7],
        evolucao: evo_(r[8], d[8]),
      });
    });
  }
  return json_({ ok: true, nextId: nextId_(sh), rows: rows });
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'Pedido inválido.' });
  }

  const str = v => String(v == null ? '' : v).trim().slice(0, MAX_LEN);
  const objetivo = str(data.objetivo);
  if (!objetivo) return json_({ ok: false, error: 'O objetivo é obrigatório.' });

  const pedido = parseDate_(data.pedido);
  const prevista = parseDate_(data.prevista);
  if (data.pedido && !pedido) return json_({ ok: false, error: 'Data do pedido inválida.' });
  if (data.prevista && !prevista) return json_({ ok: false, error: 'Data prevista inválida.' });

  const action = data.action === 'update' ? 'update' : 'add';
  const statusIn = canonStatus_(str(data.status));
  let evolucao = '';
  if (data.evolucao !== '' && data.evolucao != null) {
    const n = Number(data.evolucao);
    if (!isFinite(n) || n < 0 || n > 100) return json_({ ok: false, error: 'Evolução deve estar entre 0 e 100.' });
    evolucao = Math.round(n);
  }

  // Lock: evita dois pedidos simultâneos com o mesmo ID / a escrever na mesma linha
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    let id, row, status;

    if (action === 'update') {
      id = str(data.id);
      row = findRow_(sh, id);
      if (!row) return json_({ ok: false, error: 'Registo #' + id + ' não encontrado.' });
      // Status fora da lista (ex.: escrito à mão na Sheet) é mantido se não for alterado
      status = statusIn || sh.getRange(row, 7).getDisplayValue();
      const newId = fmtId_(str(data.newId));
      if (newId && norm_(newId) !== norm_(id)) {
        if (!validId_(newId)) return json_({ ok: false, error: 'ID inválido.' });
        if (findRow_(sh, newId)) return json_({ ok: false, error: 'O ID #' + newId + ' já existe.' });
        id = newId;
        writeId_(sh, row, id);
      }
    } else {
      const wanted = fmtId_(str(data.id));
      if (wanted) {
        if (!validId_(wanted)) return json_({ ok: false, error: 'ID inválido.' });
        if (findRow_(sh, wanted)) return json_({ ok: false, error: 'O ID #' + wanted + ' já existe.' });
        id = wanted;
      } else {
        id = nextId_(sh);
      }
      row = sh.getLastRow() + 1;
      status = statusIn || STATUS_OK[0];
      writeId_(sh, row, id);
    }

    sh.getRange(row, 2, 1, 8).setValues([[
      objetivo, pedido || '', prevista || '', str(data.quem), str(data.pais),
      status, str(data.modelo), evolucao,
    ]]);
    sh.getRange(row, 3, 1, 2).setNumberFormat('dd/MM/yyyy');
    return json_({ ok: true, action: action, id: id, row: row });
  } finally {
    lock.releaseLock();
  }
}

// "6" → "0006"; outros formatos ficam como estão
function fmtId_(v) {
  const s = String(v || '').replace(/^#/, '').trim();
  return /^\d+$/.test(s) ? String(parseInt(s, 10)).padStart(4, '0') : s;
}
function validId_(id) { return /^[A-Za-z0-9-]{1,12}$/.test(id) && norm_(id) !== '0'; }

function writeId_(sh, row, id) {
  sh.getRange(row, 1).setNumberFormat('@'); // mantém os zeros à esquerda ("0007")
  sh.getRange(row, 1).setValue(id);
}

function canonStatus_(v) {
  const k = s => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const key = k(v);
  return STATUS_OK.filter(s => k(s) === key)[0] || '';
}

// Procura a linha pelo ID (compara só os algarismos: "7" = "0007")
function findRow_(sh, id) {
  const key = norm_(id);
  if (!key) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (norm_(ids[i][0]) === key) return i + 2;
  }
  return 0;
}

function norm_(v) {
  const s = String(v == null ? '' : v).trim();
  return /^\d+$/.test(s) ? String(parseInt(s, 10)) : s.toLowerCase();
}

// Data da célula → "aaaa-mm-dd" (para o input type=date). Aceita Date ou texto dd/mm/aaaa.
function iso_(v, disp) {
  let d = v instanceof Date ? v : null;
  if (!d) {
    const m = String(disp || '').trim().match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
    if (m) d = new Date(m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }
  if (!d || isNaN(d.getTime())) return '';
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

// EVOLUÇÃO: "10" = 10%; célula formatada como % ("10%") vem como 0.1
function evo_(v, disp) {
  if (v === '' || v == null) return '';
  let n = Number(v);
  if (!isFinite(n)) n = parseFloat(String(disp).replace(',', '.'));
  if (!isFinite(n)) return '';
  if (/%/.test(String(disp)) && n <= 1) n = n * 100;
  return Math.round(n);
}

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheets().filter(s => s.getSheetId() === SHEET_GID)[0];
  if (!sh) throw new Error('Folha não encontrada (gid ' + SHEET_GID + ')');
  return sh;
}

function nextId_(sh) {
  const last = sh.getLastRow();
  let max = 0;
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(r => {
      const n = parseInt(String(r[0]).replace(/\D/g, ''), 10);
      if (n > max) max = n;
    });
  }
  return String(max + 1).padStart(4, '0');
}

// Aceita "aaaa-mm-dd" (input type=date)
function parseDate_(v) {
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
