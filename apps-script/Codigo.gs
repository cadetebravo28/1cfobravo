/**
 * API do Portal Aloja (1º CFO BM 2026 · Aspirantes 2028)
 * Lê as planilhas da turma e entrega os dados ao site em JSON.
 * Implantar como App da Web: Executar como "Eu" · Acesso "Qualquer pessoa".
 * Só os intervalos listados em FONTES podem ser lidos, e só com usuário e senha da turma.
 */
const USUARIO = 'COLOQUE_O_USUARIO';
const SENHA   = 'COLOQUE_A_SENHA';
const ESC = '1Fll0Q0Vt--jNqC1awyE2A-7JpKZcUhnpCBy-ghPPH1k';   // Escala Aspirantes 2028
const CFO = '1I01qBi7HCdUHcmlG5CIbDY3UKjMqlGFrd3cX-fG48e8';   // 1º CFO BM 2026 - BRAVO
const FX9 = '1qfS_bQqVxjGjfUHP9abjqmwg6pU9F0XgP7adtdvqSkA';   // Faxina Aloja 9
const FUN = '1K5xIBsUxgcVW48XSGLplXat0OfjCA8JMvf50RnEGr0A';   // Pelotão Bravo (FUNREBRAVO)
const ANO = '1pThqWbY11CR697tpHBBwmay680qH5hOysHM0VYDjaOc';   // Anotações Cadetes

const FONTES = {
  escala:  { id: ESC, aba: 'Escala',                      linhas: [3, 70] },
  esc_res: { id: ESC, aba: 'Dados Residência',            a1: 'A1:C70' },
  cont:    { id: ESC, aba: 'Contagem',                    a1: 'A4:L60' },
  qts:     { id: CFO, aba: 'Quadro de Trabalho Semanal',  a1: 'A1:J80' },
  ctrl:    { id: CFO, aba: 'Controle de CH',              a1: 'A1:G150' },
  marc:    { id: CFO, aba: 'Marcação de aulas',           colunas: [3, 8], desde: 2 },
  fax9:    { id: FX9, aba: 'FAXINA DIÁRIA',               a1: 'A1:D200' },
  fun:     { id: FUN, aba: 'FUNREBRAVO',                  a1: 'A1:N80' },
  fun_s:   { id: FUN, aba: 'SAÍDAS FUNREBRAVO',           a1: 'A1:H400' },
  an_resp: { id: ANO, aba: 'Respostas Revisadas',         a1: 'A1:G3000' },
  an_dados:{ id: ANO, aba: 'Dados',                       a1: 'A1:D400' },
  an_sem:  { id: ANO, aba: 'Anotações',                   a1: 'A1:F300' },
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!autorizado(p.u, p.s)) return json({ error: 'auth' });
  try {
    if (p.a === 'login') return json({ ok: true });
    if (p.a === 'v') return json(lerFonte(p.k));
    if (p.a === 'fotos') return json({ fotos: lerFotos() });
    return json({ ok: true, fontes: Object.keys(FONTES) });
  } catch (err) {
    return json({ error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    if (!autorizado(b.u, b.s)) return json({ error: 'auth' });
    if (b.a === 'foto') return json(salvarFoto(b.nr, b.img));
    return json({ error: 'ação desconhecida' });
  } catch (err) {
    return json({ error: String(err && err.message || err) });
  }
}

function autorizado(u, s) {
  return String(u || '').trim().toLowerCase() === USUARIO.toLowerCase() && String(s || '') === SENHA;
}

function lerFonte(k) {
  const f = FONTES[k];
  if (!f) throw new Error('fonte não permitida');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('v:' + k);
  if (hit) return JSON.parse(hit);
  const sh = SpreadsheetApp.openById(f.id).getSheetByName(f.aba);
  if (!sh) throw new Error('aba não encontrada: ' + f.aba);
  let rg;
  if (f.a1) rg = sh.getRange(f.a1);
  else if (f.linhas) rg = sh.getRange(f.linhas[0], 1, f.linhas[1] - f.linhas[0] + 1, Math.max(1, sh.getLastColumn()));
  else {
    const n = Math.max(1, sh.getLastRow() - f.desde + 1);
    rg = sh.getRange(f.desde, f.colunas[0], n, f.colunas[1] - f.colunas[0] + 1);
  }
  const out = { values: aparar(rg.getDisplayValues()), at: new Date().toISOString() };
  try { cache.put('v:' + k, JSON.stringify(out), 120); } catch (e) { /* grande demais para o cache */ }
  return out;
}

// Igual à API do Google Sheets: sem células vazias no fim das linhas nem linhas vazias no fim.
function aparar(v) {
  const r = v.map(row => { let n = row.length; while (n > 0 && row[n - 1] === '') n--; return row.slice(0, n); });
  let m = r.length; while (m > 0 && r[m - 1].length === 0) m--;
  return r.slice(0, m);
}

// ---- Fotos de perfil (guardadas numa planilha criada automaticamente no seu Drive) ----
function abaFotos() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('FOTOS_ID'), ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create('Portal Aloja · fotos de perfil'); props.setProperty('FOTOS_ID', ss.getId()); ss.getSheets()[0].appendRow(['nr', 'img', 'atualizado']); }
  return ss.getSheets()[0];
}

function lerFotos() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('fotos'); if (hit) return JSON.parse(hit);
  const v = abaFotos().getDataRange().getValues(); const o = {};
  v.slice(1).forEach(r => { if (r[0] && r[1]) o[String(r[0])] = String(r[1]); });
  try { cache.put('fotos', JSON.stringify(o), 60); } catch (e) {}
  return o;
}

function salvarFoto(nr, img) {
  nr = String(nr || '');
  if (!/^1\d{3}$/.test(nr)) throw new Error('número inválido');
  if (typeof img !== 'string' || !/^data:image\/jpeg;base64,[A-Za-z0-9+/=]+$/.test(img) || img.length > 48000) throw new Error('imagem inválida');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = abaFotos(); const v = sh.getDataRange().getValues();
    const i = v.findIndex((r, j) => j > 0 && String(r[0]) === nr);
    if (i > 0) sh.getRange(i + 1, 2, 1, 2).setValues([[img, new Date()]]);
    else sh.appendRow([nr, img, new Date()]);
    CacheService.getScriptCache().remove('fotos');
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
