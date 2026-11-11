/**
 * API do Portal Aloja (1º CFO BM 2026 · Aspirantes 2028)
 * Lê as planilhas da turma e entrega os dados ao site em JSON.
 * Implantar como App da Web: Executar como "Eu" · Acesso "Qualquer pessoa".
 * Só os intervalos listados em FONTES podem ser lidos, e só com usuário e senha da turma.
 */
const USUARIO = 'COLOQUE_O_USUARIO';
const SENHA   = 'COLOQUE_A_SENHA';
const ADMIN   = 'COLOQUE_A_SENHA_DE_MODERACAO';   // senha de moderação do mural
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
    if (p.a === 'feed') return json(lerFeed(p.antes));
    if (p.a === 'stories') return json(lerStories());
    if (p.a === 'perfil') return json(lerPerfil(p.nr));
    if (p.a === 'img') return json(lerImagem(p.id));
    if (p.a === 'admin') return json({ ok: conferirAdmin(p.admin) });
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
    if (b.a === 'pin') return json(conferePin(b.nr, b.pin));
    if (b.a === 'post') return json(novoPost(b));
    if (b.a === 'story') return json(novaHistoria(b));
    if (b.a === 'visto') return json(visto(b));
    if (b.a === 'bio') return json(salvaBio(b));
    if (b.a === 'reagir') return json(reagir(b));
    if (b.a === 'comentar') return json(comentar(b));
    if (b.a === 'apagar') return json(apagar(b));
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

// =====================================================================
//  REDE DO PELOTÃO: dados na planilha "Portal Aloja · mural" e fotos
//  na pasta "Portal Aloja · mídia", criadas sozinhas no seu Drive.
// =====================================================================

function muralSS() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('MURAL_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create('Portal Aloja · mural'); props.setProperty('MURAL_ID', ss.getId()); }
  return ss;
}
function aba(nome) { const ss = muralSS(); return ss.getSheetByName(nome) || ss.insertSheet(nome); }
function linhas(nome) {
  const v = aba(nome).getDataRange().getValues(); const out = [];
  v.forEach((r, i) => { if (r[0] && r[1]) { try { out.push({ row: i + 1, id: String(r[0]), d: JSON.parse(r[1]) }); } catch (e) {} } });
  return out;
}
function grava(nome, id, d) {
  const sh = aba(nome); const x = linhas(nome).find(r => r.id === id); const val = JSON.stringify(d);
  if (x) sh.getRange(x.row, 2).setValue(val); else sh.appendRow([id, val]);
}
function remove(nome, id) { const x = linhas(nome).find(r => r.id === id); if (x) aba(nome).deleteRow(x.row); }
function pasta() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('MIDIA_ID'), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); } catch (e) { f = null; } }
  if (!f) { f = DriveApp.createFolder('Portal Aloja · mídia'); props.setProperty('MIDIA_ID', f.getId()); }
  return f;
}
function salvaImagem(b64, nome) {
  if (!b64) return null;
  b64 = String(b64).replace(/^data:image\/\w+;base64,/, '');
  if (!/^[A-Za-z0-9+/=]+$/.test(b64) || b64.length > 1400000) throw new Error('imagem inválida ou grande demais');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', nome + '.jpg');
  return pasta().createFile(blob).getId();
}
function apagaImagem(id) { if (!id) return; try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
function lerImagem(id) {
  if (!/^[\w-]{10,}$/.test(String(id || ''))) throw new Error('imagem inválida');
  const f = DriveApp.getFileById(id);
  const pais = f.getParents(); let ok = false; const p = pasta().getId();
  while (pais.hasNext()) if (pais.next().getId() === p) ok = true;
  if (!ok) throw new Error('imagem não pertence ao mural');
  return { data: 'data:image/jpeg;base64,' + Utilities.base64Encode(f.getBlob().getBytes()) };
}
function novoId(pfx) { return pfx + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36); }
function txt(t, max) { return String(t || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max); }

// ---- PIN individual: criado no primeiro uso, conferido depois ----
function hashPin(nr, pin) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'aloja|' + nr + '|' + pin)); }
function conferePin(nr, pin) {
  nr = String(nr || ''); pin = String(pin || '');
  if (!/^1\d{3}$/.test(nr)) throw new Error('número inválido');
  if (!/^\d{4,6}$/.test(pin)) return { ok: false, erro: 'O PIN deve ter de 4 a 6 números.' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const x = linhas('pins').find(r => r.id === nr);
    if (!x) { grava('pins', nr, { h: hashPin(nr, pin), t: Date.now() }); return { ok: true, novo: true }; }
    return x.d.h === hashPin(nr, pin) ? { ok: true } : { ok: false, erro: 'PIN incorreto.' };
  } finally { lock.releaseLock(); }
}
function exigePin(b) { const r = conferePin(b.nr, b.pin); if (!r.ok) throw new Error(r.erro || 'PIN incorreto.'); }

function novoPost(b) {
  exigePin(b);
  const texto = txt(b.texto, 1000);
  const tipo = b.tipo === 'treino' ? 'treino' : 'foto';
  let treino = null;
  if (tipo === 'treino') {
    const tr = b.treino || {};
    const mods = ['Corrida', 'Natação', 'Musculação', 'TAF', 'Bicicleta', 'Caminhada', 'Funcional', 'Outro'];
    const dist = Math.max(0, Math.min(500, parseFloat(String(tr.dist || '0').replace(',', '.')) || 0));
    const min = Math.max(0, Math.min(1440, parseInt(tr.min, 10) || 0));
    treino = { mod: mods.indexOf(tr.mod) >= 0 ? tr.mod : 'Outro', dist: Math.round(dist * 100) / 100, min: min };
    if (!treino.dist && !treino.min) throw new Error('Informe a distância ou o tempo do treino.');
  }
  if (!texto && !b.img && !treino) throw new Error('Post vazio.');
  const id = novoId('p');
  const d = { id: id, t: Date.now(), nr: String(b.nr), aloja: txt(b.aloja, 3), tipo: tipo, treino: treino, texto: texto,
              img: salvaImagem(b.img, id), th: salvaImagem(b.th, id + '_th'), reac: {}, com: [] };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { grava('posts', id, d); } finally { lock.releaseLock(); }
  return { ok: true, post: d };
}
function editaPost(id, fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { const x = linhas('posts').find(r => r.id === id); if (!x) throw new Error('Post não encontrado.'); fn(x.d); grava('posts', id, x.d); return x.d; }
  finally { lock.releaseLock(); }
}
function reagir(b) {
  exigePin(b); const emo = ['👍', '🔥', '😂', '💪'].indexOf(b.emo) >= 0 ? b.emo : null; if (!emo) throw new Error('Reação inválida.');
  const nr = String(b.nr);
  return { ok: true, post: editaPost(String(b.id), d => { d.reac = d.reac || {}; const l = d.reac[emo] || []; const i = l.indexOf(nr); if (i >= 0) l.splice(i, 1); else l.push(nr); d.reac[emo] = l; }) };
}
function comentar(b) {
  exigePin(b); const t = txt(b.texto, 300); if (!t) throw new Error('Comentário vazio.');
  return { ok: true, post: editaPost(String(b.id), d => { d.com = (d.com || []).concat([{ id: novoId('c'), t: Date.now(), nr: String(b.nr), texto: t }]).slice(-80); }) };
}
function apagar(b) {
  const adm = String(b.admin || '') === ADMIN;
  if (!adm) exigePin(b);
  const tipo = b.tipo === 'story' ? 'stories' : 'posts'; const id = String(b.id || '');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const x = linhas(tipo).find(r => r.id === id); if (!x) return { ok: true };
    if (b.cid && tipo === 'posts') {          // apagar um comentário
      x.d.com = (x.d.com || []).filter(c => !(c.id === b.cid && (adm || c.nr === String(b.nr))));
      grava('posts', id, x.d); return { ok: true, post: x.d };
    }
    if (!adm && x.d.nr !== String(b.nr)) throw new Error('Só o autor ou a moderação podem apagar.');
    apagaImagem(x.d.img); apagaImagem(x.d.th); aba(tipo).deleteRow(x.row);
  } finally { lock.releaseLock(); }
  return { ok: true };
}
function conferirAdmin(senha) { return String(senha || '') === ADMIN; }

// =====================================================================
//  REDE DO PELOTÃO: perfis, feed e treinos
// =====================================================================
function perfis() { const o = {}; linhas('perfis').forEach(r => { o[r.id] = r.d; }); return o; }
function salvaBio(b) {
  exigePin(b);
  const bio = txt(b.bio, 160);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { grava('perfis', String(b.nr), { bio: bio, t: Date.now() }); } finally { lock.releaseLock(); }
  return { ok: true, bio: bio };
}
function mesAtual() {
  const z = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM');
  return z;
}
function mesDe(t) { return Utilities.formatDate(new Date(t), 'America/Sao_Paulo', 'yyyy-MM'); }
function rankingTreinos(todos) {
  const m = mesAtual(); const r = {};
  todos.forEach(p => { if (p.tipo === 'treino' && p.treino && mesDe(p.t) === m) {
    const x = r[p.nr] || (r[p.nr] = { nr: p.nr, n: 0, km: 0, min: 0 }); x.n++; x.km += p.treino.dist || 0; x.min += p.treino.min || 0; } });
  return Object.keys(r).map(k => r[k]).map(x => { x.km = Math.round(x.km * 10) / 10; return x; })
    .sort((a, b) => b.km - a.km || b.n - a.n || b.min - a.min).slice(0, 10);
}
function lerFeed(antes) {
  const todos = linhas('posts').map(r => r.d).sort((a, b) => b.t - a.t);
  const lim = parseInt(antes, 10) || Infinity;
  const pag = todos.filter(p => p.t < lim).slice(0, 30);
  return { posts: pag, mais: todos.filter(p => p.t < lim).length > 30, perfis: perfis(), ranking: rankingTreinos(todos), mes: mesAtual(), agora: Date.now() };
}
function lerPerfil(nr) {
  nr = String(nr || ''); if (!/^1\d{3}$/.test(nr)) throw new Error('número inválido');
  const todos = linhas('posts').map(r => r.d);
  const meus = todos.filter(p => p.nr === nr).sort((a, b) => b.t - a.t);
  const m = mesAtual(); const tm = meus.filter(p => p.tipo === 'treino' && p.treino && mesDe(p.t) === m);
  const pf = perfis()[nr] || {};
  return { nr: nr, bio: pf.bio || '', posts: meus.slice(0, 120), total: meus.length,
           treinosMes: tm.length, kmMes: Math.round(tm.reduce((a, p) => a + (p.treino.dist || 0), 0) * 10) / 10,
           minMes: tm.reduce((a, p) => a + (p.treino.min || 0), 0) };
}

// =====================================================================
//  HISTÓRIAS DE 24 HORAS
// =====================================================================
const DIA_MS = 24 * 3600 * 1000;
function lerStories() {
  const agora = Date.now(); const vivas = [];
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try { linhas('stories').slice().reverse().forEach(r => { if (agora - r.d.t > DIA_MS) { apagaImagem(r.d.img); aba('stories').deleteRow(r.row); } else vivas.push(r.d); }); }
    finally { lock.releaseLock(); }
  } else linhas('stories').forEach(r => { if (agora - r.d.t <= DIA_MS) vivas.push(r.d); });
  return { stories: vivas.sort((a, b) => a.t - b.t), agora: agora };
}
function novaHistoria(b) {
  exigePin(b); if (!b.img) throw new Error('A história precisa de uma foto.');
  const id = novoId('s');
  const d = { id: id, t: Date.now(), nr: String(b.nr), aloja: txt(b.aloja, 3), leg: txt(b.leg, 140), img: salvaImagem(b.img, id), vis: [] };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { grava('stories', id, d); } finally { lock.releaseLock(); }
  return { ok: true, story: d };
}
function visto(b) {
  const nr = String(b.nr || ''); if (!/^1\d{3}$/.test(nr)) return { ok: false };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { const x = linhas('stories').find(r => r.id === String(b.id)); if (x && x.d.nr !== nr && (x.d.vis || []).indexOf(nr) < 0) { x.d.vis = (x.d.vis || []).concat([nr]); grava('stories', x.id, x.d); } }
  finally { lock.releaseLock(); }
  return { ok: true };
}
