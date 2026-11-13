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
  marc:    { id: CFO, aba: 'Marcação de aulas',           colunas: [3, 13], desde: 2 },
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
    if (p.a === 'atividades') return json(lerAtividades(p.nr));
    if (p.a === 'post') return json({ post: lerPost(p.id) });
    if (p.a === 'online') return json(lerOnline());
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
    if (b.a === 'destacar') return json(destacar(b));
    if (b.a === 'destaque_apagar') return json(apagaDestaque(b));
    if (b.a === 'reagir') return json(reagir(b));
    if (b.a === 'comentar') return json(comentar(b));
    if (b.a === 'apagar') return json(apagar(b));
    if (b.a === 'ping') return json(presenca(b));
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
  let vals = rg.getDisplayValues();
  if (k === 'marc') vals = vals.map(r => r.map(c => /^(FALSE|FALSO)$/i.test(c) ? '' : c));   // coluna M (VA): só fica o marcado
  const out = { values: aparar(vals), at: new Date().toISOString() };
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
  avisaPelotao('post', b.nr, d);
  return { ok: true, post: d };
}
function editaPost(id, fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { const x = linhas('posts').find(r => r.id === id); if (!x) throw new Error('Post não encontrado.'); fn(x.d); grava('posts', id, x.d); return x.d; }
  finally { lock.releaseLock(); }
}
function reagir(b) {
  exigePin(b); const emo = ['❤️', '👍', '🔥', '😂', '💪'].indexOf(b.emo) >= 0 ? b.emo : null; if (!emo) throw new Error('Reação inválida.');
  const nr = String(b.nr); let novo = false;
  const p = editaPost(String(b.id), d => { d.reac = d.reac || {}; const l = d.reac[emo] || []; const i = l.indexOf(nr); d.rt = d.rt || {}; if (i >= 0) { l.splice(i, 1); delete d.rt[nr + emo]; } else { l.push(nr); d.rt[nr + emo] = Date.now(); novo = true; } d.reac[emo] = l; });
  if (novo) avisaPelotao('reac', nr, p, emo);
  return { ok: true, post: p };
}
function comentar(b) {
  exigePin(b); const t = txt(b.texto, 300); if (!t) throw new Error('Comentário vazio.');
  const p = editaPost(String(b.id), d => { d.com = (d.com || []).concat([{ id: novoId('c'), t: Date.now(), nr: String(b.nr), texto: t }]).slice(-80); });
  avisaPelotao('com', b.nr, p, t);
  return { ok: true, post: p };
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
    if (!x.d.dest) apagaImagem(x.d.img); apagaImagem(x.d.th); aba(tipo).deleteRow(x.row);
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
  const dx = linhas('destaques').find(r => r.id === nr);
  return { nr: nr, bio: pf.bio || '', destaques: dx ? dx.d.grupos || [] : [], posts: meus.slice(0, 120), total: meus.length,
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
    try { linhas('stories').slice().reverse().forEach(r => { if (agora - r.d.t > DIA_MS) { if (!r.d.dest) apagaImagem(r.d.img); aba('stories').deleteRow(r.row); } else vivas.push(r.d); }); }
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

// ---- Destaques do perfil (histórias que ficam) ----
function destacar(b) {
  exigePin(b); const nr = String(b.nr); const nome = txt(b.nome, 24) || 'Destaques';
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const st = linhas('stories').find(r => r.id === String(b.id));
    if (!st) throw new Error('História não encontrada (talvez já tenha expirado).');
    if (st.d.nr !== nr) throw new Error('Você só pode destacar as suas histórias.');
    st.d.dest = true; grava('stories', st.id, st.d);
    const x = linhas('destaques').find(r => r.id === nr); const g = x ? (x.d.grupos || []) : [];
    let grp = g.find(z => z.nome.toLowerCase() === nome.toLowerCase()); if (!grp) { grp = { nome: nome, itens: [] }; g.push(grp); }
    if (!grp.itens.some(i => i.id === st.id)) grp.itens.push({ id: st.id, img: st.d.img, leg: st.d.leg || '', t: st.d.t });
    grp.itens = grp.itens.slice(-30);
    grava('destaques', nr, { grupos: g.slice(0, 12) });
    return { ok: true, grupos: g };
  } finally { lock.releaseLock(); }
}
function apagaDestaque(b) {
  exigePin(b); const nr = String(b.nr);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const x = linhas('destaques').find(r => r.id === nr); if (!x) return { ok: true, grupos: [] };
    let g = x.d.grupos || []; const grp = g.find(z => z.nome === b.nome); if (!grp) return { ok: true, grupos: g };
    const sai = b.item ? grp.itens.filter(i => i.id === b.item) : grp.itens;
    grp.itens = b.item ? grp.itens.filter(i => i.id !== b.item) : [];
    sai.forEach(i => { if (!linhas('stories').some(r => r.id === i.id)) apagaImagem(i.img); });
    g = g.filter(z => z.itens.length);
    grava('destaques', nr, { grupos: g });
    return { ok: true, grupos: g };
  } finally { lock.releaseLock(); }
}

// ---- Atividades: curtidas e comentários nos seus posts ----
function lerAtividades(nr) {
  nr = String(nr || ''); if (!/^1\d{3}$/.test(nr)) throw new Error('número inválido');
  const ev = [];
  linhas('posts').forEach(r => { const p = r.d; if (p.nr !== nr) return;
    const rt = p.rt || {};
    Object.keys(p.reac || {}).forEach(emo => (p.reac[emo] || []).forEach(q => { if (q !== nr) ev.push({ tipo: 'reac', t: rt[q + emo] || p.t, de: q, emo: emo, post: p.id, th: p.th || p.img || null }); }));
    (p.com || []).forEach(c => { if (c.nr !== nr) ev.push({ tipo: 'com', t: c.t, de: c.nr, texto: c.texto, post: p.id, th: p.th || p.img || null }); });
  });
  linhas('posts').forEach(r => { const p = r.d; if (p.nr === nr) return;   // menções com @nome em comentários
    if (p.texto && p.texto.indexOf('@' + nr) >= 0) ev.push({ tipo: 'men', t: p.t, de: p.nr, texto: p.texto, post: p.id, th: p.th || p.img || null });
    (p.com || []).forEach(c => { if (c.nr !== nr && c.texto && c.texto.indexOf('@' + nr) >= 0) ev.push({ tipo: 'men', t: c.t, de: c.nr, texto: c.texto, post: p.id, th: p.th || p.img || null }); });
  });
  return { atividades: ev.sort((a, b) => b.t - a.t).slice(0, 60) };
}
function lerPost(id) { const x = linhas('posts').find(r => r.id === String(id)); return x ? x.d : null; }

// =====================================================================
//  NOTIFICAÇÕES NO CELULAR (OneSignal) — cada militar recebe as suas
//  1) Rode uma vez a função  instalarNotificacoes  (menu ▶ Executar) e autorize.
//  2) Para testar só no seu celular: rode  testeNotificacao .
// =====================================================================
const OS_APP = 'cb58e6e8-bfd1-4dab-a056-2f589d36b902';
const OS_KEY = 'COLE_AQUI_A_REST_API_KEY_DO_ONESIGNAL';
const SITE = 'https://cadetebravo28.github.io/1cfobravo/';
const TESTE_NR = '1023';
const MIL = {'1012':['Camila',1],'1039':['Rebecca',1],'1022':['Fernanda',1],'1050':['Ana Paula',1],'1033':['Paola',1],'1044':['Semtchuk',1],'1042':['Martini',1],'1040':['Kikuchi',7],'1052':['Moreira',7],'1053':['Henrique Cruz',7],'1027':['Parreiral',7],'1002':['Alef',7],'1006':['Bonfim',7],'1021':['Gallego',7],'1013':['Zelenski',7],'1023':['Retexin',8],'1047':['Chepiloski',8],'1016':['Douglas',8],'1025':['Gerlach',8],'1004':['Rocha',8],'1018':['Spak',8],'1011':['Stelmack',8],'1009':['André',8],'1036':['Machado',9],'1032':['Schumacher',9],'1038':['Mezzadri',9],'1045':['Sluminski',9],'1046':['Ronqui',9],'1031':['Gorri',9],'1014':['Hartemann',9],'1049':['Keven',9]};   // nº: [nome de guerra, alojamento]
const FAX8 = {"2026-11-03":"Retexin","2026-11-04":"Chepiloski","2026-11-05":"Gerlach","2026-11-06":"Spak","2026-11-09":"Retexin","2026-11-10":"Chepiloski","2026-11-11":"Gerlach","2026-11-12":"Spak","2026-11-13":"Douglas","2026-11-16":"Retexin","2026-11-17":"Chepiloski","2026-11-18":"Gerlach","2026-11-19":"Spak","2026-11-23":"Douglas","2026-11-24":"Stelmack","2026-11-25":"Retexin","2026-11-26":"Chepiloski","2026-11-27":"Gerlach","2026-11-30":"Spak"};
const FAX1 = {"2026-11-03": "Semtchuk", "2026-11-04": "Martini", "2026-11-05": "Semtchuk", "2026-11-06": "Martini", "2026-11-09": "Camila", "2026-11-10": "Rebecca", "2026-11-11": "Fernanda", "2026-11-12": "Ana Paula", "2026-11-13": "Paola", "2026-11-16": "Fernanda", "2026-11-17": "Ana Paula", "2026-11-18": "Camila", "2026-11-19": "Paola", "2026-11-23": "Martini", "2026-11-24": "Camila", "2026-11-25": "Semtchuk", "2026-11-26": "Martini", "2026-11-27": "Fernanda", "2026-11-30": "Paola"};
const JAQUETA = 15;   // mínima abaixo disso: jaqueta

function pushReq(nrs, titulo, texto, url) {
  nrs = (nrs || []).map(String).filter(n => MIL[n]); if (!nrs.length) return null;
  return { url: 'https://api.onesignal.com/notifications?c=push', method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Key ' + OS_KEY },
    payload: JSON.stringify({ app_id: OS_APP, target_channel: 'push', include_aliases: { external_id: nrs },
      headings: { en: titulo, pt: titulo }, contents: { en: texto, pt: texto }, url: url || SITE, chrome_web_icon: SITE + 'icon-192.png', ttl: 6 * 3600 }) };
}
function push(nrs, titulo, texto, url) { const r = pushReq(nrs, titulo, texto, url); if (!r) return; try { UrlFetchApp.fetch(r.url, r); } catch (e) {} }
function pushVarios(lista) { const rs = lista.filter(Boolean); for (let i = 0; i < rs.length; i += 20) { try { UrlFetchApp.fetchAll(rs.slice(i, i + 20)); } catch (e) {} } }
const nomeMil = nr => (MIL[String(nr)] || ['Alguém'])[0];

function chaveDia(desloc) { const d = new Date(Date.now() + (desloc || 0) * 864e5); return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd'); }
function rotuloDia(k) { const [y, m, d] = k.split('-').map(Number); const dt = new Date(Date.UTC(y, m - 1, d, 15));
  return ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'][dt.getUTCDay()] + ', ' + String(d).padStart(2, '0') + '/' + String(m).padStart(2, '0'); }
const p2 = n => String(n).padStart(2, '0');

// Escala de serviço: { nº: 'GBM' } no dia
function servicosDoDia(k) {
  const v = lerFonte('escala').values || []; const hdr = v[0] || []; let y = 2026, pm = 0, col = -1;
  for (let c = 3; c < hdr.length; c++) {
    const m = String(hdr[c] || '').match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/); if (!m) continue;
    const mo = +m[2]; if (m[3]) y = m[3].length === 2 ? 2000 + +m[3] : +m[3]; else if (pm && mo < pm) y++; pm = mo;
    if (y + '-' + p2(mo) + '-' + p2(+m[1]) === k) { col = c; break; }
  }
  const out = {}; if (col < 0) return out;
  for (let r = 1; r < v.length; r++) { const nr = String(parseInt(v[r][0], 10) || ''); if (!MIL[nr]) continue;
    const c = String(v[r][col] || '').trim().toUpperCase(); if (c && c !== '-' && c !== '–' && c !== 'B') out[nr] = c; }
  return out;
}
function nomeServico(c) { const chefe = /^C(?=GBM|SOB)/.test(c); const b = c.replace(/^C(?=GBM|SOB)/, '');
  const n = { GBM: 'Guarnição GBM', GBMV: 'Guarnição Vermelha (GBMV)', SOB: 'Sobreaviso', SOBV: 'Sobreaviso Vermelho' }[b] || c;
  return n + (chefe ? ' · chefe' : ''); }
// Faxina: { nº: true } no dia
function faxinaDoDia(k) {
  const out = {}; const marca = (g, a) => Object.keys(MIL).forEach(n => { if (MIL[n][1] === a && MIL[n][0].toLowerCase() === String(g).toLowerCase()) out[n] = true; });
  if (FAX8[k]) marca(FAX8[k], 8); if (FAX1[k]) marca(FAX1[k], 1);
  try { (lerFonte('fax9').values || []).slice(1).forEach(r => { const d = String(r[0] || '').match(/(\d{1,2})\/(\d{1,2})/); if (!d) return;
    if ('2026-' + p2(+d[2]) + '-' + p2(+d[1]) === k) { marca(String(r[1] || '').trim(), 9); const t = String(r[2] || '').trim(); if (t) marca(t, 9); } }); } catch (e) {}
  return out;
}
// Aulas do dia (mesmas para todos)
function aulasDia(k) {
  let rows = [];
  try { rows = (lerFonte('marc').values || []).map(r => { const d = String(r[0] || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    return d ? { k: d[3] + '-' + p2(+d[2]) + '-' + p2(+d[1]), ini: String(r[1] || ''), sg: String(r[3] || '').trim(), obs: String(r[4] || '').trim() } : null; })
    .filter(r => r && r.k === k && r.sg).sort((a, b) => a.ini.localeCompare(b.ini)); } catch (e) {}
  const prova = rows.find(r => /prova|avalia|\bv\.?a\.?\b|\bvf\b|\bvs\b|exame|\btaf\b|verifica[çc]|recupera|teste f[íi]sico/i.test(r.obs + ' ' + r.sg));
  const sig = []; rows.forEach(r => { if (sig.indexOf(r.sg) < 0) sig.push(r.sg); });
  return { primeira: rows[0] || null, prova: prova || null, sig: sig };
}
function climaDia(k) {
  try { const j = JSON.parse(UrlFetchApp.fetch('https://api.open-meteo.com/v1/forecast?latitude=-25.535&longitude=-49.206&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max&timezone=America%2FSao_Paulo&forecast_days=3').getContentText());
    const i = j.daily.time.indexOf(k); if (i < 0) return null;
    return { min: Math.round(j.daily.temperature_2m_min[i]), max: Math.round(j.daily.temperature_2m_max[i]), chuva: j.daily.precipitation_probability_max[i] }; } catch (e) { return null; }
}

// 20h: o que cada um tem amanhã
function avisoVespera() {
  const k = chaveDia(1); const sv = servicosDoDia(k), fx = faxinaDoDia(k), au = aulasDia(k), cl = climaDia(k);
  const comum = [];
  if (au.primeira) comum.push('📚 1ª aula ' + au.primeira.ini + ' · ' + au.sig.join(', '));
  if (au.prova) comum.push('📝 Avaliação: ' + au.prova.sg + (au.prova.obs ? ' · ' + au.prova.obs : ''));
  if (cl) comum.push(cl.min < JAQUETA ? '🧥 Leve a jaqueta: mínima ' + cl.min + '°' : '👕 Sem jaqueta: ' + cl.min + '° a ' + cl.max + '°' + (cl.chuva >= 60 ? ' · ☔ chuva ' + cl.chuva + '%' : ''));
  const reqs = Object.keys(MIL).map(nr => {
    const l = [];
    if (sv[nr]) l.push('🔥 Serviço: ' + nomeServico(sv[nr]));
    if (fx[nr]) l.push('🧹 Faxina do Aloja ' + MIL[nr][1]);
    const tudo = l.concat(comum); if (!tudo.length) return null;
    const tit = (sv[nr] ? '🔥 Amanhã você está de serviço' : fx[nr] ? '🧹 Amanhã é sua faxina' : 'Amanhã') + ' · ' + rotuloDia(k);
    return pushReq([nr], tit, tudo.join('\n'), SITE + '#dia');
  });
  pushVarios(reqs);
}
// 5h30: lembrete de hoje só para quem tem serviço ou faxina
function avisoManha() {
  const k = chaveDia(0); const sv = servicosDoDia(k), fx = faxinaDoDia(k);
  pushVarios(Object.keys(MIL).map(nr => {
    if (!sv[nr] && !fx[nr]) return null;
    const l = []; if (sv[nr]) l.push('🔥 Serviço hoje: ' + nomeServico(sv[nr])); if (fx[nr]) l.push('🧹 Hoje é sua faxina do Aloja ' + MIL[nr][1]);
    return pushReq([nr], 'Bom dia, Cadete ' + MIL[nr][0] + '!', l.join('\n'), SITE + '#painel');
  }));
}
function instalarNotificacoes() {
  ScriptApp.getProjectTriggers().forEach(t => { if (['avisoVespera', 'avisoManha'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('avisoVespera').timeBased().everyDays(1).atHour(20).nearMinute(0).inTimezone('America/Sao_Paulo').create();
  ScriptApp.newTrigger('avisoManha').timeBased().everyDays(1).atHour(5).nearMinute(30).inTimezone('America/Sao_Paulo').create();
  Logger.log('Pronto: avisos às 20h e às 5h30 todos os dias.');
}
function testeNotificacao() {
  const r = pushReq([TESTE_NR], '🔔 Teste do CFOBRAVO', 'Se chegou, as notificações estão funcionando!', SITE);
  const res = UrlFetchApp.fetch(r.url, r); Logger.log(res.getContentText());
}
function testeVesperaSoEu() {   // mostra no registro o que cada um receberia amanhã e manda o seu
  const k = chaveDia(1); Logger.log(JSON.stringify({ dia: k, servico: servicosDoDia(k), faxina: faxinaDoDia(k), aulas: aulasDia(k), clima: climaDia(k) }));
}

// Pelotão: curtidas, comentários e marcações
function mencoes(t) { const s = {}; String(t || '').replace(/@(1\d{3})\b/g, (m, nr) => { s[nr] = 1; }); return Object.keys(s); }
function avisaPelotao(tipo, de, post, extra) {
  try {
    const url = SITE + '#mural'; const nome = nomeMil(de); const reqs = [];
    if (tipo === 'reac' && post.nr !== String(de)) reqs.push(pushReq([post.nr], 'Pelotão', extra === '❤️' ? '❤️ ' + nome + ' curtiu sua publicação.' : extra + ' ' + nome + ' reagiu à sua publicação.', url));
    if (tipo === 'com') {
      if (post.nr !== String(de)) reqs.push(pushReq([post.nr], 'Pelotão', '💬 ' + nome + ' comentou: ' + String(extra).replace(/@(1\d{3})\b/g, (m, n) => MIL[n] ? '@' + MIL[n][0] : m), url));
      reqs.push(pushReq(mencoes(extra).filter(n => n !== String(de) && n !== post.nr), 'Pelotão', '🏷️ ' + nome + ' marcou você num comentário.', url));
    }
    if (tipo === 'post') reqs.push(pushReq(mencoes(post.texto).filter(n => n !== String(de)), 'Pelotão', '🏷️ ' + nome + ' marcou você numa publicação.', url));
    pushVarios(reqs);
  } catch (e) {}
}

// =====================================================================
//  QUEM ESTÁ ONLINE — só na memória rápida (some sozinho em 6 h)
// =====================================================================
function presenca(b) {
  const nr = String(b.nr || '');
  const c = CacheService.getScriptCache();
  if (MIL[nr]) c.put('pv_' + nr, JSON.stringify({ t: b.sai ? Date.now() - 125000 : Date.now(), inv: !!b.inv }), 21600);
  return lerOnline();
}
function lerOnline() {
  const ks = Object.keys(MIL).map(n => 'pv_' + n);
  const all = CacheService.getScriptCache().getAll(ks); const out = [];
  Object.keys(all).forEach(k => { try { const v = JSON.parse(all[k]); if (!v.inv) out.push({ nr: k.slice(3), t: v.t }); } catch (e) {} });
  return { agora: Date.now(), on: out };
}
