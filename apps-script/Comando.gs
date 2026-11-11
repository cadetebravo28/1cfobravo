// =====================================================================
//  ESCALA DE COMANDO E SUPLÊNCIA — aviso na véspera da semana nova
//  Rode uma vez  instalarAvisoComando  (▶ Executar). Não precisa reimplantar.
// =====================================================================
function avisoComando() {
  const amanha = chaveDia(1);
  let d; try { d = JSON.parse(UrlFetchApp.fetch(SITE + 'comando.json?t=' + Date.now()).getContentText()); } catch (e) { return; }
  const s = (d.semanas || []).find(x => x.ini === amanha); if (!s) return;
  const por = {};
  s.cias.forEach(c => c.linhas.forEach(r => {
    if (r.tit && r.tit.nr) (por[r.tit.nr] = por[r.tit.nr] || []).push(r.f);
    if (r.sup && r.sup.nr && !(r.tit && r.tit.nr === r.sup.nr)) (por[r.sup.nr] = por[r.sup.nr] || []).push(r.f + ' (suplente)');
  }));
  const br = k => k.slice(8, 10) + '/' + k.slice(5, 7);
  pushVarios(Object.keys(por).map(nr => pushReq([nr], '⭐ Escala de comando a partir de amanhã',
    por[nr].join(' · ') + '\nSemana ' + br(s.ini) + ' a ' + br(s.fim), SITE + '#painel')));
}
function instalarAvisoComando() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'avisoComando') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('avisoComando').timeBased().everyDays(1).atHour(19).nearMinute(30).inTimezone('America/Sao_Paulo').create();
  Logger.log('Pronto: aviso da escala de comando às 19h30 da véspera de cada semana nova.');
}
