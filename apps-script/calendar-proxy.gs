/**
 * Calendar Proxy — devolve eventos dos calendários de suporte em JSON
 * para os cards "Suporte Técnico" e "Suporte 24h" do index.html.
 *
 * INSTALAÇÃO (uma vez, com uma conta Somengil que tenha acesso aos 2 calendários):
 *   1. https://script.google.com → Novo projeto → colar este ficheiro em Code.gs
 *   2. Implementar → Nova implementação → Tipo: Aplicação Web
 *        Executar como:      Eu (a conta que está a instalar)
 *        Quem tem acesso:    Qualquer pessoa
 *   3. Autorizar o acesso ao Calendário quando pedido
 *   4. Copiar o URL que termina em /exec e colar em CALENDAR_PROXY_URL no index.html
 *
 * Só os calendários listados em CALENDARS podem ser lidos (whitelist).
 * Teste: <URL>/exec?cal=suporte_tecnico
 */

const CALENDARS = {
  suporte_tecnico: 'somengil.com_tqf200bjfjht53mc21r6ss9qsg@group.calendar.google.com',
  suporte_24h: 'c_ada11e3f2a5d6d792515b33416bb5d2ba3dfa5fe561f317b24b363e4529fb5a5@group.calendar.google.com',
};

const MAX_DAYS = 90;

function doGet(e) {
  const params = (e && e.parameter) || {};
  const key = params.cal;
  const calId = CALENDARS[key];
  if (!calId) {
    return json_({ ok: false, error: 'Calendário desconhecido: ' + key });
  }

  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) {
    return json_({ ok: false, error: 'Sem acesso ao calendário: ' + key });
  }

  const days = Math.min(Math.max(parseInt(params.days, 10) || 30, 1), MAX_DAYS);
  const now = new Date();
  const from = new Date(now.getTime() - 24 * 3600 * 1000); // inclui eventos que já começaram
  const to = new Date(now.getTime() + days * 24 * 3600 * 1000);

  const events = cal.getEvents(from, to)
    .filter(ev => ev.getEndTime() > now) // ignora os que já terminaram
    .map(ev => ({
      title: ev.getTitle() || '(sem título)',
      start: ev.getStartTime().toISOString(),
      end: ev.getEndTime().toISOString(),
      allDay: ev.isAllDayEvent(),
    }))
    .sort((a, b) => a.start < b.start ? -1 : a.start > b.start ? 1 : 0);

  return json_({ ok: true, cal: key, generatedAt: now.toISOString(), events: events });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
