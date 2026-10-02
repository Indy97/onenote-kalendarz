/* Kalendarz zadań: wtyczka OneNote (przeglądarka).
 * Zadania są zapisywane w Microsoft To Do (lista CFG.todoListName) przez Microsoft Graph.
 * Przypomnienia mailowe i w Teams wysyła osobny przepływ Power Automate (README.md, krok 4).
 * Tryb demo (?demo=1 albo brak clientId) zapisuje zadania tylko w localStorage przeglądarki.
 */
(function () {
  "use strict";

  const CFG = window.KALENDARZ_CONFIG;
  const DEMO = new URLSearchParams(location.search).has("demo") || CFG.clientId.startsWith("WSTAW");
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const AUTHORITY = "https://login.microsoftonline.com/" + CFG.tenantId;

  let inOneNote = false;
  let backend = null;
  let tasks = [];
  let view = startOfMonth(new Date());
  let selected = dateKey(new Date());
  let editingId = null;

  const $ = (id) => document.getElementById(id);
  const fmtMonth = new Intl.DateTimeFormat("pl-PL", { month: "long", year: "numeric" });
  const fmtDay = new Intl.DateTimeFormat("pl-PL", { weekday: "long", day: "numeric", month: "long" });

  // ---------- daty ----------
  function pad(n) { return String(n).padStart(2, "0"); }
  function dateKey(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function parseKey(k) { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); }
  function startOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }

  // ---------- komunikaty ----------
  function showBanner(msg, kind) {
    const b = $("banner");
    b.textContent = msg;
    b.className = "banner" + (kind === "info" ? " info" : "");
    b.hidden = false;
  }
  function hideBanner() { $("banner").hidden = true; }

  // ---------- backend: tryb demo ----------
  function DemoBackend() {
    const KEY = "kalendarz-demo";
    const load = () => { try { return JSON.parse(localStorage.getItem(KEY)) || []; } catch (e) { return []; } };
    const save = (list) => { try { localStorage.setItem(KEY, JSON.stringify(list)); } catch (e) { /* brak storage */ } };
    return {
      label: "Tryb demo (dane tylko w tej przeglądarce)",
      async start() { return true; },
      async login() { return true; },
      async list() { return load(); },
      async create(t) { const list = load(); const n = Object.assign({}, t, { id: "d" + Date.now() }); list.push(n); save(list); return n; },
      async update(id, t) { const list = load().map((x) => (x.id === id ? Object.assign({}, x, t) : x)); save(list); return list.find((x) => x.id === id); },
      async remove(id) { save(load().filter((x) => x.id !== id)); }
    };
  }

  // ---------- backend: Microsoft Graph / To Do ----------
  function GraphBackend() {
    const pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: AUTHORITY, redirectUri: new URL("redirect.html", location.href).href },
      cache: { cacheLocation: "localStorage" }
    });
    let account = null;
    let dialogToken = null; // token z okna dialogowego Office (gdy popup MSAL nie działa)
    let listId = null;

    async function tokenFromDialog() {
      return new Promise((resolve, reject) => {
        const url = new URL("dialog.html", location.href).href;
        Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error(res.error.message)); return; }
          const dlg = res.value;
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            dlg.close();
            const msg = JSON.parse(arg.message);
            if (!msg.ok) { reject(new Error(msg.error)); return; }
            dialogToken = { token: msg.token, exp: msg.expiresOn };
            resolve(msg);
          });
          dlg.addEventHandler(Office.EventType.DialogEventReceived, () => reject(new Error("Okno logowania zostało zamknięte.")));
        });
      });
    }

    async function getToken(interactive) {
      if (dialogToken && dialogToken.exp - Date.now() > 120000) return dialogToken.token;
      if (account) {
        try {
          return (await pca.acquireTokenSilent({ scopes: CFG.scopes, account })).accessToken;
        } catch (e) { /* potrzebna interakcja */ }
      }
      if (!interactive) throw new Error("login_required");
      try {
        const res = await pca.acquireTokenPopup({ scopes: CFG.scopes, account: account || undefined, prompt: account ? undefined : "select_account" });
        account = res.account;
        return res.accessToken;
      } catch (e) {
        const popupProblem = /popup|block|window/i.test(e.errorCode || e.message || "");
        if (inOneNote && popupProblem) {
          const msg = await tokenFromDialog();
          account = { username: msg.username };
          return msg.token;
        }
        throw e;
      }
    }

    async function graph(method, path, body, interactive) {
      const token = await getToken(interactive);
      const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
        method,
        headers: {
          Authorization: "Bearer " + token,
          "Content-Type": "application/json",
          Prefer: 'outlook.timezone="' + CFG.timeZone + '"'
        },
        body: body ? JSON.stringify(body) : undefined
      });
      if (res.status === 204) return null;
      const data = await res.json().catch(() => null);
      if (!res.ok) throw new Error("Graph " + res.status + ": " + (data && data.error ? data.error.message : res.statusText));
      return data;
    }

    async function ensureList() {
      if (listId) return listId;
      let url = "/me/todo/lists?$top=100";
      while (url) {
        const page = await graph("GET", url);
        const hit = page.value.find((l) => l.displayName === CFG.todoListName);
        if (hit) { listId = hit.id; return listId; }
        url = page["@odata.nextLink"];
      }
      listId = (await graph("POST", "/me/todo/lists", { displayName: CFG.todoListName })).id;
      return listId;
    }

    function stripHtml(s) {
      const d = document.createElement("div");
      d.innerHTML = s;
      return (d.textContent || "").trim();
    }

    function fromGraph(g) {
      if (!g.dueDateTime) return null; // zadania bez daty (dodane gdzie indziej) nie trafiają do kalendarza
      const notes = g.body && g.body.content ? (g.body.contentType === "html" ? stripHtml(g.body.content) : g.body.content) : "";
      return {
        id: g.id,
        title: g.title,
        notes,
        date: g.dueDateTime.dateTime.slice(0, 10),
        remind: !!g.isReminderOn,
        time: g.reminderDateTime ? g.reminderDateTime.dateTime.slice(11, 16) : CFG.defaultReminderTime,
        completed: g.status === "completed"
      };
    }

    function toGraph(t) {
      const g = {
        title: t.title,
        body: { content: t.notes || "", contentType: "text" },
        dueDateTime: { dateTime: t.date + "T00:00:00", timeZone: CFG.timeZone },
        isReminderOn: !!t.remind,
        status: t.completed ? "completed" : "notStarted"
      };
      if (t.remind) g.reminderDateTime = { dateTime: t.date + "T" + t.time + ":00", timeZone: CFG.timeZone };
      return g;
    }

    const base = () => "/me/todo/lists/" + listId + "/tasks";

    return {
      get label() { return account ? account.username : ""; },
      async start() {
        await pca.initialize();
        account = pca.getAllAccounts()[0] || null;
        if (!account) return false;
        try { await getToken(false); return true; } catch (e) { return false; }
      },
      async login() { await getToken(true); return true; },
      async list() {
        await ensureList();
        const out = [];
        let url = base() + "?$top=100";
        while (url) {
          const page = await graph("GET", url);
          page.value.forEach((g) => { const t = fromGraph(g); if (t) out.push(t); });
          url = page["@odata.nextLink"];
        }
        return out;
      },
      async create(t) { await ensureList(); return fromGraph(await graph("POST", base(), toGraph(t))); },
      async update(id, t) { await ensureList(); return fromGraph(await graph("PATCH", base() + "/" + id, toGraph(t))); },
      async remove(id) { await ensureList(); await graph("DELETE", base() + "/" + id); }
    };
  }

  // ---------- render ----------
  function tasksOn(key) {
    return tasks.filter((t) => t.date === key).sort((a, b) =>
      (a.completed - b.completed) || ((a.remind ? a.time : "99") < (b.remind ? b.time : "99") ? -1 : 1));
  }

  function renderCalendar() {
    $("month-label").textContent = fmtMonth.format(view);
    const grid = $("cal-days");
    grid.innerHTML = "";
    const first = new Date(view);
    first.setDate(1 - ((view.getDay() + 6) % 7)); // poniedziałek przed 1. dniem miesiąca
    const todayKey = dateKey(new Date());
    for (let i = 0; i < 42; i++) {
      const d = new Date(first.getFullYear(), first.getMonth(), first.getDate() + i);
      const key = dateKey(d);
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = d.getDate();
      if (d.getMonth() !== view.getMonth()) btn.classList.add("other");
      if (key === todayKey) btn.classList.add("today");
      if (key === selected) btn.classList.add("selected");
      const day = tasks.filter((t) => t.date === key);
      if (day.length) {
        const dot = document.createElement("span");
        dot.className = "dot" + (day.every((t) => t.completed) ? " done" : "");
        btn.appendChild(dot);
        btn.title = day.length + " zad.";
      }
      btn.addEventListener("click", () => selectDay(key));
      grid.appendChild(btn);
    }
  }

  function renderDay() {
    $("day-label").textContent = fmtDay.format(parseKey(selected));
    const list = $("task-list");
    list.innerHTML = "";
    const day = tasksOn(selected);
    $("empty-day").hidden = day.length > 0;
    day.forEach((t) => {
      const li = document.createElement("li");
      if (t.completed) li.classList.add("completed");

      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = t.completed;
      cb.title = "Oznacz jako wykonane";
      cb.addEventListener("change", () => save(t.id, Object.assign({}, t, { completed: cb.checked })));

      const body = document.createElement("div");
      body.className = "body";
      body.title = "Kliknij, aby edytować";
      const title = document.createElement("div");
      title.className = "title";
      title.textContent = t.title;
      body.appendChild(title);
      if (t.remind) {
        const meta = document.createElement("div");
        meta.className = "meta";
        meta.textContent = "🔔 " + t.time;
        body.appendChild(meta);
      }
      if (t.notes) {
        const notes = document.createElement("div");
        notes.className = "notes";
        notes.textContent = t.notes;
        body.appendChild(notes);
      }
      body.addEventListener("click", () => startEdit(t));

      const del = document.createElement("button");
      del.className = "del";
      del.type = "button";
      del.textContent = "×";
      del.title = "Usuń zadanie";
      del.addEventListener("click", () => removeTask(t));

      li.append(cb, body, del);
      list.appendChild(li);
    });
    $("insert-btn").disabled = !inOneNote || day.length === 0;
    $("insert-btn").title = inOneNote ? "" : "Dostępne po otwarciu wtyczki w OneNote";
  }

  function render() { renderCalendar(); renderDay(); }

  function selectDay(key) {
    selected = key;
    const d = parseKey(key);
    if (d.getMonth() !== view.getMonth() || d.getFullYear() !== view.getFullYear()) view = startOfMonth(d);
    resetForm();
    render();
  }

  // ---------- formularz ----------
  function resetForm() {
    editingId = null;
    $("task-form").reset();
    $("f-remind").checked = true;
    $("f-time").value = CFG.defaultReminderTime;
    $("f-submit").textContent = "Dodaj";
    $("f-cancel").hidden = true;
  }

  function startEdit(t) {
    editingId = t.id;
    $("f-title").value = t.title;
    $("f-notes").value = t.notes || "";
    $("f-remind").checked = t.remind;
    $("f-time").value = t.time;
    $("f-submit").textContent = "Zapisz";
    $("f-cancel").hidden = false;
    $("f-title").focus();
  }

  async function withBusy(fn) {
    document.body.style.cursor = "progress";
    try {
      if (!$("banner").classList.contains("info")) hideBanner(); // chowamy tylko stare błędy
      await fn();
    } catch (e) {
      console.error(e);
      showBanner("Błąd: " + (e.message || e));
    } finally { document.body.style.cursor = ""; }
  }

  async function save(id, t) {
    await withBusy(async () => {
      const saved = id ? await backend.update(id, t) : await backend.create(t);
      tasks = tasks.filter((x) => x.id !== id);
      if (saved) tasks.push(saved);
      render();
    });
  }

  async function removeTask(t) {
    await withBusy(async () => {
      await backend.remove(t.id);
      tasks = tasks.filter((x) => x.id !== t.id);
      if (editingId === t.id) resetForm();
      render();
    });
  }

  async function onSubmit(ev) {
    ev.preventDefault();
    const existing = editingId ? tasks.find((x) => x.id === editingId) : null;
    const t = {
      title: $("f-title").value.trim(),
      notes: $("f-notes").value.trim(),
      date: existing ? existing.date : selected,
      remind: $("f-remind").checked,
      time: $("f-time").value || CFG.defaultReminderTime,
      completed: existing ? existing.completed : false
    };
    if (!t.title) return;
    const id = editingId;
    resetForm();
    await save(id, t);
  }

  // ---------- wstawianie do OneNote ----------
  function esc(s) { return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

  async function insertIntoPage() {
    const day = tasksOn(selected);
    const items = day.map((t) =>
      "<li>" + (t.completed ? "☑ " : "☐ ") + esc(t.title) + (t.remind ? " <i>(" + t.time + ")</i>" : "") + "</li>").join("");
    const html = "<p><b>Zadania: " + esc(fmtDay.format(parseKey(selected))) + "</b></p><ul>" + items + "</ul>";
    await withBusy(() => OneNote.run(async (ctx) => {
      ctx.application.getActivePage().addOutline(40, 90, html);
      await ctx.sync();
      showBanner("Wstawiono listę zadań do bieżącej strony.", "info");
    }));
  }

  // ---------- start ----------
  async function loadAndShow() {
    $("login-view").hidden = true;
    $("app-view").hidden = false;
    $("account").textContent = backend.label;
    await withBusy(async () => { tasks = await backend.list(); });
    render();
  }

  async function start() {
    backend = DEMO ? DemoBackend() : GraphBackend();
    if (DEMO) showBanner("Tryb demo: zadania zapisują się tylko w tej przeglądarce. Uzupełnij clientId w config.js.", "info");

    $("prev-month").addEventListener("click", () => { view = new Date(view.getFullYear(), view.getMonth() - 1, 1); renderCalendar(); });
    $("next-month").addEventListener("click", () => { view = new Date(view.getFullYear(), view.getMonth() + 1, 1); renderCalendar(); });
    $("today-btn").addEventListener("click", () => selectDay(dateKey(new Date())));
    $("task-form").addEventListener("submit", onSubmit);
    $("f-cancel").addEventListener("click", resetForm);
    $("f-remind").addEventListener("change", () => { $("f-time").disabled = !$("f-remind").checked; });
    $("insert-btn").addEventListener("click", insertIntoPage);
    $("login-btn").addEventListener("click", () => withBusy(async () => { await backend.login(); await loadAndShow(); }));
    resetForm();

    let ok = false;
    try { ok = await backend.start(); } catch (e) { showBanner("Błąd inicjalizacji: " + e.message); }
    if (ok) await loadAndShow();
    else $("login-view").hidden = false;
  }

  let started = false;
  function startOnce() { if (!started) { started = true; start(); } }

  if (window.Office && Office.onReady) {
    Office.onReady((info) => {
      inOneNote = info.host === Office.HostType.OneNote;
      if (started && !$("app-view").hidden) renderDay(); // onReady przyszło po timeoucie
      startOnce();
    });
    setTimeout(startOnce, 3000); // office.js poza Office czasem nie wywołuje onReady
  } else {
    startOnce();
  }
})();
