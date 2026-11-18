/* Athena Run Club – shared data layer (used by index.html and admin.html)
   - LIVE mode: reads/writes a Google Sheet through a Google Apps Script web app (window.ARC_API)
   - DEMO mode: sample data, admin edits are stored in this browser only (localStorage)
*/
(function () {
  const API = (window.ARC_API || "").trim();
  const LS = "arc-demo-data-v1";

  const pad = n => String(n).padStart(2, "0");
  const iso = d => d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  function nextDow(dow, weeks) { const d = new Date(); d.setHours(12, 0, 0, 0); const diff = (dow - d.getDay() + 7) % 7 || 7; d.setDate(d.getDate() + diff + 7 * (weeks || 0)); return iso(d); }

  function demo() {
    return {
      demo: true,
      settings: {
        city: "",
        meetingPoint: "",
        mapUrl: "",
        whatsapp: "",
        instagram: "athena.runclub",
        whoCanJoin: "",
        whatToBring: ""
      },
      sessions: [],
      updates: []
    };
  }

  function loadLocal() { try { const raw = localStorage.getItem(LS); if (raw) return JSON.parse(raw); } catch (e) {} return demo(); }
  function saveLocal(d) { try { localStorage.setItem(LS, JSON.stringify(d)); } catch (e) {} }

  function normalize(d) {
    d = d || {};
    d.settings = d.settings || {};
    d.sessions = (d.sessions || []).map(s => ({ ...s, cancelled: s.cancelled === true || s.cancelled === "TRUE" || s.cancelled === "true" }));
    d.stories = (d.stories || []).filter(x => x && x.url);
    d.updates = (d.updates || []).map(u => ({ ...u, pinned: u.pinned === true || u.pinned === "TRUE" || u.pinned === "true" }));
    d.sessions.forEach(s => { if (s.date instanceof Date) s.date = iso(s.date); s.date = String(s.date || "").slice(0, 10); s.time = String(s.time || "").slice(0, 5); });
    d.updates.forEach(u => { u.date = String(u.date || "").slice(0, 10); });
    return d;
  }

  async function load() {
    if (API) {
      try {
        const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), 9000);
        const r = await fetch(API + (API.includes("?") ? "&" : "?") + "action=get&t=" + Date.now(), { signal: ctl.signal });
        clearTimeout(t);
        const j = await r.json();
        if (j && j.ok !== false) return normalize({ ...j, demo: false, live: true });
      } catch (e) { console.warn("ARC: live data unavailable", e); }
      return normalize({ settings: {}, sessions: [], updates: [], demo: false, live: false, error: true });
    }
    return normalize(loadLocal());
  }

  async function call(body) {
    if (API) {
      const r = await fetch(API, { method: "POST", body: JSON.stringify(body) }); // text/plain: no CORS preflight
      const j = await r.json();
      if (!j.ok) throw new Error(j.error || "error");
      return normalize({ ...j.data, demo: false, live: true });
    }
    // demo mode
    const d = loadLocal();
    if (body.op === "ping") return normalize(d);
    if (body.op === "save") {
      if (body.type === "settings") d.settings = { ...d.settings, ...body.item };
      else {
        const list = d[body.type]; const i = list.findIndex(x => x.id === body.item.id);
        if (i >= 0) list[i] = body.item; else list.push(body.item);
      }
    }
    if (body.op === "delete") d[body.type] = d[body.type].filter(x => x.id !== body.id);
    if (body.op === "reset") { localStorage.removeItem(LS); return normalize(demo()); }
    saveLocal(d);
    return normalize(d);
  }

  async function instagram(fromApi) {
    const FEED = (window.ARC_IG_FEED || "").trim();
    if (FEED) {
      try {
        const j = await (await fetch(FEED)).json();
        return (j.posts || []).map(p => ({
          permalink: p.permalink, caption: p.prunedCaption || p.caption || "", timestamp: p.timestamp, mediaType: p.mediaType,
          image: (p.sizes && (p.sizes.medium || p.sizes.large || p.sizes.full) || {}).mediaUrl || (p.mediaType === "VIDEO" ? p.thumbnailUrl : p.mediaUrl) || ""
        })).filter(p => p.image);
      } catch (e) { console.warn("ARC: Instagram feed unavailable", e); }
    }
    return (fromApi || []).filter(p => p.image);
  }

  window.ARC = {
    instagram,
    mode: API ? "live" : "demo",
    load, call,
    uid: () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    today: () => iso(new Date())
  };
})();
