/* ============================================================
   ESP · Painel Professor v4.1-stable
   ------------------------------------------------------------
   - Login Microsoft 365 (redirect flow)
   - OneDrive Graph integração
   - Painel de Dia + Registos rápidos
   - Admin modal com abas
   - XLSX import/export + Backup
   ============================================================ */

const MSAL_CONFIG = {
  auth: {
    clientId: "c5573063-8a04-40d3-92bf-eb229ad4701c",
    authority: "https://login.microsoftonline.com/d650692c-6e73-48b3-af84-e3497ff3e1f1",
    redirectUri: "https://bibliotecaesparedes-hub.github.io/esp-painel-professor-v4.1/"
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
};

const MSAL_SCOPES = { scopes: ["Files.ReadWrite.All", "User.Read", "openid", "profile", "offline_access"] };
const SITE_ID = "esparedes-my.sharepoint.com,540a0485-2578-481e-b4d8-220b41fb5c43,7335dc42-69c8-42d6-8282-151e3783162d";
const CFG_PATH = "/Documents/GestaoAlunos-OneDrive/config_especial.json";
const REG_PATH = "/Documents/GestaoAlunos-OneDrive/2registos_alunos.json";
const BACKUP_FOLDER = "/Documents/GestaoAlunos-OneDrive/backup";

let msalApp, account, accessToken;
const state = { config: null, reg: null };
const $ = s => document.querySelector(s);

function updateSync(msg) { const el = $("#syncIndicator"); if (el) el.textContent = msg; }
function log(msg) { const el = $("#importLog"); if (el) el.innerHTML = `<div>${new Date().toLocaleTimeString()} — ${msg}</div>` + el.innerHTML; }

/* ============================
   MSAL Login (Redirect Flow)
   ============================ */
async function initMsal() {
  try {
    msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
    const redirectResponse = await msalApp.handleRedirectPromise();
    if (redirectResponse && redirectResponse.account) {
      account = redirectResponse.account;
      msalApp.setActiveAccount(account);
      await acquireToken();
      onLogin();
      return;
    }

    const accs = msalApp.getAllAccounts();
    if (accs.length) {
      account = accs[0];
      msalApp.setActiveAccount(account);
      await acquireToken();
      onLogin();
    }
  } catch (err) {
    console.error("Erro MSAL:", err);
    Swal.fire("Erro", "Falha ao inicializar autenticação.", "error");
  }
}

async function acquireToken() {
  try {
    const r = await msalApp.acquireTokenSilent(MSAL_SCOPES);
    accessToken = r.accessToken;
  } catch {
    msalApp.acquireTokenRedirect(MSAL_SCOPES);
  }
}

function ensureLogin() {
  try {
    msalApp.loginRedirect(MSAL_SCOPES);
  } catch (e) {
    console.error(e);
    Swal.fire("Erro", "Falha ao iniciar sessão.", "error");
  }
}

async function onLogout() {
  try { await msalApp.logoutRedirect(); }
  catch (e) { console.warn(e); }
}

function onLogin() {
  $("#btnMsLogin").style.display = "none";
  $("#btnMsLogout").style.display = "inline-flex";
  $("#sessNome").textContent = account.username;
  updateSync("🔁 A carregar...");
  loadConfigAndReg();
}

/* ============================
   Graph API helpers
   ============================ */
async function graphLoad(path) {
  if (!accessToken) await acquireToken();
  const r = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${path}:/content`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (r.ok) return JSON.parse(await r.text());
  if (r.status === 404) return null;
  throw new Error("Graph load failed: " + r.status);
}

async function graphSave(path, obj) {
  if (!accessToken) await acquireToken();
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
  const r = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${path}:/content`, {
    method: "PUT", headers: { Authorization: `Bearer ${accessToken}` }, body: blob
  });
  if (!r.ok) throw new Error("Graph save failed: " + r.status);
}

/* ============================
   Load Config & Registos
   ============================ */
async function loadConfigAndReg() {
  try {
    const cfg = await graphLoad(CFG_PATH);
    const reg = await graphLoad(REG_PATH);
    state.config = cfg || { professores: [], alunos: [], disciplinas: [], grupos: [] };
    state.reg = reg || { versao: "v1", registos: [] };
    localStorage.setItem("esp_config", JSON.stringify(state.config));
    localStorage.setItem("esp_reg", JSON.stringify(state.reg));
    updateSync("💾 Sincronizado");
    renderDay();
  } catch (err) {
    console.warn("Offline fallback:", err);
    const c = localStorage.getItem("esp_config");
    const r = localStorage.getItem("esp_reg");
    if (c) state.config = JSON.parse(c);
    if (r) state.reg = JSON.parse(r);
    updateSync("⚠ Offline");
    renderDay();
  }
}

/* ============================
   Painel do Dia
   ============================ */
function renderDay() {
  const date = $("#dataHoje").value || new Date().toISOString().slice(0, 10);
  $("#dataHoje").value = date;
  const out = $("#sessoesHoje");
  out.innerHTML = "";

  if (!state.config || !state.config.professores) {
    out.innerHTML = "<div>⚠️ Configuração não carregada.</div>";
    return;
  }

  const prof = state.config.professores.find(p => p.email?.toLowerCase() === account?.username?.toLowerCase());
  $("#sessNome").textContent = prof ? prof.nome : "—";

  if (!prof) { out.innerHTML = "<div>Professor não reconhecido.</div>"; return; }

  const grupos = (state.config.grupos || []).filter(g => g.professorId === prof.id);
  if (!grupos.length) { out.innerHTML = "<div>Sem sessões.</div>"; return; }

  grupos.forEach(g => {
    const disc = (state.config.disciplinas || []).find(d => d.id === g.disciplinaId) || { nome: g.disciplinaId };
    const card = document.createElement("div");
    card.className = "session card";
    card.innerHTML = `
      <div><strong>${disc.nome}</strong> | Sala: ${g.sala || "-"} (${g.inicio || "08:15"} - ${g.fim || "09:05"})</div>
      <div style="margin-top:8px;display:flex;gap:6px">
        <input class="input lessonNumber" placeholder="Nº Lição" style="width:90px">
        <input class="input sumario" placeholder="Sumário">
        <button class="btn presencaP">Presente</button>
        <button class="btn ghost presencaF">Falta</button>
      </div>`;
    out.appendChild(card);

    card.querySelector(".presencaP").addEventListener("click", () => quickSaveAttendance(g, card, true));
    card.querySelector(".presencaF").addEventListener("click", () => quickSaveAttendance(g, card, false));
  });
}

function makeRegId() { return "R" + Date.now(); }

async function quickSaveAttendance(group, card, present) {
  try {
    const date = $("#dataHoje").value;
    const lesson = card.querySelector(".lessonNumber").value.trim();
    const sumario = card.querySelector(".sumario").value.trim();
    const reg = { id: makeRegId(), data: date, professorId: group.professorId, disciplinaId: group.disciplinaId, numeroLicao: lesson, sumario, presenca: present };
    state.reg.registos.push(reg);
    await graphSave(REG_PATH, state.reg);
    Swal.fire("Guardado", "Registo gravado com sucesso.", "success");
  } catch (e) {
    Swal.fire("Erro", "Falha ao guardar.", "error");
  }
}

/* ============================
   Manual Registo
   ============================ */
async function manualReg() {
  const { value: vals } = await Swal.fire({
    title: "Novo Registo Manual",
    html: `<input id="swDisc" class="swal2-input" placeholder="Disciplina">
           <input id="swSum" class="swal2-input" placeholder="Sumário">`,
    focusConfirm: false,
    preConfirm: () => ({
      disc: document.getElementById("swDisc").value,
      sum: document.getElementById("swSum").value
    })
  });
  if (vals) {
    const reg = { id: makeRegId(), data: new Date().toISOString().slice(0,10), disciplina: vals.disc, sumario: vals.sum };
    state.reg.registos.push(reg);
    await graphSave(REG_PATH, state.reg);
    Swal.fire("Guardado", "Registo manual criado.", "success");
  }
}

/* ============================
   DOM Ready
   ============================ */
document.addEventListener("DOMContentLoaded", async () => {
  $("#btnMsLogin").addEventListener("click", ensureLogin);
  $("#btnMsLogout").addEventListener("click", onLogout);
  $("#btnRefreshDay").addEventListener("click", renderDay);
  $("#btnManualReg").addEventListener("click", manualReg);

  const theme = localStorage.getItem("esp_theme") || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  document.documentElement.setAttribute("data-theme", theme);
  $("#themeToggle").addEventListener("click", () => {
    const cur = document.documentElement.getAttribute("data-theme");
    const next = cur === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem("esp_theme", next);
  });

  await initMsal();
});
