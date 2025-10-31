/* ESP Painel v4.1 - complete with admin modal, XLSX import, redirect MSAL, robust saves */

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
const state = { config:null, reg:null, prof:null };
const $ = s => document.querySelector(s);

function log(m){ const el = $("#importLog"); if(el) el.innerHTML = `<div>${new Date().toLocaleTimeString()} — ${m}</div>` + el.innerHTML; }

/* ---------------- MSAL redirect flow ---------------- */
async function initMsal(){
  msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
  const redirectResp = await msalApp.handleRedirectPromise();
  if(redirectResp && redirectResp.account){ account = redirectResp.account; msalApp.setActiveAccount(account); await acquireToken(); onLogin(); return; }
  const accs = msalApp.getAllAccounts();
  if(accs.length){ account = accs[0]; msalApp.setActiveAccount(account); await acquireToken(); onLogin(); }
}

async function acquireToken(){
  try{
    const r = await msalApp.acquireTokenSilent(MSAL_SCOPES);
    accessToken = r.accessToken;
  }catch(e){
    console.warn('Silent token failed',e);
    try{ await msalApp.acquireTokenRedirect(MSAL_SCOPES); }catch(err){ console.error(err); }
  }
}

function ensureLogin(){ try{ msalApp.loginRedirect(MSAL_SCOPES); }catch(e){ console.error('Login redirect failed',e); Swal.fire('Erro','Falha ao iniciar sessão','error'); } }

async function onLogout(){ try{ await msalApp.logoutRedirect(); }catch(e){ console.warn(e); } }

function onLogin(){ $("#btnMsLogin").style.display='none'; $("#btnMsLogout").style.display='inline-flex'; $("#sessNome").textContent = account.username; updateSync('🔁 sincronizando...'); loadConfigAndReg(); }

/* ---------------- Graph helpers ---------------- */
async function graphLoad(path){
  if(!accessToken) await acquireToken();
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${path}:/content`;
  const r = await fetch(url, { headers:{ Authorization:`Bearer ${accessToken}` } });
  if(r.ok){ const t = await r.text(); return JSON.parse(t); }
  if(r.status===404) return null;
  throw new Error('Graph load failed: '+r.status);
}
async function graphSave(path,obj){
  if(!accessToken) await acquireToken();
  const blob = new Blob([JSON.stringify(obj,null,2)],{type:'application/json'});
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${path}:/content`;
  const r = await fetch(url,{ method:'PUT', headers:{ Authorization:`Bearer ${accessToken}` }, body: blob });
  if(!r.ok) throw new Error('Graph save failed: '+r.status);
  return await r.json();
}
async function graphList(folderPath){
  if(!accessToken) await acquireToken();
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${folderPath}:/children`;
  const r = await fetch(url,{ headers:{ Authorization:`Bearer ${accessToken}` } });
  if(r.ok) return await r.json();
  if(r.status===404) return { value: [] };
  throw new Error('Graph list failed: '+r.status);
}
async function graphDelete(itemId){
  if(!accessToken) await acquireToken();
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/items/${itemId}`;
  const r = await fetch(url,{ method:'DELETE', headers:{ Authorization:`Bearer ${accessToken}` } });
  if(!r.ok) throw new Error('Graph delete failed: '+r.status);
  return true;
}

/* ---------------- Config & reg load ---------------- */
async function loadConfigAndReg(){
  try{
    const cfg = await graphLoad(CFG_PATH);
    const reg = await graphLoad(REG_PATH);
    state.config = cfg || { professores:[],alunos:[],disciplinas:[],grupos:[],calendario:{} };
    state.reg = reg || { versao:'v1', registos:[] };
    localStorage.setItem('esp_config', JSON.stringify(state.config));
    localStorage.setItem('esp_reg', JSON.stringify(state.reg));
    updateSync('💾 guardado');
    renderDay();
  }catch(e){
    console.warn('Fallback to cache',e);
    const c = localStorage.getItem('esp_config');
    if(c) state.config = JSON.parse(c);
    const r = localStorage.getItem('esp_reg');
    if(r) state.reg = JSON.parse(r);
    if(!state.config) state.config = { professores:[],alunos:[],disciplinas:[],grupos:[],calendario:{} };
    if(!state.reg) state.reg = { versao:'v1', registos:[] };
    updateSync('⚠ offline');
    renderDay();
  }
}
