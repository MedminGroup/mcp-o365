import { createServer, IncomingMessage, ServerResponse } from 'node:http';
import { writeFileSync, readFileSync, existsSync } from 'node:fs';
import { exec } from 'node:child_process';

const DEVICE_CODE_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode';
const TOKEN_URL       = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const GRAPH_ME_URL    = 'https://graph.microsoft.com/v1.0/me';

const SCOPES = [
  'openid', 'profile', 'email',
  'Calendars.ReadWrite', 'Mail.ReadWrite', 'Mail.Send',
  'Files.ReadWrite', 'Sites.Read.All', 'Contacts.ReadWrite',
  'People.Read', 'User.Read', 'OnlineMeetings.Read',
  'OnlineMeetingTranscript.Read.All', 'offline_access',
].join(' ');

// ── Auth state ────────────────────────────────────────────────────────────────

type AuthState =
  | { status: 'idle' }
  | { status: 'pending'; userCode: string; verificationUri: string; expiresInMinutes: number }
  | { status: 'complete'; username: string; name: string }
  | { status: 'error'; message: string };

let state: AuthState = { status: 'idle' };
let closeServer: (() => void) | null = null;

// ── Helpers ───────────────────────────────────────────────────────────────────

function openBrowser(url: string): void {
  const cmd = process.platform === 'darwin' ? `open "${url}"`
            : process.platform === 'win32'  ? `start "" "${url}"`
            : `xdg-open "${url}"`;
  exec(cmd, () => {});
}

async function post(url: string, params: Record<string, string>): Promise<Record<string, unknown>> {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(params).toString(),
  });
  return res.json() as Promise<Record<string, unknown>>;
}

function parseIdToken(token: string): Record<string, string> {
  try {
    const part = token.split('.')[1];
    return JSON.parse(
      Buffer.from(part + '='.repeat((4 - (part.length % 4)) % 4), 'base64').toString('utf-8'),
    ) as Record<string, string>;
  } catch { return {}; }
}

function findFreePort(start: number): Promise<number> {
  return new Promise((resolve, reject) => {
    const net = require('node:net') as typeof import('node:net');
    const s = net.createServer();
    s.listen(start, '127.0.0.1', () => {
      const addr = s.address() as { port: number };
      s.close(() => resolve(addr.port));
    });
    s.on('error', () =>
      start < 3860
        ? findFreePort(start + 1).then(resolve, reject)
        : reject(new Error('No free port found in range 3847–3860')),
    );
  });
}

// ── Device-code flow ──────────────────────────────────────────────────────────

function startDeviceCodeFlow(clientId: string, cachePath: string): void {
  (async () => {
    const dc = await post(DEVICE_CODE_URL, { client_id: clientId, scope: SCOPES });

    if (!dc.device_code) {
      state = { status: 'error', message: String(dc.error_description ?? dc.error ?? 'Failed to get device code') };
      return;
    }

    state = {
      status: 'pending',
      userCode: String(dc.user_code),
      verificationUri: String(dc.verification_uri),
      expiresInMinutes: Math.floor(Number(dc.expires_in ?? 900) / 60),
    };

    let interval = Number(dc.interval ?? 5);

    while (true) {
      await new Promise(r => setTimeout(r, interval * 1000));

      const data = await post(TOKEN_URL, {
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        device_code: String(dc.device_code),
        client_id: clientId,
      });

      if (data.access_token) {
        const accessToken = String(data.access_token);
        const claims = data.id_token ? parseIdToken(String(data.id_token)) : {};
        let username = String(claims.preferred_username ?? claims.upn ?? claims.email ?? '');
        let name     = String(claims.name ?? '');

        if (!username) {
          try {
            const me = await fetch(GRAPH_ME_URL, {
              headers: { Authorization: `Bearer ${accessToken}` },
            }).then(r => r.json()) as Record<string, string>;
            username = me.userPrincipalName ?? me.mail ?? me.id ?? 'unknown';
            name     = me.displayName ?? username;
          } catch { username = 'unknown'; name = 'unknown'; }
        }
        if (!name) name = username;

        const cache = existsSync(cachePath)
          ? JSON.parse(readFileSync(cachePath, 'utf-8')) as { accounts: Record<string, unknown> }
          : { accounts: {} };
        delete cache.accounts['unknown'];
        cache.accounts[username.toLowerCase()] = {
          username, name,
          tenantId: claims.tid ?? 'common',
          accessToken,
          refreshToken: data.refresh_token ? String(data.refresh_token) : '',
          expiresAt: Math.floor(Date.now() / 1000) + Number(data.expires_in ?? 3600),
        };
        writeFileSync(cachePath, JSON.stringify(cache, null, 2), { mode: 0o600, encoding: 'utf-8' });

        state = { status: 'complete', username, name };
        // Close the server after a short delay so the page can show success
        setTimeout(() => closeServer?.(), 8000);
        return;
      }

      const error = String(data.error ?? '');
      if (error === 'authorization_pending') continue;
      if (error === 'slow_down')    { interval += 5; continue; }
      if (error === 'expired_token') {
        state = { status: 'error', message: 'Sign-in code expired. Please refresh and try again.' };
        return;
      }
      state = { status: 'error', message: String(data.error_description ?? data.error ?? 'Authentication error') };
      return;
    }
  })().catch(err => {
    state = { status: 'error', message: err instanceof Error ? err.message : String(err) };
  });
}

// ── HTTP request handler ──────────────────────────────────────────────────────

function handleRequest(
  clientId: string,
  cachePath: string,
  req: IncomingMessage,
  res: ServerResponse,
): void {
  const url = req.url ?? '/';

  // API: start auth
  if (url === '/api/start' && req.method === 'POST') {
    if (state.status === 'idle' || state.status === 'error') {
      state = { status: 'idle' };
      startDeviceCodeFlow(clientId, cachePath);
    }
    // Wait briefly for the device code to arrive
    setTimeout(() => {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      if (state.status === 'pending') {
        res.end(JSON.stringify({
          userCode: state.userCode,
          verificationUri: state.verificationUri,
          expiresInMinutes: state.expiresInMinutes,
        }));
      } else if (state.status === 'error') {
        res.end(JSON.stringify({ error: state.message }));
      } else {
        res.end(JSON.stringify({ error: 'Not ready yet — try again in a moment.' }));
      }
    }, 1500);
    return;
  }

  // API: poll status
  if (url === '/api/status' && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(state));
    return;
  }

  // Serve HTML
  if (url === '/' || url === '/index.html') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(HTML);
    return;
  }

  res.writeHead(404);
  res.end('Not found');
}

// ── Main export ───────────────────────────────────────────────────────────────

export async function runSetup(clientId: string, cachePath: string): Promise<void> {
  const port = await findFreePort(3847);
  const url  = `http://localhost:${port}`;

  const server = createServer(handleRequest.bind(null, clientId, cachePath));

  await new Promise<void>(resolve => server.listen(port, '127.0.0.1', resolve));

  closeServer = () => server.close();

  process.stderr.write(`\n  Setup wizard running at ${url}\n  Press Ctrl+C to cancel\n\n`);
  setTimeout(() => openBrowser(url), 400);

  await new Promise<void>(resolve => server.on('close', resolve));
}

// ── HTML ──────────────────────────────────────────────────────────────────────

const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Medmin — Microsoft 365 Setup</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif;
  background:#eef2f7;
  min-height:100vh;
  display:flex;
  flex-direction:column;
  align-items:center;
  justify-content:center;
  padding:24px;
  color:#1a2744;
}
.card{background:white;border-radius:16px;box-shadow:0 4px 32px rgba(0,0,0,.12);width:100%;max-width:520px;overflow:hidden}
.hd{background:#1B3A6B;padding:28px 32px;color:white}
.hd .logo{font-size:11px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;opacity:.65;margin-bottom:8px}
.hd h1{font-size:21px;font-weight:600;line-height:1.3}
.bd{padding:32px}
.ft{padding:14px 32px;border-top:1px solid #edf2f7;font-size:12px;color:#a0aec0;text-align:center}
.panel{display:none}.panel.on{display:block}
p{font-size:15px;line-height:1.65;color:#4a5568;margin-bottom:14px}
ul.feats{list-style:none;margin-bottom:24px}
ul.feats li{font-size:14px;color:#4a5568;padding:5px 0 5px 22px;position:relative}
ul.feats li::before{content:'✓';position:absolute;left:0;color:#2B6CB0;font-weight:700}
.btn{
  display:inline-flex;align-items:center;justify-content:center;gap:8px;
  background:#1B3A6B;color:white;border:none;border-radius:8px;
  padding:13px 24px;font-size:15px;font-weight:600;cursor:pointer;
  width:100%;transition:background .15s;text-decoration:none
}
.btn:hover{background:#2c4f8f}
.btn:disabled{background:#a0aec0;cursor:not-allowed}
.btn-ghost{background:white;color:#1B3A6B;border:2px solid #1B3A6B;margin-top:10px}
.btn-ghost:hover{background:#ebf0f9}
.codebox{
  background:#EBF4FF;border:2px solid #BEE3F8;border-radius:10px;
  padding:20px;text-align:center;margin:18px 0
}
.codelabel{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#2B6CB0;margin-bottom:8px}
.codevalue{font-size:38px;font-weight:700;letter-spacing:.18em;color:#1B3A6B;font-family:'Courier New',monospace}
.statusline{display:flex;align-items:center;gap:10px;font-size:14px;color:#718096;margin-top:16px}
.spin{display:inline-block;width:14px;height:14px;border:2px solid #CBD5E0;border-radius:50%;border-top-color:#2B6CB0;animation:sp .8s linear infinite;flex-shrink:0}
.spin-lg{width:18px;height:18px;border:2px solid rgba(255,255,255,.35);border-top-color:white}
@keyframes sp{to{transform:rotate(360deg)}}
.success-ring{width:64px;height:64px;background:#C6F6D5;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 20px;font-size:28px}
.acct{background:#F7FAFC;border:1px solid #E2E8F0;border-radius:8px;padding:12px 16px;margin:14px 0}
.acct-name{font-weight:600;color:#1a2744}
.acct-email{color:#718096;font-size:13px;margin-top:2px}
.tip{font-size:13px;color:#744210;margin-top:14px;padding:11px 14px;background:#FFFBEB;border-radius:8px;border-left:3px solid #F6AD55;line-height:1.5}
.err{font-size:13px;color:#C53030;font-family:monospace;background:#FFF5F5;border-radius:6px;padding:10px 12px;margin-top:10px;line-height:1.5}
</style>
</head>
<body>
<div class="card">
  <div class="hd">
    <div class="logo">Medmin</div>
    <h1>Microsoft 365 Setup</h1>
  </div>

  <div class="bd">

    <div class="panel on" id="p-welcome">
      <p>Connect your Microsoft 365 account to let Claude analyse your Teams meeting recordings.</p>
      <ul class="feats">
        <li>Access your calendar to find recorded meetings</li>
        <li>Fetch live transcripts from Teams</li>
        <li>Analyse speaking patterns, decisions, and action items</li>
        <li>Your credentials stay on this device &mdash; never uploaded</li>
      </ul>
      <button class="btn" id="btn-connect" onclick="startSignIn(this)">
        Connect Microsoft 365 account
      </button>
    </div>

    <div class="panel" id="p-auth">
      <p>A Microsoft sign-in page has opened in your browser.<br>Enter the code below when prompted:</p>
      <div class="codebox">
        <div class="codelabel">Your sign-in code</div>
        <div class="codevalue" id="user-code">&#8212;&#8212;&#8212;&#8212;</div>
      </div>
      <a class="btn" id="ms-link" href="#" target="_blank" rel="noopener">
        Open Microsoft sign-in page
      </a>
      <div class="statusline">
        <span class="spin"></span>
        <span>Waiting for you to sign in&hellip;</span>
      </div>
    </div>

    <div class="panel" id="p-complete">
      <div class="success-ring">&#10003;</div>
      <p style="text-align:center;font-weight:600;font-size:17px;color:#276749;margin-bottom:6px">
        Account connected!
      </p>
      <div class="acct">
        <div class="acct-name" id="acct-name"></div>
        <div class="acct-email" id="acct-email"></div>
      </div>
      <p>You&rsquo;re all set. <strong>Restart Claude Code</strong>, then try:</p>
      <p style="font-style:italic;color:#2B6CB0">&ldquo;Analyse last week&rsquo;s [meeting name] for [your name]&rdquo;</p>
      <p class="tip">
        Requires transcription to have been active during the meeting &mdash;
        in Teams click <strong>&bull;&bull;&bull; &rarr; Start transcription</strong> before your next meeting.
      </p>
    </div>

    <div class="panel" id="p-error">
      <p style="font-weight:600;color:#C53030">Something went wrong</p>
      <div class="err" id="err-msg"></div>
      <button class="btn btn-ghost" onclick="reset()">Try again</button>
    </div>

  </div>
  <div class="ft">Credentials stored at ~/.mcp-o365-token-cache.json &mdash; never transmitted to Medmin</div>
</div>

<script>
var polling=false;
function show(id){document.querySelectorAll('.panel').forEach(function(p){p.classList.remove('on')});document.getElementById('p-'+id).classList.add('on')}
async function startSignIn(btn){
  btn.disabled=true;
  btn.innerHTML='<span class="spin spin-lg"></span>&nbsp;Connecting&hellip;';
  try{
    var r=await fetch('/api/start',{method:'POST'});
    var d=await r.json();
    if(d.error)throw new Error(d.error);
    document.getElementById('user-code').textContent=d.userCode;
    document.getElementById('ms-link').href=d.verificationUri;
    show('auth');
    window.open(d.verificationUri,'_blank');
    startPoll();
  }catch(e){
    btn.disabled=false;
    btn.textContent='Connect Microsoft 365 account';
    document.getElementById('err-msg').textContent=e.message;
    show('error');
  }
}
function startPoll(){if(polling)return;polling=true;poll()}
async function poll(){
  if(!polling)return;
  try{
    var r=await fetch('/api/status');
    var d=await r.json();
    if(d.status==='complete'){
      polling=false;
      document.getElementById('acct-name').textContent=d.name;
      document.getElementById('acct-email').textContent=d.username;
      show('complete');
      return;
    }
    if(d.status==='error'){
      polling=false;
      document.getElementById('err-msg').textContent=d.message;
      show('error');
      return;
    }
  }catch(e){}
  setTimeout(poll,2500);
}
function reset(){polling=false;show('welcome');document.getElementById('user-code').textContent='\u2014\u2014\u2014\u2014'}
</script>
</body>
</html>`;
