import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const auth = readFileSync(new URL("../auth.js", import.meta.url), "utf8");
const html = readFileSync(new URL("../index.html", import.meta.url), "utf8");
const worker = readFileSync(new URL("../service-worker.js", import.meta.url), "utf8");
const app = readFileSync(new URL("../app.js", import.meta.url), "utf8");

assert.match(auth, /masol-pumps-auth-session-v1/);
assert.match(auth, /c5c2fdc13e5a50820763d623d5cfd1d4085525010980ae35d682a44bc9a80671/);
assert.match(auth, /SESSION_DURATION_MS = 8 \* 60 \* 60 \* 1000/);
assert.match(auth, /LOCK_DURATION_MS = 5 \* 60 \* 1000/);
assert.match(auth, /MAX_ATTEMPTS = 5/);
assert.match(auth, /loadScript\("app\.js\?v=21"\)/);
assert.match(auth, /id = "logoutButton"/);
assert.match(html, /body class="auth-pending"/);
assert.match(html, /auth\.js\?v=21/);
assert.doesNotMatch(html, /<script src="app\.js/);
assert.match(worker, /masol-lecturas-v21/);
assert.match(worker, /auth\.js\?v=21/);
assert.match(app, /document\.readyState === "complete"/);
assert.match(app, /registerServiceWorker/);

console.log("Validación de autenticación compartida: OK");
