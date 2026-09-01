import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";

const auth = readFileSync(new URL("../auth.js", import.meta.url), "utf8");
const app = readFileSync(new URL("../app.js", import.meta.url), "utf8");
const manifest = JSON.parse(readFileSync(new URL("../manifest.webmanifest", import.meta.url), "utf8"));
const worker = readFileSync(new URL("../service-worker.js", import.meta.url), "utf8");

assert.match(auth, /MASOL_INSTALL_STATE/);
assert.match(auth, /beforeinstallprompt/);
assert.match(app, /requestAppInstall/);
assert.match(app, /masol-install-prompt-ready/);
assert.match(app, /Añadir a pantalla de inicio/);
assert.match(app, /display-mode: standalone/);
assert.equal(manifest.id, "./");
assert.ok(manifest.icons.some((icon) => icon.src === "icon-192.png" && icon.sizes === "192x192"));
assert.ok(manifest.icons.some((icon) => icon.src === "icon-512.png" && icon.sizes === "512x512"));
assert.ok(existsSync(new URL("../icon-192.png", import.meta.url)));
assert.ok(existsSync(new URL("../icon-512.png", import.meta.url)));
assert.match(worker, /icon-192\.png/);
assert.match(worker, /icon-512\.png/);

console.log("Validación de instalación PWA: OK");
