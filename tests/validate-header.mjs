import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";

const app = readFileSync(new URL("../app.js", import.meta.url), "utf8");
const html = readFileSync(new URL("../index.html", import.meta.url), "utf8");
const worker = readFileSync(new URL("../service-worker.js", import.meta.url), "utf8");

assert.doesNotMatch(html, /id="saveButton"/);
assert.doesNotMatch(app, /saveButton/);
assert.match(html, /class="brand-logo" src="logo-masol\.png[^\"]*" alt="MASOL [^"]+"/);
assert.match(html, /auth\.js\?v=21/);
assert.doesNotMatch(html, /<script src="app\.js/);
assert.match(worker, /masol-lecturas-v21/);
assert.match(worker, /logo-masol\.png/);
assert.ok(existsSync(new URL("../logo-masol.png", import.meta.url)));

console.log("Validación de cabecera y acciones: OK");
