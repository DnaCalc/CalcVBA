// Build first: cargo build -p calcvba-spreadsheet-compat --no-default-features --example browser_probe --target wasm32-unknown-unknown
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
const bytes = readFileSync(process.argv[2] ?? 'target/wasm32-unknown-unknown/debug/examples/browser_probe.wasm');
const module = new WebAssembly.Module(bytes);
assert.deepEqual(WebAssembly.Module.imports(module), [], 'pure UDF probe needs no browser clock or native imports');
const instance = new WebAssembly.Instance(module, {});
assert.equal(instance.exports.calcvba_add2_probe(19, 23), 42);
assert.equal(instance.exports.calcvba_add2_probe(-1, 3), 2);
assert.equal(instance.exports.calcvba_budget_probe(), 1, 'On Error must not catch the host budget');
assert.equal(instance.exports.calcvba_matrix_probe(), 42);
console.log('UDF WebAssembly probes: Add2, mixed matrix, noncatchable budget passed');
