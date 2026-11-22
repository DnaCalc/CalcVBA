# Run CalcVBA in a browser worker

Build the VM3 WebAssembly module and place it beside the JavaScript modules.

```sh
cargo rustc -p calcvba-host --no-default-features --features browser-worker \
  --target wasm32-unknown-unknown --lib --crate-type cdylib
cp target/wasm32-unknown-unknown/debug/calcvba_host.wasm examples/browser/
```

Serve the directory over HTTPS or loopback HTTP. For modal dialogs, send
`Cross-Origin-Opener-Policy: same-origin` and
`Cross-Origin-Embedder-Policy: require-corp` and supply an `onDialog` handler.
Open `probe.html` to check arithmetic, the local clock, dialog replies and load
refusals. The page can also run without cross-origin isolation.

```js
import { runVba } from "./runner.mjs";

const text = await runVba(source, {
    onPrint: console.log,
    onDialog: async dialog => replyFromPage(dialog),
    locale: "en-ZA",
    timeZone: "Africa/Johannesburg",
    settingsKey: "my-project",
});
```

The source contains a parameterless public `Main` procedure. The worker checks
it in Windows-faithful mode and runs it on VM3. `runVba` resolves to the text
of a function result or an empty string after a Sub. Pass an `AbortSignal` as
`signal` to cancel the worker.

A dialog handler receives a `message` with prompt and VBA style, or an `input`
with prompt and default value. Return the selected VBA button number for a
message and the entered text for input. An optional time zone takes an IANA
name. Without a time zone the clock follows the browser's local zone. Locale
formatting uses the en-ZA model when requested explicitly or through the
browser's language, otherwise en-US.

Supply a nonempty `settingsKey` to enable VBA settings. Workers with the same
key on the same origin share an IndexedDB store. Runs with that key execute
serially and report their result after the settings transaction completes.
A dialog handler must reply before starting another run with the same key.
