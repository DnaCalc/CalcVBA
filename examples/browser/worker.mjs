self.onmessage = async ({ data }) => {
    try {
        const decoder = new TextDecoder();
        const encoder = new TextEncoder();
        const formatter = data.timeZone && new Intl.DateTimeFormat("en-CA", {
            timeZone: data.timeZone, year: "numeric", month: "numeric", day: "numeric",
            hour: "numeric", minute: "numeric", second: "numeric", hourCycle: "h23",
        });
        const read = (ptr, length) => decoder.decode(new Uint8Array(wasm.memory.buffer, ptr, length));
        const ask = dialog => {
            const buffer = new SharedArrayBuffer(8 + 1024);
            const result = new Int32Array(buffer, 0, 2);
            self.postMessage({ kind: "dialog", dialog, buffer });
            while (Atomics.load(result, 0) === 0) {
                Atomics.wait(result, 0, 0);
            }
            return { buffer, result: result[1] };
        };
        const imports = {
            calcvba_browser: {
                report(kind, ptr, len) {
                    self.postMessage({ kind: ["print", "result", "error"][kind], text: read(ptr, len) });
                },
                message(ptr, len, style) {
                    return ask({ kind: "message", prompt: read(ptr, len), style }).result;
                },
                input(ptr, len, def, defLen, output, capacity) {
                    const reply = ask({
                        kind: "input", prompt: read(ptr, len), defaultValue: read(def, defLen),
                    });
                    const length = Math.min(capacity, reply.result);
                    new Uint8Array(wasm.memory.buffer, output, length)
                        .set(new Uint8Array(reply.buffer, 8, length));
                    return length;
                },
                local_time_millis() {
                    const date = new Date();
                    if (!formatter) return date.getTime() - date.getTimezoneOffset() * 60000;
                    const parts = Object.fromEntries(formatter.formatToParts(date)
                        .map(({ type, value }) => [type, Number(value)]));
                    return Date.UTC(parts.year, parts.month - 1, parts.day,
                        parts.hour, parts.minute, parts.second, date.getMilliseconds());
                },
            },
        };
        const response = await fetch(new URL("./calcvba_host.wasm", import.meta.url));
        if (!response.ok) throw new Error("Could not load the CalcVBA WebAssembly module");
        const { instance } = await WebAssembly.instantiate(await response.arrayBuffer(), imports);
        const wasm = instance.exports;
        const bytes = encoder.encode(data.source);
        const pointer = wasm.calcvba_source_alloc(bytes.length) >>> 0;
        new Uint8Array(wasm.memory.buffer, pointer, bytes.length).set(bytes);
        const locale = data.locale ?? navigator.language;
        wasm.calcvba_run_source(pointer, bytes.length, data.dialogs && self.crossOriginIsolated,
            locale.toLowerCase() === "en-za");
    } catch (error) {
        self.postMessage({ kind: "error", text: String(error) });
    }
};
