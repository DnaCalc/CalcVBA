async function runWithSettings(key, execute) {
    const database = await new Promise((resolve, reject) => {
        const request = indexedDB.open(`calcvba-settings:${key}`, 1);
        request.onupgradeneeded = () => request.result.createObjectStore("state");
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
    try {
        return await new Promise((resolve, reject) => {
            const transaction = database.transaction("state", "readwrite");
            const store = transaction.objectStore("state");
            const request = store.get("settings");
            let result;
            // Run synchronously in the active IndexedDB transaction. Concurrent
            // runs with the same key see each other's committed settings.
            request.onsuccess = () => {
                try {
                    const snapshot = request.result ?? "";
                    result = execute(snapshot);
                    if (result.settings !== undefined && result.settings !== snapshot) {
                        store.put(result.settings, "settings");
                    }
                } catch (error) {
                    reject(error);
                    transaction.abort();
                }
            };
            transaction.oncomplete = () => resolve(result.message);
            transaction.onabort = () => reject(transaction.error ?? new Error("Settings transaction was aborted"));
        });
    } finally {
        database.close();
    }
}

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
        let message, settings;
        const imports = {
            calcvba_browser: {
                report(kind, ptr, len) {
                    const text = read(ptr, len);
                    if (kind === 0) {
                        self.postMessage({ kind: "print", text });
                    } else if (kind === 3) {
                        settings = text;
                    } else {
                        message = { kind: kind === 1 ? "result" : "error", text };
                    }
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
        const allocate = text => {
            const bytes = encoder.encode(text);
            const pointer = wasm.calcvba_source_alloc(bytes.length) >>> 0;
            new Uint8Array(wasm.memory.buffer, pointer, bytes.length).set(bytes);
            return [pointer, bytes.length];
        };
        const execute = snapshot => {
            const locale = data.locale ?? navigator.language;
            const [pointer, length] = allocate(data.source);
            const initialSettings = snapshot === undefined ? [0, 0] : allocate(snapshot);
            wasm.calcvba_run_source(pointer, length, data.dialogs && self.crossOriginIsolated,
                locale.toLowerCase() === "en-za", ...initialSettings);
            if (!message) throw new Error("CalcVBA did not report a result");
            return { message, settings };
        };
        const result = data.settingsKey === undefined
            ? execute().message
            : await runWithSettings(data.settingsKey, execute);
        self.postMessage(result);
    } catch (error) {
        self.postMessage({ kind: "error", text: String(error) });
    }
};
