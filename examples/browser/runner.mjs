// Run one source module with a parameterless Main on a dedicated Web Worker.
export function runVba(source, { onDialog, onPrint, locale, timeZone, settingsKey, signal } = {}) {
    return new Promise((resolve, reject) => {
        if (signal?.aborted) {
            reject(signal.reason);
            return;
        }
        if (settingsKey !== undefined && (typeof settingsKey !== "string" || !settingsKey)) {
            reject(new TypeError("settingsKey must be a nonempty string"));
            return;
        }
        const worker = new Worker(new URL("./worker.mjs", import.meta.url), { type: "module" });
        const dialogs = Boolean(onDialog && globalThis.crossOriginIsolated
            && typeof SharedArrayBuffer === "function");
        const finish = (complete, value) => {
            signal?.removeEventListener("abort", abort);
            worker.terminate();
            complete(value);
        };
        const abort = () => finish(reject, signal.reason);
        signal?.addEventListener("abort", abort, { once: true });
        worker.onerror = event => finish(reject, new Error(event.message));
        worker.onmessage = async ({ data }) => {
            try {
                switch (data.kind) {
                    case "result":
                        finish(resolve, data.text);
                        break;
                    case "error":
                        finish(reject, new Error(data.text));
                        break;
                    case "print":
                        onPrint?.(data.text);
                        break;
                    case "dialog": {
                        const reply = new Int32Array(data.buffer, 0, 2);
                        const value = await onDialog(data.dialog);
                        if (data.dialog.kind === "message") {
                            if (!Number.isInteger(value) || value < 1 || value > 7) {
                                throw new TypeError("MsgBox needs a VBA button number from 1 through 7");
                            }
                            reply[1] = value;
                        } else {
                            if (typeof value !== "string") {
                                throw new TypeError("InputBox needs a string response");
                            }
                            const bytes = new Uint8Array(data.buffer, 8);
                            const encoded = new TextEncoder().encode(value.slice(0, 254));
                            bytes.set(encoded);
                            reply[1] = encoded.length;
                        }
                        Atomics.store(reply, 0, 1);
                        Atomics.notify(reply, 0);
                        break;
                    }
                }
            } catch (error) {
                finish(reject, error);
            }
        };
        worker.postMessage({ source, dialogs, locale, timeZone, settingsKey });
    });
}
