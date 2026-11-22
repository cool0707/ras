// Makes the playground cross-origin isolated on hosts that cannot send the headers themselves
// (GitHub Pages): the responses get COOP / COEP, so that the page has SharedArrayBuffer, with
// which a running program waits for the answer of a dialog (msgBox / inputBox) shown by the page.
// The playground loads only files of its own origin.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => event.waitUntil(self.clients.claim()));
self.addEventListener("fetch", (event) => {
    const request = event.request;
    if (request.cache === "only-if-cached" && request.mode !== "same-origin") return;
    if (new URL(request.url).origin !== self.location.origin) return;
    event.respondWith(
        fetch(request).then((response) => {
            if (response.status === 0) return response;
            const headers = new Headers(response.headers);
            headers.set("Cross-Origin-Embedder-Policy", "require-corp");
            headers.set("Cross-Origin-Opener-Policy", "same-origin");
            headers.set("Cross-Origin-Resource-Policy", "same-origin");
            return new Response(response.body, { status: response.status, statusText: response.statusText, headers });
        }),
    );
});
