// Portal do Recanto da Siriema: guarda a tela no aparelho para abrir sem internet.
const CACHE = 'siriema-v4';
const ARQUIVOS = ['./', './index.html', './manifest.json', './icone-180.png', './icone-192.png', './icone-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Os dados vêm do Apps Script e são guardados pelo próprio aplicativo: aqui só passam pela rede.
  if (url.hostname.includes('script.google') || url.hostname.includes('googleusercontent')) return;
  const pagina = req.mode === 'navigate' || url.pathname.endsWith('/') || url.pathname.endsWith('.html');
  if (pagina) {
    // Página: tenta a versão nova; sem internet, abre a guardada.
    e.respondWith(fetch(req).then((r) => { const c = r.clone(); caches.open(CACHE).then((x) => x.put('./index.html', c)); return r; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  // Ícones e fontes: usa o guardado e atualiza quando der.
  e.respondWith(caches.match(req).then((guardado) => {
    const rede = fetch(req).then((r) => { if (r && (r.ok || r.type === 'opaque')) { const c = r.clone(); caches.open(CACHE).then((x) => x.put(req, c)); } return r; }).catch(() => guardado);
    return guardado || rede;
  }));
});
