// ─── SERENA CHECKLIST VEICULAR — SERVICE WORKER ───────────────────────────────
const CACHE_NAME = 'serena-checklist-v1';
const ASSETS = [
  './checklist_veicular.html',
  './manifest.json'
];

// ─── INSTALL: faz cache dos arquivos principais ───────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// ─── ACTIVATE: limpa caches antigos ──────────────────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// ─── FETCH: serve do cache, tenta rede em paralelo (stale-while-revalidate) ──
self.addEventListener('fetch', event => {
  // Requisições para o Microsoft Graph vão direto para a rede (nunca cacheia)
  if (event.request.url.includes('graph.microsoft.com')) {
    event.respondWith(fetch(event.request));
    return;
  }

  event.respondWith(
    caches.open(CACHE_NAME).then(async cache => {
      const cached = await cache.match(event.request);
      const networkFetch = fetch(event.request).then(response => {
        if (response && response.status === 200) {
          cache.put(event.request, response.clone());
        }
        return response;
      }).catch(() => null);

      return cached || networkFetch;
    })
  );
});

// ─── BACKGROUND SYNC: reenvio de registros pendentes ─────────────────────────
self.addEventListener('sync', event => {
  if (event.tag === 'sync-checklists') {
    event.waitUntil(syncPendingRecords());
  }
});

async function syncPendingRecords() {
  // Abre o IndexedDB para buscar registros pendentes
  const db = await openDB();
  const pending = await getPending(db);

  for (const item of pending) {
    try {
      const resp = await fetch(item.url, {
        method: 'POST',
        headers: item.headers,
        body: item.body
      });
      if (resp.ok) {
        await markSynced(db, item.id);
        // Notifica o app que o registro foi sincronizado
        const clients = await self.clients.matchAll();
        clients.forEach(c => c.postMessage({ type: 'SYNC_OK', id: item.id }));
      }
    } catch (e) {
      // Mantém pendente para próxima tentativa
      console.log('[SW] Sync falhou, tentará novamente:', item.id);
    }
  }
}

// ─── INDEXEDDB helpers ────────────────────────────────────────────────────────
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('serena_sync', 1);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('pending')) {
        db.createObjectStore('pending', { keyPath: 'id' });
      }
    };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = () => reject(req.error);
  });
}

function getPending(db) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction('pending', 'readonly');
    const req = tx.objectStore('pending').getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function markSynced(db, id) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction('pending', 'readwrite');
    tx.objectStore('pending').delete(id);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}
