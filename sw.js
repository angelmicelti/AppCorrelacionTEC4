/**
 * Service Worker para Matriz de correlación TEC4
 *
 * Estrategia de caching:
 *   - App shell (HTML, CSS, JS embebido): network-first con fallback a caché
 *   - CDN libraries (ExcelJS, SheetJS, Firebase): stale-while-revalidate
 *   - Firebase API calls: siempre network (no se cachean)
 *   - Navegación offline: servir la página cacheada
 *
 * Versiones:
 *   v1 — versión inicial
 *   v2 — bug fixes (race conditions, parseNotaEspanola, migración, etc.)
 *   v3 — panel de diagnóstico de auth + detección file:// + mejor manejo de errores
 *   v4 — fix: login no ocultaba pantalla de auth si recargarDatosTrasLogin fallaba
 *   v5 — limpieza: eliminado panel de diagnóstico de acceso
 *   v6 — fix: getTecScores crash si inputs no existen + try/catch en showResults
 *   v7 — matriz de cálculo como acordeón replegable
 *   v8 — matriz de cálculo: solo encabezados por defecto + botón toggle
 *   v9 — selector de materia con radio buttons
 *   v10 — eliminar selector de materia: resultados de todas las materias activas apiladas
 *   v11 — matrices se actualizan en tiempo real al activar/desactivar materias
 *   v12 — resultados también se actualizan en tiempo real al cambiar materias
 *   v13 — confirmar: desmarcar materia elimina su tabla de resultados
 *   v14 — botones de exportación por materia (JSON y Excel) en cada bloque de resultados
 *   v15 — notas en color: rojo < 5, verde >= 5 (formulario y resultados)
 *   v16 — fix: tildes en nombres de archivos de exportación
 *   v17 — mejoras de responsividad móvil: pestañas scrollables, contenedores de
 *         tabla con overflow-x:auto, inputs font-size:16px (sin zoom iOS),
 *         safe-area-inset para notch, modales full-screen en <480px, touch
 *         targets más grandes, breakpoints múltiples (1024/768/480/380px)
 *   v18 — rendimiento táctil: eliminado listener touchend preventivo (bloqueaba
 *         scroll 300ms), mouseover/mouseout desactivados en móvil, throttle rAF
 *         de scroll/resize, debounce 400ms en guardado de notas, delegación
 *         de eventos en pestañas, DocumentFragment en checkboxes, CSS will-change
 *         + contain + content-visibility, transiciones hover desactivadas en móvil
 *   v19 — fix scroll táctil en tablas: touch-action: pan-x en contenedores
 *         scrollables y pan-x pan-y en celdas (manipulation rompía el inicio
 *         del scroll al poner el dedo sobre una celda con listener click).
 *         will-change quitado de celdas para evitar capas GPU que capturan
 *         el gesto táctil.
 *   v20 — fix DEFINITIVO scroll táctil: eliminado position:sticky de th
 *         (causa principal del bloqueo en iOS Safari cuando el dedo está
 *         sobre un header), touch-action: pan-x pan-y forzado en table/thead/
 *         th/tbody, -webkit-overflow-scrolling: touch con !important, quitado
 *         will-change:auto y backface-visibility:hidden de th.
 */

const CACHE_VERSION = 'v20';
const CACHE_NAME = 'matriz-tec4-' + CACHE_VERSION;
const CDN_CACHE_NAME = 'matriz-tec4-cdn-' + CACHE_VERSION;
const APP_SHELL = [
    './',
    './matriz_correlacion_TEC4.html',
    './manifest.json',
    './icon-192.png',
    './icon-512.png',
    './icon-192-maskable.png',
    './icon-512-maskable.png',
    './apple-touch-icon.png',
    './favicon.png',
    './favicon.ico'
];

// Recursos de CDN que podemos cachear
const CDN_ASSETS = [
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
    'https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js',
    'https://www.gstatic.com/firebasejs/10.12.2/firebase-database-compat.js'
];

// ===== INSTALACIÓN =====
self.addEventListener('install', function(event) {
    event.waitUntil(
        Promise.all([
            // Pre-cachear el app shell
            caches.open(CACHE_NAME).then(function(cache) {
                return cache.addAll(APP_SHELL);
            }),
            // Pre-cachear los recursos de CDN
            caches.open(CDN_CACHE_NAME).then(function(cache) {
                return cache.addAll(CDN_ASSETS);
            })
        ]).then(function() {
            // Forzar activación inmediata
            return self.skipWaiting();
        })
    );
});

// ===== ACTIVACIÓN =====
self.addEventListener('activate', function(event) {
    event.waitUntil(
        caches.keys().then(function(cacheNames) {
            return Promise.all(
                cacheNames.map(function(cacheName) {
                    // Eliminar cachés de versiones anteriores
                    if (cacheName.startsWith('matriz-tec4-') && cacheName !== CACHE_NAME && cacheName !== CDN_CACHE_NAME) {
                        console.log('[SW] Eliminando caché antiguo:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        }).then(function() {
            // Tomar control inmediato de todas las pestañas
            return self.clients.claim();
        })
    );
});

// ===== FETCH =====
self.addEventListener('fetch', function(event) {
    const request = event.request;
    const url = new URL(request.url);

    // 1. NUNCA interceptar Firebase API (necesita conexión en tiempo real)
    if (url.hostname.includes('firebasedatabase.app') ||
        url.hostname.includes('firebaseio.com') ||
        url.hostname.includes('gstatic.com') && url.pathname.includes('firebase')) {
        return; // Dejar que el navegador lo gestione normalmente
    }

    // 2. Recursos de CDN: stale-while-revalidate
    if (CDN_ASSETS.indexOf(request.url) !== -1 || url.hostname.includes('cdn.jsdelivr.net')) {
        event.respondWith(staleWhileRevalidate(request, CDN_CACHE_NAME));
        return;
    }

    // 3. Navegación (peticiones de documentos HTML): network-first con fallback a caché
    if (request.mode === 'navigate') {
        event.respondWith(
            fetch(request)
                .then(function(response) {
                    // Actualizar el caché con la versión nueva
                    var responseClone = response.clone();
                    caches.open(CACHE_NAME).then(function(cache) {
                        cache.put(request, responseClone);
                    });
                    return response;
                })
                .catch(function() {
                    // Sin conexión: servir desde caché
                    return caches.match(request)
                        .then(function(cached) {
                            return cached || caches.match('./matriz_correlacion_TEC4.html');
                        });
                })
        );
        return;
    }

    // 4. App shell (imágenes, manifest, etc.): cache-first con revalidación
    //    EXCEPTO el HTML principal que ya se maneja con network-first arriba
    //    (request.mode === 'navigate'), así que aquí solo llegan assets.
    if (request.method === 'GET' && url.origin === self.location.origin) {
        event.respondWith(
            caches.match(request)
                .then(function(cached) {
                    if (cached) {
                        // Devolver del caché y actualizar en background
                        fetch(request).then(function(response) {
                            if (response && response.status === 200) {
                                caches.open(CACHE_NAME).then(function(cache) {
                                    cache.put(request, response);
                                });
                            }
                        }).catch(function() { /* offline, no importa */ });
                        return cached;
                    }
                    // No está en caché: intentar red
                    return fetch(request).then(function(response) {
                        if (response && response.status === 200) {
                            var responseClone = response.clone();
                            caches.open(CACHE_NAME).then(function(cache) {
                                cache.put(request, responseClone);
                            });
                        }
                        return response;
                    }).catch(function() {
                        // Sin conexión y sin caché
                        return new Response('', { status: 504, statusText: 'Offline' });
                    });
                })
        );
        return;
    }
});

/**
 * Estrategia stale-while-revalidate: devolver del caché inmediatamente
 * y actualizar en background.
 */
function staleWhileRevalidate(request, cacheName) {
    return caches.open(cacheName).then(function(cache) {
        return cache.match(request).then(function(cachedResponse) {
            var fetchPromise = fetch(request).then(function(networkResponse) {
                if (networkResponse && networkResponse.status === 200) {
                    cache.put(request, networkResponse.clone());
                }
                return networkResponse;
            }).catch(function() {
                // Sin conexión: si hay caché, ya se devolvió; si no, error
                return cachedResponse || new Response('', { status: 504 });
            });
            // Devolver caché inmediatamente si existe, si no esperar a la red
            return cachedResponse || fetchPromise;
        });
    });
}

// ===== MENSAJES desde la app =====
self.addEventListener('message', function(event) {
    if (event.data && event.data.type === 'SKIP_WAITING') {
        self.skipWaiting();
    }
    if (event.data && event.data.type === 'GET_VERSION') {
        event.ports[0].postMessage({ version: CACHE_VERSION });
    }
});
