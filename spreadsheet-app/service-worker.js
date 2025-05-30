const CACHE_NAME = 'excel-lite-cache-v1';

const STATIC_ASSETS = [
  './',
  './index.html',
  './css/styles.css',
  './js/app.js',
  './js/spreadsheet.js',
  './js/formulas.js',
  './js/import-export.js',
  './js/charts.js',
  './js/conditional-formatting.js',
  './js/utils.js',
  './js/themes.js',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-solid-900.woff2',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-regular-400.woff2'
];

// Install event - cache the static assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Caching static assets');
        // Use Promise.allSettled to continue even if some assets fail to cache
        return Promise.allSettled(
          STATIC_ASSETS.map(url => 
            fetch(url, { cache: 'no-cache' })
              .then(response => {
                if (response.ok) {
                  return cache.put(url, response);
                }
                console.warn(`Failed to cache: ${url}, status: ${response.status}`);
              })
              .catch(error => {
                console.warn(`Failed to fetch for caching: ${url}`, error);
              })
          )
        );
      })
      .then(() => self.skipWaiting()) // Activate new service worker immediately
  );
});

// Activate event - clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.filter(cacheName => {
          return cacheName !== CACHE_NAME;
        }).map(cacheName => {
          console.log('Deleting old cache:', cacheName);
          return caches.delete(cacheName);
        })
      );
    }).then(() => self.clients.claim()) // Take control of all clients
  );
});

// Fetch event - network-first strategy with fallback to cache
self.addEventListener('fetch', event => {
  // Skip non-GET requests and browser extensions
  if (event.request.method !== 'GET' || event.request.url.startsWith('chrome-extension://')) {
    return;
  }

  // Skip icon requests that aren't available yet
  if (event.request.url.includes('/icons/') && event.request.url.includes('icon-')) {
    // For missing icons, we'll return a default icon if we have one
    event.respondWith(
      fetch(event.request)
        .catch(() => {
          // If we have a default icon, use it
          return caches.match('/favicon.ico')
            .then(response => {
              if (response) {
                return response;
              }
              // Otherwise return a simple transparent png
              return new Response(
                new Uint8Array([
                  0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
                  0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
                  0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                  0x42, 0x60, 0x82
                ]).buffer,
                {
                  status: 200,
                  headers: { 'Content-Type': 'image/png' }
                }
              );
            });
        })
    );
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Clone the response to store in cache
        const responseToCache = response.clone();
        
        // Only cache successful responses
        if (response.status === 200) {
          caches.open(CACHE_NAME)
            .then(cache => {
              cache.put(event.request, responseToCache);
            });
        }
        
        return response;
      })
      .catch(error => {
        console.log('Fetch failed, falling back to cache:', error);
        return caches.match(event.request)
          .then(cachedResponse => {
            if (cachedResponse) {
              return cachedResponse;
            }
            
            // If no cache match and it's an HTML request, return the offline page
            if (event.request.headers.get('accept') && 
                event.request.headers.get('accept').includes('text/html')) {
              return caches.match('/index.html');
            }
            
            return new Response('Network error', {
              status: 408,
              headers: { 'Content-Type': 'text/plain' }
            });
          });
      })
  );
});

// Handle messages from the main thread
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// Periodic sync for background updates (when supported)
self.addEventListener('periodicsync', event => {
  if (event.tag === 'sync-spreadsheet-data') {
    event.waitUntil(syncSpreadsheetData());
  }
});

// Background sync function
async function syncSpreadsheetData() {
  // This would implement logic to sync with a backend server
  console.log('Background sync triggered');
}
