const CACHE_NAME = 'presensi-ai-cache-v1';

// Daftar file dan otak AI yang WAJIB disimpan ke memori HP
const urlsToCache = [
    '/',
    '/index.html',
    '/dashboard.html',
    '/liveness.html',
    '/admin.html',
    'https://arngoding.github.io/aset-presensi-230/face-api.min.js',
    'https://arngoding.github.io/aset-presensi-230/weights/tiny_face_detector_model-weights_manifest.json',
    'https://arngoding.github.io/aset-presensi-230/weights/tiny_face_detector_model.weights.bin',
    'https://arngoding.github.io/aset-presensi-230/weights/face_landmark_68_model-weights_manifest.json',
    'https://arngoding.github.io/aset-presensi-230/weights/face_landmark_68_model.weights.bin',
    'https://arngoding.github.io/aset-presensi-230/weights/face_expression_model-weights_manifest.json',
    'https://arngoding.github.io/aset-presensi-230/weights/face_expression_model.weights.bin',
    'https://arngoding.github.io/aset-presensi-230/weights/face_recognition_model-weights_manifest.json',
    'https://arngoding.github.io/aset-presensi-230/weights/face_recognition_model.weights.bin'
];

// Proses Mengunduh & Menyimpan saat pertama kali diakses
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => {
            console.log('Cache AI Berhasil Disimpan!');
            return cache.addAll(urlsToCache);
        })
    );
});

// Menghapus cache versi lama ketika ada versi baru
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cacheName => {
                    if (cacheName !== CACHE_NAME) {
                        console.log('Menghapus Cache Lama:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
});

// Memotong permintaan internet dan menggantinya dengan file lokal HP
self.addEventListener('fetch', event => {
    // Khusus untuk halaman HTML, gunakan metode Network First (agar selalu dapat update terbaru)
    if (event.request.mode === 'navigate' || (event.request.headers.get('accept') && event.request.headers.get('accept').includes('text/html'))) {
        event.respondWith(
            fetch(event.request).catch(() => caches.match(event.request))
        );
    } else {
        // Untuk aset lain gunakan Cache First
        event.respondWith(
            caches.match(event.request).then(response => {
                return response || fetch(event.request);
            })
        );
    }
});