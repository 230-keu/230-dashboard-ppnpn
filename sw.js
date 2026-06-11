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

// Memotong permintaan internet dan menggantinya dengan file lokal HP
self.addEventListener('fetch', event => {
    event.respondWith(
        caches.match(event.request).then(response => {
            // Jika file ada di memori HP, pakai yang lokal. Jika tidak, ambil dari internet.
            return response || fetch(event.request);
        })
    );
});