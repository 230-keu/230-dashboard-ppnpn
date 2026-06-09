export default async function handler(req, res) {
  const URL_GAS = process.env.URL_GAS;
  
  if (!URL_GAS) {
    return res.status(500).json({ status: "error", message: "Kunci URL_GAS belum diatur di Vercel!" });
  }

  // Pengaturan CORS agar aman
  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // Ambil parameter query jika ada (misal: ?action=get_users)
  const queryString = req.url.includes('?') ? req.url.substring(req.url.indexOf('?')) : '';
  const targetUrl = URL_GAS + queryString;

  try {
    const options = {
      method: req.method,
      headers: { 'Content-Type': 'application/json' }
    };

    if (req.method === 'POST') {
      // Teruskan body data dari frontend ke Google Apps Script
      options.body = typeof req.body === 'object' ? JSON.stringify(req.body) : req.body;
    }

    const response = await fetch(targetUrl, options);
    const data = await response.json();
    
    // Kembalikan hasil dari Google Sheets ke Frontend
    return res.status(200).json(data);
  } catch (error) {
    return res.status(500).json({ status: "error", message: "Gagal terhubung ke server API Proxy: " + error.message });
  }
}