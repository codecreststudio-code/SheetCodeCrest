export default function handler(req, res) {
  res.setHeader("Content-Type", "application/xml; charset=utf-8")
  res.setHeader("Cache-Control", "public, max-age=3600")
  res.status(200).send(
    '<?xml version="1.0" encoding="UTF-8"?>
' +
    '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
' +
    '  <url>
' +
    '    <loc>https://sheetcodecrest.vercel.app/</loc>
' +
    '    <lastmod>2026-05-27</lastmod>
' +
    '    <changefreq>weekly</changefreq>
' +
    '    <priority>1.0</priority>
' +
    '  </url>
' +
    '</urlset>'
  )
}