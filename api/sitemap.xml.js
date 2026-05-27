export default function handler(req, res) {
  res.setHeader("Content-Type", "application/xml; charset=utf-8")
  res.setHeader("Cache-Control", "public, max-age=3600")
  res.status(200).send(
    `<?xml version="1.0" encoding="UTF-8"?>
    <urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
      <url>
        <loc>https://sheetcodecrest.vercel.app/</loc>
        <lastmod>2026-05-27</lastmod>
        <changefreq>weekly</changefreq>
        <priority>1.0</priority>
      </url>
      <url>
        <loc>https://sheetcodecrest.vercel.app/tools</loc>
        <lastmod>2026-05-27</lastmod>
        <changefreq>weekly</changefreq>
        <priority>0.8</priority>
      </url>
      <url>
        <loc>https://sheetcodecrest.vercel.app/about</loc>
        <lastmod>2026-05-27</lastmod>
        <changefreq>monthly</changefreq>
        <priority>0.5</priority>
      </url>
      <!-- Add all your other pages here -->
    </urlset>`
  )
}