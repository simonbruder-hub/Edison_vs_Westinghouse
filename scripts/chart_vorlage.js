// Grafik Weltindustrieproduktion im Farbschema der Slidesgo-Vorlage (transparent, Navy-Töne).
const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const W = 1118, H = 612, m = { l: 84, r: 40, t: 76, b: 64 };
const INK = '#0A1A35';
const years = [1860, 1880, 1900, 1913];
const series = [
  { name: 'USA', color: '#0A1C5E', w: 5, dash: '', v: [7.2, 14.7, 23.6, 32.0] },
  { name: 'Grossbritannien', color: '#4E5B9B', w: 5, dash: '', v: [19.9, 22.9, 18.5, 13.6] },
  { name: 'Deutschland', color: '#7D86A8', w: 3.5, dash: '10 8', v: [4.9, 8.5, 13.2, 14.8] },
];
const x = y => m.l + (y - 1860) / (1913 - 1860) * (W - m.l - m.r);
const yv = v => H - m.b - v / 35 * (H - m.t - m.b);
const F = 'Source Sans 3, Source Sans Pro, Arial, sans-serif';
let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" font-family="${F}">`;
s += `<rect x="${x(1886)}" y="${m.t}" width="${x(1892) - x(1886)}" height="${H - m.t - m.b}" fill="#FFFFFF" opacity="0.45"/>`;
s += `<text x="${(x(1886) + x(1892)) / 2}" y="${m.t - 12}" font-size="28" fill="${INK}" text-anchor="middle" font-style="italic">Stromkrieg 1886–1892</text>`;
for (const g of [0, 10, 20, 30]) {
  s += `<line x1="${m.l}" x2="${W - m.r}" y1="${yv(g)}" y2="${yv(g)}" stroke="${INK}" stroke-opacity="${g ? 0.15 : 0.6}" stroke-width="1.5"/>`;
  s += `<text x="${m.l - 12}" y="${yv(g) + 7}" font-size="28" fill="${INK}" text-anchor="end">${g} %</text>`;
}
for (const y of years) s += `<text x="${x(y)}" y="${H - m.b + 40}" font-size="28" fill="${INK}" text-anchor="middle">${y}</text>`;
for (const se of series) {
  s += `<polyline fill="none" stroke="${se.color}" stroke-width="${se.w}" stroke-dasharray="${se.dash}" stroke-linejoin="round" points="${se.v.map((v, i) => `${x(years[i])},${yv(v)}`).join(' ')}"/>`;
  se.v.forEach((v, i) => s += `<circle cx="${x(years[i])}" cy="${yv(v)}" r="8" fill="${se.color}" stroke="#E4E9F2" stroke-width="3"/>`);
}
// Direktbeschriftung Deutschland (USA und GB sind auf der Folie beschriftet)
s += `<text x="${x(1900)}" y="${yv(13.2) + 42}" font-size="27" fill="${INK}" text-anchor="middle">Deutschland</text>`;
s += `</svg>`;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ deviceScaleFactor: 2, viewport: { width: W, height: H } });
  await p.setContent(`<html><body style="margin:0;background:transparent">${s}</body></html>`);
  await p.locator('svg').screenshot({ path: process.argv[2], omitBackground: true }); await b.close();
})();
