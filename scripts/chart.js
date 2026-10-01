const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const W = 680, H = 320, m = { l: 48, r: 150, t: 44, b: 40 };
const years = [1860, 1880, 1900, 1913];
const series = [
  { name: 'USA', color: '#0E7C9E', v: [7.2, 14.7, 23.6, 32.0] },
  { name: 'Deutschland', color: '#7E5AA6', v: [4.9, 8.5, 13.2, 14.8] },
  { name: 'Grossbritannien', color: '#C8641E', v: [19.9, 22.9, 18.5, 13.6] },
];
const x = y => m.l + (y - 1860) / (1913 - 1860) * (W - m.l - m.r);
const yv = v => H - m.b - v / 35 * (H - m.t - m.b);
const fmt = v => v.toFixed(1).replace('.', ',');
let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" font-family="Arial, Liberation Sans, sans-serif">`;
s += `<rect width="${W}" height="${H}" fill="#fff"/>`;
// war of currents band
s += `<rect x="${x(1886)}" y="${m.t}" width="${x(1892) - x(1886)}" height="${H - m.t - m.b}" fill="#F2B33D" opacity="0.22"/>`;
s += `<text x="${(x(1886) + x(1892)) / 2}" y="${m.t - 6}" font-size="11" fill="#55534E" text-anchor="middle">Stromkrieg 1886–1892</text>`;
for (const g of [0, 10, 20, 30]) {
  s += `<line x1="${m.l}" x2="${W - m.r}" y1="${yv(g)}" y2="${yv(g)}" stroke="${g ? '#E6E4DF' : '#9A978F'}" stroke-width="1"/>`;
  s += `<text x="${m.l - 8}" y="${yv(g) + 4}" font-size="11" fill="#6B6862" text-anchor="end">${g} %</text>`;
}
for (const y of years) s += `<text x="${x(y)}" y="${H - m.b + 18}" font-size="11" fill="#6B6862" text-anchor="middle">${y}</text>`;
// legend
let lx = m.l;
for (const se of series) {
  s += `<line x1="${lx}" x2="${lx + 18}" y1="14" y2="14" stroke="${se.color}" stroke-width="3"/><circle cx="${lx + 9}" cy="14" r="4" fill="${se.color}" stroke="#fff" stroke-width="2"/>`;
  s += `<text x="${lx + 24}" y="18" font-size="12" fill="#2A2926">${se.name}</text>`;
  lx += 24 + se.name.length * 6.3 + 26;
}
for (const se of series) {
  s += `<polyline fill="none" stroke="${se.color}" stroke-width="2.5" stroke-linejoin="round" points="${se.v.map((v, i) => `${x(years[i])},${yv(v)}`).join(' ')}"/>`;
  se.v.forEach((v, i) => s += `<circle cx="${x(years[i])}" cy="${yv(v)}" r="4.5" fill="${se.color}" stroke="#fff" stroke-width="2"/>`);
}
// direct end labels (nudged to avoid collision)
const endY = { 'USA': yv(32.0), 'Deutschland': yv(14.8) - 7, 'Grossbritannien': yv(13.6) + 9 };
for (const se of series) s += `<text x="${x(1913) + 10}" y="${endY[se.name] + 4}" font-size="12" fill="#2A2926"><tspan font-weight="bold">${fmt(se.v[3])} %</tspan> ${se.name}</text>`;
s += `</svg>`;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ deviceScaleFactor: 3, viewport: { width: W, height: H } });
  await p.setContent(`<html><body style="margin:0">${s}</body></html>`);
  await p.locator('svg').screenshot({ path: process.argv[2] }); await b.close();
})();
