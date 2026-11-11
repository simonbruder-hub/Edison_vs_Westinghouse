// Grafik «Anteil an der Weltindustrieproduktion» (Daten: Bairoch 1982) für Word und Folien.
// Aufruf: node scripts/chart.js   → assets/chart_weltindustrie.png (weiss) und …_folie.png (transparent)
const path = require('path');
const { chromium } = require('playwright');

const W = 1520, H = 600, m = { l: 96, r: 372, t: 72, b: 66 };
const INK = '#0A1A35';
const years = [1860, 1880, 1900, 1913];
// Farben: geprüfte Palette (Blaugrün, Kupfer, Violett), abgedunkelt für Kontrast >= 3,5:1 auf dem Folienhintergrund.
// Zweite Kodierung für den Schwarz-Weiss-Druck: Strichmuster und Beschriftung direkt an der Linie.
const series = [
  { name: 'USA', color: '#0B6682', w: 5.5, dash: '', v: [7.2, 14.7, 23.6, 32.0], dy: 0 },
  { name: 'Deutschland', color: '#6C4A93', w: 4.5, dash: '2 11', v: [4.9, 8.5, 13.2, 14.8], dy: -14 },
  { name: 'Grossbritannien', color: '#B0531A', w: 5, dash: '20 11', v: [19.9, 22.9, 18.5, 13.6], dy: 26 },
];
const x = y => m.l + (y - 1860) / (1913 - 1860) * (W - m.l - m.r);
const yv = v => H - m.b - v / 35 * (H - m.t - m.b);
const fmt = v => v.toFixed(1).replace('.', ',');

function svg(surface) {
  const F = "'Source Sans 3', Arial, sans-serif";
  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" font-family="${F}">`;
  s += `<rect x="${x(1886)}" y="${m.t}" width="${x(1892) - x(1886)}" height="${H - m.t - m.b}" fill="${INK}" opacity="0.09"/>`;
  s += `<text x="${(x(1886) + x(1892)) / 2}" y="${m.t - 16}" font-size="30" font-style="italic" fill="${INK}" text-anchor="middle">Stromkrieg 1886–1892</text>`;
  for (const g of [0, 10, 20, 30]) {
    s += `<line x1="${m.l}" x2="${x(1913)}" y1="${yv(g)}" y2="${yv(g)}" stroke="${INK}" stroke-opacity="${g ? 0.16 : 0.65}" stroke-width="1.5"/>`;
    s += `<text x="${m.l - 14}" y="${yv(g) + 11}" font-size="32" fill="${INK}" text-anchor="end">${g} %</text>`;
  }
  for (const y of years) s += `<text x="${x(y)}" y="${H - m.b + 44}" font-size="32" fill="${INK}" text-anchor="middle">${y}</text>`;
  for (const se of series) {
    s += `<polyline fill="none" stroke="${se.color}" stroke-width="${se.w}" stroke-dasharray="${se.dash}" stroke-linecap="round" stroke-linejoin="round" points="${se.v.map((v, i) => `${x(years[i])},${yv(v)}`).join(' ')}"/>`;
    se.v.forEach((v, i) => s += `<circle cx="${x(years[i])}" cy="${yv(v)}" r="9" fill="${se.color}" stroke="${surface}" stroke-width="3"/>`);
    s += `<text x="${x(1913) + 22}" y="${yv(se.v[3]) + 11 + se.dy}" font-size="33" fill="${INK}"><tspan font-weight="700">${fmt(se.v[3])} %</tspan> ${se.name}</text>`;
  }
  return s + '</svg>';
}

(async () => {
  const browser = await chromium.launch();
  for (const [file, surface, transparent] of [['chart_weltindustrie.png', '#FFFFFF', false], ['chart_weltindustrie_folie.png', '#D2D7EA', true]]) {
    const page = await browser.newPage({ deviceScaleFactor: 2, viewport: { width: W, height: H } });
    await page.setContent(`<html><body style="margin:0;background:${transparent ? 'transparent' : surface}">${svg(surface)}</body></html>`);
    await page.evaluate(() => document.fonts.ready);
    await page.locator('svg').screenshot({ path: path.join(__dirname, '..', 'assets', file), omitBackground: transparent });
    await page.close();
  }
  await browser.close();
  console.log('OK Grafiken');
})();
