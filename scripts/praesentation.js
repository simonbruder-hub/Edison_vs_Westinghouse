// Erzeugt abgabe/Praesentation_Stromkrieg_Bruder.pptx (16:9, 10 Folien + Anhang, mit Sprechernotizen).
// Zitatnummern = Nummern im Handout (assets/zitierreihenfolge.json, von handout.js geschrieben).
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const REFS = require('./quellen');

const ROOT = path.join(__dirname, '..');
const ORDER = require(path.join(ROOT, 'assets', 'zitierreihenfolge.json'));
const OUT = path.join(ROOT, 'abgabe', 'Praesentation_Stromkrieg_Bruder.pptx');

const C = {
  dark: '15171C', ink: '1E1E1E', muted: '5F5D58', light: 'F3F5F6', tint: 'E6F1F5',
  teal: '0E7C9E', copper: 'C8641E', amber: 'F2B33D', violet: '7E5AA6', white: 'FFFFFF', grayline: 'C9D3D8',
};
const HF = 'Cambria', BF = 'Calibri';

const n = k => { const i = ORDER.indexOf(k); if (i < 0) throw new Error('Nicht im Handout zitiert: ' + k); return i + 1; };
function cites(...keys) {
  const nums = keys.map(n).sort((a, b) => a - b), parts = [];
  for (let i = 0; i < nums.length;) {
    let j = i; while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++;
    if (j - i >= 2) parts.push(`[${nums[i]}]–[${nums[j]}]`); else for (let k = i; k <= j; k++) parts.push(`[${nums[k]}]`);
    i = j + 1;
  }
  return parts.join(', ');
}

async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: '#' + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
  pres.author = 'Simon Bruder';
  pres.title = 'Der Stromkrieg – Edison vs. Westinghouse';

  const title = (s, text, color = C.ink) => s.addText(text, { x: 0.5, y: 0.32, w: 9, h: 0.75, fontFace: HF, fontSize: 26, bold: true, color, margin: 0, valign: 'middle', isTextBox: true });
  const source = (s, text, color = C.muted) => s.addText('Quellen: ' + text + ' (Nummern wie im Handout)', { x: 0.5, y: 5.2, w: 7.2, h: 0.25, fontFace: BF, fontSize: 9, color, margin: 0, isTextBox: true });
  const pageNo = (s, i, color = C.muted) => s.addText(String(i), { x: 9.0, y: 5.2, w: 0.5, h: 0.25, fontFace: BF, fontSize: 9, color, align: 'right', margin: 0, isTextBox: true });
  const circleIcon = async (s, name, x, y, d, fill, fg = C.white) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
    const p = d * 0.25;
    s.addImage({ data: await icon(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  };

  // ---------- 1 Einstieg ----------
  {
    const s = pres.addSlide(); s.background = { color: C.dark };
    await circleIcon(s, 'FaBolt', 0.5, 0.45, 0.7, C.amber, C.dark);
    s.addText('Referat · Modul «Weltmacht USA» · FHNW', { x: 1.4, y: 0.55, w: 7, h: 0.5, fontFace: BF, fontSize: 13, color: C.amber, margin: 0, valign: 'middle', isTextBox: true });
    s.addText('Der Stromkrieg', { x: 0.5, y: 1.35, w: 9, h: 0.95, fontFace: HF, fontSize: 48, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText('Edison vs. Westinghouse – der Kampf um Amerikas Stromnetz und die Rolle von Nikola Tesla', { x: 0.5, y: 2.3, w: 8.6, h: 0.7, fontFace: BF, fontSize: 18, color: 'C9CDD3', margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.3, w: 9, h: 1.35, fill: { color: '22252C' }, line: { color: '22252C' } });
    s.addText([
      { text: 'Auburn (New York), 6. August 1890: ', options: { bold: true, color: C.amber } },
      { text: 'Die erste Hinrichtung mit Strom – betrieben von einem Westinghouse-Generator.', options: { color: C.white, breakLine: true } },
      { text: 'Wie wird eine Hinrichtung zur Werbekampagne?', options: { italic: true, color: 'C9CDD3' } },
    ], { x: 0.75, y: 3.4, w: 8.5, h: 1.15, fontFace: BF, fontSize: 17, margin: 0, valign: 'middle', paraSpaceAfter: 6, isTextBox: true });
    s.addText('Simon Bruder · 10. Dezember 2026 · Dozent: Stephan Schwarz', { x: 0.5, y: 4.95, w: 7, h: 0.3, fontFace: BF, fontSize: 11, color: '9AA0A8', margin: 0, isTextBox: true });
    s.addNotes(`[0:00–0:45] Einstieg
Auburn, New York, 6. August 1890: William Kemmler wird als erster Mensch auf dem elektrischen Stuhl hingerichtet. Der Generator stammt von Westinghouse – beschafft hatte ihn aber das Lager um Edison, um Wechselstrom als «Todesstrom» abzustempeln.
Leitfrage: Warum setzte sich in den USA trotzdem der Wechselstrom durch – und was zeigt der Konflikt über den Aufstieg der USA zur führenden Industriemacht?
These in einem Satz: Entschieden haben nicht Propaganda oder Genies, sondern Kosten, Kapital und Netzlogik.
Quellen: ${cites('moran', 'essig')}.`);
  }

  // ---------- 2 Technik ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, 'Hohe Spannung macht Strom transportierbar');
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.1, h: 0.9, fill: { color: C.tint }, line: { color: C.tint } });
    s.addText([{ text: 'Leitungsverlust = I', options: {} }, { text: '2', options: { superscript: true } }, { text: ' · R', options: {} }],
      { x: 0.65, y: 1.3, w: 3.9, h: 0.9, fontFace: HF, fontSize: 22, bold: true, color: C.teal, margin: 0, valign: 'middle', isTextBox: true });
    s.addText([
      { text: 'Gleichstrom (Edison): ', options: { bold: true, color: C.copper } },
      { text: 'rund 110 V, um 1890 nicht transformierbar – Reichweite ein Stadtquartier.', options: { breakLine: true } },
      { text: 'Wechselstrom (Westinghouse, Tesla): ', options: { bold: true, color: C.teal } },
      { text: 'Der Transformator setzt die Spannung im Netz hoch und beim Kunden herunter.' },
    ], { x: 0.5, y: 2.4, w: 4.1, h: 1.6, fontFace: BF, fontSize: 15, color: C.ink, margin: 0, valign: 'top', paraSpaceAfter: 10, isTextBox: true });
    // Kette 10x -> 1/10 -> 1/100
    const steps = [['10 ×', 'Spannung'], ['1/10', 'Strom'], ['1/100', 'Verlust']];
    steps.forEach(([big, lab], i) => {
      const x = 5.05 + i * 1.55;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.3, w: 1.25, h: 1.45, fill: { color: C.light }, line: { color: C.light }, rectRadius: 0.08 });
      s.addText(big, { x, y: 1.4, w: 1.25, h: 0.75, fontFace: HF, fontSize: 28, bold: true, color: i === 2 ? C.teal : C.ink, align: 'center', margin: 0, valign: 'middle', isTextBox: true });
      s.addText(lab, { x, y: 2.15, w: 1.25, h: 0.45, fontFace: BF, fontSize: 13, color: C.muted, align: 'center', margin: 0, isTextBox: true });
      if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 1.31, y: 1.9, w: 0.2, h: 0.25, fill: { color: C.grayline }, line: { color: C.grayline } });
    });
    s.addText('bei gleicher übertragener Leistung', { x: 5.05, y: 2.85, w: 4.35, h: 0.3, fontFace: BF, fontSize: 11, italic: true, color: C.muted, margin: 0, isTextBox: true });
    // Reichweite
    const reach = [[C.copper, 'Pearl Street 1882', 'Versorgungsgebiet 0,65 km²'], [C.teal, 'Niagara–Buffalo 1896', 'Fernleitung über 30 km']];
    reach.forEach(([col, h, t], i) => {
      const y = 3.35 + i * 0.8;
      s.addShape(pres.shapes.OVAL, { x: 5.05, y: y + 0.12, w: 0.35, h: 0.35, fill: { color: col }, line: { color: col } });
      s.addText([{ text: h, options: { bold: true, breakLine: true } }, { text: t, options: { color: C.muted } }], { x: 5.55, y, w: 3.9, h: 0.65, fontFace: BF, fontSize: 14, color: C.ink, margin: 0, valign: 'middle', isTextBox: true });
    });
    source(s, cites('pearl', 'edpCW', 'jonnes')); pageNo(s, 2);
    s.addNotes(`[0:45–1:45] Technik
Kernidee in einem Satz: Strom lässt sich nur mit hoher Spannung billig über weite Strecken schicken.
Der Verlust in der Leitung wächst mit dem Quadrat der Stromstärke. Verzehnfacht man die Spannung, braucht man für dieselbe Leistung nur ein Zehntel des Stroms – und hat nur noch ein Hundertstel des Verlusts.
Um 1890 liess sich nur Wechselstrom mit Transformatoren hoch- und heruntersetzen. Edisons Gleichstrom brauchte deshalb viele kleine Kraftwerke mitten in der Stadt und viel Kupfer.
Beispiel: Pearl Street versorgte 1882 rund 0,65 km²; 1896 floss Wechselstrom von den Niagarafällen über 30 km nach Buffalo.`);
  }

  // ---------- 3 Akteure ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, 'Firmen und Investoren – nicht zwei Genies');
    const people = [
      ['FaLightbulb', C.copper, 'Thomas A. Edison', '1847–1931', 'Erfinder-Unternehmer (Glühlampe). Verteidigt seine Gleichstrom-Investitionen – fürchtet aber auch ehrlich die Hochspannung.'],
      ['FaBullhorn', C.copper, 'Harold P. Brown', 'Elektroingenieur', 'Wortführer der Kampagne gegen den Wechselstrom, unterstützt aus Edisons Labor.'],
      ['FaIndustry', C.teal, 'George Westinghouse', '1846–1914', 'Industrieller (Druckluftbremse). Kauft Patente zu und baut ab 1886 Wechselstromnetze.'],
      ['FaCogs', C.teal, 'Nikola Tesla', '1856–1943', 'Ingenieur, 1884 kurz bei Edison. Seine Motorpatente (1888) machen Wechselstrom fabriktauglich.'],
    ];
    for (let i = 0; i < people.length; i++) {
      const [ic, col, name, yrs, role] = people[i];
      const x = 0.5 + i * 2.3;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.1, h: 3.3, fill: { color: C.light }, line: { color: C.light } });
      await circleIcon(s, ic, x + 0.2, 1.5, 0.65, col);
      s.addText(name, { x: x + 0.2, y: 2.3, w: 1.75, h: 0.55, fontFace: BF, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: 'top', isTextBox: true });
      s.addText(yrs, { x: x + 0.2, y: 2.82, w: 1.75, h: 0.3, fontFace: BF, fontSize: 11, color: C.muted, margin: 0, isTextBox: true });
      s.addText(role, { x: x + 0.2, y: 3.15, w: 1.75, h: 1.35, fontFace: BF, fontSize: 12, color: C.ink, margin: 0, valign: 'top', isTextBox: true });
    }
    [[C.copper, 'Gleichstrom-Lager', 0.5], [C.teal, 'Wechselstrom-Lager', 2.8]].forEach(([col, t, x]) => {
      s.addShape(pres.shapes.OVAL, { x, y: 4.8, w: 0.18, h: 0.18, fill: { color: col }, line: { color: col } });
      s.addText(t, { x: x + 0.27, y: 4.74, w: 2, h: 0.3, fontFace: BF, fontSize: 11, color: C.muted, margin: 0, isTextBox: true });
    });
    source(s, cites('edpCW', 'carlson', 'hughes58')); pageNo(s, 3);
    s.addNotes(`[1:45–3:00] Akteure
Wichtig für die These: Im Kern stritten zwei Aktiengesellschaften mit ihren Investoren, nicht bloss zwei Erfinder.
Edison hatte viel Geld und Ruf in Gleichstrom-Kraftwerke gesteckt. Seine Warnungen vor Hochspannung waren aber nicht nur Taktik – er hielt sie auch für gefährlich.
Harold P. Brown war das öffentliche Gesicht der Kampagne gegen Wechselstrom; Edisons Labor half ihm im Hintergrund.
Westinghouse war Industrieller, kein Erfinder-Star: Er kaufte die nötigen Patente zu, darunter 1888 Teslas Motorpatente.
Tesla lieferte mit dem Wechselstrommotor das fehlende Stück, damit Wechselstrom auch in Fabriken nutzbar wurde.`);
  }

  // ---------- 4 Verlauf I ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, '1882–1890: Vom Wettbewerb zum Propagandakrieg');
    [[0.5, 3.4, 'Phase 1 · Wettbewerb 1882–1887', C.muted], [4.1, 5.4, 'Phase 2 · Propagandakrieg 1888–1890', C.copper]].forEach(([x, w, t, col]) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.25, w, h: 0.4, fill: { color: col }, line: { color: col }, rectRadius: 0.2 });
      s.addText(t, { x, y: 1.25, w, h: 0.4, fontFace: BF, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
    });
    s.addShape(pres.shapes.LINE, { x: 0.6, y: 2.55, w: 8.8, h: 0, line: { color: C.grayline, width: 2 } });
    const ev = [
      ['1882', C.muted, 'Pearl Street (Manhattan): Edisons Gleichstrom für rund 400 Lampen'],
      ['1886', C.muted, 'Great Barrington: erstes Wechselstromnetz der USA mit Transformatoren'],
      ['1888', C.copper, 'Broschüre «A Warning»; Brown tötet öffentlich Hunde mit Wechselstrom'],
      ['1889', C.copper, 'Elektrischer Stuhl mit Westinghouse-Generator; Pressekrieg in der North American Review'],
      ['1890', C.copper, 'Kemmler: erste Hinrichtung mit Strom – ein qualvolles Debakel'],
    ];
    ev.forEach(([y, col, t], i) => {
      const x = 0.5 + i * 1.82;
      s.addText(y, { x, y: 1.85, w: 1.65, h: 0.5, fontFace: HF, fontSize: 24, bold: true, color: col, margin: 0, isTextBox: true });
      s.addShape(pres.shapes.OVAL, { x: x + 0.02, y: 2.44, w: 0.22, h: 0.22, fill: { color: col }, line: { color: C.white, width: 2 } });
      s.addText(t, { x, y: 2.85, w: 1.65, h: 1.9, fontFace: BF, fontSize: 13, color: C.ink, margin: 0, valign: 'top', isTextBox: true });
    });
    source(s, cites('pearl', 'jonnes', 'hughes58', 'moran', 'edison89', 'west89', 'essig')); pageNo(s, 4);
    s.addNotes(`[3:00–4:15] Verlauf, Teil 1
Phase 1: Offener Wettbewerb. 1882 startet Edisons Pearl Street Station, 1886 läuft in Great Barrington das erste Wechselstromnetz der USA mit Transformatoren, im selben Jahr gründet Westinghouse seine Elektrofirma.
Phase 2: Als Wechselstrom Kunden gewinnt, eskaliert es. 1888 warnt Edisons Firma mit der Broschüre «A Warning», Harold Brown tötet öffentlich Hunde mit Wechselstrom.
1889 führt New York den elektrischen Stuhl ein – bewusst mit einem Westinghouse-Generator. Edison und Westinghouse streiten sich zudem öffentlich in Artikeln der North American Review.
1890 die Hinrichtung von Kemmler: kein «sauberer» Tod, sondern ein Debakel – zurück zum Einstieg.`);
  }

  // ---------- 5 Verlauf II ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, '1891–1896: Markt und Kapital entscheiden');
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.25, w: 9, h: 0.4, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.2 });
    s.addText('Phase 3 · Entscheidung 1891–1896', { x: 0.5, y: 1.25, w: 9, h: 0.4, fontFace: BF, fontSize: 13, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
    const cards = [
      ['1891', 'Westinghouse gerät in eine Finanzkrise; Tesla verzichtet auf seine Lizenzgebühren.'],
      ['1892', 'J. P. Morgan fusioniert Edison General Electric und Thomson-Houston zu General Electric – Edison verliert die Kontrolle, GE setzt auf Wechselstrom.'],
      ['1893', 'Westinghouse beleuchtet die Weltausstellung in Chicago – deutlich günstiger als GE.'],
      ['1896', 'Strom von den Niagarafällen erreicht Buffalo; GE und Westinghouse schliessen ein Patentabkommen.'],
    ];
    cards.forEach(([y, t], i) => {
      const x = 0.5 + (i % 2) * 4.6, yy = 1.9 + Math.floor(i / 2) * 1.6;
      const hi = y === '1892';
      s.addShape(pres.shapes.RECTANGLE, { x, y: yy, w: 4.4, h: 1.4, fill: { color: hi ? C.tint : C.light }, line: { color: hi ? C.tint : C.light } });
      s.addText(y, { x: x + 0.2, y: yy, w: 1.1, h: 1.4, fontFace: HF, fontSize: 28, bold: true, color: C.teal, margin: 0, valign: 'middle', isTextBox: true });
      s.addText(t, { x: x + 1.35, y: yy + 0.1, w: 2.9, h: 1.2, fontFace: BF, fontSize: 13, color: C.ink, margin: 0, valign: 'middle', isTextBox: true });
    });
    source(s, cites('carlson', 'edpCW', 'jonnes')); pageNo(s, 5);
    s.addNotes(`[4:15–5:30] Verlauf, Teil 2
Phase 3: Jetzt entscheiden Markt und Kapital.
1891 ist Westinghouse finanziell unter Druck; Tesla verzichtet auf seine Lizenzgebühren.
Der Wendepunkt ist 1892: Der Bankier J. P. Morgan fusioniert Edisons Firma mit Thomson-Houston zu General Electric. Edison verliert die Kontrolle – und GE setzt ebenfalls auf Wechselstrom. Damit ist der Systemstreit faktisch entschieden.
1893 beleuchtet Westinghouse die Weltausstellung in Chicago, 1896 fliesst Strom von den Niagarafällen nach Buffalo. Im selben Jahr teilen GE und Westinghouse ihre Patente – aus Rivalen wird ein Duopol.`);
  }

  // ---------- 6 Warum ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, 'Kosten und Kapital schlugen Propaganda');
    const cols = [
      ['FaCoins', C.teal, 'Kosten', 'Wechselstromnetze waren billiger zu bauen und zu betreiben; neue Versorger wählten sie zunehmend.'],
      ['FaBullhorn', C.copper, 'Propaganda scheitert', 'Cole und Chandler: «kompetitives Impression Management». Wechselstrom liess sich in der Öffentlichkeit nicht mit dem Tod verbinden.'],
      ['FaGlobeEurope', C.violet, 'Schweizer Bezug', '1891: Drehstrom über 175 km von Lauffen nach Frankfurt (AEG und MFO Oerlikon). Im selben Jahr gründet Charles E. L. Brown in Baden die BBC – heute ABB.'],
    ];
    for (let i = 0; i < cols.length; i++) {
      const [ic, col, h, t] = cols[i];
      const x = 0.5 + i * 3.05;
      await circleIcon(s, ic, x, 1.3, 0.6, col);
      s.addText(h, { x, y: 2.05, w: 2.85, h: 0.4, fontFace: BF, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(t, { x, y: 2.45, w: 2.85, h: 1.6, fontFace: BF, fontSize: 13, color: C.ink, margin: 0, valign: 'top', isTextBox: true });
    }
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.2, w: 9, h: 0.8, fill: { color: C.light }, line: { color: C.light } });
    s.addText([
      { text: 'Kein K.-o.-Sieg: ', options: { bold: true } },
      { text: 'Umformer koppelten neue Wechselstrom- an alte Gleichstromnetze; Manhattans letzte Gleichstromversorgung endete erst 2007.' },
    ], { x: 0.7, y: 4.2, w: 8.6, h: 0.8, fontFace: BF, fontSize: 13, color: C.ink, margin: 0, valign: 'middle', isTextBox: true });
    source(s, cites('edpCW', 'cole', 'steig', 'davidBunn', 'millard', 'lee')); pageNo(s, 6);
    s.addNotes(`[5:30–6:30] Warum Wechselstrom gewann
Erstens Kosten: Wenige grosse Kraftwerke und dünnere Leitungen waren billiger als viele kleine Gleichstromzentralen.
Zweitens: Die Angstkampagne scheiterte. Cole und Chandler (2019) beschreiben sie als «kompetitives Impression Management» – Edisons Lager wollte das Bild des Konkurrenzprodukts beschädigen, schaffte es aber nicht, Wechselstrom in der Öffentlichkeit dauerhaft mit dem Tod zu verknüpfen.
Drittens Europa und die Schweiz: 1891 übertragen AEG und die Maschinenfabrik Oerlikon Drehstrom über 175 km. Der Oerlikon-Ingenieur Charles Brown gründet im selben Jahr in Baden die BBC – heute ABB.
Und: Es war kein K.-o.-Sieg. Umformer verbanden alte und neue Netze; in Manhattan lief Gleichstrom bis 2007.
Deutung: Paul David nennt das Pfadabhängigkeit – ein Standard setzt sich in einem kurzen, offenen Zeitfenster durch und verfestigt sich dann.`);
  }

  // ---------- 7 Weltindustrie ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, 'Die USA werden zur grössten Industriemacht');
    const labels = ['1860', '1880', '1900', '1913'];
    s.addChart(pres.charts.LINE, [
      { name: 'USA', labels, values: [7.2, 14.7, 23.6, 32.0] },
      { name: 'Grossbritannien', labels, values: [19.9, 22.9, 18.5, 13.6] },
      { name: 'Deutschland', labels, values: [4.9, 8.5, 13.2, 14.8] },
    ], {
      x: 0.4, y: 1.2, w: 6.1, h: 3.85,
      chartColors: [C.teal, C.copper, C.violet], lineSize: 3, lineDataSymbol: 'circle', lineDataSymbolSize: 8,
      showLegend: true, legendPos: 't', legendFontFace: BF, legendFontSize: 12, legendColor: C.ink,
      showTitle: true, title: 'Anteil an der Weltindustrieproduktion in %', titleFontFace: BF, titleFontSize: 13, titleColor: C.muted,
      valAxisMinVal: 0, valAxisMaxVal: 35, valAxisMajorUnit: 10, valAxisLabelFormatCode: '0" %"',
      valAxisLabelColor: C.muted, catAxisLabelColor: C.muted, valAxisLabelFontFace: BF, catAxisLabelFontFace: BF,
      valAxisLabelFontSize: 11, catAxisLabelFontSize: 11,
      valGridLine: { color: 'E6E4DF', size: 0.75 }, catGridLine: { style: 'none' },
    });
    s.addText('1913 – zwei Jahrzehnte nach dem Stromkrieg', { x: 6.9, y: 1.4, w: 2.6, h: 0.4, fontFace: BF, fontSize: 14, color: C.muted, margin: 0, isTextBox: true });
    s.addText('32 %', { x: 6.9, y: 1.8, w: 2.6, h: 0.95, fontFace: HF, fontSize: 54, bold: true, color: C.teal, margin: 0, isTextBox: true });
    s.addText('der Weltindustrieproduktion stammen aus den USA – mehr als aus Grossbritannien und Deutschland zusammen (28,4 %).', { x: 6.9, y: 2.8, w: 2.6, h: 1.4, fontFace: BF, fontSize: 14, color: C.ink, margin: 0, valign: 'top', isTextBox: true });
    s.addText('Stichjahre; Abstände auf der Zeitachse nicht massstäblich.', { x: 6.9, y: 4.4, w: 2.6, h: 0.5, fontFace: BF, fontSize: 10, italic: true, color: C.muted, margin: 0, isTextBox: true });
    source(s, cites('hughes91', 'bairoch')); pageNo(s, 7);
    s.addNotes(`[6:30–7:30] Bedeutung für die USA (1)
Hier der Bezug zum Modul «Weltmacht USA». Die Daten stammen von Paul Bairoch.
1880 lag Grossbritannien noch klar vorne. Zwischen 1880 und 1900 – also genau im Jahrzehnt des Stromkriegs – ziehen die USA vorbei. 1913 stellen sie knapp ein Drittel der Weltindustrieproduktion, mehr als Grossbritannien und Deutschland zusammen.
Vorsicht bei der Deutung: Der Stromkrieg hat das nicht allein verursacht. Aber billiger, übertragbarer Strom wurde zu einer Basistechnologie dieses Aufstiegs (Hughes).`);
  }

  // ---------- 8 Fabrik und Konzern ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, 'Strom baut die Fabrik und den Grosskonzern um');
    const stats = [
      ['FaIndustry', '5 % → 78 %', 'Anteil der Elektromotoren an der Antriebsleistung der US-Industrie, 1899 → 1929', cites('devine')],
      ['FaChartLine', '1,3 % → 3,1 %', 'Wachstum des Outputs pro Arbeitsstunde pro Jahr, vor → nach 1919 – sobald die Fabriken um den Motor herum neu organisiert waren', cites('david90')],
      ['FaLandmark', '1896', 'GE gehört zu den zwölf ersten Werten des Dow Jones; mit Westinghouse bildet es ein Duopol mit Patentpool', ''],
    ];
    for (let i = 0; i < stats.length; i++) {
      const [ic, big, t, src] = stats[i];
      const x = 0.5 + i * 3.05;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.85, h: 3.6, fill: { color: C.light }, line: { color: C.light } });
      await circleIcon(s, ic, x + 0.25, 1.5, 0.55, C.teal);
      s.addText(big, { x: x + 0.25, y: 2.2, w: 2.45, h: 0.75, fontFace: HF, fontSize: 26, bold: true, color: C.teal, margin: 0, valign: 'middle', isTextBox: true });
      s.addText(t, { x: x + 0.25, y: 3.0, w: 2.45, h: 1.45, fontFace: BF, fontSize: 13, color: C.ink, margin: 0, valign: 'top', isTextBox: true });
      s.addText(src, { x: x + 0.25, y: 4.5, w: 2.45, h: 0.3, fontFace: BF, fontSize: 10, color: C.muted, margin: 0, isTextBox: true });
    }
    source(s, cites('devine', 'david90', 'hughes83')); pageNo(s, 8);
    s.addNotes(`[7:30–8:30] Bedeutung für die USA (2)
Fabriken: 1899 lieferten Elektromotoren knapp 5 Prozent der Antriebsleistung der US-Industrie, 1929 rund 78 Prozent (Devine).
Spannend ist die Verzögerung: Die Produktivität wuchs vor 1919 um 1,3 Prozent pro Jahr, danach um 3,1 Prozent. Paul David erklärt das so: Erst als Fabriken nicht mehr um eine zentrale Dampfmaschine, sondern um viele Elektromotoren herum gebaut wurden, zahlte sich die neue Technik aus.
Big Business: GE war 1896 einer der zwölf ersten Werte im Dow Jones. Mit Westinghouse bildete es ein Duopol mit Patentpool – typisch für das Gilded Age. An den Niagarafällen zog billige Wasserkraft zudem Aluminium- und Chemieindustrie an (Hughes, Networks of Power).`);
  }

  // ---------- 9 Mythen ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, 'Drei Mythen – und was die Quellen sagen');
    s.addText('Mythos', { x: 1.1, y: 1.2, w: 3.6, h: 0.35, fontFace: BF, fontSize: 13, bold: true, color: C.copper, margin: 0, isTextBox: true });
    s.addText('Befund', { x: 5.8, y: 1.2, w: 3.7, h: 0.35, fontFace: BF, fontSize: 13, bold: true, color: C.teal, margin: 0, isTextBox: true });
    const rows = [
      ['Edison liess 1903 den Elefanten Topsy töten, um Wechselstrom zu diskreditieren (so etwa ABB).', 'Die Tötung geschah über zehn Jahre nach dem Stromkrieg – ohne Edisons Beteiligung.', cites('abb', 'topsy')],
      ['Der Stromkrieg war ein Duell «Tesla gegen Edison» (so etwa der Film The Current War).', 'Gestritten haben vor allem Unternehmen und ihre Investoren.', cites('edpCW')],
      ['Tesla zerriss aus Grossmut seinen Vertrag mit Westinghouse.', 'Belegt sind nur der Verzicht auf Lizenzgebühren (1891) und der Verkauf der Patente für 216 000 Dollar.', cites('carlson')],
    ];
    for (let i = 0; i < rows.length; i++) {
      const [m, b, src] = rows[i];
      const y = 1.6 + i * 1.18;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9, h: 1.05, fill: { color: C.light }, line: { color: C.light } });
      await circleIcon(s, 'FaTimes', 0.62, y + 0.3, 0.42, C.copper);
      s.addText(m, { x: 1.1, y, w: 3.75, h: 1.05, fontFace: BF, fontSize: 13, color: C.ink, margin: 0, valign: 'middle', isTextBox: true });
      s.addShape(pres.shapes.RIGHT_ARROW, { x: 4.95, y: y + 0.4, w: 0.3, h: 0.25, fill: { color: C.grayline }, line: { color: C.grayline } });
      await circleIcon(s, 'FaCheck', 5.32, y + 0.3, 0.42, C.teal);
      s.addText([{ text: b + ' ' }, { text: src, options: { color: C.muted, fontSize: 11 } }], { x: 5.85, y, w: 3.55, h: 1.05, fontFace: BF, fontSize: 13, color: C.ink, margin: 0, valign: 'middle', isTextBox: true });
    }
    source(s, cites('abb', 'topsy', 'edpCW', 'carlson')); pageNo(s, 9);
    s.addNotes(`[8:30–9:30] Mythen und Quellenkritik
Zum Abschluss die Quellenkritik – das ist mir wichtig, weil viele populäre Darstellungen hier danebenliegen.
Erstens Topsy: Ein ABB-Beitrag behauptet, Edison habe zur Abschreckung einen Elefanten töten lassen. Die Edison Papers der Rutgers University zeigen: Das war 1903, über zehn Jahre nach dem Stromkrieg, und Edison war nicht beteiligt.
Zweitens «Tesla gegen Edison»: Filme machen daraus ein Duell der Genies; tatsächlich stritten Firmen und Investoren.
Drittens der zerrissene Vertrag: Dafür gibt es keinen Beleg – belegt ist nur der Verzicht auf Lizenzgebühren und der spätere Patentverkauf (Carlson).`);
  }

  // ---------- 10 Fazit ----------
  {
    const s = pres.addSlide(); s.background = { color: C.dark };
    title(s, 'Fazit: der erste grosse Standardkrieg', C.white);
    const pts = [
      ['1', 'Gewonnen hat das System, das sich billiger skalieren liess – nicht die bessere Propaganda.'],
      ['2', 'Mitgewonnen hat der finanzstarke Grosskonzern: GE und Westinghouse als Muster des US-Industriekapitalismus.'],
      ['3', 'Ironie: Mit Stromrichtern kehrt der Gleichstrom zurück – 1954 als erste kommerzielle Hochspannungs-Gleichstrom-Übertragung (Gotland), gebaut von ASEA, heute ABB.'],
    ];
    pts.forEach(([num, t], i) => {
      const y = 1.35 + i * 1.12;
      s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.12, w: 0.6, h: 0.6, fill: { color: C.amber }, line: { color: C.amber } });
      s.addText(num, { x: 0.5, y: y + 0.12, w: 0.6, h: 0.6, fontFace: HF, fontSize: 20, bold: true, color: C.dark, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
      s.addText(t, { x: 1.35, y, w: 8.1, h: 0.85, fontFace: BF, fontSize: 17, color: C.white, margin: 0, valign: 'middle', isTextBox: true });
    });
    s.addText('Danke – Fragen?', { x: 0.5, y: 4.65, w: 9, h: 0.45, fontFace: HF, fontSize: 20, bold: true, color: C.amber, margin: 0, isTextBox: true });
    source(s, cites('mcnichol', 'gotland'), '9AA0A8'); pageNo(s, 10, '9AA0A8');
    s.addNotes(`[9:30–10:00] Fazit
Zurück zur Leitfrage: Wechselstrom gewann, weil er sich billiger skalieren liess und weil das grosse Kapital – J. P. Morgan und GE – am Ende darauf setzte. Die Angstkampagne hat daran nichts geändert.
Für die «Weltmacht USA» heisst das: Der Stromkrieg steht am Anfang eines Musters – Standardkriege, die mit Kapital entschieden werden, und Grosskonzerne, die danach den Markt aufteilen. McNichol nennt ihn deshalb den ersten Standardkrieg.
Und die Pointe: Heute kehrt der Gleichstrom für lange Strecken zurück – mit Hochspannungs-Gleichstrom-Übertragung, zuerst 1954 auf Gotland, gebaut von ASEA, heute wie die BBC aus Baden Teil von ABB.
Danke – Fragen?`);
  }

  // ---------- Anhang: Quellen (Auswahl) ----------
  {
    const s = pres.addSlide(); s.background = { color: C.white };
    title(s, 'Quellen (Auswahl, IEEE)');
    const pick = ['edpCW', 'jonnes', 'carlson', 'hughes58', 'cole', 'david92', 'bairoch', 'devine', 'david90', 'topsy'].sort((a, b) => n(a) - n(b));
    const parts = [];
    pick.forEach((k, i) => {
      const txt = REFS[k].replace(/\[\[|\]\]/g, '').replace(/"([^"]*)"/g, '“$1”');
      parts.push({ text: `[${n(k)}]  `, options: { bold: true } });
      const segs = txt.split(/(\*[^*]+\*)/g).filter(Boolean);
      segs.forEach((seg, j) => {
        const last = j === segs.length - 1 && i < pick.length - 1;
        parts.push({ text: seg.startsWith('*') ? seg.slice(1, -1) : seg, options: { italic: seg.startsWith('*'), breakLine: last } });
      });
    });
    s.addText(parts, { x: 0.5, y: 1.15, w: 9, h: 3.9, fontFace: BF, fontSize: 10.5, color: C.ink, margin: 0, valign: 'top', paraSpaceAfter: 5, isTextBox: true });
    s.addText('Vollständiges Verzeichnis (30 Einträge) im Handout.', { x: 0.5, y: 5.2, w: 7, h: 0.25, fontFace: BF, fontSize: 9, color: C.muted, margin: 0, isTextBox: true });
    s.addNotes('Anhang – nur bei Nachfragen zeigen. Die Nummern entsprechen dem Literaturverzeichnis im Handout.');
  }

  await pres.writeFile({ fileName: OUT });
  console.log('OK', OUT);
})();
