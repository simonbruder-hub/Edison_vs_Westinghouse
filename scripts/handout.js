// Erzeugt abgabe/Handout_Stromkrieg_Bruder.docx (A4, Arial 11, IEEE-Zitierweise).
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Footer,
  AlignmentType, WidthType, ShadingType, BorderStyle, LevelFormat, PageNumber, TabStopType,
  HeadingLevel,
} = require('docx');
const REFS = require('./quellen');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'abgabe', 'Ausarbeitung_Stromkrieg_lang.docx');
const FONT = 'Arial';
const INK = '1E1E1E', MUTED = '5F5D58', TEAL = '0E7C9E', COPPER = 'C8641E', TINT = 'EEF5F8', HEAD = 'DCEBF1';
const PAGE_W = 11906, MARGIN = 1134, CONTENT_W = PAGE_W - 2 * MARGIN; // 9638 DXA

// ---------- IEEE-Zitate: Nummer nach erster Nennung ----------
const order = [];
function cite(keys) {
  const nums = keys.map(k => {
    if (!REFS[k]) throw new Error('Unbekannte Quelle: ' + k);
    if (!order.includes(k)) order.push(k);
    return order.indexOf(k) + 1;
  }).sort((a, b) => a - b);
  const parts = [];
  for (let i = 0; i < nums.length;) {
    let j = i;
    while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++;
    if (j - i >= 2) parts.push(`[${nums[i]}]–[${nums[j]}]`);
    else for (let k = i; k <= j; k++) parts.push(`[${nums[k]}]`);
    i = j + 1;
  }
  return parts.join(', ');
}

// Markup: {key,key} = Zitat, **fett**, *kursiv*, [[Platzhalter]] = gelb
function runs(text, base = {}) {
  text = text.replace(/\{([\w,]+)\}/g, (_, k) => cite(k.split(',')));
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|\[\[[^\]]+\]\])/g;
  let last = 0, m;
  const push = (t, o = {}) => t && out.push(new TextRun({ text: t, font: FONT, color: INK, ...base, ...o }));
  while ((m = re.exec(text))) {
    push(text.slice(last, m.index));
    const t = m[0];
    if (t.startsWith('**')) push(t.slice(2, -2), { bold: true });
    else if (t.startsWith('[[')) push('[' + t.slice(2, -2) + ']', { highlight: 'yellow' });
    else push(t.slice(1, -1), { italics: true });
    last = m.index + t.length;
  }
  push(text.slice(last));
  return out;
}

const P = (text, opts = {}) => new Paragraph({ children: runs(text, opts.run), spacing: { after: 100, line: 264 }, ...opts.para });
const H = text => new Paragraph({
  heading: HeadingLevel.HEADING_1, keepNext: true,
  spacing: { before: 220, after: 80 },
  children: [new TextRun({ text, font: FONT, bold: true, size: 26, color: TEAL })],
});
const B = text => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, children: runs(text), spacing: { after: 70, line: 259 } });

const border = { style: BorderStyle.SINGLE, size: 4, color: 'C9D3D8' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, header, rows, size = 19) {
  const cell = (t, w, isHead) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders,
    shading: isHead ? { type: ShadingType.CLEAR, color: 'auto', fill: HEAD } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: [new Paragraph({ children: runs(t, { size, bold: isHead || undefined }), spacing: { after: 0, line: 245 } })],
  });
  return new Table({
    width: { size: widths.reduce((a, b) => a + b), type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], true)) }),
      ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], false)) })),
    ],
  });
}
const caption = t => new Paragraph({ children: runs(t, { size: 17, color: MUTED }), spacing: { before: 60, after: 120 } });

// ---------- Inhalt (Reihenfolge = Zitierreihenfolge) ----------
const body = [];

body.push(new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: 'Der Stromkrieg (War of Currents)', font: FONT, bold: true, size: 36, color: INK })] }));
body.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'Edison vs. Westinghouse – der Kampf um Amerikas Stromnetz und die Rolle von Nikola Tesla', font: FONT, size: 22, color: TEAL })] }));
body.push(new Paragraph({
  spacing: { after: 160 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'C9D3D8', space: 6 } },
  children: [new TextRun({ text: 'Handout zum Referat · Modul «Weltmacht USA» (FHNW) · Dozent: Stephan Schwarz · Simon Bruder · 10.12.2026', font: FONT, size: 18, color: MUTED })],
}));

// Leitfrage und These als getönter Kasten
const boxCell = new TableCell({
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
  shading: { type: ShadingType.CLEAR, color: 'auto', fill: TINT },
  margins: { top: 120, bottom: 80, left: 160, right: 160 },
  children: [
    P('**Leitfrage:** Warum setzte sich in den USA der Wechselstrom gegen Edisons Gleichstrom durch – und was zeigt dieser Konflikt über den Aufstieg der USA zur führenden Industriemacht?'),
    P('**These:** Den «Stromkrieg» (etwa 1886–1892) entschieden nicht Propaganda oder ein einzelnes Genie, sondern Kosten, Kapital und Netzlogik. Die anschliessende Konsolidierung zu Grosskonzernen wie General Electric wurde zum Muster des amerikanischen Industriekapitalismus.'),
  ],
});
body.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W], rows: [new TableRow({ children: [boxCell] })] }));

body.push(H('1  Der technische Kern: Warum die Spannung entscheidet'));
body.push(P('Strom lässt sich nur mit hoher Spannung billig über weite Strecken schicken – und die liess sich um 1890 nur beim Wechselstrom mit Transformatoren hoch- und heruntersetzen. Denn der Leitungsverlust wächst mit dem Quadrat der Stromstärke (Verlust = *I*² · *R*): Zehnfache Spannung bedeutet bei gleicher Leistung ein Zehntel des Stroms und ein Hundertstel des Verlusts.'));
body.push(table([1700, 3969, 3969], ['', 'Gleichstrom (Edison)', 'Wechselstrom (Westinghouse, Tesla)'], [
  ['**Spannung**', 'rund 110 V, nicht transformierbar', 'transformierbar: hoch im Netz, tief beim Kunden'],
  ['**Reichweite**', 'ein Stadtquartier, viele kleine Zentralen (Pearl Street 1882: 0,65 km²)', 'Dutzende Kilometer, wenige grosse Kraftwerke (Niagara–Buffalo 1896: über 30 km)'],
  ['**Schwachpunkt**', 'hoher Kupferbedarf', 'Hochspannung gefährlich; brauchbarer Motor erst ab 1888'],
]));
body.push(caption('Quellen: {pearl,edpCW,jonnes}.'));

body.push(H('2  Akteure und Interessen'));
body.push(P('Im Kern stritten zwei Aktiengesellschaften mit ihren Investoren – nicht bloss zwei Erfinder {edpCW}.'));
body.push(table([3000, 6638], ['Akteur', 'Rolle und Interesse'], [
  ['**Thomas A. Edison** (1847–1931)', 'Erfinder-Unternehmer (Glühlampe, Pearl Street 1882); verteidigt seine Gleichstrom-Investitionen, fürchtet aber auch ehrlich die Hochspannung'],
  ['**George Westinghouse** (1846–1914)', 'Industrieller (Druckluftbremse); kauft Patente zu und baut ab 1886 Wechselstromnetze'],
  ['**Nikola Tesla** (1856–1943)', 'Ingenieur, 1884 kurz bei Edison; seine Motorpatente (1888) machen Wechselstrom fabriktauglich {carlson}'],
  ['**Harold P. Brown**', 'Wortführer der Kampagne gegen den Wechselstrom, unterstützt aus Edisons Labor {hughes58}'],
]));

body.push(H('3  Verlauf in drei Phasen (1882–1896)'));
body.push(P('Auf einen offenen Wettbewerb (1882–1887) folgte ein Propagandakrieg (1888–1890); entschieden wurde der Konflikt danach durch Markt und Kapital (1891–1896).'));
body.push(table([1500, 8138], ['Datum', 'Ereignis'], [
  ['4.9.1882', 'Edisons Pearl Street Station in Manhattan liefert Gleichstrom für rund 400 Lampen {pearl}.'],
  ['1886', 'Gründung von Westinghouse Electric; in Great Barrington läuft das erste Wechselstromnetz der USA mit Transformatoren {jonnes}.'],
  ['1888', 'Edisons Firma warnt mit der Broschüre «A Warning»; Harold P. Brown tötet öffentlich Hunde mit Wechselstrom {hughes58}. Westinghouse erwirbt Teslas Motorpatente {carlson}.'],
  ['1889', 'New York führt den elektrischen Stuhl ein – mit einem Westinghouse-Generator, beschafft vom Edison-Lager {moran}. Pressekrieg in der *North American Review* {edison89,west89}.'],
  ['6.8.1890', 'Erste Hinrichtung auf dem elektrischen Stuhl (William Kemmler) – ein qualvolles Debakel {essig}.'],
  ['1891', 'Westinghouse in der Finanzkrise; Tesla verzichtet auf seine Lizenzgebühren {carlson}.'],
  ['15.4.1892', 'Unter J. P. Morgan fusionieren Edison General Electric und Thomson-Houston zu General Electric (GE); Edison verliert die Kontrolle {edpCW}.'],
  ['1.5.1893', 'Westinghouse beleuchtet die Weltausstellung in Chicago, deutlich günstiger als GE {jonnes}.'],
  ['16.11.1896', 'Strom von den Niagarafällen erreicht Buffalo; im selben Jahr schliessen GE und Westinghouse ein Patentabkommen {jonnes,carlson}.'],
]));

body.push(H('4  Warum gewann der Wechselstrom?'));
body.push(P('Den Ausschlag gaben Kosten und Kapital: Nach 1892 setzte selbst Edisons frühere Firma, nun GE, auf Wechselstrom.'));
body.push(B('**Kosten:** Wechselstromnetze waren billiger zu bauen und zu betreiben; neue Versorger wählten sie zunehmend {edpCW}.'));
body.push(B('**Propaganda scheitert:** Cole und Chandler deuten Edisons Kampagne als «kompetitives Impression Management» – als Versuch, das Bild des Konkurrenzprodukts beim gemeinsamen Publikum zu beschädigen. Trotz Tierversuchen und Flugschriften gelang es nicht, Wechselstrom in der öffentlichen Wahrnehmung mit dem Tod zu verbinden {cole}.'));
body.push(B('**Europa:** 1891 übertrugen AEG und Maschinenfabrik Oerlikon Drehstrom über 175 km (Lauffen–Frankfurt); der Oerlikon-Ingenieur Charles E. L. Brown gründete im selben Jahr in Baden die BBC, heute ABB {steig}.'));
body.push(B('**Kein K.-o.-Sieg:** Umformer koppelten neue Wechselstrom- an alte Gleichstromnetze {davidBunn,millard}; Manhattans letzte Gleichstromversorgung endete erst 2007 {lee}.'));
body.push(P('**Deutung:** Paul David liest den Fall als Lehrstück der Pfadabhängigkeit: Standards entscheiden sich in kurzen, offenen Momenten und verfestigen sich danach {david92}. Edison handelte aus gemischten Motiven – geschäftlich, aber auch aus Sorge vor Hochspannung {israel}.'));

body.push(H('5  Bedeutung für den Aufstieg der USA'));
body.push(P('Der Stromkrieg fällt in die Jahrzehnte, in denen die USA Grossbritannien als grösste Industriemacht überholten; billiger, übertragbarer Strom wurde zu einer Basistechnologie dieses Aufstiegs {hughes91}. 1913 stellten die USA knapp ein Drittel der Weltindustrieproduktion – mehr als Grossbritannien und Deutschland zusammen {bairoch}.'));
const png = fs.readFileSync(path.join(ROOT, 'assets', 'chart_weltindustrie.png'));
body.push(new Paragraph({
  keepNext: true, spacing: { before: 60, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new ImageRun({ type: 'png', data: png, transformation: { width: 480, height: 226 }, altText: { title: 'Anteil an der Weltindustrieproduktion', description: 'Liniendiagramm 1860–1913: USA steigen von 7,2 auf 32,0 Prozent, Grossbritannien fällt von 22,9 (1880) auf 13,6 Prozent, Deutschland steigt auf 14,8 Prozent.', name: 'chart' } })],
}));
body.push(caption('Abb. 1: Anteil an der Weltindustrieproduktion in Prozent, Stichjahre 1860, 1880, 1900 und 1913 (Linien verbinden die Stichjahre). Daten: Bairoch {bairoch}.'));
body.push(B('**Fabriken:** Elektromotoren lieferten 1899 knapp 5 %, 1929 rund 78 % der Antriebsleistung der US-Industrie {devine}. Der Output pro Arbeitsstunde wuchs vor 1919 um 1,3 %, danach um 3,1 % pro Jahr – sobald die Fabriken um den Motor herum neu organisiert waren {david90}.'));
body.push(B('**Big Business:** GE zählte 1896 zu den zwölf ersten Werten des Dow Jones; mit Westinghouse bildete es ein Duopol mit Patentpool – typisch für das «Gilded Age».'));
body.push(B('**Netze:** Grosskraftwerke und Fernleitungen begründeten das heutige Verbundnetz {edpCW}; an den Niagarafällen zog billige Wasserkraft Aluminium- und Elektrochemie-Industrie an {hughes83}.'));

body.push(H('6  Mythen und Quellenkritik'));
body.push(P('Populäre Darstellungen erzählen ein Duell zweier Genies (z. B. {tagi,abb,spiegel}); die Forschung ist nüchterner.'));
body.push(B('**«Tesla gegen Edison»:** Filme wie *The Current War* (2017/2019) spitzen den Konflikt auf die Erfinder zu; gestritten haben vor allem Unternehmen und Investoren {edpCW}.'));
body.push(B('**Elefant Topsy (1903):** Die Tötung gilt oft als Höhepunkt von Edisons Kampagne – auch bei ABB {abb}. Sie geschah aber über zehn Jahre nach dem Stromkrieg und ohne Edisons Beteiligung {topsy}.'));
body.push(B('**Der zerrissene Vertrag:** Belegt sind nur Teslas Verzicht auf Lizenzgebühren (1891) und der spätere Verkauf der Patente für 216 000 Dollar {carlson}.'));
body.push(P('Zum Einstieg eignen sich die Dokumentationen {pbsEdison,pbsTesla,terraX}; auch sie sind populäre Formate und quellenkritisch zu lesen.'));

body.push(H('7  Fazit und Ausblick'));
body.push(P('Der Stromkrieg war der erste grosse Standardkrieg der Industriegeschichte {mcnichol}. Gewonnen hat das System, das sich billiger skalieren liess – und mit ihm der finanzstarke Grosskonzern. Darin liegt seine Bedeutung für die «Weltmacht USA».'));
body.push(P('**Ironie der Geschichte:** Dem Gleichstrom fehlte nur die Technik, seine Spannung zu wandeln. Mit Stromrichtern kehrt er zurück – die erste kommerzielle Hochspannungs-Gleichstrom-Übertragung (Gotland) ging 1954 in Betrieb, gebaut von ASEA, heute wie die BBC Teil von ABB {gotland}.'));

// ---------- Literaturverzeichnis ----------
const refs = [new Paragraph({
  heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { after: 140 },
  children: [new TextRun({ text: 'Literatur- und Quellenverzeichnis (IEEE)', font: FONT, bold: true, size: 26, color: TEAL })],
})];
order.forEach((k, i) => refs.push(new Paragraph({
  tabStops: [{ type: TabStopType.LEFT, position: 567 }],
  indent: { left: 567, hanging: 567 }, spacing: { after: 90, line: 252 },
  children: [new TextRun({ text: `[${i + 1}]\t`, font: FONT, size: 19, color: INK }), ...runs(REFS[k].replace(/"([^"]*)"/g, '\u201C$1\u201D'), { size: 19 })],
})));
const unused = Object.keys(REFS).filter(k => !order.includes(k));
if (unused.length) throw new Error('Nicht zitiert: ' + unused.join(', '));

const doc = new Document({
  creator: 'Simon Bruder', title: 'Der Stromkrieg (War of Currents) – Handout',
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: 'Simon Bruder · Der Stromkrieg · Seite ', font: FONT, size: 16, color: MUTED }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: MUTED }),
    ] })] }) },
    children: [...body, ...refs],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  // Zitierreihenfolge für die Folien (gleiche Nummern wie im Handout)
  fs.writeFileSync(path.join(ROOT, 'assets', 'zitierreihenfolge.json'), JSON.stringify(order, null, 2));
  console.log('OK', OUT, order.length, 'Quellen');
});
