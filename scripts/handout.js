// Erzeugt abgabe/Ausarbeitung_Stromkrieg_lang.docx (A4, Arial 11, IEEE-Zitierweise).
// Die lange Fassung hat eine eigene Nummerierung; Handout und Folien richten sich nach latex/handout/main.tex.
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Footer,
  AlignmentType, WidthType, ShadingType, BorderStyle, LevelFormat, PageNumber, TabStopType,
  HeadingLevel,
} = require('docx');
const REFS = require('./quellen.json');

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
    else if (t.startsWith('[[')) push('[' + t.slice(2, -2) + ']', { shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'FFF200' } });
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
  children: [new TextRun({ text: 'Ausführliche Ausarbeitung zum Referat · Modul «Weltmacht USA» (FHNW) · Dozent: Stephan Schwarz · Simon Bruder · 10.12.2026', font: FONT, size: 18, color: MUTED })],
}));

// Leitfrage und These als getönter Kasten (Wortlaut wie im Handout und auf Folie 4)
const boxCell = new TableCell({
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' } },
  shading: { type: ShadingType.CLEAR, color: 'auto', fill: TINT },
  margins: { top: 120, bottom: 80, left: 160, right: 160 },
  children: [
    P('**Leitfrage:** Warum setzte sich in den USA der Wechselstrom gegen Edisons Gleichstrom durch – und was zeigt der Konflikt über den Aufstieg der USA zur führenden Industriemacht?'),
    P('**These:** Entschieden haben nicht Propaganda oder ein einzelnes Genie, sondern Kosten, Kapital und Technik. Aus dem Streit gingen Grosskonzerne hervor, die typisch für den amerikanischen Industriekapitalismus wurden.'),
  ],
});
body.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W], rows: [new TableRow({ children: [boxCell] })] }));
body.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
body.push(P('**Zur Datierung:** Der «Stromkrieg» (englisch «War of the Currents») fällt in die späten 1880er- und frühen 1890er-Jahre {edpCW}. Diese Arbeit rechnet von 1886 (erste Wechselstromnetze) bis 1892 (Gründung von General Electric); Vorgeschichte und Folgen reichen von 1882 bis 1896.'));

body.push(H('1  Der technische Kern: Warum die Spannung entscheidet'));
body.push(P('Strom lässt sich nur mit hoher Spannung billig über weite Strecken schicken – und die liess sich um 1890 nur beim Wechselstrom mit Transformatoren hoch- und heruntersetzen. Denn der Leitungsverlust wächst mit dem Quadrat der Stromstärke (Verlust = *I*² · *R*): Zehnfache Spannung bedeutet bei gleicher Leistung ein Zehntel des Stroms und ein Hundertstel des Verlusts.'));
body.push(table([1700, 3969, 3969], ['', 'Gleichstrom (Edison)', 'Wechselstrom (Westinghouse, Thomson-Houston)'], [
  ['**Spannung**', 'rund 110 V, damals nicht transformierbar', 'transformierbar: hoch im Netz, tief beim Kunden'],
  ['**Reichweite**', 'ein Stadtquartier, viele kleine Zentralen (Pearl Street 1882: 0,65 km²)', 'Dutzende Kilometer, wenige grosse Kraftwerke (Niagara–Buffalo 1896: rund 40 km)'],
  ['**Schwachpunkt**', 'hoher Kupferbedarf', 'Hochspannung gefährlich; brauchbarer Motor erst ab 1888'],
]));
body.push(caption('Tab. 1: Die beiden Systeme im Vergleich. Quellen: {pearl,jonnes}.'));

body.push(H('2  Akteure und Interessen'));
body.push(P('Im Kern stritten drei Elektrokonzerne und ihre Geldgeber – nicht bloss zwei Erfinder {jonnes}.'));
body.push(table([3000, 6638], ['Akteur', 'Rolle und Interesse'], [
  ['**Thomas A. Edison** (1847–1931)', 'Erfinder-Unternehmer (Glühlampe, Pearl Street 1882); verteidigt seine Gleichstrom-Investitionen, fürchtet aber auch ehrlich die Hochspannung {israel}'],
  ['**George Westinghouse** (1846–1914)', 'Industrieller (Druckluftbremse); kauft Patente zu und baut ab 1886 Wechselstromnetze {jonnes}'],
  ['**Nikola Tesla** (1856–1943)', 'Ingenieur, 1884 kurz bei Edison; seine Motorpatente (1888) machen Wechselstrom fabriktauglich {carlson}'],
  ['**Charles A. Coffin** (1844–1926)', 'Chef von Thomson-Houston, dem dritten grossen Anbieter; wird 1892 erster Präsident von General Electric {jonnes}'],
  ['**J. P. Morgan** (1837–1913)', 'Bankier und früher Geldgeber Edisons; seine Bank fädelt 1892 die Fusion zu General Electric mit ein {jonnes}'],
  ['**Harold P. Brown**', 'Elektroingenieur und Wortführer der Kampagne gegen den Wechselstrom, unterstützt aus Edisons Labor {hughes58}; nicht verwandt mit Charles E. L. Brown, dem Gründer der BBC'],
]));
body.push(caption('Tab. 2: Akteure.'));

body.push(H('3  Verlauf in drei Phasen (1882–1896)'));
body.push(P('Auf einen offenen Wettbewerb (1882–1887) folgte eine Angstkampagne (1888–1890); entschieden wurde der Konflikt danach durch Markt und Kapital (1891–1896).'));
body.push(table([1500, 8138], ['Datum', 'Ereignis'], [
  ['4.9.1882', 'Edison eröffnet in New York sein erstes dauerhaftes kommerzielles Zentralkraftwerk (Pearl Street): Gleichstrom für anfangs rund 85 Kunden und 400 Lampen {pearl}.'],
  ['1886', 'Gründung von Westinghouse Electric; in Great Barrington läuft das erste Wechselstromnetz der USA mit Transformatoren {jonnes}.'],
  ['1888', 'Westinghouse erwirbt Teslas Motorpatente {carlson}. Die Edison Electric Light Company warnt mit der Broschüre «A Warning» vor dem Wechselstrom {jonnes}; Harold P. Brown tötet öffentlich Hunde mit Wechselstrom {hughes58}.'],
  ['1889', 'Im Staat New York tritt das Gesetz über die Hinrichtung durch Strom in Kraft. Brown beschafft dafür über Dritte gebrauchte Westinghouse-Generatoren – gegen den Willen von Westinghouse {hughes58,moran}. Pressekrieg in der *North American Review* {edison89,west89}.'],
  ['6.8.1890', 'Erste Hinrichtung auf dem elektrischen Stuhl (William Kemmler) – ein qualvolles Debakel {essig}.'],
  ['1891', 'Westinghouse gerät in Finanznot; Tesla verzichtet auf seine Lizenzgebühren {carlson}.'],
  ['15.4.1892', 'Edison General Electric und der Rivale Thomson-Houston fusionieren zu General Electric (GE), eingefädelt von den Banken Drexel, Morgan & Co. und Lee, Higginson & Co. Präsident wird Charles A. Coffin von Thomson-Houston; Edisons Name verschwindet {jonnes}.'],
  ['1.5.1893', 'Westinghouse unterbietet GE und beleuchtet die Weltausstellung in Chicago {jonnes}.'],
  ['16.11.1896', 'Strom von den Niagarafällen erreicht das rund 40 km entfernte Buffalo; im selben Jahr legen GE und Westinghouse ihre Patente zusammen {jonnes,carlson}.'],
]));
body.push(caption('Tab. 3: Chronologie.'));

body.push(H('4  Warum gewann der Wechselstrom?'));
body.push(P('Den Ausschlag gaben Kosten und Kapital: Seit 1892 setzte mit GE auch Edisons frühere Firma auf Wechselstrom {jonnes}.'));
body.push(B('**Kosten:** Wechselstromnetze kamen mit wenigen grossen Kraftwerken und dünneren Leitungen aus; Edisons Gleichstrom brauchte dagegen alle paar Kilometer ein Kraftwerk {jonnes}.'));
body.push(B('**Angstkampagne scheitert:** Cole und Chandler deuten die Kampagne des Edison-Lagers als «competitive impression management» – als Versuch, das Bild des Konkurrenzprodukts beim gemeinsamen Publikum zu beschädigen. Westinghouse ignorierte die Angriffe zuerst, bestritt sie dann und machte sie schliesslich mit besserer Technik gegenstandslos {cole}.'));
body.push(B('**Europa:** 1891 übertrugen die Maschinenfabrik Oerlikon und die AEG Drehstrom über 175 km von Lauffen nach Frankfurt. Der Oerlikon-Ingenieur Charles E. L. Brown gründete im selben Jahr in Baden die BBC {steig}, die sich 1988 mit der schwedischen ASEA zu ABB zusammenschloss {abbHistory}.'));
body.push(B('**Kein K.-o.-Sieg:** Umformer koppelten neue Wechselstrom- an alte Gleichstromnetze {davidBunn,millard}; Manhattans letzte Gleichstromversorgung endete erst 2007 {lee}.'));
body.push(P('**Deutung:** Paul David liest den Fall als Lehrstück der Pfadabhängigkeit: Standards entscheiden sich in kurzen, offenen Momenten und verfestigen sich danach {david92}.'));

body.push(H('5  Bedeutung für den Aufstieg der USA'));
body.push(P('Der Stromkrieg fällt in die Jahrzehnte, in denen die USA Grossbritannien als grösste Industriemacht überholten: zwischen 1880 und 1900 (Abb. 1). 1913 stellten die USA knapp ein Drittel der Weltindustrieproduktion – mehr als Grossbritannien und Deutschland zusammen {bairoch}. Die Grafik zeigt einen zeitlichen Zusammenhang, keinen Beweis für eine Ursache: Billiger, übertragbarer Strom war eine von mehreren Basistechnologien dieses Aufstiegs {hughes91}.'));
const png = fs.readFileSync(path.join(ROOT, 'assets', 'chart_weltindustrie.png'));
body.push(new Paragraph({
  keepNext: true, spacing: { before: 60, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new ImageRun({ type: 'png', data: png, transformation: { width: 480, height: 190 }, altText: { title: 'Anteil an der Weltindustrieproduktion', description: 'Liniendiagramm 1860–1913: USA steigen von 7,2 auf 32,0 Prozent, Grossbritannien fällt von 22,9 (1880) auf 13,6 Prozent, Deutschland steigt auf 14,8 Prozent.', name: 'chart' } })],
}));
body.push(caption('Abb. 1: Anteil an der Weltindustrieproduktion in Prozent. Die Linien verbinden die Stichjahre 1860, 1880, 1900 und 1913; es sind Schätzwerte. Eigene Darstellung nach {bairoch}.'));
body.push(B('**Fabriken:** Elektromotoren lieferten 1899 knapp 5 %, 1929 rund 78 % der Antriebsleistung der US-Industrie. Der Output pro Arbeitsstunde wuchs vor 1919 um 1,3 %, danach um 3,1 % pro Jahr {devine}. Paul David erklärt die Verzögerung so: Der Motor zahlte sich erst aus, als die Fabriken um ihn herum neu organisiert waren {david90}.'));
body.push(B('**Big Business:** GE und Westinghouse legten 1896 ihre Patente zusammen und beherrschten den Markt als Duopol {jonnes} – typisch für die Konzernbildung des «Gilded Age».'));
body.push(B('**Netze:** Grosskraftwerke und Fernleitungen wuchsen zu regionalen Verbundnetzen zusammen {hughes83}; an den Niagarafällen zog billige Wasserkraft die Aluminium- und Elektrochemie-Industrie an {jonnes}.'));

body.push(H('6  Mythen und Quellenkritik'));
body.push(P('Populäre Darstellungen erzählen ein Duell zweier Genies (z. B. {tagi,abb,spiegel}); die Forschung ist nüchterner.'));
body.push(B('**«Tesla gegen Edison»:** Filme wie *The Current War* (2017/2019) spitzen den Konflikt auf die Erfinder zu. Gestritten haben Unternehmen und ihre Geldgeber {jonnes}; Tesla lieferte Westinghouse 1888 die Motorpatente {carlson}.'));
body.push(B('**Elefantin Topsy (1903):** Der ABB-Artikel schreibt, Edison habe öffentlich «Experimente mit Hunden, Pferden und sogar einem Elefanten» vorgeführt {abb}. Die Tierversuche von 1888 führte aber Harold P. Brown durch {hughes58}, und Topsy starb erst 1903 – über zehn Jahre nach dem Stromkrieg und ohne Edisons Beteiligung; nur seine Filmfirma filmte die Tötung {topsy}.'));
body.push(B('**Der zerrissene Vertrag:** Belegt sind nur Teslas Verzicht auf Lizenzgebühren (1891) und der spätere Verkauf der Patente für 216 000 Dollar {carlson}.'));
body.push(P('Zum Einstieg eignen sich die Dokumentationen {pbsEdison,pbsTesla,terraX}; auch sie sind populäre Formate und quellenkritisch zu lesen.'));

body.push(H('7  Fazit und Ausblick'));
body.push(P('Der Stromkrieg war ein früher, prägender Standardkrieg der Industriegeschichte; McNichol spricht sogar vom «first standards war» {mcnichol}. Gewonnen hat das System, das sich billiger skalieren liess – und mit ihm der finanzstarke Grosskonzern. Darin liegt seine Bedeutung für die «Weltmacht USA».'));
body.push(P('**Ironie der Geschichte:** Dem Gleichstrom fehlte um 1890 nur die Technik, seine Spannung zu wandeln. Mit Stromrichtern kehrt er zurück: Die erste kommerzielle Hochspannungs-Gleichstrom-Übertragung moderner Bauart ging 1954 zwischen dem schwedischen Festland und Gotland in Betrieb, gebaut von ASEA {gotland}. Auch Akkus, Solarzellen und Elektronik arbeiten mit Gleichstrom.'));
body.push(P('**Standardkriege heute:** Weltweit kommt Wechselstrom aus der Steckdose, aber mit verschiedenen Spannungen und Steckern {iec}. In den USA wird der Ladestecker der Firma Tesla seit Ende 2023 zur Norm SAE J3400 {nacs}; in der EU ist USB-C für Handys seit Ende 2024 Pflicht {eu}.'));

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

const doc = new Document({
  creator: 'Simon Bruder', title: 'Der Stromkrieg (War of Currents) – Ausarbeitung',
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
  execFileSync('python3', [path.join(__dirname, 'zip_ohne_ordner.py'), OUT]); // Ordner-Einträge stören OneDrive
  console.log('OK', OUT, order.length, 'Quellen');
});
