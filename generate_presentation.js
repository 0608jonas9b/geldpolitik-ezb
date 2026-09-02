/**
 * Praesentation: "Die Geldpolitik der EZB"
 * Zielgruppe: Berufsschule, Politikunterricht | Dauer: max. 20 Minuten
 * Aufbau: Deckblatt + Gliederung + 17 Themenfolien
 */
const pptxgen = require("pptxgenjs");

/* ----------------------------- Design-System ----------------------------- */
const NAVY      = "10224E"; // dominante Farbe (EU-Blau, dunkel)
const NAVY_MID  = "1C3670"; // Karten auf dunklem Grund
const BLUE      = "2E5AAC"; // Stuetzfarbe
const GOLD      = "FFC72C"; // Akzent (EU-Gold)
const GOLD_DARK = "9A6F00"; // Gold-Ton, lesbar auf hellem Grund
const INK       = "1A1D26"; // Fliesstext
const MUTED     = "5A6274"; // Sekundaertext
const CARD      = "F1F4FA"; // helle Karte
const CARD_LINE = "D8E0EF";
const WHITE     = "FFFFFF";
const LIGHTTXT  = "C9D5EE"; // Text auf dunklem Grund
const GREEN     = "1B7A5A"; // semantisch: expansiv
const RED       = "B3321F"; // semantisch: restriktiv

const HEAD = "Cambria";
const BODY = "Calibri";

const W = 13.333, H = 7.5;
const LM = 0.62;                 // linker Rand
const CW = W - 2 * LM;           // nutzbare Breite = 12.093

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Politikunterricht";
pres.title  = "Die Geldpolitik der EZB";

/* ------------------------------- Helfer --------------------------------- */
const shadow = () => ({ type: "outer", color: "8A96AD", blur: 7, offset: 1, angle: 90, opacity: 0.22 });

function card(slide, o) {
  slide.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill || CARD },
    line: o.line === null ? { color: o.fill || CARD, width: 0 } : { color: o.line || CARD_LINE, width: 1 },
    rectRadius: 0.08,
    shadow: o.noShadow ? undefined : shadow(),
  });
}

function circle(slide, o) {
  slide.addShape(pres.ShapeType.ellipse, {
    x: o.x, y: o.y, w: o.d, h: o.d,
    fill: { color: o.fill || GOLD },
    line: { color: o.fill || GOLD, width: 0 },
  });
  slide.addText(o.txt, {
    x: o.x, y: o.y, w: o.d, h: o.d,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
    fontFace: o.face || HEAD, fontSize: o.size || 15, bold: true,
    color: o.txtColor || NAVY,
  });
}

function txt(slide, text, o) {
  slide.addText(text, Object.assign({
    isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK,
    valign: "top", align: "left",
  }, o));
}

/** Kopfzeile: kleines Kapitel-Label + Folientitel */
function head(slide, eyebrow, title, dark) {
  txt(slide, eyebrow.toUpperCase(), {
    x: LM, y: 0.36, w: CW, h: 0.26,
    fontSize: 11, bold: true, charSpacing: 1.4,
    color: dark ? GOLD : GOLD_DARK, fontFace: BODY,
  });
  txt(slide, title, {
    x: LM, y: 0.62, w: CW, h: 0.74,
    fontSize: 30, bold: true, fontFace: HEAD,
    color: dark ? WHITE : NAVY, valign: "middle",
  });
}

/** Fusszeile mit Foliennummer */
function foot(slide, label) {
  txt(slide, "Die Geldpolitik der EZB", {
    x: LM, y: 6.96, w: 6.0, h: 0.28, fontSize: 9.5, color: MUTED,
  });
  txt(slide, label, {
    x: W - LM - 4.0, y: 6.96, w: 4.0, h: 0.28, fontSize: 9.5, color: MUTED, align: "right",
  });
}

/** Aufzaehlung als Textarray */
function bullets(items, opt) {
  return items.map((t, i) => ({
    text: t,
    options: Object.assign({ bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 5 }, opt || {}),
  }));
}

let n = 0;
const NUM = () => `Folie ${++n} von 17`;

/* ============================ 1 · DECKBLATT ============================== */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  // Motiv: Euro-Zeichen im Kreis aus goldenen Punkten (Anlehnung an den Sternenkranz)
  const cx = 10.60, cy = 3.60, r = 1.85, d = 0.17;
  for (let i = 0; i < 12; i++) {
    const a = (i / 12) * 2 * Math.PI - Math.PI / 2;
    s.addShape(pres.ShapeType.ellipse, {
      x: cx + r * Math.cos(a) - d / 2, y: cy + r * Math.sin(a) - d / 2, w: d, h: d,
      fill: { color: GOLD }, line: { color: GOLD, width: 0 },
    });
  }
  s.addShape(pres.ShapeType.ellipse, {
    x: cx - 1.15, y: cy - 1.15, w: 2.30, h: 2.30,
    fill: { color: NAVY_MID }, line: { color: GOLD, width: 1.25 },
  });
  txt(s, "€", {
    x: cx - 1.15, y: cy - 1.15, w: 2.30, h: 2.30,
    align: "center", valign: "middle", fontFace: HEAD, fontSize: 88, bold: true, color: GOLD,
  });

  txt(s, "POLITIKUNTERRICHT · BERUFSSCHULE", {
    x: 0.90, y: 1.62, w: 7.2, h: 0.30, fontSize: 12.5, bold: true, charSpacing: 1.6, color: GOLD,
  });
  txt(s, "Die Geldpolitik\nder EZB", {
    x: 0.90, y: 1.98, w: 7.2, h: 1.66, fontSize: 44, bold: true, fontFace: HEAD, color: WHITE, valign: "top",
  });
  txt(s, "Inhalte, Strategie und Instrumente – und warum der Zins uns alle betrifft", {
    x: 0.90, y: 3.78, w: 7.0, h: 0.85, fontSize: 17, color: LIGHTTXT, lineSpacingMultiple: 1.2,
  });
  txt(s, [
    { text: "Vorgestellt von: ", options: { bold: true, color: WHITE } },
    { text: "[Name eintragen]", options: { color: LIGHTTXT, breakLine: true } },
    { text: "Fach / Klasse: ", options: { bold: true, color: WHITE } },
    { text: "Politik · [Klasse eintragen]", options: { color: LIGHTTXT, breakLine: true } },
    { text: "Datum: ", options: { bold: true, color: WHITE } },
    { text: "[TT.MM.JJJJ]", options: { color: LIGHTTXT } },
  ], { x: 0.90, y: 4.92, w: 7.0, h: 1.10, fontSize: 13.5, paraSpaceAfter: 4 });

  txt(s, "Vortragsdauer: ca. 20 Minuten", {
    x: 0.90, y: 6.55, w: 7.0, h: 0.30, fontSize: 11, color: "8FA2C8", italic: true,
  });

  s.addNotes("Begruessung. Einstiegsfrage ans Publikum: Wer hat schon einmal von 'Leitzins' in den Nachrichten gehoert? "
    + "Kurz das Ziel der Praesentation nennen: Wir schauen uns an, was die EZB macht, mit welchen Werkzeugen sie arbeitet "
    + "und wie ihre Zinsentscheidungen bei uns allen ankommen. (ca. 30 Sekunden)");
}

/* ============================ 2 · GLIEDERUNG ============================= */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Überblick", "Gliederung", false);

  const parts = [
    { t: "Grundlagen", z: "ca. 6 Minuten", i: ["Warum uns die EZB alle betrifft", "Was ist Geld?", "Wer schafft das Geld?", "Die EZB: Aufbau und Auftrag", "Das Ziel: Preisstabilität"] },
    { t: "Strategie und Instrumente", z: "ca. 6 Minuten", i: ["Die Strategie der EZB", "Der Werkzeugkasten im Überblick", "Leitzinsen", "Offenmarktgeschäfte", "Mindestreserve"] },
    { t: "Der Zins und seine Wirkung", z: "ca. 8 Minuten", i: ["Der Transmissionsprozess", "Expansiv oder restriktiv?", "Praxisbeispiel: die Zinswende", "Haushalte, Unternehmen, Staat", "Fazit und Diskussion"] },
  ];
  const cw = 3.83, gap = 0.30;
  parts.forEach((p, k) => {
    const x = LM + k * (cw + gap), y = 1.72, h = 4.00;
    card(s, { x, y, w: cw, h });
    circle(s, { x: x + 0.32, y: y + 0.32, d: 0.62, txt: String(k + 1), size: 20 });
    txt(s, `Teil ${k + 1}`, { x: x + 1.08, y: y + 0.40, w: cw - 1.4, h: 0.26, fontSize: 11, bold: true, color: GOLD_DARK, charSpacing: 1.2 });
    txt(s, p.t, { x: x + 1.08, y: y + 0.66, w: cw - 1.35, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
    txt(s, bullets(p.i), { x: x + 0.34, y: y + 1.30, w: cw - 0.68, h: 1.95, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.1 });
    txt(s, p.z, { x: x + 0.34, y: y + 3.38, w: cw - 0.68, h: 0.28, fontSize: 11.5, italic: true, color: GOLD_DARK });
  });

  card(s, { x: LM, y: 5.95, w: CW, h: 0.62, fill: NAVY, line: null });
  circle(s, { x: LM + 0.24, y: 6.10, d: 0.32, txt: "›", size: 14 });
  txt(s, "Roter Faden: vom Geld über die Instrumente der EZB bis zur Wirkung des Zinses auf Haushalte, Unternehmen und Staat.", {
    x: LM + 0.72, y: 5.95, w: CW - 1.0, h: 0.62, fontSize: 12.5, color: WHITE, valign: "middle",
  });

  foot(s, "Gliederung");
  s.addNotes("Gliederung kurz vorstellen, jeden Teil in einem Satz ankuendigen. Nicht vorlesen! "
    + "Hinweis geben, dass Fragen am Ende gesammelt beantwortet werden. (ca. 30 Sekunden)");
}

/* =================== FOLIE 1 · Warum uns die EZB betrifft ================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 1 · Grundlagen", "Warum uns die EZB alle betrifft", false);
  txt(s, "Vier Alltagssituationen, in denen Entscheidungen aus Frankfurt am Main direkt spürbar werden:", {
    x: LM, y: 1.42, w: CW, h: 0.40, fontSize: 14.5, color: MUTED,
  });

  const items = [
    ["Einkaufen", "Steigen die Preise schneller als der Lohn, kann man sich für dasselbe Geld weniger leisten. Genau das soll die EZB verhindern."],
    ["Sparen", "Ob es für das Ersparte 0 % oder 3 % Zinsen gibt, hängt vor allem an den Leitzinsen der EZB."],
    ["Kredit aufnehmen", "Ratenkauf, Autokredit, später die Baufinanzierung: Der Zins bestimmt, wie teuer geliehenes Geld ist."],
    ["Ausbildung und Job", "Günstige Kredite erleichtern Investitionen der Betriebe – und damit Arbeits- und Ausbildungsplätze."],
  ];
  const cw = 5.90, ch = 1.80;
  items.forEach((it, k) => {
    const x = LM + (k % 2) * (cw + 0.29), y = 2.00 + Math.floor(k / 2) * 1.95;
    card(s, { x, y, w: cw, h: ch });
    circle(s, { x: x + 0.30, y: y + 0.28, d: 0.62, txt: String(k + 1), size: 19 });
    txt(s, it[0], { x: x + 1.06, y: y + 0.30, w: cw - 1.4, h: 0.34, fontSize: 16, bold: true, fontFace: HEAD, color: NAVY });
    txt(s, it[1], { x: x + 1.06, y: y + 0.72, w: cw - 1.4, h: 0.92, fontSize: 13, color: INK, lineSpacingMultiple: 1.1 });
  });

  card(s, { x: LM, y: 6.00, w: CW, h: 0.62, fill: NAVY, line: null });
  txt(s, "Hinter all dem steht eine Institution: die Europäische Zentralbank (EZB).", {
    x: LM + 0.35, y: 6.00, w: CW - 0.7, h: 0.62, fontSize: 14, bold: true, color: WHITE, valign: "middle",
  });

  foot(s, NUM());
  s.addNotes("Einstieg ueber den Alltag. Klasse aktivieren: 'Wer von euch spart auf etwas Groesseres?' "
    + "Die vier Beispiele nur kurz antippen - sie werden in Teil 3 vertieft. (ca. 60 Sekunden)");
}

/* ======================= FOLIE 2 · Was ist Geld? ========================= */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 1 · Grundlagen", "Was ist Geld?", false);

  txt(s, "Drei Funktionen des Geldes", { x: LM, y: 1.48, w: 6.4, h: 0.34, fontSize: 16, bold: true, fontFace: HEAD, color: NAVY });
  const fn = [
    ["Tauschmittel", "Geld wird überall als Bezahlung akzeptiert. Ohne Geld müsste Ware gegen Ware getauscht werden – man bräuchte immer den passenden Partner."],
    ["Recheneinheit", "Alle Güter werden in Euro bewertet. So lassen sich ganz unterschiedliche Dinge überhaupt vergleichen."],
    ["Wertaufbewahrung", "Geld lässt sich sparen: Kaufkraft wird in die Zukunft verschoben – solange die Inflation niedrig bleibt."],
  ];
  fn.forEach((f, k) => {
    const y = 1.92 + k * 1.10;
    circle(s, { x: LM + 0.06, y: y + 0.06, d: 0.55, txt: String(k + 1), size: 16 });
    txt(s, f[0], { x: LM + 0.82, y: y + 0.02, w: 5.5, h: 0.30, fontSize: 15, bold: true, color: NAVY });
    txt(s, f[1], { x: LM + 0.82, y: y + 0.34, w: 5.5, h: 0.70, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08 });
  });
  card(s, { x: LM, y: 5.35, w: 6.4, h: 0.85, fill: NAVY, line: null });
  txt(s, "„Geld ist alles, was die Geldfunktionen erfüllt.“", {
    x: LM + 0.32, y: 5.35, w: 5.8, h: 0.85, fontSize: 14.5, italic: true, color: WHITE, valign: "middle",
  });

  const rx = 7.20, rw = 5.51;
  card(s, { x: rx, y: 1.48, w: rw, h: 4.72 });
  txt(s, "Zwei Arten von Geld", { x: rx + 0.28, y: 1.72, w: rw - 0.56, h: 0.34, fontSize: 16, bold: true, fontFace: HEAD, color: NAVY });
  const arten = [
    ["Bargeld", "gesetzliches Zahlungsmittel", "Münzen geben die Staaten aus (mit Genehmigung der EZB), Banknoten nur die Zentralbank – sie hat das Notenmonopol."],
    ["Buchgeld (Giralgeld)", "kein gesetzliches Zahlungsmittel", "Guthaben auf dem Girokonto, jederzeit verfügbar: Überweisung, Karte, Bezahl-App. Es entsteht in den Geschäftsbanken."],
  ];
  arten.forEach((a, k) => {
    const y = 2.20 + k * 1.68;
    card(s, { x: rx + 0.26, y, w: rw - 0.52, h: 1.52, fill: WHITE, line: CARD_LINE, noShadow: true });
    txt(s, a[0], { x: rx + 0.50, y: y + 0.16, w: rw - 1.0, h: 0.30, fontSize: 14.5, bold: true, color: NAVY });
    txt(s, a[1], { x: rx + 0.50, y: y + 0.46, w: rw - 1.0, h: 0.26, fontSize: 11.5, italic: true, color: GOLD_DARK });
    txt(s, a[2], { x: rx + 0.50, y: y + 0.74, w: rw - 1.0, h: 0.68, fontSize: 12, color: INK, lineSpacingMultiple: 1.05 });
  });
  txt(s, "Der weitaus größte Teil der Geldmenge im Euroraum ist Buchgeld – nicht Bargeld.", {
    x: rx + 0.26, y: 5.62, w: rw - 0.52, h: 0.45, fontSize: 12.5, bold: true, color: NAVY,
  });

  foot(s, NUM());
  s.addNotes("Kernaussage: Geld ist keine Sache, sondern eine Funktion. Beispiel Tauschwirtschaft nennen "
    + "(Baecker will keine Schuhe - 'Problem des doppelten Zufalls'). Bei den Geldarten betonen: "
    + "Was auf dem Konto liegt, ist Buchgeld und kein gesetzliches Zahlungsmittel. (ca. 90 Sekunden)");
}

/* ==================== FOLIE 3 · Wer schafft das Geld? ==================== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 1 · Grundlagen", "Wer schafft das Geld?", false);

  const steps = [
    ["EZB schafft die Basis", "Sie gibt Zentralbankgeld aus: Bargeld und Guthaben der Banken bei der Notenbank."],
    ["Banken leihen sich Geld", "Gegen Sicherheiten und gegen Zins holen sie sich Zentralbankgeld bei der EZB."],
    ["Banken vergeben Kredite", "Der Betrag wird dem Konto gutgeschrieben – dabei entsteht neues Buchgeld."],
    ["Die Geldmenge wächst", "Über viele Runden entsteht aus der Basis ein Vielfaches an Geldmenge."],
  ];
  const bw = 2.82, bh = 2.20, by = 1.78;
  steps.forEach((st, k) => {
    const x = LM + k * (bw + 0.25);
    card(s, { x, y: by, w: bw, h: bh });
    circle(s, { x: x + 0.18, y: by + 0.18, d: 0.46, txt: String(k + 1), size: 14 });
    txt(s, st[0], { x: x + 0.18, y: by + 0.70, w: bw - 0.36, h: 0.60, fontSize: 13.5, bold: true, color: NAVY });
    txt(s, st[1], { x: x + 0.18, y: by + 1.32, w: bw - 0.36, h: 0.80, fontSize: 11.5, color: INK, lineSpacingMultiple: 1.05 });
    if (k < 3) {
      txt(s, "›", { x: x + bw - 0.05, y: by + bh / 2 - 0.22, w: 0.35, h: 0.45, fontSize: 22, bold: true, color: GOLD, align: "center", valign: "middle" });
    }
  });

  card(s, { x: LM, y: 4.15, w: 5.90, h: 1.55, fill: NAVY, line: null });
  txt(s, "325.000 €", { x: LM + 0.32, y: 4.30, w: 5.3, h: 0.62, fontSize: 38, bold: true, fontFace: HEAD, color: GOLD });
  txt(s, "Geldmenge, die im Rechenbeispiel des Skripts aus 100.000 € Zentralbankgeld entstehen kann – bei 30 % Bargeldquote und 10 % Mindestreserve.", {
    x: LM + 0.32, y: 4.94, w: 5.3, h: 0.68, fontSize: 12, color: LIGHTTXT, lineSpacingMultiple: 1.05,
  });

  card(s, { x: 6.81, y: 4.15, w: 5.90, h: 1.55 });
  txt(s, "Merke", { x: 7.13, y: 4.30, w: 5.3, h: 0.30, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Die EZB steuert nur die Geldbasis. Wie viel Geld am Ende wirklich umläuft, hängt auch vom Verhalten der Banken (Kreditvergabe) und der Kunden (Bargeldhaltung, Kreditnachfrage) ab.", {
    x: 7.13, y: 4.64, w: 5.3, h: 0.95, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });

  txt(s, "Vereinfachte Darstellung des Geldschöpfungsprozesses; Geldschöpfungsmultiplikator m = (1 + c) / (c + r).", {
    x: LM, y: 5.86, w: CW, h: 0.32, fontSize: 10.5, italic: true, color: MUTED,
  });

  foot(s, NUM());
  s.addNotes("Wichtigster Satz: Den groessten Teil des Geldes schaffen nicht die Notenbanken, sondern die "
    + "Geschaeftsbanken durch Kreditvergabe. Das Zahlenbeispiel nur als Groessenordnung nennen, nicht vorrechnen. "
    + "Ueberleitung: Wer kontrolliert diesen Prozess? (ca. 75 Sekunden)");
}

/* ================== FOLIE 4 · Die EZB - Aufbau und Auftrag =============== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 1 · Grundlagen", "Die EZB: Aufbau und Auftrag", false);

  const org = [
    ["EZB-Rat", "Oberstes geldpolitisches Gremium: das Direktorium und die Präsidentinnen und Präsidenten der nationalen Zentralbanken des Euroraums. Er beschließt die Leitzinsen – in der Regel alle sechs Wochen."],
    ["Direktorium", "Präsident/in, Vizepräsident/in und vier weitere Mitglieder. Es führt die Beschlüsse des EZB-Rats im Tagesgeschäft aus."],
    ["Nationale Zentralbanken", "Zum Beispiel die Deutsche Bundesbank. Sie setzen die Geldpolitik vor Ort um und sind integraler Teil des Systems."],
  ];
  const orgY = [1.52, 3.26, 4.62], orgH = [1.62, 1.24, 1.24];
  org.forEach((o, k) => {
    const y = orgY[k], x = LM, w = 7.30;
    card(s, { x, y, w, h: orgH[k] });
    circle(s, { x: x + 0.26, y: y + 0.24, d: 0.50, txt: String(k + 1), size: 15 });
    txt(s, o[0], { x: x + 0.92, y: y + 0.18, w: w - 1.2, h: 0.32, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY });
    txt(s, o[1], { x: x + 0.92, y: y + 0.52, w: w - 1.2, h: orgH[k] - 0.62, fontSize: 12, color: INK, lineSpacingMultiple: 1.06 });
  });

  const rx = 8.15, rw = 4.56;
  card(s, { x: rx, y: 1.52, w: rw, h: 4.34, fill: NAVY, line: null });
  txt(s, "Unabhängig – und trotzdem gebunden", { x: rx + 0.30, y: 1.76, w: rw - 0.60, h: 0.62, fontSize: 15.5, bold: true, fontFace: HEAD, color: GOLD });
  txt(s, bullets([
    "Die EZB darf keine Weisungen von Regierungen annehmen (Art. 130 AEUV).",
    "Ihr Ziel steht im EU-Vertrag – es hängt nicht von Wahlergebnissen ab.",
    "Sie muss sich gegenüber dem Europäischen Parlament und der Öffentlichkeit erklären.",
    "Warum? Regierungen könnten versucht sein, kurz vor Wahlen billiges Geld zu fordern – auf Kosten stabiler Preise.",
  ]), { x: rx + 0.30, y: 2.46, w: rw - 0.60, h: 3.28, fontSize: 12.5, color: LIGHTTXT, lineSpacingMultiple: 1.1 });

  card(s, { x: LM, y: 6.00, w: CW, h: 0.62 });
  txt(s, [
    { text: "Eurosystem", options: { bold: true, color: NAVY } },
    { text: " = EZB + die nationalen Zentralbanken der Euro-Länder.   •   " },
    { text: "ESZB", options: { bold: true, color: NAVY } },
    { text: " = EZB + die Zentralbanken aller EU-Staaten." },
  ], { x: LM + 0.35, y: 6.00, w: CW - 0.7, h: 0.62, fontSize: 13, color: INK, valign: "middle" });

  foot(s, NUM());
  s.addNotes("Aufbau nur grob erklaeren - Pruefungsrelevant ist: Der EZB-Rat entscheidet ueber die Zinsen. "
    + "Bei der Unabhaengigkeit die politische Debatte antippen: Sie ist gewollt, wird aber auch kritisiert, "
    + "weil niemand die EZB abwaehlen kann. (ca. 75 Sekunden)");
}

/* ===================== FOLIE 5 · Ziel: Preisstabilitaet ================== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 1 · Grundlagen", "Das Ziel: Preisstabilität", false);

  card(s, { x: LM, y: 1.52, w: 4.20, h: 3.42, fill: NAVY, line: null });
  txt(s, "2 %", { x: LM + 0.32, y: 1.78, w: 3.56, h: 1.05, fontSize: 62, bold: true, fontFace: HEAD, color: GOLD, valign: "middle" });
  txt(s, "symmetrisches Inflationsziel für den Euroraum – mittelfristig, gemessen am harmonisierten Verbraucherpreisindex (HVPI).", {
    x: LM + 0.32, y: 2.92, w: 3.56, h: 0.90, fontSize: 12.5, color: WHITE, lineSpacingMultiple: 1.08,
  });
  txt(s, "Symmetrisch heißt: Abweichungen nach oben und nach unten sind gleich unerwünscht.", {
    x: LM + 0.32, y: 3.92, w: 3.56, h: 0.85, fontSize: 11.5, italic: true, color: LIGHTTXT, lineSpacingMultiple: 1.05,
  });

  card(s, { x: LM, y: 5.08, w: 4.20, h: 1.32 });
  txt(s, "Konkret heißt das", { x: LM + 0.30, y: 5.24, w: 3.60, h: 0.30, fontSize: 14.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Ein Warenkorb, der heute 100 € kostet, darf in einem Jahr rund 102 € kosten – nicht 108 €.", {
    x: LM + 0.30, y: 5.58, w: 3.60, h: 0.72, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });

  const rx = 5.10, rw = 7.61;
  card(s, { x: rx, y: 1.52, w: rw, h: 2.38 });
  txt(s, "Warum nicht 0 % Inflation?", { x: rx + 0.32, y: 1.72, w: rw - 0.64, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, bullets([
    "Sicherheitsabstand zur Deflation: Fallende Preise lassen Käufe aufschieben – die Wirtschaft schrumpft.",
    "Der Preisindex misst leicht zu hoch, ein kleiner Puffer gleicht das aus.",
    "Nur wer nicht schon bei null steht, kann die Zinsen im Abschwung noch senken.",
  ]), { x: rx + 0.32, y: 2.14, w: rw - 0.64, h: 1.60, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08 });

  card(s, { x: rx, y: 4.10, w: rw, h: 2.30 });
  txt(s, "Vorrang und Nachrang (Art. 127 AEUV)", { x: rx + 0.32, y: 4.30, w: rw - 0.64, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Vorrangiges Ziel ist die Preisstabilität. Nur soweit dieses Ziel nicht beeinträchtigt wird, unterstützt die EZB die allgemeine Wirtschaftspolitik der EU – also Wachstum, hohe Beschäftigung und sozialen Zusammenhalt.", {
    x: rx + 0.32, y: 4.68, w: rw - 0.64, h: 0.95, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });
  txt(s, "Diese klare Rangfolge ist eine bewusste Entscheidung: Die US-Notenbank Fed verfolgt stabile Preise und hohe Beschäftigung dagegen gleichrangig.", {
    x: rx + 0.32, y: 5.62, w: rw - 0.64, h: 0.62, fontSize: 11.5, italic: true, color: MUTED, lineSpacingMultiple: 1.05,
  });

  foot(s, NUM());
  s.addNotes("Zentrale Folie fuer den Auftrag der EZB. Die 2 % gut erklaeren: Es geht nicht um sinkende Preise, "
    + "sondern um langsam und berechenbar steigende Preise. Vergleich mit der Fed nennen - zeigt, dass die "
    + "Rangfolge eine politische Entscheidung ist. (ca. 75 Sekunden)");
}

/* ======================= FOLIE 6 · Strategie der EZB ===================== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 2 · Strategie und Instrumente", "Die Strategie: Wie die EZB entscheidet", false);

  card(s, { x: LM, y: 1.52, w: CW, h: 0.80, fill: NAVY, line: null });
  txt(s, "Vorrangiges Ziel: Preisstabilität – rund 2 % Inflation auf mittlere Sicht", {
    x: LM + 0.35, y: 1.52, w: CW - 0.70, h: 0.80, fontSize: 16, bold: true, fontFace: HEAD, color: WHITE, valign: "middle",
  });

  const pil = [
    ["Säule 1", "Wirtschaftliche Analyse", ["Konjunktur, Aufträge und Arbeitsmarkt", "Löhne, Gewinne und Unternehmenskosten", "Energie- und Rohstoffpreise, Wechselkurs"], "Blick vor allem auf kurz- und mittelfristige Risiken für die Preise."],
    ["Säule 2", "Monetäre und finanzielle Analyse", ["Entwicklung der Geldmengen M1, M2 und M3", "Kreditvergabe der Geschäftsbanken", "Lage an den Finanz- und Immobilienmärkten"], "Blick vor allem auf längerfristige Inflationsrisiken."],
  ];
  const pw = 5.92;
  pil.forEach((p, k) => {
    const x = LM + k * (pw + 0.25), y = 2.60, h = 2.55;
    card(s, { x, y, w: pw, h });
    circle(s, { x: x + 0.30, y: y + 0.26, d: 0.54, txt: String(k + 1), size: 16 });
    txt(s, p[0], { x: x + 1.00, y: y + 0.26, w: pw - 1.3, h: 0.24, fontSize: 11, bold: true, color: GOLD_DARK, charSpacing: 1.2 });
    txt(s, p[1], { x: x + 1.00, y: y + 0.50, w: pw - 1.3, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
    txt(s, bullets(p[2]), { x: x + 0.34, y: y + 1.05, w: pw - 0.68, h: 0.95, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.05 });
    txt(s, p[3], { x: x + 0.34, y: y + 2.02, w: pw - 0.68, h: 0.42, fontSize: 11.5, italic: true, color: MUTED });
  });

  txt(s, "▼", { x: LM + pw / 2 - 0.25, y: 5.22, w: 0.50, h: 0.28, fontSize: 14, color: GOLD, align: "center" });
  txt(s, "▼", { x: LM + pw + 0.25 + pw / 2 - 0.25, y: 5.22, w: 0.50, h: 0.28, fontSize: 14, color: GOLD, align: "center" });

  card(s, { x: LM, y: 5.55, w: CW, h: 1.00 });
  txt(s, "Gesamtbeurteilung im EZB-Rat → Entscheidung über die Leitzinsen", {
    x: LM + 0.35, y: 5.68, w: CW - 0.70, h: 0.34, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY,
  });
  txt(s, "Entschieden wird datenabhängig von Sitzung zu Sitzung. Maßstab ist nicht der Preis von heute, sondern die erwartete Inflation in ein bis zwei Jahren.", {
    x: LM + 0.35, y: 6.02, w: CW - 0.70, h: 0.42, fontSize: 12.5, color: INK,
  });

  foot(s, NUM());
  s.addNotes("Bild vom Tempel nutzen: zwei Saeulen tragen das eine Ziel. Wichtig: Die EZB schaut nicht nur auf "
    + "die aktuelle Inflationsrate, sondern auf viele Indikatoren und auf die Zukunft. (ca. 75 Sekunden)");
}

/* ======================== FOLIE 7 · Werkzeugkasten ======================= */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 2 · Strategie und Instrumente", "Der Werkzeugkasten der EZB", false);
  txt(s, "Die EZB kann Geld teurer oder billiger machen – und die Menge an Zentralbankgeld steuern. Dafür hat sie drei klassische Instrumente:", {
    x: LM, y: 1.42, w: CW, h: 0.42, fontSize: 14.5, color: MUTED,
  });

  const tools = [
    ["Leitzinsen", "Der Preis des Geldes", "Die EZB legt fest, zu welchem Zins sich Banken Geld leihen oder Geld anlegen können. Das wichtigste und schnellste Instrument."],
    ["Offenmarktgeschäfte", "Die Menge des Geldes", "Über befristete Refinanzierungsgeschäfte gegen Sicherheiten versorgt die EZB die Banken mit Zentralbankgeld."],
    ["Mindestreserve", "Die Bindung von Geld", "Banken müssen einen Teil der Kundeneinlagen bei der Zentralbank halten. Das begrenzt ihren Spielraum für Kredite."],
  ];
  const cw = 3.83;
  tools.forEach((t, k) => {
    const x = LM + k * (cw + 0.30), y = 2.05, h = 3.05;
    card(s, { x, y, w: cw, h });
    circle(s, { x: x + 0.32, y: y + 0.32, d: 0.70, txt: String(k + 1), size: 22 });
    txt(s, t[0], { x: x + 0.32, y: y + 1.20, w: cw - 0.64, h: 0.40, fontSize: 16.5, bold: true, fontFace: HEAD, color: NAVY });
    txt(s, t[1], { x: x + 0.32, y: y + 1.62, w: cw - 0.64, h: 0.28, fontSize: 12, italic: true, color: GOLD_DARK });
    txt(s, t[2], { x: x + 0.32, y: y + 1.98, w: cw - 0.64, h: 0.92, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08 });
  });

  card(s, { x: LM, y: 5.35, w: CW, h: 1.05, fill: NAVY, line: null });
  txt(s, "Dazu kommen unkonventionelle Instrumente", { x: LM + 0.35, y: 5.48, w: CW - 0.70, h: 0.30, fontSize: 14, bold: true, color: GOLD });
  txt(s, "In Krisen kauft die EZB zusätzlich Anleihen von Staaten und Unternehmen oder vergibt sehr langfristige Kredite an Banken. So beeinflusst sie auch die langfristigen Zinsen – dort, wo der Leitzins allein nicht hinreicht.", {
    x: LM + 0.35, y: 5.78, w: CW - 0.70, h: 0.52, fontSize: 12.5, color: LIGHTTXT, lineSpacingMultiple: 1.05,
  });

  foot(s, NUM());
  s.addNotes("Uebersichtsfolie - hier den Aufbau der naechsten drei Folien ankuendigen. Die Merkformel nennen: "
    + "Preis, Menge, Bindung. (ca. 45 Sekunden)");
}

/* ======================== FOLIE 8 · Die Leitzinsen ======================= */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Instrument 1", "Die drei Leitzinsen und der Zinskanal", false);

  card(s, { x: LM, y: 1.52, w: 6.00, h: 4.88, fill: CARD });
  const rows = [
    ["Spitzenrefinanzierungsfazilität", "2,65 %", "Obergrenze: Zu diesem Zins bekommen Banken über Nacht Geld von der Zentralbank.", false],
    ["Hauptrefinanzierungssatz", "2,40 %", "Der eigentliche Leitzins: Preis für das wöchentliche Refinanzierungsgeschäft.", true],
    ["Einlagefazilität", "2,25 %", "Untergrenze: Zu diesem Zins parken Banken überschüssiges Geld über Nacht bei der Zentralbank.", false],
  ];
  rows.forEach((r, k) => {
    const y = 1.88 + k * 1.35, x = LM + 0.24, w = 5.52, hi = r[3];
    card(s, { x, y, w, h: 1.20, fill: hi ? NAVY : WHITE, line: hi ? null : CARD_LINE, noShadow: !hi });
    txt(s, r[0], { x: x + 0.26, y: y + 0.16, w: w - 1.60, h: 0.34, fontSize: 13.5, bold: true, color: hi ? WHITE : NAVY });
    txt(s, r[1], { x: x + w - 1.55, y: y + 0.12, w: 1.30, h: 0.42, fontSize: 20, bold: true, fontFace: HEAD, color: hi ? GOLD : NAVY, align: "right" });
    txt(s, r[2], { x: x + 0.26, y: y + 0.54, w: w - 0.52, h: 0.56, fontSize: 11.5, color: hi ? LIGHTTXT : INK, lineSpacingMultiple: 1.05 });
  });
  txt(s, "Stand: 09.07.2026 (Quelle: Vorlesungsskript). Die Sätze ändern sich mit jedem Zinsbeschluss.", {
    x: LM + 0.24, y: 5.96, w: 5.52, h: 0.34, fontSize: 10.5, italic: true, color: MUTED,
  });

  const rx = 6.90, rw = 5.81;
  card(s, { x: rx, y: 1.52, w: rw, h: 2.28 });
  txt(s, "Ein Kanal für den Geldmarkt", { x: rx + 0.32, y: 1.74, w: rw - 0.64, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Zwischen Ober- und Untergrenze bewegt sich der Zins, zu dem sich Banken untereinander über Nacht Geld leihen (€STR). Die EZB steuert damit den kurzfristigen Zins sehr genau – den langfristigen Zins nur indirekt.", {
    x: rx + 0.32, y: 2.12, w: rw - 0.64, h: 1.10, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });
  txt(s, "Die Initiative bei den Fazilitäten geht immer von den Geschäftsbanken aus.", {
    x: rx + 0.32, y: 3.22, w: rw - 0.64, h: 0.40, fontSize: 11.5, italic: true, color: MUTED,
  });

  card(s, { x: rx, y: 4.00, w: rw, h: 2.40, fill: NAVY, line: null });
  txt(s, "Warum uns das betrifft", { x: rx + 0.32, y: 4.22, w: rw - 0.64, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: GOLD });
  txt(s, "An diesen Sätzen hängen die Zinsen am Markt – zum Beispiel der EURIBOR, an den viele Kredite gekoppelt sind. Über die Banken landet die Entscheidung schließlich bei Sparbuch, Dispo, Autokredit und Baufinanzierung.", {
    x: rx + 0.32, y: 4.60, w: rw - 0.64, h: 1.15, fontSize: 12.5, color: LIGHTTXT, lineSpacingMultiple: 1.08,
  });
  txt(s, "Seit September 2024 liegt der Hauptrefinanzierungssatz nur noch 0,15 Prozentpunkte über dem Einlagesatz – gesteuert wird vor allem über die Einlagefazilität.", {
    x: rx + 0.32, y: 5.72, w: rw - 0.64, h: 0.56, fontSize: 11.5, italic: true, color: "9FB0D0", lineSpacingMultiple: 1.05,
  });

  foot(s, NUM());
  s.addNotes("Die drei Saetze in der richtigen Reihenfolge erklaeren: oben teuer leihen, unten sicher parken, "
    + "in der Mitte das eigentliche Geschaeft. Merksatz: Der Leitzins ist der Grosshandelspreis fuer Geld - "
    + "alles andere baut darauf auf. (ca. 90 Sekunden)");
}

/* ==================== FOLIE 9 · Offenmarktgeschaefte ===================== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Instrument 2", "Offenmarktgeschäfte: Geld auf Zeit", false);

  const left = [
    ["Hauptrefinanzierungsgeschäfte (HRG)", ["wöchentlich angeboten, Laufzeit eine Woche", "hier gilt der Leitzins der EZB", "wichtigste Quelle für Zentralbankgeld"]],
    ["Längerfristige Refinanzierungsgeschäfte (LRG)", ["monatlich angeboten, Laufzeit drei Monate", "sichern die Basisversorgung mit Liquidität", "geben den Banken Planungssicherheit"]],
  ];
  left.forEach((l, k) => {
    const y = 1.52 + k * 2.06, x = LM, w = 5.90;
    card(s, { x, y, w, h: 1.92 });
    txt(s, l[0], { x: x + 0.32, y: y + 0.18, w: w - 0.64, h: 0.60, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY });
    txt(s, bullets(l[1]), { x: x + 0.32, y: y + 0.82, w: w - 0.64, h: 1.00, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.05 });
  });
  card(s, { x: LM, y: 5.66, w: 5.90, h: 0.76, fill: NAVY, line: null });
  txt(s, "Immer gegen Sicherheiten: Die Banken hinterlegen Wertpapiere und kaufen sie später zurück (Pfandgeschäft bzw. Repo).", {
    x: LM + 0.32, y: 5.66, w: 5.26, h: 0.76, fontSize: 12.5, color: WHITE, valign: "middle", lineSpacingMultiple: 1.05,
  });

  const rx = 6.81, rw = 5.90;
  card(s, { x: rx, y: 1.52, w: rw, h: 2.58 });
  txt(s, "So läuft ein Geschäft ab", { x: rx + 0.32, y: 1.72, w: rw - 0.64, h: 0.32, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY });
  const steps = ["Die EZB schreibt einen Betrag aus (Tender).", "Die Banken geben ihre Gebote ab.", "Zuteilung des Zentralbankgeldes gegen Sicherheiten.", "Rückzahlung mit Zins nach Ablauf der Frist."];
  steps.forEach((st, k) => {
    const y = 2.14 + k * 0.46;
    circle(s, { x: rx + 0.34, y: y + 0.02, d: 0.32, txt: String(k + 1), size: 11 });
    txt(s, st, { x: rx + 0.80, y, w: rw - 1.15, h: 0.36, fontSize: 12.5, color: INK, valign: "middle" });
  });

  card(s, { x: rx, y: 4.30, w: rw, h: 2.10 });
  txt(s, "Zwei Tenderverfahren", { x: rx + 0.32, y: 4.50, w: rw - 0.64, h: 0.32, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, [
    { text: "Mengentender: ", options: { bold: true, color: NAVY } },
    { text: "Der Zins steht fest, die Banken nennen nur die gewünschte Menge. Zugeteilt wird nach Quote.", options: { breakLine: true } },
    { text: "Zinstender: ", options: { bold: true, color: NAVY, paraSpaceBefore: 6 } },
    { text: "Die Banken nennen Menge und Zins, den sie zu zahlen bereit sind. Wer mehr bietet, wird zuerst bedient." },
  ], { x: rx + 0.32, y: 4.88, w: rw - 0.64, h: 1.30, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08 });

  foot(s, NUM());
  s.addNotes("Vergleich anbieten: Die EZB verleiht Geld wie ein Pfandleihhaus - nur gegen Sicherheiten und nur "
    + "auf Zeit. Tenderverfahren kurz halten, das ist Detailwissen. (ca. 75 Sekunden)");
}

/* ======================== FOLIE 10 · Mindestreserve ====================== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Instrument 3", "Die Mindestreserve", false);

  card(s, { x: LM, y: 1.52, w: 6.40, h: 2.58 });
  txt(s, "Was steckt dahinter?", { x: LM + 0.32, y: 1.72, w: 5.76, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Banken müssen einen Teil bestimmter Kundeneinlagen als Guthaben bei ihrer nationalen Zentralbank halten – derzeit 1 % für Einlagen mit Laufzeit unter zwei Jahren.", {
    x: LM + 0.32, y: 2.08, w: 5.76, h: 0.80, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });
  txt(s, "Erfüllt werden muss die Vorgabe nur im Monatsdurchschnitt. Bei Verstoß drohen Sanktionen. Seit September 2023 wird das Guthaben nicht mehr verzinst.", {
    x: LM + 0.32, y: 2.96, w: 5.76, h: 0.76, fontSize: 12, color: MUTED, lineSpacingMultiple: 1.06,
  });

  card(s, { x: LM, y: 4.30, w: 6.40, h: 2.10 });
  txt(s, "Zwei Funktionen", { x: LM + 0.32, y: 4.50, w: 5.76, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, [
    { text: "Anbindung: ", options: { bold: true, color: NAVY } },
    { text: "Die Banken bleiben dauerhaft auf Zentralbankgeld angewiesen – und damit auf die EZB.", options: { breakLine: true } },
    { text: "Zinsglättung: ", options: { bold: true, color: NAVY } },
    { text: "Die Durchschnittserfüllung dämpft tägliche Schwankungen der Geldmarktsätze.", options: { breakLine: true } },
    { text: "In der Praxis ändert die EZB diesen Satz nur selten – ihr wichtigster Hebel bleibt der Zins.", options: { italic: true, color: MUTED, paraSpaceBefore: 6 } },
  ], { x: LM + 0.32, y: 4.88, w: 5.76, h: 1.35, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08 });

  const rx = 7.30, rw = 5.41;
  const eff = [
    ["↑", "Reservesatz steigt", ["bindet Zentralbankgeld bei der Notenbank", "Banken können weniger Kredite vergeben", "die Geldmenge wächst langsamer", "dämpfend für die Inflation"], RED],
    ["↓", "Reservesatz sinkt", ["setzt Zentralbankgeld wieder frei", "mehr Spielraum für die Kreditvergabe", "die Geldmenge wächst schneller", "anregend für die Konjunktur"], GREEN],
  ];
  eff.forEach((e, k) => {
    const y = 1.52 + k * 2.55;
    card(s, { x: rx, y, w: rw, h: 2.30 });
    circle(s, { x: rx + 0.30, y: y + 0.26, d: 0.58, txt: e[0], fill: e[3], txtColor: WHITE, size: 18, face: BODY });
    txt(s, e[1], { x: rx + 1.02, y: y + 0.34, w: rw - 1.32, h: 0.38, fontSize: 15, bold: true, fontFace: HEAD, color: e[3], valign: "middle" });
    txt(s, bullets(e[2]), { x: rx + 0.32, y: y + 0.98, w: rw - 0.64, h: 1.18, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.05 });
  });

  foot(s, NUM());
  s.addNotes("Bild nutzen: Die Mindestreserve ist wie ein Sicherheitsgurt - sie bindet einen Teil des Geldes. "
    + "Wichtig fuer die Ueberleitung: Heute ist sie eher technisches Beiwerk, der Zins ist das entscheidende "
    + "Instrument. Damit beginnt Teil 3. (ca. 60 Sekunden)");
}

/* ====================== FOLIE 11 · Transmissionsprozess ================== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 3 · Der Zins und seine Wirkung", "Wie eine Zinsentscheidung bei uns ankommt", false);
  txt(s, "Der Weg vom Beschluss in Frankfurt bis zum Preis im Laden heißt Transmissionsprozess:", {
    x: LM, y: 1.42, w: CW, h: 0.36, fontSize: 14.5, color: MUTED,
  });

  const chain = [
    ["EZB", "beschließt den Leitzins"],
    ["Banken", "ihre Refinanzierung wird teurer oder billiger"],
    ["Kunden", "Kredit- und Sparzinsen für Haushalte und Betriebe ändern sich"],
    ["Nachfrage", "Konsum und Investitionen steigen oder sinken"],
    ["Preise", "die Inflationsrate reagiert – mit Verzögerung"],
  ];
  const bw = 2.17, gap = 0.31, by = 2.00, bh = 2.20;
  chain.forEach((c, k) => {
    const x = LM + k * (bw + gap);
    card(s, { x, y: by, w: bw, h: bh, fill: k === 0 || k === 4 ? NAVY : CARD, line: k === 0 || k === 4 ? null : CARD_LINE });
    const dark = (k === 0 || k === 4);
    circle(s, { x: x + 0.18, y: by + 0.20, d: 0.46, txt: String(k + 1), size: 14, fill: dark ? GOLD : NAVY, txtColor: dark ? NAVY : WHITE });
    txt(s, c[0], { x: x + 0.18, y: by + 0.78, w: bw - 0.36, h: 0.36, fontSize: 14.5, bold: true, fontFace: HEAD, color: dark ? WHITE : NAVY });
    txt(s, c[1], { x: x + 0.18, y: by + 1.16, w: bw - 0.36, h: 0.88, fontSize: 11.5, color: dark ? LIGHTTXT : INK, lineSpacingMultiple: 1.05 });
    if (k < 4) {
      txt(s, "›", { x: x + bw, y: by + 0.85, w: gap, h: 0.45, fontSize: 22, bold: true, color: GOLD, align: "center", valign: "middle", margin: 0 });
    }
  });

  card(s, { x: LM, y: 4.55, w: 5.90, h: 1.85 });
  txt(s, "Zeit ist der Knackpunkt", { x: LM + 0.32, y: 4.75, w: 5.26, h: 0.32, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Zinsentscheidungen wirken nicht sofort. Bis eine Änderung voll bei Preisen und Beschäftigung ankommt, vergehen meist vier bis acht Quartale. Die EZB muss deshalb vorausschauend handeln – und riskiert, zu spät oder zu stark zu reagieren.", {
    x: LM + 0.32, y: 5.12, w: 5.26, h: 1.15, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });

  card(s, { x: 6.81, y: 4.55, w: 5.90, h: 1.85 });
  txt(s, "Wo der Weg stocken kann", { x: 7.13, y: 4.75, w: 5.26, h: 0.32, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Geben die Banken günstige Konditionen nicht weiter, oder ist die Stimmung so schlecht, dass trotz niedriger Zinsen niemand investiert, verpufft die Wirkung. Ökonomen sprechen dann von einer Liquiditätsfalle.", {
    x: 7.13, y: 5.12, w: 5.26, h: 1.15, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });

  foot(s, NUM());
  s.addNotes("Kernfolie von Teil 3. Die Kette einmal langsam durchgehen und dabei auf die Folie zeigen. "
    + "Betonen: Die EZB kann Preise nicht direkt festlegen - sie wirkt nur indirekt ueber die Banken. "
    + "(ca. 90 Sekunden)");
}

/* ==================== FOLIE 12 · Expansiv vs. restriktiv ================= */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 3 · Der Zins und seine Wirkung", "Zwei Richtungen der Geldpolitik", false);

  const dirs = [
    ["↓", "Expansive Geldpolitik", "Leitzins runter – Geld wird billiger", GREEN, [
      ["Anlass", "schwache Konjunktur, Inflation dauerhaft unter 2 %"],
      ["Wirkung", "Kredite werden günstiger – Investitionen, Konsum und Beschäftigung steigen"],
      ["Risiko", "billiges Geld treibt Inflation und Preisblasen; Sparer verlieren real an Wert"],
    ]],
    ["↑", "Restriktive Geldpolitik", "Leitzins rauf – Geld wird teurer", RED, [
      ["Anlass", "Inflation deutlich über 2 %, Wirtschaft überhitzt"],
      ["Wirkung", "Kredite werden teurer – Investitionen und Konsum gehen zurück, der Preisdruck sinkt"],
      ["Risiko", "Wachstum bricht ein, Arbeitslosigkeit steigt, Staat und Betriebe zahlen mehr Zinsen"],
    ]],
  ];
  const cw = 5.90;
  dirs.forEach((d, k) => {
    const x = LM + k * (cw + 0.29), y = 1.52, h = 4.45;
    card(s, { x, y, w: cw, h });
    circle(s, { x: x + 0.30, y: y + 0.26, d: 0.62, txt: d[0], fill: d[3], txtColor: WHITE, size: 20, face: BODY });
    txt(s, d[1], { x: x + 1.06, y: y + 0.26, w: cw - 1.36, h: 0.34, fontSize: 16.5, bold: true, fontFace: HEAD, color: d[3] });
    txt(s, d[2], { x: x + 1.06, y: y + 0.62, w: cw - 1.36, h: 0.28, fontSize: 12.5, italic: true, color: MUTED });
    d[4].forEach((b, i) => {
      const by = y + 1.28 + i * 1.05;
      txt(s, b[0], { x: x + 0.32, y: by, w: 1.05, h: 0.26, fontSize: 11, bold: true, color: GOLD_DARK, charSpacing: 1.0 });
      txt(s, b[1], { x: x + 0.32, y: by + 0.26, w: cw - 0.64, h: 0.70, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.06 });
    });
  });

  card(s, { x: LM, y: 6.12, w: CW, h: 0.56, fill: NAVY, line: null });
  txt(s, "Die Richtung wählt die EZB immer mit Blick auf ihr Ziel: rund 2 % Inflation auf mittlere Sicht.", {
    x: LM + 0.35, y: 6.12, w: CW - 0.70, h: 0.56, fontSize: 13, bold: true, color: WHITE, valign: "middle",
  });

  foot(s, NUM());
  s.addNotes("Gegenueberstellung ist pruefungsrelevant. Merksatz anbieten: Expansiv = Gas geben, "
    + "restriktiv = bremsen. Wichtig: Beide Richtungen haben Verlierer - das leitet zur Bewertung am Ende ueber. "
    + "(ca. 90 Sekunden)");
}

/* ==================== FOLIE 13 · Praxisbeispiel Zinswende ================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 3 · Der Zins und seine Wirkung", "Praxisbeispiel: die Zinswende", false);

  s.addChart(pres.ChartType.bar, [{
    name: "Inflationsrate",
    labels: ["2020", "2021", "2022", "2023", "2024", "2025"],
    values: [0.3, 2.6, 8.4, 5.4, 2.4, 2.1],
  }], {
    x: LM, y: 1.50, w: 6.60, h: 4.30,
    barDir: "col", barGapWidthPct: 55,
    chartColors: [BLUE, BLUE, RED, RED, GOLD, GOLD],
    varyColors: true,
    showTitle: true, title: "Inflation im Euroraum in Prozent", titleFontSize: 14,
    titleColor: NAVY, titleFontFace: HEAD,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11,
    dataLabelColor: INK, dataLabelFontFace: BODY, dataLabelFormatCode: '0.0" %"',
    showLegend: false,
    catAxisLabelColor: MUTED, catAxisLabelFontSize: 11, catAxisLabelFontFace: BODY,
    valAxisLabelColor: MUTED, valAxisLabelFontSize: 10, valAxisLabelFontFace: BODY,
    valAxisMaxVal: 10, valAxisMajorUnit: 2,
    valGridLine: { color: "E4E9F2", size: 1 },
    catGridLine: { style: "none" },
    plotArea: { fill: { color: WHITE } },
  });
  txt(s, "Quelle: Eurostat / EZB, HVPI Euroraum, Jahresdurchschnitt; Wert für 2025 vorläufig. Zielwert der EZB: 2 %.", {
    x: LM, y: 5.86, w: 6.60, h: 0.36, fontSize: 10.5, italic: true, color: MUTED, lineSpacingMultiple: 1.0,
  });

  const rx = 7.45, rw = 5.26;
  card(s, { x: rx, y: 1.50, w: rw, h: 4.90, fill: NAVY, line: null });
  txt(s, "Die Antwort der EZB", { x: rx + 0.32, y: 1.72, w: rw - 0.64, h: 0.34, fontSize: 15.5, bold: true, fontFace: HEAD, color: GOLD });
  const mile = [
    ["bis 2021", "Einlagesatz bei −0,50 % – Geld ist extrem billig, die Inflation liegt jahrelang unter 2 %."],
    ["Juli 2022", "erste Zinserhöhung seit elf Jahren – Reaktion auf explodierende Energiepreise."],
    ["Sept. 2023", "Einlagesatz bei 4,00 % – der höchste Wert in der Geschichte der EZB."],
    ["ab Juni 2024", "schrittweise Senkungen, weil sich die Inflation dem Ziel wieder nähert."],
    ["Juli 2026", "Einlagesatz 2,25 %, Leitzins 2,40 % – die Geldpolitik ist wieder im Normalbereich."],
  ];
  mile.forEach((m, k) => {
    const y = 2.14 + k * 0.83;
    circle(s, { x: rx + 0.32, y: y + 0.04, d: 0.30, txt: String(k + 1), size: 11 });
    txt(s, m[0], { x: rx + 0.74, y, w: rw - 1.06, h: 0.24, fontSize: 11.5, bold: true, color: GOLD });
    txt(s, m[1], { x: rx + 0.74, y: y + 0.24, w: rw - 1.06, h: 0.54, fontSize: 11.5, color: LIGHTTXT, lineSpacingMultiple: 1.03 });
  });

  foot(s, NUM());
  s.addNotes("Hier wird die Theorie konkret. Die Grafik sprechen lassen: 2022 mehr als das Vierfache des "
    + "Zielwerts. Erklaeren, warum die EZB erst spaet reagiert hat (Annahme: Energiepreise sind nur voruebergehend) "
    + "- das ist ein guter Kritikpunkt fuer die Diskussion. (ca. 90 Sekunden)");
}

/* ==================== FOLIE 14 · Wirkung: Private Haushalte ============== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 3 · Wirkung auf die Akteure", "Was der Zins mit privaten Haushalten macht", false);

  const items = [
    ["Sparen und Anlegen", "Steigende Zinsen bringen wieder Erträge auf Tagesgeld und Sparbuch. Entscheidend ist der Realzins: Zins minus Inflation. Bei 2 % Zins und 3 % Inflation verliert das Ersparte trotzdem an Wert."],
    ["Kredite und Konsum", "Dispo, Ratenkauf und Autokredit werden teurer. Bei einer Baufinanzierung über 300.000 € bedeutet ein Prozentpunkt mehr Zins rund 250 € höhere Zinskosten im Monat."],
    ["Wohnen und Mieten", "Teure Baukredite bremsen den Wohnungsbau. Werden weniger Wohnungen gebaut, kann das die Mieten mittelfristig weiter nach oben treiben."],
    ["Ausbildung und Einkommen", "Bremst die EZB die Nachfrage, sinken Umsätze und Einstellungen. Für Auszubildende zeigt sich das zuerst bei Übernahmen und offenen Stellen."],
  ];
  const cw = 5.90, ch = 1.80;
  items.forEach((it, k) => {
    const x = LM + (k % 2) * (cw + 0.29), y = 1.52 + Math.floor(k / 2) * 1.95;
    card(s, { x, y, w: cw, h: ch });
    circle(s, { x: x + 0.30, y: y + 0.26, d: 0.58, txt: String(k + 1), size: 17 });
    txt(s, it[0], { x: x + 1.02, y: y + 0.28, w: cw - 1.32, h: 0.32, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
    txt(s, it[1], { x: x + 1.02, y: y + 0.68, w: cw - 1.32, h: 1.00, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.06 });
  });

  card(s, { x: LM, y: 5.65, w: CW, h: 0.95, fill: NAVY, line: null });
  txt(s, "Kurz gesagt", { x: LM + 0.35, y: 5.78, w: CW - 0.70, h: 0.28, fontSize: 13, bold: true, color: GOLD });
  txt(s, "Steigende Zinsen begünstigen alle, die Geld anlegen – und belasten alle, die Geld leihen müssen. Eine Zinsentscheidung verteilt Einkommen zwischen Sparern und Schuldnern um.", {
    x: LM + 0.35, y: 6.06, w: CW - 0.70, h: 0.48, fontSize: 12.5, color: LIGHTTXT,
  });

  foot(s, NUM());
  s.addNotes("Direkter Bezug zur Klasse. Das Rechenbeispiel 300.000 Euro / 250 Euro im Monat unbedingt nennen - "
    + "es macht die Groessenordnung greifbar. Den Realzins kurz erklaeren, das ist der haeufigste Denkfehler. "
    + "(ca. 90 Sekunden)");
}

/* ====================== FOLIE 15 · Wirkung: Unternehmen ================== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 3 · Wirkung auf die Akteure", "Was der Zins mit Unternehmen macht", false);

  card(s, { x: LM, y: 1.52, w: 6.40, h: 2.38 });
  txt(s, "Jede Investition wird neu gerechnet", { x: LM + 0.32, y: 1.72, w: 5.76, h: 0.34, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Ein Betrieb investiert nur, wenn die erwartete Rendite über den Finanzierungskosten liegt. Steigt der Zins, fallen Projekte aus der Rechnung: die neue Maschine, die Photovoltaikanlage auf dem Hallendach, die zusätzliche Filiale.", {
    x: LM + 0.32, y: 2.12, w: 5.76, h: 1.10, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });
  txt(s, "Faustregel: Der Zins ist die Messlatte, die jedes Projekt überspringen muss.", {
    x: LM + 0.32, y: 3.26, w: 5.76, h: 0.42, fontSize: 12, italic: true, color: GOLD_DARK,
  });

  card(s, { x: LM, y: 4.05, w: 6.40, h: 2.35 });
  txt(s, "Kleine Betriebe trifft es zuerst", { x: LM + 0.32, y: 4.25, w: 5.76, h: 0.34, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, "Handwerk, Bau und Einzelhandel finanzieren sich stark über Bankkredite und haben weniger Rücklagen. Große Konzerne können sich am Kapitalmarkt häufig günstiger Geld beschaffen – eine Zinserhöhung wirkt deshalb nicht auf alle gleich.", {
    x: LM + 0.32, y: 4.65, w: 5.76, h: 1.45, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08,
  });

  const rx = 7.30, rw = 5.41;
  card(s, { x: rx, y: 1.52, w: rw, h: 2.38 });
  txt(s, "Besonders zinsempfindlich", { x: rx + 0.32, y: 1.72, w: rw - 0.64, h: 0.34, fontSize: 15.5, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, bullets([
    "Bauwirtschaft und Immobilienbranche",
    "Automobil- und Maschinenbau (Großaufträge auf Kredit)",
    "Start-ups und junge Unternehmen ohne Eigenkapital",
    "Betriebe mit hohem Lagerbestand",
  ]), { x: rx + 0.32, y: 2.12, w: rw - 0.64, h: 1.60, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.08 });

  card(s, { x: rx, y: 4.05, w: rw, h: 2.35, fill: NAVY, line: null });
  txt(s, "Folge für den Arbeitsmarkt", { x: rx + 0.32, y: 4.25, w: rw - 0.64, h: 0.34, fontSize: 15.5, bold: true, fontFace: HEAD, color: GOLD });
  txt(s, "Weniger Investitionen bedeuten weniger Aufträge, weniger neue Stellen und weniger Ausbildungsplätze. Umgekehrt kann billiges Geld einen Aufschwung verstärken – genau das ist von der EZB beabsichtigt.", {
    x: rx + 0.32, y: 4.65, w: rw - 0.64, h: 1.45, fontSize: 12.5, color: LIGHTTXT, lineSpacingMultiple: 1.08,
  });

  foot(s, NUM());
  s.addNotes("Bezug zum Ausbildungsbetrieb der Klasse herstellen: Auch dort werden Investitionen ueber Kredite "
    + "finanziert. Die Verbindung Investition - Auftraege - Ausbildungsplaetze deutlich machen. (ca. 75 Sekunden)");
}

/* ============== FOLIE 16 · Wirkung: Staat und Aussenwirtschaft =========== */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Teil 3 · Wirkung auf die Akteure", "Was der Zins mit Staat und Außenhandel macht", false);

  const cols = [
    ["Der Staat als größter Schuldner", "Bund, Länder und Kommunen finanzieren sich über Anleihen. Steigt der Zins, steigen die Zinsausgaben im Haushalt – Geld, das dann für Schulen, Straßen oder Digitalisierung fehlt."],
    ["Wechselkurs und Export", "Höhere Zinsen machen Anlagen in Euro attraktiver, der Euro wertet auf. Importe wie Öl und Rohstoffe werden billiger, deutsche Exporte für Kunden außerhalb der Eurozone aber teurer."],
    ["Banken und Vermögen", "Steigende Zinsen erhöhen die Zinsmarge der Banken, entwerten aber ältere Anleihen in ihren Büchern. Auch Aktien- und Immobilienpreise reagieren: niedrige Zinsen treiben sie, hohe bremsen sie."],
  ];
  const cw = 3.83;
  cols.forEach((c, k) => {
    const x = LM + k * (cw + 0.30), y = 1.52, h = 3.24;
    card(s, { x, y, w: cw, h });
    circle(s, { x: x + 0.32, y: y + 0.30, d: 0.62, txt: String(k + 1), size: 19 });
    txt(s, c[0], { x: x + 0.32, y: y + 1.06, w: cw - 0.64, h: 0.66, fontSize: 15, bold: true, fontFace: HEAD, color: NAVY });
    txt(s, c[1], { x: x + 0.32, y: y + 1.72, w: cw - 0.64, h: 1.34, fontSize: 12, color: INK, lineSpacingMultiple: 1.06 });
  });

  card(s, { x: LM, y: 5.00, w: CW, h: 1.55, fill: NAVY, line: null });
  txt(s, "Ein Zins für viele Länder – der Dauerkonflikt der Währungsunion", {
    x: LM + 0.35, y: 5.16, w: CW - 0.70, h: 0.34, fontSize: 15.5, bold: true, fontFace: HEAD, color: GOLD,
  });
  txt(s, "Die EZB setzt einen einzigen Leitzins für alle Mitgliedstaaten fest. Hoch verschuldete Staaten und schwache Volkswirtschaften trifft eine Zinserhöhung deutlich härter als Länder mit stabilen Finanzen. Ein Zinssatz, der für das eine Land passt, kann für das andere zu hoch oder zu niedrig sein – das ist der Preis der gemeinsamen Währung.", {
    x: LM + 0.35, y: 5.54, w: CW - 0.70, h: 0.90, fontSize: 12.5, color: LIGHTTXT, lineSpacingMultiple: 1.08,
  });

  foot(s, NUM());
  s.addNotes("Hier wird die politische Dimension sichtbar. Der Punkt 'ein Zins fuer viele Laender' eignet sich "
    + "besonders gut fuer die anschliessende Diskussion und ist der Uebergang zum Fazit. (ca. 75 Sekunden)");
}

/* ========================= FOLIE 17 · Fazit ============================== */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  head(s, "Fazit", "Ein Ziel, viele Zielkonflikte", true);

  const take = [
    ["Die EZB hat einen klaren Auftrag", "Preisstabilität bei rund 2 % Inflation. Alles andere ist ausdrücklich nachrangig – so steht es im EU-Vertrag."],
    ["Ihr stärkstes Werkzeug ist der Zins", "Über Banken, Kredite und Nachfrage wirkt er auf die Preise – langsam, indirekt und nie punktgenau."],
    ["Jede Entscheidung hat Gewinner und Verlierer", "Was Sparerinnen und Sparer freut, belastet Kreditnehmer, Betriebe und den Staat. Geldpolitik ist immer eine Abwägung."],
  ];
  take.forEach((t, k) => {
    const y = 1.58 + k * 1.60;
    card(s, { x: LM, y, w: 7.00, h: 1.46, fill: NAVY_MID, line: null, noShadow: true });
    circle(s, { x: LM + 0.28, y: y + 0.34, d: 0.60, txt: String(k + 1), size: 19 });
    txt(s, t[0], { x: LM + 1.06, y: y + 0.20, w: 5.60, h: 0.56, fontSize: 15.5, bold: true, fontFace: HEAD, color: GOLD });
    txt(s, t[1], { x: LM + 1.06, y: y + 0.78, w: 5.60, h: 0.58, fontSize: 12.5, color: LIGHTTXT, lineSpacingMultiple: 1.08 });
  });

  const rx = 7.95, rw = 4.76;
  card(s, { x: rx, y: 1.58, w: rw, h: 4.66, fill: WHITE, line: null });
  txt(s, "Fragen zur Diskussion", { x: rx + 0.32, y: 1.84, w: rw - 0.64, h: 0.36, fontSize: 16, bold: true, fontFace: HEAD, color: NAVY });
  txt(s, bullets([
    "Sollte die EZB neben stabilen Preisen auch für Beschäftigung und Klimaschutz zuständig sein?",
    "Ist es demokratisch vertretbar, dass eine unabhängige Institution so stark in unser Leben eingreift?",
    "Wer trägt die Last einer Zinserhöhung – und wer profitiert davon?",
    "Was würdet ihr an der Stelle des EZB-Rats heute entscheiden?",
  ]), { x: rx + 0.32, y: 2.36, w: rw - 0.64, h: 3.60, fontSize: 13, color: INK, lineSpacingMultiple: 1.12 });

  txt(s, "Quellen: Vorlesungsskript „Geld und Geldpolitik“ (Prof. Dr. L. Rebeggiani); Europäische Zentralbank; Deutsche Bundesbank; Eurostat.", {
    x: LM, y: 6.52, w: CW, h: 0.32, fontSize: 10, italic: true, color: "8FA2C8",
  });
  txt(s, "Vielen Dank für eure Aufmerksamkeit.", {
    x: LM, y: 6.90, w: 7.0, h: 0.32, fontSize: 12, bold: true, color: GOLD,
  });
  txt(s, NUM(), { x: W - LM - 4.0, y: 6.90, w: 4.0, h: 0.32, fontSize: 9.5, color: "8FA2C8", align: "right" });

  s.addNotes("Die drei Kernaussagen wiederholen - das bleibt haengen. Danach die Diskussionsfragen oeffnen und "
    + "die Klasse einbeziehen. Zum Schluss auf Rueckfragen eingehen. (ca. 90 Sekunden plus Diskussion)");
}

pres.writeFile({ fileName: "/home/user/geldpolitik-ezb/Geldpolitik_der_EZB_Berufsschule.pptx" })
  .then(f => console.log("Erstellt:", f));
