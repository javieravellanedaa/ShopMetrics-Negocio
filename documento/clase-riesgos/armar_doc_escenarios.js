const fs = require("fs");
const {Document, Packer, Paragraph, TextRun, PageBreak, Table, TableRow, TableCell, WidthType, ShadingType,
       HeadingLevel, AlignmentType, BorderStyle, LevelFormat} = require("docx");
const J = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3];

const n = (x, d=2) => (typeof x === "number")
  ? x.toLocaleString("es-AR", {minimumFractionDigits: d, maximumFractionDigits: d}) : String(x);
const pct = x => (x*100).toLocaleString("es-AR", {minimumFractionDigits: 2, maximumFractionDigits: 2}) + " %";
const PF = J.pf, AM = J.amort;
const ANIOS = ["2026", "2027", "2028"];

const P = (t, o={}) => new Paragraph({spacing: {after: 120}, ...o, children: [new TextRun({text: t, font: "Arial", size: 20, ...(o.run||{})})]});
const H1 = t => new Paragraph({heading: HeadingLevel.HEADING_1, spacing: {before: 360, after: 160}, children: [new TextRun({text: t, font: "Arial", size: 30, bold: true})]});
const H2 = t => new Paragraph({heading: HeadingLevel.HEADING_2, spacing: {before: 280, after: 120}, children: [new TextRun({text: t, font: "Arial", size: 24, bold: true})]});
const B = (t) => new Paragraph({numbering: {reference: "vinetas", level: 0}, spacing: {after: 80}, children: [new TextRun({text: t, font: "Arial", size: 20})]});

const border = {style: BorderStyle.SINGLE, size: 4, color: "999999"};
const borders = {top: border, bottom: border, left: border, right: border};
function celda(t, w, opts={}) {
  const runs = Array.isArray(t) ? t : [t];
  return new TableCell({width: {size: w, type: WidthType.DXA}, borders,
    shading: opts.fill ? {type: ShadingType.CLEAR, color: "auto", fill: opts.fill} : undefined,
    margins: {top: 60, bottom: 60, left: 90, right: 90},
    children: runs.map(r => new Paragraph({alignment: opts.align || AlignmentType.LEFT, spacing: {after: 0},
      children: [new TextRun({text: String(r), font: "Arial", size: opts.size || 17, bold: !!opts.bold})]}))});
}
function tabla(anchos, filas, cab) {
  const rows = [];
  if (cab) rows.push(new TableRow({tableHeader: true, children: cab.map((c,i) => celda(c, anchos[i], {bold: true, fill: "E2EFD9"}))}));
  for (const f of filas) rows.push(new TableRow({children: f.map((c,i) => celda(c.t ?? c, anchos[i], c.o || {}))}));
  return new Table({columnWidths: anchos, width: {size: anchos.reduce((a,b)=>a+b,0), type: WidthType.DXA}, rows});
}

// ------------------------------------------------------------- texto comun
const A4W = 11906 - 2*1134;   // ancho util con margenes de 2 cm
const kids = [];
kids.push(new Paragraph({spacing: {after: 80}, children: [new TextRun({text: "Escenarios de riesgo: de dónde sale cada número", font: "Arial", size: 40, bold: true})]}));
kids.push(P("Presupuesto financiero ShopMetrics.xlsx · hojas Escenario 1, Escenario 2, Escenario 3 y Plan de Contingencia", {run: {color: "555555", size: 18}}));
kids.push(P("Cada escenario es el presupuesto financiero traído entero por referencia, con un riesgo aplicado encima. Ninguna celda tiene un número escrito a mano: todas son fórmulas que apuntan a la hoja \"Presupuesto financiero\", multiplican o suman el efecto del riesgo, y recalculan el resto. Si mañana cambia una venta en el modelo de ingresos, los cuatro escenarios cambian solos."));
kids.push(P("La regla de armado es la de la clase: lo que se sigue de aplicar el riesgo va en el escenario y no en el plan de contingencia. Si bajan las ventas, bajan los costos variables en la misma proporción, no se incorpora el recurso que ese nivel de actividad ya no justifica y no se arma su puesto. El plan de contingencia, en cambio, son acciones nuevas que se agregan sobre el peor escenario."));

kids.push(H1("1. Las fórmulas, iguales en los tres escenarios"));
kids.push(P("Todas las filas se calculan igual en las tres hojas. Lo único que cambia de un escenario a otro es el factor que multiplica las ventas y los montos que se suman o restan, y esos salen de la matriz de riesgos. Las referencias son a la hoja \"Presupuesto financiero\" (PF), que a su vez trae los valores de los modelos de ingresos, egresos e inversión."));
kids.push(tabla([2200, 3300, 4138], [
  ["Ingresos", "= PF!fila 12 × factor del riesgo", "PF!E12:G12 viene de Mod. ingresos!C30:E30. El factor (0,90 / 0,85 / 0,60) es la caída de ventas que la matriz cuantifica para ese riesgo. El año 2026 no se toca: el riesgo se aplica a los ejercicios 2 y 3."],
  ["Costos fijos", "= PF!fila 13 (+ medida del plan)", "PF!E13:G13 viene de Mod. egresos!B10:D10. En los escenarios no cambia; sólo el plan de contingencia le suma el costo de sus medidas."],
  ["Costos variables", "= PF!fila 14 × mismo factor + reposición", "PF!E14:G14 viene de Mod. egresos!B11:D11. Bajan en la misma proporción que las ventas porque son variables; se les suma la reposición y reinstalación de equipamiento que el riesgo provoca."],
  ["Costos de RRHH", "= PF!fila 15 − puesto no incorporado", "PF!E15:G15 viene de Mod. egresos!B12:D12. Cuando el nivel de ventas no justifica el segundo ejecutivo comercial, se resta su costo anual: Costos RRHH fila 19, sueldo bruto 1.000 × coeficiente de cargas 1,2945 × 12 meses = 15.534."],
  ["UAII", "= Ingresos − Costos fijos − Costos variables − Costos de RRHH", "Misma fórmula que PF!fila 16."],
  ["Impuesto Ingresos Brutos", "= Ingresos × 3 %", "Misma alícuota que PF!fila 17 (celda C9 de cada escenario)."],
  ["Monto imponible", "= UAII − IIB − Amortización del ejercicio", "Igual que PF!L23:N23. La amortización sale de Amortizaciones!I5:K5: " + n(AM[0]) + " / " + n(AM[1]) + " / " + n(AM[2]) + ". Queda en la fila 22 de cada escenario, fuera del cuadro, como en el presupuesto."],
  ["Impuesto a las ganancias", "2026: 0 · 2027: MAX(0; MI 2026) × 35 % · 2028: MAX(0; MI 2027) × 35 %", "Diferido un año, igual que PF!F19 y PF!G19 (=MAX(0,$L$23)*0.35). Se recalcula dentro del escenario y no se copia del presupuesto: si baja la utilidad, baja el impuesto. El MAX deja en cero el año con quebranto."],
  ["UDII", "= UAII − IIB − Ganancias", "Misma fórmula que PF!fila 20."],
  ["Inversión", "= PF!fila 21 (+ puesto no armado)", "PF!D21:G21 viene de Mod. inversión (fila 20 del bloque J:M). Si no se incorpora el recurso, no se arma su puesto: notebook 940, escritorio 100, silla 85 y periféricos según Amortizaciones, unos 1.200, que se suman porque la inversión está en negativo."],
  ["Flujo de Fondos", "= UDII + Inversión", "Misma fórmula que PF!fila 22. El año Inicio es sólo la inversión inicial, PF!D21 = " + n(PF["Inversión"][0], 0) + "."],
  ["Tasa de corte", "= PF!G33", n(PF.tasa*100, 0) + " %. La misma del presupuesto: no se mueve para que los riesgos muerdan."],
  ["VAN", "= VNA(tasa; FF 2026:2028) + FF Inicio", "Misma fórmula que PF!G34: el desembolso inicial no se descuenta, se suma."],
  ["TIR", "= TIR(FF Inicio:2028)", "Misma fórmula que PF!G35."],
], ["Renglón", "Cómo se calcula en la hoja", "De dónde sale"]));

// ------------------------------------------------------------ el presupuesto
kids.push(H1("2. El punto de partida: el presupuesto financiero"));
kids.push(P("Estos son los valores de la hoja \"Presupuesto financiero\" que los escenarios toman como base. Son los que ya estaban; ningún escenario los modifica."));
const fila = (k, d=2) => [k, ...PF[k].map(x => ({t: n(x,d), o:{align: AlignmentType.RIGHT}}))];
kids.push(tabla([2600, 1700, 1800, 1800, 1738], [
  fila("Ingresos"), fila("Costos fijos"), fila("Costos variables"), fila("Costos de RRHH"),
  fila("UAII"), fila("IIB"), fila("Ganancias"), fila("UDI"), fila("Inversión"), fila("FF"),
], ["Presupuesto financiero", "Inicio", "2026", "2027", "2028"]));
kids.push(P("Tasa de corte " + n(PF.tasa*100,0) + " %  ·  VAN " + n(PF.VAN,0) + "  ·  TIR " + pct(PF.TIR), {run: {bold: true}, spacing: {before: 120, after: 200}}));

// ------------------------------------------------------------ cada escenario
const SUP = {
  "Escenario 1": {factor: 0.90, cv_extra: 3000, rrhh: 0, inv: 0,
    porque: ["El factor 0,90 sale del efecto cuantitativo de R1 en la matriz: una caída de 72 horas bonifica el abono del período y una de cada diez cuentas se da de baja en el trimestre, de modo que el ingreso de 2027 y 2028 queda un 10 % por debajo.",
             "Los 3.000 por ejercicio son la recuperación: horas de DevOps y créditos a los comercios afectados. Van en costos variables porque dependen del incidente, no de la estructura.",
             "RRHH e inversión no cambian: una caída no reduce la dotación; si algo, la exige más."]},
  "Escenario 2": {factor: 0.85, cv_extra: 12000, rrhh: 0, inv: 0,
    porque: ["El factor 0,85 sale del efecto cuantitativo de R2: los comercios que se dan de baja antes del año dejan de pagar el abono que la proyección daba por cobrado, un 15 % del ingreso recurrente de 2027 y 2028.",
             "Los 12.000 por ejercicio son la reposición y el reacondicionamiento del equipamiento retirado, que vuelve en comodato pero hay que volver a instalarlo en otro local.",
             "RRHH no baja: sostener la base rotando exige el mismo trabajo de instalación y posventa que crecer."]},
  "Escenario 3": {factor: 0.60, cv_extra: 15000, rrhh: 15534, inv: 1200,
    porque: ["Es el escenario combinado que pide la clase cuando ningún riesgo suelto rompe el proyecto: con 61,7 puntos de colchón sobre la tasa de corte, hay que apilar riesgos. Se juntan R1 (caída de la plataforma), R2 (bajas anticipadas) y R10 (competidor con precio agresivo): la caída hace perder cuentas, las bajas no se reponen y el competidor presiona el precio. El 40 % de caída es el efecto de los tres a la vez, no la suma aritmética de cada uno por separado.",
             "Los 15.000 por ejercicio son reposición y reinstalación de equipamiento, más altos que en el escenario 2 porque las bajas son más.",
             "A ese nivel de altas no se incorpora el segundo ejecutivo comercial: se restan 15.534 por ejercicio (Costos RRHH fila 19: 1.000 × 1,2945 × 12). El técnico de instalación se mantiene, porque con rotación hay más altas y retiros que atender, no menos.",
             "Tampoco se arma el puesto de trabajo de ese recurso: notebook, escritorio, silla y periféricos, unos 1.200 que se suman a la inversión de 2027 y 2028 porque está en negativo."]},
};
const ORDEN = ["Ingresos","Costos fijos","Costos variables","Costos de RRHH","UAII","IIB","Monto imponible","Ganancias","UDII","Inversión","FF"];
const NOM = {IIB: "Impuesto Ingresos Brutos", Ganancias: "Impuesto a las ganancias 35 %", FF: "Flujo de Fondos"};
let num = 3;
for (const nombre of ["Escenario 1","Escenario 2","Escenario 3"]) {
  const E = J.escenarios[nombre], s = SUP[nombre];
  kids.push(H1(`${num++}. ${nombre} · ${E.titulo.replace(/^RIESGO: /, "")}`));
  kids.push(H2("Qué se aplica y por qué"));
  for (const t of s.porque) kids.push(B(t));
  kids.push(H2("La cuenta, renglón por renglón"));
  kids.push(P("Cada celda muestra la operación y su resultado. Las bases son las del presupuesto financiero de la sección 2."));
  const v = k => E.filas[k].map(p => p[1]);
  const base = k => PF[{Ingresos:"Ingresos","Costos fijos":"Costos fijos","Costos variables":"Costos variables","Costos de RRHH":"Costos de RRHH","Inversión":"Inversión"}[k]];
  const filas = [];
  for (const k of ORDEN) {
    const r = [NOM[k] || k];
    for (let i = 1; i <= 3; i++) {
      const val = v(k)[i], yr = ANIOS[i-1], f = i === 1 ? 1 : s.factor;
      let cuenta = "";
      if (k === "Ingresos") cuenta = i === 1 ? "sin cambio" : `${n(base(k)[i],0)} × ${n(f)}`;
      else if (k === "Costos fijos") cuenta = "sin cambio";
      else if (k === "Costos variables") cuenta = i === 1 ? "sin cambio" : `${n(base(k)[i])} × ${n(f)}${s.cv_extra ? " + " + n(s.cv_extra,0) : ""}`;
      else if (k === "Costos de RRHH") cuenta = (i === 1 || !s.rrhh) ? "sin cambio" : `${n(base(k)[i])} − ${n(s.rrhh,0)}`;
      else if (k === "UAII") cuenta = `${n(v("Ingresos")[i])} − ${n(v("Costos fijos")[i])} − ${n(v("Costos variables")[i])} − ${n(v("Costos de RRHH")[i])}`;
      else if (k === "IIB") cuenta = `${n(v("Ingresos")[i])} × 3 %`;
      else if (k === "Monto imponible") cuenta = `${n(v("UAII")[i])} − ${n(v("IIB")[i])} − ${n(AM[i-1])}`;
      else if (k === "Ganancias") cuenta = i === 1 ? "primer ejercicio: 0" : `MAX(0; ${n(v("Monto imponible")[i-1])}) × 35 %`;
      else if (k === "UDII") cuenta = `${n(v("UAII")[i])} − ${n(v("IIB")[i])} − ${n(v("Ganancias")[i])}`;
      else if (k === "Inversión") cuenta = (i === 1 || !s.inv) ? "sin cambio" : `${n(base(k)[i],0)} + ${n(s.inv,0)}`;
      else if (k === "FF") cuenta = `${n(v("UDII")[i])} + (${n(v("Inversión")[i],0)})`;
      r.push({t: [cuenta, "= " + n(val)], o: {align: AlignmentType.RIGHT, size: 15}});
    }
    filas.push(r);
  }
  kids.push(tabla([2000, 2546, 2546, 2546], filas, ["Renglón", "2026", "2027", "2028"]));
  const ff = v("FF");
  kids.push(P(`Flujo de fondos: ${n(E.filas.FF[0][1],0)} · ${n(ff[1],0)} · ${n(ff[2],0)} · ${n(ff[3],0)}`, {spacing: {before: 160, after: 40}}));
  kids.push(P(`Tasa de corte ${n(E.tasa[1]*100,0)} %  ·  VAN = VNA(30 %; 2026:2028) + Inicio = ${n(E.VAN[1],0)}  ·  TIR = ${pct(E.TIR[1])}`, {run: {bold: true}}));
  const resiste = E.TIR[1] >= E.tasa[1];
  kids.push(P(resiste ? `La TIR queda por encima de la tasa de corte y el VAN positivo: el proyecto absorbe este riesgo.`
                      : `La TIR cae por debajo de la tasa de corte y el VAN queda negativo: el proyecto no resiste. Es el peor de los tres y es el que sostiene el plan de contingencia.`));
}

kids.push(new Paragraph({children: [new PageBreak()]}));
kids.push(H1("6. Resumen"));
kids.push(tabla([3000, 1600, 1600, 1600, 1838], [
  ["Presupuesto financiero", {t: n(PF.tasa*100,0)+" %", o:{align:AlignmentType.RIGHT}}, {t: n(PF.VAN,0), o:{align:AlignmentType.RIGHT}}, {t: pct(PF.TIR), o:{align:AlignmentType.RIGHT}}, "el proyecto sano"],
  ...["Escenario 1","Escenario 2","Escenario 3"].map(k => { const E = J.escenarios[k]; const ok = E.TIR[1] >= E.tasa[1];
    return [k, {t: n(E.tasa[1]*100,0)+" %", o:{align:AlignmentType.RIGHT}}, {t: n(E.VAN[1],0), o:{align:AlignmentType.RIGHT}}, {t: pct(E.TIR[1]), o:{align:AlignmentType.RIGHT}}, ok ? "resiste" : "no resiste"]; }),
], ["Hoja", "Tasa", "VAN", "TIR", "Lectura"]));
kids.push(P("Dos resisten y uno no, igual que en el ejemplo de la cátedra. El plan de contingencia se arma sobre el escenario 3 y lo lleva a 31,34 %: apenas por encima de la tasa de corte y muy por debajo del 91,70 % del presupuesto, porque un plan nivela, no mejora.", {spacing: {before: 160}}));

const doc = new Document({
  numbering: {config: [{reference: "vinetas", levels: [{level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: {paragraph: {indent: {left: 500, hanging: 260}}}}]}]},
  styles: {default: {document: {run: {font: "Arial", size: 20}}}},
  sections: [{properties: {page: {margin: {top: 1134, bottom: 1134, left: 1134, right: 1134}}}, children: kids}],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("  escrito:", OUT, b.length, "bytes"); });
