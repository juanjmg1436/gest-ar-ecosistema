const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, ShadingType, Footer, PageNumber,
  VerticalAlign, HeadingLevel
} = require('docx');
const fs = require('fs');
const path = require('path');

// ---- Paleta (la misma de las laminas y del material conceptual) ----
const C = {
  dark:   '0F172A',
  text:   '334155',
  muted:  '64748B',
  soft:   '94A3B8',
  line:   'E2E8F0',
  bg:     'F8FAFC',
  bgalt:  'F1F5F9',
  azul:   '2563EB',
  verde:  '059669',
  viole:  '9333EA',
  ambar:  'B45309',
};
const F = 'Segoe UI';

const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const hair = (color) => ({ style: BorderStyle.SINGLE, size: 2, color });

const run = (text, o = {}) => new TextRun({
  text, font: F, size: o.size || 21, color: o.color || C.text,
  bold: !!o.bold, italics: !!o.italics, allCaps: !!o.caps,
  characterSpacing: o.spacing || 0, language: { value: 'es-AR' },
});

const p = (text, o = {}) => new Paragraph({
  children: Array.isArray(text) ? text : [run(text, o)],
  alignment: o.align,
  spacing: { before: o.before || 0, after: o.after === undefined ? 140 : o.after, line: o.line || 290 },
  border: o.border,
  keepNext: !!o.keepNext,
  keepLines: !!o.keepLines,
  widowControl: true,
});

const bullet = (text, o = {}) => new Paragraph({
  children: Array.isArray(text) ? text : [run(text, o)],
  bullet: { level: 0 },
  keepLines: true,
  widowControl: true,
  spacing: { before: 0, after: 70, line: 280 },
});

const eyebrow = (text, color) => p([run(text, { caps: true, bold: true, size: 16, color: color || C.muted, spacing: 14 })], { after: 90, keepNext: true, keepLines: true });

const h1 = (num, text, color) => new Paragraph({
  children: [run(num + '  ', { bold: true, size: 28, color }), run(text, { bold: true, size: 28, color: C.dark })],
  heading: HeadingLevel.HEADING_1,
  keepNext: true, keepLines: true, widowControl: true,
  spacing: { before: 380, after: 170, line: 300 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.line, space: 8 } },
});

const h2 = (text, color) => new Paragraph({
  children: [run(text, { bold: true, size: 22, color: color || C.dark })],
  heading: HeadingLevel.HEADING_2,
  keepNext: true, keepLines: true, widowControl: true,
  spacing: { before: 240, after: 90, line: 290 },
});

const sp = (h) => new Paragraph({ children: [], keepNext: true, spacing: { after: h === undefined ? 160 : h } });

// ── Ficha de toma: encabezado de color + dos columnas (pantalla / voz) ──────
const toma = (n, tiempo, titulo, color, pantalla, voz, prep) => new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  columnWidths: [4100, 5100],
  borders: {
    top: hair(C.line), bottom: hair(C.line), left: hair(C.line), right: hair(C.line),
    insideHorizontal: hair(C.line), insideVertical: hair(C.line),
  },
  rows: [
    new TableRow({
      cantSplit: true,
      children: [new TableCell({
        columnSpan: 2,
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: color },
        margins: { top: 130, bottom: 130, left: 190, right: 190 },
        children: [
          p([
            run('TOMA ' + n, { bold: true, size: 22, color: 'FFFFFF' }),
            run('   ·   ' + tiempo, { size: 19, color: 'FFFFFF' }),
          ], { after: 20, keepNext: true }),
          p([run(titulo, { bold: true, size: 19, color: 'FFFFFF' })], { after: 0, keepNext: true }),
        ],
      })],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          verticalAlign: VerticalAlign.TOP,
          margins: { top: 150, bottom: 150, left: 190, right: 150 },
          children: [eyebrow('En pantalla', color), ...pantalla],
        }),
        new TableCell({
          verticalAlign: VerticalAlign.TOP,
          shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.bg },
          margins: { top: 150, bottom: 150, left: 190, right: 150 },
          children: [
            eyebrow('Voz en off', color),
            ...voz.map(t => p([run('«' + t + '»', { italics: true, size: 20, color: C.dark })], { after: 90, line: 275 })),
          ],
        }),
      ],
    }),
    ...(prep ? [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        columnSpan: 2,
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.bgalt },
        margins: { top: 120, bottom: 120, left: 190, right: 190 },
        children: [p([
          run('Preparar antes: ', { bold: true, size: 18, color: C.dark }),
          run(prep, { size: 18, color: C.text }),
        ], { after: 0, line: 260 })],
      })],
    })] : []),
  ],
});

// ── Planilla de rodaje ─────────────────────────────────────────────────────
const planilla = (filas) => new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  columnWidths: [900, 1900, 4200, 2200],
  borders: {
    top: hair(C.line), bottom: hair(C.line), left: NONE, right: NONE,
    insideHorizontal: hair(C.line), insideVertical: NONE,
  },
  rows: [
    new TableRow({
      tableHeader: true, cantSplit: true,
      children: ['Toma', 'Tiempo', 'Qué se ve', 'Duración'].map((t, i) => new TableCell({
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.bgalt },
        margins: { top: 110, bottom: 110, left: i === 0 ? 140 : 150, right: 140 },
        children: [p([run(t, { caps: true, bold: true, size: 16, color: C.muted, spacing: 12 })], { after: 0, keepNext: true })],
      })),
    }),
    ...filas.map(([a, b, c, d]) => new TableRow({
      cantSplit: true,
      children: [a, b, c, d].map((v, i) => new TableCell({
        margins: { top: 110, bottom: 110, left: i === 0 ? 140 : 150, right: 140 },
        children: [p([run(v, { size: 18, bold: i === 0, color: i === 0 ? C.dark : C.text })], { after: 0, line: 250 })],
      })),
    })),
  ],
});

const destacado = (children, color) => new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: { top: NONE, bottom: NONE, right: NONE,
             left: { style: BorderStyle.SINGLE, size: 18, color },
             insideHorizontal: NONE, insideVertical: NONE },
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.bg },
    margins: { top: 170, bottom: 170, left: 220, right: 200 },
    children,
  })] })],
});

// ============================ CONTENIDO ============================
const children = [];

children.push(eyebrow('Video de presentación · Guion técnico'));
children.push(p([run('GEST-AR en cuatro minutos', { bold: true, size: 40, color: C.dark })], { after: 80, line: 340 }));
children.push(p([run('Screencast con voz en off: qué se ve, qué se hace y qué se dice, toma por toma', { size: 23, color: C.muted })],
  { after: 150, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.dark, space: 10 } } }));
children.push(p([run('Juan Manuel Gómez  ·  Profesor en Ciencias Económicas  ·  Licenciado en Educación', { size: 17, color: C.muted })], { after: 30, line: 250 }));
children.push(p([run('Escuela Normal Superior N.º 9 y Anexo Comercial Juan XXIII — Wanda, Misiones  ·  2026', { size: 17, color: C.muted })], { after: 60, line: 250 }));

// ── 1. Antes de grabar ──
children.push(h1('1', 'Antes de grabar', C.azul));
children.push(p('El video dura cuatro minutos, pero se gana o se pierde en la preparación. La regla es una sola: en cámara no se completa ningún formulario largo ni se espera ninguna carga. Todo lo que no sea el gesto que querés mostrar tiene que estar hecho de antemano.'));

children.push(h2('Los datos', C.azul));
children.push(bullet([run('Usá un contribuyente de demostración, nunca datos de alumnos. ', { bold: true, size: 21, color: C.dark }), run('En la base hay 37 perfiles reales con nombre y CUIT. Creá uno aparte para filmar.', { size: 21 })]));
children.push(bullet([run('En TRIBUT.AR: ', { bold: true, size: 21, color: C.dark }), run('empresa ya constituida, con régimen elegido e impuestos inscriptos, pero con el punto de venta todavía SIN habilitar — ese es el paso que se muestra en cámara.', { size: 21 })]));
children.push(bullet([run('En PyMEZ 360: ', { bold: true, size: 21, color: C.dark }), run('productos y clientes ya cargados, y la empresa SIN vincular, para que el bloqueo de la primera toma aparezca de verdad.', { size: 21 })]));
children.push(bullet([run('Ventas del mes anterior ya registradas en PyMEZ, ', { bold: true, size: 21, color: C.dark }), run('para que la declaración jurada de la toma 6 traiga números reales y no ceros.', { size: 21 })]));

children.push(h2('La grabación', C.verde));
children.push(bullet('OBS a 1920×1080 y 30 fps. Ventana del navegador maximizada, sin barra de marcadores ni pestañas que muestren datos personales.'));
children.push(bullet('Zoom del navegador al 110–125 %: buena parte del público lo va a ver en el celular, y el texto de las tablas es chico.'));
children.push(bullet('Silenciá notificaciones de Windows y del correo antes de empezar.'));
children.push(bullet([run('Grabá la voz por separado y en una sola toma. ', { bold: true, size: 21, color: C.dark }), run('Es mucho más fácil repetir el audio que volver a filmar la pantalla.', { size: 21 })]));
children.push(bullet('Ensayá el recorrido completo dos veces antes de grabar. La tercera es la buena.'));
children.push(sp(120));
children.push(destacado([
  p([run('Si algo sale mal, no repitas los cuatro minutos.', { bold: true, size: 22, color: C.dark })], { after: 60, line: 270 }),
  p([run('Grabá cada toma como un clip independiente y unilos al final. Ocho clips cortos se rehacen sin drama; una toma única de cuatro minutos se rompe siempre en el minuto tres.', { size: 19 })], { after: 0, line: 265 }),
], C.ambar));

// ── 2. Planilla ──
children.push(h1('2', 'Planilla de rodaje', C.verde));
children.push(p('Ocho tomas, cuatro minutos exactos. El corazón del video son las tomas 3 a 6: es lo único que ninguna otra herramienta hace.', { after: 130, keepNext: true }));
children.push(planilla([
  ['1', '0:00 – 0:22', 'El sistema no lo deja facturar', '22 s'],
  ['2', '0:22 – 0:48', 'Qué es GEST-AR: las tres apps', '26 s'],
  ['3', '0:48 – 1:20', 'TRIBUT.AR: el trámite y el código', '32 s'],
  ['4', '1:20 – 1:50', 'PyMEZ 360: se destraba y factura', '30 s'],
  ['5', '1:50 – 2:20', 'El asiento aparece solo', '30 s'],
  ['6', '2:20 – 2:55', 'La DDJJ se arma con sus ventas', '35 s'],
  ['7', '2:55 – 3:30', 'Por qué funciona: el enfoque', '35 s'],
  ['8', '3:30 – 4:00', 'Cierre y acceso', '30 s'],
]));

// ── 3. Guion ──
children.push(h1('3', 'El guion, toma por toma', C.viole));

children.push(toma('1', '0:00 – 0:22  ·  22 s', 'El sistema no lo deja facturar', C.azul,
  [
    p('PyMEZ 360, pestaña Ventas. Clic en “Nueva venta”.'),
    p('Aparece el bloqueo: no se puede registrar la venta sin punto de venta habilitado.'),
    p([run('Quedate dos segundos quieto sobre el mensaje. ', { bold: true, size: 20, color: C.dark }), run('Sin mover el mouse. Que se lea.', { size: 20 })], { after: 0 }),
  ],
  [
    'Este alumno está por hacer su primera venta. Y el sistema no lo deja.',
    'No es un error. Todavía no habilitó el punto de venta que lo autoriza a facturar.',
    'En GEST-AR las cosas pasan en el orden en que pasan en la realidad.',
  ],
  'La empresa en PyMEZ tiene que estar sin vincular. Si ya la vinculaste para probar, desvinculala antes.'));
children.push(sp(180));

children.push(toma('2', '0:22 – 0:48  ·  26 s', 'Qué es GEST-AR', C.verde,
  [
    p('Corte al portal. Logo y las tres aplicaciones.'),
    p('Pasá el mouse por cada tarjeta mientras la nombrás, sin hacer clic.'),
    p([run('No abras menús. ', { bold: true, size: 20, color: C.dark }), run('Acá solo se presenta, no se recorre.', { size: 20 })], { after: 0 }),
  ],
  [
    'GEST-AR es un ecosistema de simuladores para enseñar gestión administrativa.',
    'Tres aplicaciones que trabajan sobre una misma empresa ficticia: TRIBUT.AR, el marco fiscal. PyMEZ 360, la operación diaria. Sueldos 360, las relaciones laborales.',
    'No son tres ejercicios sueltos: comparten los datos.',
  ],
  null));
children.push(sp(180));

children.push(toma('3', '0:48 – 1:20  ·  32 s', 'TRIBUT.AR: el trámite y el código', C.azul,
  [
    p('TRIBUT.AR. Mostrá al pasar la empresa ya constituida y sus impuestos inscriptos.'),
    p('Entrá a Puntos de venta y habilitá uno.'),
    p('Aparece el código, formato PV1-XXXXXX. Copialo con el botón.'),
    p([run('Pausa de dos segundos sobre el código.', { bold: true, size: 20, color: C.dark })], { after: 0 }),
  ],
  [
    'Volvamos al principio. En TRIBUT.AR la empresa se constituye, elige su régimen e inscribe los impuestos que le corresponden.',
    'Y habilita un punto de venta. Eso genera un código.',
    'Hasta acá, para el alumno, esto es un trámite más.',
  ],
  'Dejá el punto de venta creado pero sin habilitar, así el clic en cámara es uno solo.'));
children.push(sp(180));

children.push(toma('4', '1:20 – 1:50  ·  30 s', 'PyMEZ 360: se destraba', C.verde,
  [
    p('PyMEZ 360 → Mi Empresa → Habilitación de punto de venta. Pegá el código. Confirmación en verde.'),
    p('Volvé a Ventas y cargá la venta que antes no podía hacer. Emitila.'),
    p([run('Este es el momento del video. ', { bold: true, size: 20, color: C.dark }), run('No lo apures.', { size: 20 })], { after: 0 }),
  ],
  [
    'Pega el código en PyMEZ… y recién ahora puede facturar.',
    'El trámite dejó de ser un trámite: era la condición para trabajar.',
    'Esa es la diferencia entre leerlo en un apunte y necesitarlo.',
  ],
  'Tené el producto y el cliente ya cargados: en cámara solo elegís de la lista y confirmás.'));
children.push(sp(180));

children.push(toma('5', '1:50 – 2:20  ·  30 s', 'El asiento aparece solo', C.ambar,
  [
    p('El comprobante emitido y, debajo, el asiento que generó la venta.'),
    p('Señalá con el mouse cada línea mientras la nombrás: Deudores por ventas, Ventas, IVA Débito Fiscal.'),
    p('Cerrá con un vistazo rápido al Libro Diario y al Mayor de la cuenta Ventas.'),
  ],
  [
    'Y con la venta, aparece el asiento. Nadie se lo pidió: es la consecuencia contable de lo que acaba de hacer.',
    'Deudores por ventas al debe. Ventas e IVA débito fiscal al haber.',
    'La contabilidad deja de ser un ejercicio en papel: es el registro de una decisión que tomó él.',
  ],
  null));
children.push(sp(180));

children.push(toma('6', '2:20 – 2:55  ·  35 s', 'La declaración jurada se arma con sus ventas', C.viole,
  [
    p('TRIBUT.AR → Misiones → Ingresos Brutos → Nueva DDJJ.'),
    p('Pegá el código de sincronización de la empresa. La base imponible se carga sola.'),
    p([run('Si podés, dejá el Mayor de Ventas abierto en otra pestaña y mostrá que es el mismo número. ', { bold: true, size: 20, color: C.dark }), run('Es el argumento más fuerte del video.', { size: 20 })], { after: 0 }),
  ],
  [
    'Un mes después, la declaración jurada de Ingresos Brutos.',
    'El alumno importa sus ventas, y la base imponible ya está. Es el saldo de su cuenta Ventas.',
    'Nadie le dictó ese número: salió de lo que él vendió. Y si se equivocó en marzo, lo va a ver en abril.',
  ],
  'Necesitás ventas cargadas en el período que vas a declarar. Sin eso, la DDJJ viene en cero y la toma no dice nada.'));
children.push(sp(180));

children.push(toma('7', '2:55 – 3:30  ·  35 s', 'Por qué funciona', C.viole,
  [
    p('Corte a la lámina del enfoque neuro-tecno-pedagógico (ya la tenés en Laminas_Exposicion_Docentes).'),
    p('Zoom lento sobre el bucle Desafío → Decisión → Consecuencia → Feedback → Logro, y después sobre la fila de recursos y mecanismos.'),
    p([run('Imagen fija. ', { bold: true, size: 20, color: C.dark }), run('Acá la voz lleva el peso.', { size: 20 })], { after: 0 }),
  ],
  [
    'Nada de esto es casual. Cada recurso de la plataforma responde a un mecanismo cognitivo.',
    'Las etapas que se desbloquean dosifican la carga: nadie liquida IVA antes de saber facturar. El error sin costo real habilita explorar. Los datos que se arrastran fuerzan la transferencia.',
    'No se gamifica el premio: se gamifica el proceso de pensar.',
  ],
  null));
children.push(sp(180));

children.push(toma('8', '3:30 – 4:00  ·  30 s', 'Cierre y acceso', C.azul,
  [
    p('Las tres aplicaciones en pantalla, con sus direcciones o los códigos QR de la presentación.'),
    p('Placa final con tu nombre, la institución y el año.'),
    p([run('Dejá la placa tres segundos en silencio antes de cortar.', { bold: true, size: 20, color: C.dark })], { after: 0 }),
  ],
  [
    'GEST-AR funciona en línea, sin instalación, con datos ficticios y sin validez fiscal.',
    'Sirve para el aula, para diplomaturas y para formación profesional.',
    'El ciclo administrativo completo: del hecho económico al estado contable.',
  ],
  'Los QR ya están generados en la carpeta de la presentación del módulo de impuestos.'));

// ── 4. Notas ──
children.push(h1('4', 'Notas de locución', C.ambar));
children.push(p('El texto está calculado a un ritmo de unas 150 palabras por minuto, que es una locución tranquila. Si al ensayar te sobra tiempo, no aceleres: usá el aire para dejar respirar las pausas marcadas en cada toma.'));
children.push(bullet('Las frases cortas del guion son cortes deliberados. Bajá el tono al final de cada una en lugar de encadenarlas.'));
children.push(bullet('Las tres frases que sostienen el video son: “Y el sistema no lo deja”, “era la condición para trabajar” y “nadie le dictó ese número”. Decilas más lento que el resto.'));
children.push(bullet('No leas los números en pantalla. Si el espectador ya los ve, la voz solo tiene que explicar por qué están ahí.'));

children.push(h2('Variante de 90 segundos', C.verde));
children.push(p('Para redes o para abrir una reunión, hay una versión corta que se arma con las mismas tomas, sin volver a grabar: tomas 1, 4, 6 y 8. Queda el bloqueo, la factura que se destraba, la declaración jurada que se arma sola y el cierre. Es el argumento completo en minuto y medio.'));

// ============================ DOCUMENTO ============================
const doc = new Document({
  creator: 'Juan Manuel Gómez',
  title: 'GEST-AR en cuatro minutos — guion técnico',
  description: 'Guion técnico del video de presentación del ecosistema GEST-AR: pantalla, acción y voz en off toma por toma',
  subject: 'Guion de video institucional del ecosistema GEST-AR',
  keywords: 'GEST-AR, guion, video, screencast, TRIBUT.AR, PyMEZ 360, Sueldos 360',
  lastModifiedBy: 'Juan Manuel Gómez',
  styles: {
    default: { document: { run: { font: F, size: 21, color: C.text, language: { value: 'es-AR' } } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: F, size: 28, bold: true, color: C.dark, language: { value: 'es-AR' } },
        paragraph: { spacing: { before: 380, after: 170, line: 300 }, keepNext: true, keepLines: true, outlineLevel: 0 },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: F, size: 22, bold: true, color: C.dark, language: { value: 'es-AR' } },
        paragraph: { spacing: { before: 240, after: 90, line: 290 }, keepNext: true, keepLines: true, outlineLevel: 1 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1247, bottom: 1134, left: 1304, right: 1304, footer: 567 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 0 },
          children: [
            new TextRun({ text: 'GEST-AR · Guion del video de presentación   ', font: F, size: 15, color: C.soft }),
            new TextRun({ children: [PageNumber.CURRENT], font: F, size: 15, color: C.muted, bold: true }),
          ],
        })],
      }),
    },
    children,
  }],
});

const out = process.argv[2];

// La libreria emite sus propios Heading1/Heading2 ademas de los definidos aca:
// dos estilos con el mismo w:styleId es OOXML invalido y puede hacer que Word
// pida reparar el archivo. Nos quedamos con el ultimo (el propio).
const dedupe = (xml, ids) => {
  for (const id of ids) {
    const re = new RegExp('<w:style [^>]*w:styleId="' + id + '"[^]*?</w:style>', 'g');
    const todos = xml.match(re) || [];
    if (todos.length > 1) { let n = 0; xml = xml.replace(re, (m) => (++n === todos.length ? m : '')); }
  }
  return xml;
};

const docxDir = path.dirname(require.resolve('docx'));
const JSZip = require(require.resolve('jszip', { paths: [docxDir] }));

Packer.toBuffer(doc)
  .then(b => JSZip.loadAsync(b))
  .then(async z => {
    const st = await z.file('word/styles.xml').async('string');
    z.file('word/styles.xml', dedupe(st, ['Heading1', 'Heading2']));
    return z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  })
  .then(b => { fs.writeFileSync(out, b); console.log('OK ' + out + ' — ' + b.length + ' bytes'); });
