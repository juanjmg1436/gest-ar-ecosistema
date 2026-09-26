const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, ShadingType, Footer, PageNumber,
  VerticalAlign, HeadingLevel
} = require('docx');
const fs = require('fs');
const path = require('path');

// ---- Paleta (tomada de las laminas de exposicion) ----
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
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE,
                    insideHorizontal: NONE, insideVertical: NONE };
const hair = (color) => ({ style: BorderStyle.SINGLE, size: 2, color });

// ---- Helpers ----
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
  shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
});

const bullet = (text, o = {}) => new Paragraph({
  children: Array.isArray(text) ? text : [run(text, o)],
  bullet: { level: 0 },
  keepNext: !!o.keepNext,
  keepLines: true,
  widowControl: true,
  spacing: { before: 0, after: 70, line: 280 },
});

const eyebrow = (text, color) => p([run(text, { caps: true, bold: true, size: 16, color: color || C.muted, spacing: 14 })], { after: 90 });

// Titulo de seccion: numero en color + texto, con filete inferior
const h1 = (num, text, color) => new Paragraph({
  children: [
    run(num + '  ', { bold: true, size: 28, color: color }),
    run(text, { bold: true, size: 28, color: C.dark }),
  ],
  heading: HeadingLevel.HEADING_1,
  keepNext: true,
  keepLines: true,
  spacing: { before: 380, after: 170, line: 300 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.line, space: 8 } },
});

const h2 = (text, color) => new Paragraph({
  children: [run(text, { bold: true, size: 22, color: color || C.dark })],
  heading: HeadingLevel.HEADING_2,
  keepNext: true,
  keepLines: true,
  widowControl: true,
  spacing: { before: 240, after: 90, line: 290 },
});

// Ficha de aplicacion: encabezado de color + cuerpo
const ficha = (nombre, bajada, color, cuerpo) => new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: {
    top: hair(C.line), bottom: hair(C.line), left: hair(C.line), right: hair(C.line),
    insideHorizontal: hair(C.line), insideVertical: NONE,
  },
  rows: [
    new TableRow({ cantSplit: true,
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: color },
        margins: { top: 150, bottom: 150, left: 200, right: 200 },
        children: [
          p([run(nombre, { bold: true, size: 24, color: 'FFFFFF' })], { after: 30, keepNext: true }),
          p([run(bajada, { size: 19, color: 'FFFFFF' })], { after: 0, keepNext: true }),
        ],
      })],
    }),
    new TableRow({ cantSplit: true,
      children: [new TableCell({
        margins: { top: 170, bottom: 150, left: 200, right: 200 },
        children: cuerpo,
      })],
    }),
  ],
});

// Tabla de dos columnas
const tabla2 = (encabezados, filas) => new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  columnWidths: [3200, 6000],
  borders: {
    top: hair(C.line), bottom: hair(C.line), left: NONE, right: NONE,
    insideHorizontal: hair(C.line), insideVertical: NONE,
  },
  rows: [
    new TableRow({ cantSplit: true,
      tableHeader: true,
      children: encabezados.map((t, i) => new TableCell({
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.bgalt },
        margins: { top: 110, bottom: 110, left: i === 0 ? 140 : 160, right: 140 },
        children: [p([run(t, { caps: true, bold: true, size: 16, color: C.muted, spacing: 12 })], { after: 0, keepNext: true })],
      })),
    }),
    ...filas.map(([a, b]) => new TableRow({ cantSplit: true,
      children: [
        new TableCell({
          margins: { top: 120, bottom: 120, left: 140, right: 140 },
          children: [p([run(a, { bold: true, size: 19, color: C.dark })], { after: 0, line: 260 })],
        }),
        new TableCell({
          margins: { top: 120, bottom: 120, left: 160, right: 140 },
          children: [p([run(b, { size: 19 })], { after: 0, line: 260 })],
        }),
      ],
    })),
  ],
});

// Bucle de 5 pasos, en una fila
const bucle = (pasos) => new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: noBorders,
  rows: [new TableRow({ cantSplit: true,
    children: pasos.map(([n, titulo, texto]) => new TableCell({
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.bg },
      verticalAlign: VerticalAlign.TOP,
      margins: { top: 130, bottom: 130, left: 110, right: 110 },
      borders: { top: hair(C.line), bottom: hair(C.line), left: hair(C.line), right: hair(C.line) },
      children: [
        p([run(n, { bold: true, size: 17, color: C.azul })], { after: 40 }),
        p([run(titulo, { bold: true, size: 18, color: C.dark })], { after: 40, line: 240 }),
        p([run(texto, { size: 16, color: C.muted })], { after: 0, line: 230 }),
      ],
    })),
  })],
});

// Bloque destacado
const destacado = (children, color) => new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: { top: NONE, bottom: NONE, right: NONE,
             left: { style: BorderStyle.SINGLE, size: 18, color: color },
             insideHorizontal: NONE, insideVertical: NONE },
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: C.bg },
    margins: { top: 170, bottom: 170, left: 220, right: 200 },
    children: children,
  })] })],
});

const sp = (h) => new Paragraph({ children: [], keepNext: true, spacing: { after: h === undefined ? 160 : h } });

// ============================ CONTENIDO ============================
const children = [];

// --- Encabezado ---
children.push(eyebrow('Material complementario de la presentación'));
children.push(p([run('GEST-AR: cómo está pensado el ecosistema', { bold: true, size: 40, color: C.dark })], { after: 80, line: 340 }));
children.push(p([run('El enfoque neuro-tecno-pedagógico y las aplicaciones que lo componen', { size: 23, color: C.muted })],
  { after: 150, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.dark, space: 10 } } }));
children.push(p([run('Juan Manuel Gómez  ·  Profesor en Ciencias Económicas  ·  Licenciado en Educación', { size: 17, color: C.muted })], { after: 30, line: 250 }));
children.push(p([run('Escuela Normal Superior N.º 9 y Anexo Comercial Juan XXIII — Wanda, Misiones  ·  2026', { size: 17, color: C.muted })], { after: 60, line: 250 }));

// --- 1 ---
children.push(h1('1', 'Punto de partida', C.azul));
children.push(p('La gestión administrativa suele enseñarse en dos tiempos separados: primero la definición —qué es el IVA, qué es un asiento, qué es un recibo de sueldo— y bastante después, si el tiempo alcanza, la práctica. El resultado conocido es un estudiante capaz de explicar correctamente un concepto que no sabría por dónde empezar a ejecutar.'));
children.push(p('GEST-AR invierte ese orden. El estudiante entra a un entorno donde ya hay una empresa que atender, y los conceptos aparecen cuando la operación los necesita: se explica qué es un punto de venta en el momento exacto en que el sistema no lo deja facturar sin tenerlo habilitado. La definición llega como respuesta a un problema que el estudiante ya tiene, y no antes.'));
children.push(p('Este material acompaña la presentación de la plataforma. Desarrolla dos cosas: por qué el ecosistema está diseñado como está —el enfoque neuro-tecno-pedagógico— y en qué consiste cada aplicación que lo integra.'));

// --- 2 ---
children.push(h1('2', 'El enfoque neuro-tecno-pedagógico', C.viole));
children.push(p('No alcanza con digitalizar el contenido: un apunte en pantalla sigue siendo un apunte. Cada recurso de la plataforma existe porque activa un mecanismo cognitivo concreto. El enfoque cruza tres planos que, por separado, no producen aprendizaje.'));

children.push(h2('Neuro — cómo aprende el cerebro', C.viole));
children.push(bullet('La atención es un recurso limitado: se sostiene con novedad y desafío, no con repetición.'));
children.push(bullet('Lo que se decide se recuerda mejor que lo que se escucha: la acción consolida la memoria.'));
children.push(bullet('El error sin costo real es un momento de alta plasticidad: corregir lo que uno mismo hizo deja una huella que la explicación anticipada no deja.'));

children.push(h2('Tecno — un entorno que responde', C.verde));
children.push(bullet('Simula consecuencias reales sin riesgo real: se puede equivocar y volver a intentar.'));
children.push(bullet('La respuesta es inmediata. Entre decidir y ver el efecto pasan segundos, no una semana hasta la devolución del docente.'));
children.push(bullet('Los datos persisten y se integran entre aplicaciones: lo que hace hoy le condiciona el mes que viene.'));

children.push(h2('Pedagógico — una secuencia con sentido', C.azul));
children.push(bullet('Andamiaje: cada etapa habilita la siguiente, para no saturar la memoria de trabajo.'));
children.push(bullet('Aprendizaje situado: la tarea es la del puesto de trabajo, no una versión escolar de esa tarea.'));
children.push(bullet('Evaluación auténtica: se evalúa la operación completa, no la definición memorizada.'));

children.push(h2('El bucle que sostiene la atención', C.dark));
children.push(p('Los tres planos se articulan en un ciclo que se repite en cada actividad de la plataforma:', { after: 130, keepNext: true }));
children.push(bucle([
  ['1', 'Desafío', 'Una consigna con dificultad justa'],
  ['2', 'Decisión', 'El estudiante elige y se compromete'],
  ['3', 'Consecuencia', 'El sistema reacciona al instante'],
  ['4', 'Feedback', 'Ve el error y entiende por qué'],
  ['5', 'Logro', 'Refuerza y habilita el desafío siguiente'],
]));
children.push(sp(140));
children.push(p('Cerrado el bucle, vuelve a empezar con más complejidad, apoyado en lo que ya consolidó. Esa es la razón por la que el estudiante sostiene la tarea: no porque la plataforma lo entretenga, sino porque cada vuelta le devuelve una pregunta un poco mejor que la anterior.'));

children.push(h2('Cada recurso responde a un mecanismo, no a una moda', C.dark));
children.push(p('La gamificación de la plataforma no es decorativa. Cada elemento se justifica por lo que provoca:', { after: 130, keepNext: true }));
children.push(tabla2(['Recurso', 'Mecanismo cognitivo que activa'], [
  ['Puntos, niveles y desafíos', 'Convierten una tarea larga en metas cortas: sostienen la atención.'],
  ['Etapas que se desbloquean', 'Dosifican la carga cognitiva: nadie liquida IVA antes de saber facturar.'],
  ['El error no tiene costo real', 'Habilita el error productivo: se explora sin miedo a arruinar nada.'],
  ['Datos que se arrastran', 'Fuerzan la transferencia: la venta de hoy reaparece en la declaración jurada del mes.'],
  ['Comprobantes y libros con formato real', 'Acortan la distancia entre la tarea escolar y la tarea profesional.'],
]));
children.push(sp(150));
children.push(destacado([
  p([run('No se gamifica el premio: se gamifica el proceso de pensar.', { bold: true, size: 22, color: C.dark })], { after: 60, line: 270 }),
  p([run('El puntaje no reemplaza al contenido: ordena el recorrido, marca el progreso y le devuelve al estudiante una pregunta mejor que la anterior.', { size: 19, color: C.text })], { after: 0, line: 265 }),
], C.viole));

// --- 3 ---
children.push(h1('3', 'Las aplicaciones del ecosistema', C.verde));
children.push(p('Tres simuladores sostienen el ciclo administrativo completo. Cada uno cubre un dominio y puede usarse por separado, pero están pensados para trabajar sobre una misma empresa ficticia y compartir sus datos.'));
children.push(sp(60));

children.push(ficha('TRIBUT.AR', 'El marco fiscal donde todo empieza', C.azul, [
  p('Simula la relación entre la empresa y el fisco. Es el punto de partida del recorrido: sin constituir la empresa y definir su situación fiscal, no hay operación posible aguas abajo.'),
  eyebrow('Qué hace el estudiante', C.azul),
  bullet('Da de alta una empresa ficticia: datos registrales, domicilio fiscal, actividad y régimen (monotributo o régimen general).'),
  bullet('Inscribe los impuestos nacionales y provinciales que le corresponden según esa elección.'),
  bullet('Habilita el punto de venta que después le permitirá facturar.'),
  bullet('Liquida el IVA mes a mes, calcula el impuesto a las Ganancias y presenta la declaración jurada de Ingresos Brutos, modelada sobre la normativa de Misiones.'),
  bullet('Genera los VEP y registra los pagos en la billetera fiscal.'),
  p([run('Qué concepto instala: ', { bold: true, size: 19, color: C.dark }),
     run('que la obligación fiscal no es un trámite de fin de año, sino la consecuencia directa de cada operación que la empresa realiza.', { size: 19 })],
    { after: 0, before: 120, line: 265 }),
]));
children.push(sp(180));

children.push(ficha('PyMEZ 360', 'La operación diaria de la empresa', C.verde, [
  p('Simula la gestión comercial y contable del día a día. Es donde ocurren los hechos económicos y donde se ve, en el mismo instante, el asiento que cada uno genera.'),
  eyebrow('Qué hace el estudiante', C.verde),
  bullet('Emite facturas A, B y C, y registra compras con cómputo del crédito fiscal.'),
  bullet('Administra clientes, proveedores, stock, cobros, pagos y tesorería (caja y bancos).'),
  bullet('Registra las retenciones de Ingresos Brutos sufridas en los cobros.'),
  bullet('Consulta el Libro Diario, el Mayor, el balance de sumas y saldos y el resultado del período.'),
  bullet('Devenga Ingresos Brutos con cada venta y cancela el impuesto desde el panel correspondiente.'),
  p([run('Qué concepto instala: ', { bold: true, size: 19, color: C.dark }),
     run('que la contabilidad no es un ejercicio en papel. El asiento es la huella de una decisión que el estudiante acaba de tomar.', { size: 19 })],
    { after: 0, before: 120, line: 265 }),
]));
children.push(sp(180));

children.push(ficha('Sueldos 360', 'Las relaciones laborales', C.viole, [
  p('Simula la administración de la nómina, desde el alta del trabajador hasta la obligación con el organismo recaudador.'),
  eyebrow('Qué hace el estudiante', C.viole),
  bullet('Da de alta empleados por categoría del convenio colectivo.'),
  bullet('Liquida el mes: remuneraciones, descuentos del trabajador (jubilación, obra social) y contribuciones patronales.'),
  bullet('Emite e imprime los recibos de sueldo.'),
  bullet('Obtiene la información que conforma la declaración de cargas sociales (F.931).'),
  p([run('Qué concepto instala: ', { bold: true, size: 19, color: C.dark }),
     run('que el costo laboral no es el sueldo de bolsillo, y que cada descuento tiene un destinatario y un vencimiento.', { size: 19 })],
    { after: 0, before: 120, line: 265 }),
]));

children.push(h2('Cómo se integran: el dato que viaja', C.dark));
children.push(p('Lo que convierte estas aplicaciones en un ecosistema —y no en tres ejercicios independientes— es que la información pasa de una a otra, igual que en una empresa real.'));
children.push(bullet([run('De TRIBUT.AR a PyMEZ 360. ', { bold: true, size: 21, color: C.dark }),
  run('TRIBUT.AR emite un código de punto de venta. Hasta que el estudiante lo pega en PyMEZ, la aplicación no le permite registrar ventas. El bloqueo es deliberado: reproduce que no se puede facturar sin autorización fiscal, y es lo que lo obliga a volver al trámite en lugar de saltearlo.', { size: 21 })]));
children.push(bullet([run('De PyMEZ 360 a TRIBUT.AR. ', { bold: true, size: 21, color: C.dark }),
  run('Un código de sincronización permite importar las ventas del período y precargar la base imponible del IVA y de Ingresos Brutos. La declaración jurada sale de lo que el estudiante efectivamente vendió, no de un enunciado inventado por el docente.', { size: 21 })]));
children.push(bullet([run('De Sueldos 360 a PyMEZ 360. ', { bold: true, size: 21, color: C.dark }),
  run('La liquidación del mes se importa y se transforma en los asientos de devengamiento de sueldos y de cargas patronales.', { size: 21 })]));
children.push(sp(100));
children.push(p('La consecuencia pedagógica es la más importante del diseño: no hay forma de resolver bien la última etapa si las anteriores están mal. El error no se corrige con una marca del docente en el margen; aparece solo, más adelante, en el número que no cierra.'));

children.push(h2('Puertas de entrada y recursos de apoyo', C.dark));
children.push(bullet([run('Gestión 360 Emprende. ', { bold: true, size: 21, color: C.dark }),
  run('Gestión administrativa en su versión simple, para proyectos escolares y microemprendimientos: registro de operaciones, ingresos y egresos, análisis de costos y toma de decisiones comerciales. Es la puerta de entrada cuando todavía no hay régimen fiscal ni nómina que administrar.', { size: 21 })]));
children.push(bullet([run('Campus Virtual GEST-AR. ', { bold: true, size: 21, color: C.dark }),
  run('Cursos autoasistidos organizados en módulos progresivos, con actividades de verificación, recursos digitales y posibles certificaciones. Es donde se apoya el marco teórico que los simuladores ponen a prueba.', { size: 21 })]));
children.push(bullet([run('Emprendeplan. ', { bold: true, size: 21, color: C.dark }),
  run('Guía en nueve secciones —idea, mercado, estrategia comercial, plan operativo, plan financiero, equipo, cronograma, riesgos e indicadores— para formular el proyecto antes de gestionarlo. Se completa en el navegador y se exporta en PDF.', { size: 21 })]));

// --- 4 ---
children.push(h1('4', 'El motor contable: de la operación al asiento', C.ambar));
children.push(p('Detrás de las tres aplicaciones hay un solo motor: cada hecho económico se convierte en un asiento por partida doble. El estudiante registra una venta y ve, en el mismo momento, el registro que ese hecho produce.'));
children.push(sp(60));
children.push(tabla2(['Ejemplo — venta a crédito con IVA', 'Importe'], [
  ['Debe — Deudores por ventas', '121.000,00'],
  ['Haber — Ventas', '100.000,00'],
  ['Haber — IVA Débito Fiscal', '21.000,00'],
]));
children.push(sp(150));
children.push(p('De ahí en más, todo lo que el estudiante puede mirar es consecuencia de lo que hizo: el Libro Diario con los asientos en orden, el Mayor con el saldo de cada cuenta, el balance de sumas y saldos, el resultado del período y las declaraciones juradas.'));
children.push(sp(60));
children.push(destacado([
  p([run('La obligación nace con el hecho que la genera y sólo desaparece cuando se paga.', { bold: true, size: 22, color: C.dark })], { after: 100, line: 270 }),
  bullet([run('IVA: ', { bold: true, size: 19, color: C.dark }), run('se devenga en cada venta y en cada compra → se liquida el período → se paga, o queda un saldo técnico a favor.', { size: 19 })]),
  bullet([run('Ingresos Brutos: ', { bold: true, size: 19, color: C.dark }), run('nace con la venta, sobre base neta de IVA → se declara y se paga a la DGR.', { size: 19 })]),
  bullet([run('Sueldos: ', { bold: true, size: 19, color: C.dark }), run('se devengan con la liquidación → se cancelan en dos destinos distintos: el neto al trabajador y el F.931 al organismo recaudador.', { size: 19 })]),
], C.ambar));
children.push(sp(120));
children.push(p('El mismo principio, tres veces. Esa repetición no es redundancia: es lo que permite que el estudiante generalice el patrón y lo reconozca frente a una obligación que nunca vio.'));

// --- 5 ---
children.push(h1('5', 'Qué se evalúa', C.azul));
children.push(p('La evaluación no pregunta por la definición: mira la operación. La evidencia del aprendizaje son los productos que el propio sistema generó —el libro diario, la declaración jurada, el recibo de sueldo, el comprobante de pago— y la pregunta es si el ciclo cierra: si lo que se devengó se canceló, si lo que se vendió aparece en la declaración, si el saldo de la cuenta explica lo que pasó.'));
children.push(p('Eso vuelve visible el proceso y no sólo el resultado. Un error de criterio deja rastro en la cadena, se puede señalar dónde se produjo, y el estudiante puede volver a intentarlo sin que nada se rompa de manera definitiva.'));
children.push(sp(150));
children.push(destacado([
  p([run('«Cuando el entorno digital está bien diseñado, el estudiante no percibe que está aprendiendo: percibe que está resolviendo.»', { bold: true, italics: true, size: 22, color: C.dark })], { after: 0, line: 275 }),
], C.azul));

// ============================ DOCUMENTO ============================
const doc = new Document({
  creator: 'Juan Manuel Gómez',
  title: 'GEST-AR: cómo está pensado el ecosistema',
  description: 'Material conceptual complementario a la presentación del ecosistema digital educativo GEST-AR',
  subject: 'Enfoque neuro-tecno-pedagógico y aplicaciones del ecosistema GEST-AR',
  keywords: 'GEST-AR, neuro-tecno-pedagógico, TRIBUT.AR, PyMEZ 360, Sueldos 360, educación técnica',
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
            new TextRun({ text: 'GEST-AR · Material conceptual complementario   ', font: F, size: 15, color: C.soft }),
            new TextRun({ children: [PageNumber.CURRENT], font: F, size: 15, color: C.muted, bold: true }),
          ],
        })],
      }),
    },
    children: children,
  }],
});

const out = process.argv[2];

// La libreria emite sus propios Heading1/Heading2 y ademas agrega los definidos aca:
// quedan dos estilos con el mismo w:styleId, lo que es OOXML invalido y puede
// hacer que Word pida reparar el archivo. Nos quedamos con el ultimo (el propio).
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
