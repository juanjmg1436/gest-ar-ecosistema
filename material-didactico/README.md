# Material didáctico GEST-AR

Material de clase y de exposición docente del ecosistema GEST-AR.
Autor: Juan Manuel Gómez · Profesor en Ciencias Económicas.

Vive acá, junto al portal, porque es su contracara: el portal muestra las
aplicaciones y este material explica cómo se usan y por qué están pensadas así.

## Qué hay

### `Consignas_Clase/`
`GESTAR_Consignas.docx` — guía de actividades para el estudiante, con espacio
para nombre y curso. Es el documento que se reparte impreso.

### `Laminas_Exposicion_Docentes/`
Dos láminas para exponer ante colegas:

| Lámina | Tema |
|---|---|
| `GESTAR_1_Circuito_Contable` | El circuito que recorre una operación entre las tres apps |
| `GESTAR_2_Enfoque_NeuroTecnoPedagogico` | El enfoque pedagógico detrás del ecosistema |

**De cada una hay un `.svg` y un `.jpg`. El `.svg` es la fuente: editalo ahí.**
El `.jpg` es sólo la exportación para pegar en una presentación o imprimir.
Si cambiás el SVG, volvé a exportar el JPG — si no, quedan diciendo cosas distintas.

### `Material_Conceptual/`
- `GESTAR_Enfoque_y_Aplicaciones.docx` — material complementario de la presentación.
- `GESTAR_Guion_Video.docx` — guion técnico del video de presentación, toma por toma.
- `generar_documento.cjs` y `generar_guion_video.cjs` — **los scripts que generan
  esos dos Word.** El texto vive en el script, no en el `.docx`.

**Para cambiar el contenido, editá el `.cjs` y regeneralo**, no el Word:

```bash
npm install docx
node generar_documento.cjs
node generar_guion_video.cjs
```

Así el texto queda versionado en un archivo que git puede comparar. Si editás el
`.docx` a mano, el script lo pisa la próxima vez que corra.

### `Presentacion_GESTAR_Modulo_Impuestos/`
`GESTAR_Modulo_Impuestos.html` — presentación del módulo de impuestos, se abre en
cualquier navegador y se exporta a PDF con Ctrl+P → "Guardar como PDF". Incluye
los cuatro códigos QR que llevan al portal y a cada aplicación.

## Qué NO está acá

Los archivos de **EmprendePlan** (la app de escritorio, sus `.html`, `.py` y `.bat`)
quedaron afuera a propósito: en el portal está marcada como de **uso interno de la
escuela**, y este repositorio es público.
