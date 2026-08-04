/************************************************
 * SISTEMA DE DEVOLUCIÓN DE TRÁMITES (API GITHUB PAGES)
 * Codigo.gs - Optimizado & Corregido
 ************************************************/

const HOJA_DEVOLUCIONES = "DEVOLUCIONES";
const HOJA_DETALLE = "DEVOLUCIONES_DETALLES";
const HOJA_FUNCIONARIOS = "FUNCIONARIOS";
const HOJA_MOTIVOS = "MOTIVOS";

/************************************************
 * MANEJADOR DE PETICIONES GET
 ************************************************/
function doGet(e) {
  try {
    const accion = e.parameter.accion;
    let respuesta = {};

    switch (accion) {
      case "listarFuncionarios":
        respuesta = { ok: true, datos: listarFuncionarios() };
        break;
      case "obtenerMotivos":
        respuesta = { ok: true, datos: obtenerMotivos() };
        break;
      case "dashboard":
        respuesta = { ok: true, datos: dashboard() };
        break;
      case "buscarPlaca":
        respuesta = { ok: true, datos: buscarPlaca(e.parameter.placa) };
        break;
      case "consultarDevolucion":
        respuesta = { ok: true, datos: consultarDevolucion(e.parameter.id) };
        break;
      default:
        respuesta = { ok: false, mensaje: "Acción GET no válida o no especificada." };
    }

    return responderJSON(respuesta);

  } catch (error) {
    return responderJSON({ ok: false, mensaje: error.message });
  }
}

/************************************************
 * MANEJADOR DE PETICIONES POST
 ************************************************/
function doPost(e) {
  try {
    // Maneja payload formateado desde fetch cliente (JSON string)
    const contenido = JSON.parse(e.postData.contents);
    const accion = contenido.accion;
    const payload = contenido.payload;

    let respuesta = {};

    switch (accion) {
      case "guardarDevolucion":
        respuesta = guardarDevolucion(payload);
        break;
      case "actualizarDevolucion":
        respuesta = actualizarDevolucion(payload);
        break;
      case "eliminarDevolucion":
        respuesta = eliminarDevolucion(payload.id);
        break;
      case "generarPDF":
        respuesta = generarPDFDevolucion(payload.id);
        break;
      default:
        respuesta = { ok: false, mensaje: "Acción POST no válida o no especificada." };
    }

    return responderJSON(respuesta);

  } catch (error) {
    return responderJSON({ ok: false, mensaje: error.message });
  }
}

/************************************************
 * FUNCIONES AUXILIARES Y DE RED
 ************************************************/
function responderJSON(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

function obtenerHoja(nombre) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja) throw new Error("No existe la hoja: " + nombre);
  return hoja;
}

function siguienteID(hoja) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) return 1;

  const ids = hoja
    .getRange(2, 1, ultimaFila - 1, 1)
    .getValues()
    .flat()
    .map(Number)
    .filter(id => !isNaN(id) && id > 0);

  return ids.length ? Math.max(...ids) + 1 : 1;
}

function fechaActual() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/************************************************
 * LISTAR DATOS
 ************************************************/
function listarFuncionarios() {
  const hoja = obtenerHoja(HOJA_FUNCIONARIOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];

  return hoja
    .getRange(2, 1, ultimaFila - 1, 1)
    .getDisplayValues()
    .flat()
    .filter(nombre => nombre.trim() !== "");
}

function obtenerMotivos() {
  const hoja = obtenerHoja(HOJA_MOTIVOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];

  return hoja
    .getRange(2, 1, ultimaFila - 1, 1)
    .getDisplayValues()
    .flat()
    .filter(motivo => motivo.trim() !== "");
}

/************************************************
 * VALIDACIONES
 ************************************************/
function validarCabecera(datos) {
  if (!datos) throw new Error("No se recibieron datos.");
  if (!datos.fecha) throw new Error("Debe indicar la fecha.");
  if (!datos.funcionario) throw new Error("Debe seleccionar un funcionario.");
  if (!datos.placa) throw new Error("Debe ingresar la placa.");
  if (!datos.cedula) throw new Error("Debe ingresar la cédula.");
  if (!datos.nombre) throw new Error("Debe ingresar el nombre del ciudadano.");
  if (!datos.detalle || datos.detalle.length === 0) {
    throw new Error("Debe agregar al menos un motivo de devolución.");
  }

  datos.placa = datos.placa.toUpperCase().trim();
  datos.nombre = datos.nombre.toUpperCase().trim();
  datos.cedula = String(datos.cedula).trim();

  return datos;
}

/************************************************
 * GUARDAR DEVOLUCIÓN (CON LOCK DE SEGURIDAD)
 ************************************************/
function guardarDevolucion(datos) {
  const lock = LockService.getScriptLock();
  // Esperar hasta 10 segundos para bloquear proceso concurrente
  if (!lock.tryLock(10000)) {
    return { ok: false, mensaje: "El sistema está ocupado. Intenta de nuevo en unos segundos." };
  }

  try {
    datos = validarCabecera(datos);

    const hojaCabecera = obtenerHoja(HOJA_DEVOLUCIONES);
    const hojaDetalle = obtenerHoja(HOJA_DETALLE);

    const idDevolucion = siguienteID(hojaCabecera);

    // Guardar Cabecera
    hojaCabecera.appendRow([
      idDevolucion,
      datos.fecha,
      datos.funcionario,
      datos.placa,
      datos.cedula,
      datos.nombre
    ]);

    // Guardar Detalle
    let idDetalle = siguienteID(hojaDetalle);
    const registrosDetalle = datos.detalle.map(item => [
      idDetalle++,
      idDevolucion,
      item.motivo,
      item.observacion || ""
    ]);

    hojaDetalle
      .getRange(hojaDetalle.getLastRow() + 1, 1, registrosDetalle.length, registrosDetalle[0].length)
      .setValues(registrosDetalle);

    return {
      ok: true,
      id: idDevolucion,
      mensaje: "La devolución fue guardada correctamente."
    };

  } catch (error) {
    return { ok: false, mensaje: error.message };
  } finally {
    lock.releaseLock();
  }
}

/************************************************
 * CONSULTAS Y BÚSQUEDAS
 ************************************************/
function buscarDevolucion(id) {
  const hojaCabecera = obtenerHoja(HOJA_DEVOLUCIONES);
  const hojaDetalle = obtenerHoja(HOJA_DETALLE);

  const cabecera = hojaCabecera.getDataRange().getValues();
  const detalle = hojaDetalle.getDataRange().getValues();

  const respuesta = { detalle: [] };

  for (let i = 1; i < cabecera.length; i++) {
    if (Number(cabecera[i][0]) === Number(id)) {
      respuesta.id = cabecera[i][0];
      // Si la fecha viene como Date Object de Apps Script, la formateamos como string YYYY-MM-DD
      respuesta.fecha = cabecera[i][1] instanceof Date 
        ? Utilities.formatDate(cabecera[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd") 
        : cabecera[i][1];
      respuesta.funcionario = cabecera[i][2];
      respuesta.placa = cabecera[i][3];
      respuesta.cedula = cabecera[i][4];
      respuesta.nombre = cabecera[i][5];
      break;
    }
  }

  if (!respuesta.id) return null;

  for (let i = 1; i < detalle.length; i++) {
    if (Number(detalle[i][1]) === Number(id)) {
      respuesta.detalle.push({
        id: detalle[i][0],
        motivo: detalle[i][2],
        observacion: detalle[i][3]
      });
    }
  }

  return respuesta;
}

function consultarDevolucion(id) {
  return buscarDevolucion(id);
}

function buscarPlaca(placa) {
  placa = String(placa).trim().toUpperCase();
  const hoja = obtenerHoja(HOJA_DEVOLUCIONES);
  const datos = hoja.getDataRange().getValues();
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][3]).trim().toUpperCase() === placa) {
      resultados.push({
        id: datos[i][0],
        fecha: datos[i][1] instanceof Date 
          ? Utilities.formatDate(datos[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd") 
          : datos[i][1],
        funcionario: datos[i][2],
        placa: datos[i][3],
        cedula: datos[i][4],
        nombre: datos[i][5]
      });
    }
  }
  return resultados; // Retorna array de coincidencias
}

/************************************************
 * ELIMINAR DEVOLUCIÓN (OPTIMIZADO EN MEMORIA)
 ************************************************/
function eliminarDevolucion(id) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, mensaje: "Sistema ocupado." };

  try {
    const hojaCabecera = obtenerHoja(HOJA_DEVOLUCIONES);
    const hojaDetalle = obtenerHoja(HOJA_DETALLE);

    // Filtrar Cabecera
    const datosCabecera = hojaCabecera.getDataRange().getValues();
    const nuevaCabecera = datosCabecera.filter((fila, index) => index === 0 || Number(fila[0]) !== Number(id));

    // Filtrar Detalle
    const datosDetalle = hojaDetalle.getDataRange().getValues();
    const nuevoDetalle = datosDetalle.filter((fila, index) => index === 0 || Number(fila[1]) !== Number(id));

    // Reescribir hojas
    hojaCabecera.clearContents();
    if (nuevaCabecera.length > 0) {
      hojaCabecera.getRange(1, 1, nuevaCabecera.length, nuevaCabecera[0].length).setValues(nuevaCabecera);
    }

    hojaDetalle.clearContents();
    if (nuevoDetalle.length > 0) {
      hojaDetalle.getRange(1, 1, nuevoDetalle.length, nuevoDetalle[0].length).setValues(nuevoDetalle);
    }

    return { ok: true, mensaje: "Registro eliminado correctamente." };
  } catch(error) {
    return { ok: false, mensaje: error.message };
  } finally {
    lock.releaseLock();
  }
}

/************************************************
 * ACTUALIZAR DEVOLUCIÓN
 ************************************************/
function actualizarDevolucion(datos) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, mensaje: "Sistema ocupado." };

  try {
    datos = validarCabecera(datos);
    if (!datos.id) throw new Error("No se recibió el ID de la devolución.");

    const hojaCabecera = obtenerHoja(HOJA_DEVOLUCIONES);
    const hojaDetalle = obtenerHoja(HOJA_DETALLE);

    const cabecera = hojaCabecera.getDataRange().getValues();
    let filaCabecera = -1;

    for (let i = 1; i < cabecera.length; i++) {
      if (Number(cabecera[i][0]) === Number(datos.id)) {
        filaCabecera = i + 1;
        break;
      }
    }

    if (filaCabecera === -1) throw new Error("No existe la devolución.");

    // Actualizar fila en Cabecera
    hojaCabecera.getRange(filaCabecera, 1, 1, 6).setValues([[
      datos.id,
      datos.fecha,
      datos.funcionario,
      datos.placa,
      datos.cedula,
      datos.nombre
    ]]);

    // Limpiar detalle viejo y reescribir en memoria
    const datosDetalle = hojaDetalle.getDataRange().getValues();
    const detalleLimpio = datosDetalle.filter((fila, idx) => idx === 0 || Number(fila[1]) !== Number(datos.id));

    let idDetalle = siguienteID(hojaDetalle);
    datos.detalle.forEach(item => {
      detalleLimpio.push([
        idDetalle++,
        datos.id,
        item.motivo,
        item.observacion || ""
      ]);
    });

    hojaDetalle.clearContents();
    hojaDetalle.getRange(1, 1, detalleLimpio.length, detalleLimpio[0].length).setValues(detalleLimpio);

    return { ok: true, mensaje: "Devolución actualizada correctamente." };

  } catch (error) {
    return { ok: false, mensaje: error.message };
  } finally {
    lock.releaseLock();
  }
}

/************************************************
 * DASHBOARD Y MÉTRICAS
 ************************************************/
function totalDevoluciones() {
  const hoja = obtenerHoja(HOJA_DEVOLUCIONES);
  return Math.max(0, hoja.getLastRow() - 1);
}

function totalMotivos() {
  const hoja = obtenerHoja(HOJA_DETALLE);
  return Math.max(0, hoja.getLastRow() - 1);
}

function dashboard() {
  return {
    devoluciones: totalDevoluciones(),
    motivos: totalMotivos(),
    funcionarios: listarFuncionarios().length,
    fecha: fechaActual()
  };
}

/************************************************
 * GENERACIÓN DE PDF DEVOLUCIÓN (OPTIMIZADO)
 ************************************************/
function generarPDFDevolucion(idDevolucion) {
  try {
    const datos = buscarDevolucion(idDevolucion);

    if (!datos) {
      throw new Error("No se encontró la devolución.");
    }

    const ID_ENCABEZADO = "1wPY_QJ4G_W7rz5bdkz7ObGc0L0cN9_ML";
    const ID_PIE = "1m-KztMZ-KSlX-tu4BS61qrRQ-9TX8YpR";

    // URLs directas de transmisión pública de Google Drive (evita errores de Base64 en PDF)
    const urlEncabezado = "https://lh3.googleusercontent.com/d/" + ID_ENCABEZADO;
    const urlPie = "https://lh3.googleusercontent.com/d/" + ID_PIE;

    // Formatear Fecha
    let fecha = datos.fecha;
    if (fecha instanceof Date) {
      fecha = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else if (typeof fecha === "string" && fecha.includes("-")) {
      fecha = fecha.split("-").reverse().join("/");
    }

    // Filas de los detalles
    let filas = "";
    if (datos.detalle && datos.detalle.length > 0) {
      datos.detalle.forEach(function(f) {
        filas += `
        <tr>
            <td>${f.motivo}</td>
            <td>${f.observacion || ""}</td>
        </tr>`;
      });
    }

    // Plantilla HTML para la conversión a PDF
    const html = `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  @page {
    size: letter;
    margin: 12mm;
  }
  body {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 10pt;
    color: #222222;
    margin: 0;
    padding: 0;
  }
  table {
    border-collapse: collapse;
    width: 100%;
  }
  .img-banner {
    width: 100%;
    max-width: 100%;
    height: auto;
    display: block;
  }
  .info td {
    border: 1px solid #bdbdbd;
    padding: 6px 8px;
  }
  .detalle th {
    background-color: #0c4da2;
    color: #ffffff;
    border: 1px solid #0c4da2;
    padding: 6px 8px;
    text-align: left;
  }
  .detalle td {
    border: 1px solid #999999;
    padding: 6px 8px;
  }
  .titulo {
    text-align: center;
    font-size: 15pt;
    font-weight: bold;
    margin-top: 15px;
    margin-bottom: 3px;
    color: #111111;
  }
  .subtitulo {
    text-align: center;
    font-size: 11pt;
    margin-bottom: 15px;
    color: #555555;
  }
  .label {
    background-color: #f2f2f2;
    font-weight: bold;
    width: 20%;
  }
  .linea {
    width: 220px;
    border-top: 1px solid #000000;
    margin: 0 auto 5px auto;
  }
  .tabla-firmas {
    margin-top: 50px;
    margin-bottom: 20px;
  }
</style>
</head>
<body>

  <!-- ENCABEZADO -->
  <table width="100%" style="margin-bottom: 10px;">
    <tr>
      <td align="center">
        <img src="${urlEncabezado}" class="img-banner" />
      </td>
    </tr>
  </table>

  <div class="titulo">COMPROBANTE DE DEVOLUCIÓN DE TRÁMITE</div>
  <div class="subtitulo">Radicado No. ${datos.id}</div>

  <!-- INFORMACIÓN GENERAL -->
  <table class="info">
    <tr>
      <td class="label">Fecha</td>
      <td>${fecha}</td>
      <td class="label">Placa</td>
      <td><b>${datos.placa}</b></td>
    </tr>
    <tr>
      <td class="label">Funcionario</td>
      <td>${datos.funcionario}</td>
      <td class="label">Cédula</td>
      <td>${datos.cedula}</td>
    </tr>
    <tr>
      <td class="label">Ciudadano</td>
      <td colspan="3">${datos.nombre}</td>
    </tr>
  </table>

  <br>

  <!-- MOTIVOS Y DETALLES -->
  <table class="detalle">
    <thead>
      <tr>
        <th width="40%">Motivo</th>
        <th>Observación</th>
      </tr>
    </thead>
    <tbody>
      ${filas}
    </tbody>
  </table>

  <!-- FIRMAS -->
  <table class="tabla-firmas">
    <tr>
      <td align="center" style="vertical-align: top;">
        <div class="linea"></div>
        <b>${datos.funcionario}</b><br>
        Funcionario Responsable
      </td>
      <td align="center" style="vertical-align: top;">
        <div class="linea"></div>
        <b>${datos.nombre}</b><br>
        Ciudadano / Recibido
      </td>
    </tr>
  </table>

  <!-- PIE DE PÁGINA -->
  <table width="100%">
    <tr>
      <td align="center">
        <img src="${urlPie}" class="img-banner" />
      </td>
    </tr>
  </table>

</body>
</html>
`;

    const pdf = HtmlService
      .createHtmlOutput(html)
      .getAs(MimeType.PDF)
      .setName("Devolucion_" + datos.placa + "_" + datos.id + ".pdf");

    return {
      ok: true,
      nombreArchivo: pdf.getName(),
      base64: Utilities.base64Encode(pdf.getBytes())
    };

  } catch (error) {
    return {
      ok: false,
      mensaje: error.message
    };
  }
}
