/************************************************
 * SISTEMA DE DEVOLUCIÓN DE TRÁMITES (VERSIÓN API PARA GITHUB PAGES)
 * Codigo.gs
 ************************************************/
const encabezadoURL = "https://raw.githubusercontent.com/arieltobon2026/DEVOLUCION/main/img/Encabezado.png";
const pieURL = "https://raw.githubusercontent.com/arieltobon2026/DEVOLUCION/main/img/PiePagina.png";

const HOJA_DEVOLUCIONES = "DEVOLUCIONES";
const HOJA_DETALLE = "DEVOLUCIONES_DETALLES";
const HOJA_FUNCIONARIOS = "FUNCIONARIOS";
const HOJA_MOTIVOS = "MOTIVOS";

/************************************************
 * MANEJADOR DE PETICIONES GET
 ************************************************/
function doGet(e) {
  try {
    const accion = e ? e.parameter.accion : null;
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
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No se recibieron datos POST.");
    }

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
 * MANEJADOR PREFLIGHT (CORS)
 ************************************************/
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

/************************************************
 * FUNCIÓN AUXILIAR PARA DEVOLVER EN FORMATO JSON
 ************************************************/
function responderJSON(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

/************************************************
 * OBTENER HOJA
 ************************************************/
function obtenerHoja(nombre) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja) throw new Error("No existe la hoja: " + nombre);
  return hoja;
}

/************************************************
 * OBTENER SIGUIENTE ID (CORREGIDO)
 ************************************************/
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

/************************************************
 * FECHA ACTUAL
 ************************************************/
function fechaActual() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/************************************************
 * LISTAR FUNCIONARIOS
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

/************************************************
 * LISTAR MOTIVOS DE DEVOLUCIÓN
 ************************************************/
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
 * VALIDAR DATOS DE CABECERA
 ************************************************/
function validarCabecera(datos) {
  if (!datos) throw new Error("No se recibieron datos.");
  if (!datos.fecha) throw new Error("Debe indicar la fecha.");
  if (!datos.funcionario) throw new Error("Debe seleccionar un funcionario.");
  if (!datos.placa) throw new Error("Debe ingresar la placa.");
  if (!datos.cedula) throw new Error("Debe ingresar la cédula.");
  if (!datos.nombre) throw new Error("Debe ingresar el nombre del ciudadano.");
  if (!datos.detalle || !Array.isArray(datos.detalle) || datos.detalle.length === 0) {
    throw new Error("Debe agregar al menos un motivo de devolución.");
  }

  datos.placa = datos.placa.toUpperCase().trim();
  datos.nombre = datos.nombre.toUpperCase().trim();
  datos.cedula = String(datos.cedula).trim();

  return datos;
}

/************************************************
 * GUARDAR DEVOLUCIÓN
 ************************************************/
function guardarDevolucion(datos) {
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

    // Preparar y Guardar Detalle
    let idDetalle = siguienteID(hojaDetalle);
    const registrosDetalle = [];

    datos.detalle.forEach(function(item) {
      registrosDetalle.push([
        idDetalle++,
        idDevolucion,
        item.motivo,
        item.observacion || ""
      ]);
    });

    hojaDetalle
      .getRange(
        hojaDetalle.getLastRow() + 1,
        1,
        registrosDetalle.length,
        registrosDetalle[0].length
      )
      .setValues(registrosDetalle);

    return {
      ok: true,
      id: idDevolucion,
      mensaje: "La devolución fue guardada correctamente."
    };

  } catch (error) {
    return { ok: false, mensaje: error.message };
  }
}

/************************************************
 * BUSCAR DEVOLUCIÓN POR ID
 ************************************************/
function buscarDevolucion(id) {
  const hojaCabecera = obtenerHoja(HOJA_DEVOLUCIONES);
  const hojaDetalle = obtenerHoja(HOJA_DETALLE);

  const cabecera = hojaCabecera.getDataRange().getValues();
  const detalle = hojaDetalle.getDataRange().getValues();

  let respuesta = null;

  for (let i = 1; i < cabecera.length; i++) {
    if (Number(cabecera[i][0]) === Number(id)) {
      respuesta = {
        id: cabecera[i][0],
        fecha: cabecera[i][1] instanceof Date ? Utilities.formatDate(cabecera[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd") : cabecera[i][1],
        funcionario: cabecera[i][2],
        placa: cabecera[i][3],
        cedula: cabecera[i][4],
        nombre: cabecera[i][5],
        detalle: []
      };
      break;
    }
  }

  if (!respuesta) return null;

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

/************************************************
 * BUSCAR POR PLACA (DEVUELVE HISTORIAL)
 ************************************************/
function buscarPlaca(placa) {
  placa = String(placa).trim().toUpperCase();
  const hoja = obtenerHoja(HOJA_DEVOLUCIONES);
  const datos = hoja.getDataRange().getValues();
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][3]).trim().toUpperCase() === placa) {
      resultados.push(buscarDevolucion(datos[i][0]));
    }
  }
  return resultados;
}

/************************************************
 * ELIMINAR DEVOLUCIÓN
 ************************************************/
function eliminarDevolucion(id) {
  const hojaCabecera = obtenerHoja(HOJA_DEVOLUCIONES);
  const hojaDetalle = obtenerHoja(HOJA_DETALLE);

  const cabecera = hojaCabecera.getDataRange().getValues();
  for (let i = cabecera.length - 1; i >= 1; i--) {
    if (Number(cabecera[i][0]) === Number(id)) {
      hojaCabecera.deleteRow(i + 1);
      break;
    }
  }

  const detalle = hojaDetalle.getDataRange().getValues();
  for (let i = detalle.length - 1; i >= 1; i--) {
    if (Number(detalle[i][1]) === Number(id)) {
      hojaDetalle.deleteRow(i + 1);
    }
  }

  return { ok: true, mensaje: "Registro eliminado correctamente." };
}

/************************************************
 * ACTUALIZAR DEVOLUCIÓN
 ************************************************/
function actualizarDevolucion(datos) {
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

    if (filaCabecera === -1) throw new Error("No existe la devolución especificada.");

    hojaCabecera.getRange(filaCabecera, 1, 1, 6).setValues([[
      datos.id,
      datos.fecha,
      datos.funcionario,
      datos.placa,
      datos.cedula,
      datos.nombre
    ]]);

    // Eliminar detalle anterior
    const detalle = hojaDetalle.getDataRange().getValues();
    for (let i = detalle.length - 1; i >= 1; i--) {
      if (Number(detalle[i][1]) === Number(datos.id)) {
        hojaDetalle.deleteRow(i + 1);
      }
    }

    // Insertar nuevo detalle
    let idDetalle = siguienteID(hojaDetalle);
    const registros = [];

    datos.detalle.forEach(function(item) {
      registros.push([
        idDetalle++,
        datos.id,
        item.motivo,
        item.observacion || ""
      ]);
    });

    if (registros.length > 0) {
      hojaDetalle
        .getRange(hojaDetalle.getLastRow() + 1, 1, registros.length, registros[0].length)
        .setValues(registros);
    }

    return { ok: true, mensaje: "Devolución actualizada correctamente." };

  } catch (error) {
    return { ok: false, mensaje: error.message };
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
 * FUNCIÓN AUXILIAR PARA OBTENER LAS IMÁGENES
 ************************************************/
function obtenerBlobImagen(url) {
  try {
    // Si es URL de GitHub Pages, la convertimos a rawusercontent por seguridad
    let urlRaw = url.replace("github.io", "raw.githubusercontent.com").replace("/DEVOLUCION/", "/DEVOLUCION/main/");
    
    const respuesta = UrlFetchApp.fetch(urlRaw, {
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" }
    });
    
    if (respuesta.getResponseCode() === 200) {
      return respuesta.getBlob();
    } else {
      // Intentar con la URL original por si no es GitHub Pages
      const respOriginal = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (respOriginal.getResponseCode() === 200) return respOriginal.getBlob();
    }
  } catch (e) {
    Logger.log("Error al cargar imagen (" + url + "): " + e.message);
  }
  return null;
}

/************************************************
 * GENERAR PDF CON ENCABEZADO Y PIE DE PÁGINA
 ************************************************/
function generarPDFDevolucion(id) {
  try {
    const datos = buscarDevolucion(id);
    if (!datos) throw new Error("No se encontró la devolución ID: " + id);

    const doc = DocumentApp.create("DEVOLUCION_" + id);
    const body = doc.getBody();

    // ==========================================
    // 1. IMAGEN DE ENCABEZADO (PARTE SUPERIOR)
    // ==========================================
    const imgEncabezadoBlob = obtenerBlobImagen(encabezadoURL);
    if (imgEncabezadoBlob) {
      const header = doc.getHeader() || doc.addHeader();
      const pHeader = header.appendParagraph("");
      pHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      const img = pHeader.appendImage(imgEncabezadoBlob);
      // Ajustar dimensiones según necesites (Ancho, Alto en píxeles)
      img.setWidth(500).setHeight(80); 
    }

    // ==========================================
    // 2. CONTENIDO DEL DOCUMENTO
    // ==========================================
    const titulo = body.appendParagraph("SECRETARIA DE MOVILIDAD\nMUNICIPIO DE LA CEJA - ANTIOQUIA\nCOMPROBANTE DE DEVOLUCIÓN DE TRÁMITE #" + datos.id);
    titulo.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titulo.getAttributes()[DocumentApp.Attribute.FONT_SIZE] = 14;

    body.appendParagraph(""); // Espacio

    // Tabla de Datos
    const tablaDatos = body.appendTable([
      ["Fecha", datos.fecha, "Placa", datos.placa],
      ["Cédula", datos.cedula, "", ""],
      ["Ciudadano", datos.nombre, "", ""]
    ]);

    body.appendParagraph("\nMotivos de Devolución").setBold(true);

    // Tabla de Detalle
    const tablaDetalle = body.appendTable();
    const headerRow = tablaDetalle.appendTableRow();
    headerRow.appendTableCell("Motivo").setBold(true);
    headerRow.appendTableCell("Observación").setBold(true);

    datos.detalle.forEach(function(item) {
      const fila = tablaDetalle.appendTableRow();
      fila.appendTableCell(item.motivo);
      fila.appendTableCell(item.observacion);
    });

    body.appendParagraph("\n\n\n");

    // Firmas
    const tablaFirmas = body.appendTable([
      ["___________________________________", "___________________________________"],
      [datos.funcionario + "\nFuncionario Responsable", datos.nombre + "\nCiudadano / Recibido"]
    ]);
    tablaFirmas.setBorderWidth(0);

    // ==========================================
    // 3. IMAGEN DE PIE DE PÁGINA (PARTE INFERIOR)
    // ==========================================
    const imgPieBlob = obtenerBlobImagen(pieURL);
    if (imgPieBlob) {
      const footer = doc.getFooter() || doc.addFooter();
      const pFooter = footer.appendParagraph("");
      pFooter.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      const imgPie = pFooter.appendImage(imgPieBlob);
      // Ajustar dimensiones (Ancho, Alto en píxeles)
      imgPie.setWidth(500).setHeight(50);
    }

    doc.saveAndClose();

    // ==========================================
    // 4. EXPORTAR A BASE64 Y BORRAR TEMPORAL
    // ==========================================
    const pdfBlob = DriveApp.getFileById(doc.getId()).getBlob();
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    // Borrar el borrador de Google Docs
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return {
      ok: true,
      nombreArchivo: "DEVOLUCION_" + id + ".pdf",
      base64: pdfBase64
    };

  } catch (error) {
    return { ok: false, mensaje: "Error al generar PDF: " + error.message };
  }
}
