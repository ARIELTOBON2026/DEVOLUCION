/************************************************
 * SISTEMA DE DEVOLUCIÓN DE TRÁMITES (VERSIÓN API PARA GITHUB PAGES)
 * Codigo.gs
 ************************************************/

const HOJA_DEVOLUCIONES = "DEVOLUCIONES";
const HOJA_DETALLE = "DEVOLUCIONES_DETALLES";
const HOJA_FUNCIONARIOS = "FUNCIONARIOS";
const HOJA_MOTIVOS = "MOTIVOS";

/************************************************
 * MANEJADOR DE PETICIONES GET (Para consultar datos)
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
 * MANEJADOR DE PETICIONES POST (Para enviar/guardar datos)
 ************************************************/
function doPost(e) {
  try {
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
  const hoja = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(nombre);

  if (!hoja) {
    throw new Error("No existe la hoja: " + nombre);
  }

  return hoja;
}

/************************************************
 * OBTENER SIGUIENTE ID
 ************************************************/
function siguienteID(hoja) {
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila <= 1) {
    return 1;
  }

  const ids = hoja
    .getRange(2, 1, ultimaFila - 1, 1)
    .getValues()
    .flat()
    .map(Number)
    .filter(id => !isNaN(id));

  return ids.length ? Math.max(...ids) + 1 : 1;
}

/************************************************
 * FECHA ACTUAL
 ************************************************/
function fechaActual() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
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
    .filter(nombre => nombre !== "");
}

/************************************************
 * LISTAR MOTIVOS DE DEVOLUCIÓN
 ************************************************/
function obtenerMotivos() {
  const hoja = obtenerHoja(HOJA_MOTIVOS);
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila <= 1) return [];

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
  if (!datos.detalle || datos.detalle.length === 0) {
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
        item.observacion
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
    return {
      ok: false,
      mensaje: error.message
    };
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

  const respuesta = { detalle: [] };

  for (let i = 1; i < cabecera.length; i++) {
    if (Number(cabecera[i][0]) === Number(id)) {
      respuesta.id = cabecera[i][0];
      respuesta.fecha = cabecera[i][1];
      respuesta.funcionario = cabecera[i][2];
      respuesta.placa = cabecera[i][3];
      respuesta.cedula = cabecera[i][4];
      respuesta.nombre = cabecera[i][5];
      break;
    }
  }

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
 * LISTAR DEVOLUCIONES
 ************************************************/
function listarDevoluciones() {
  const hoja = obtenerHoja(HOJA_DEVOLUCIONES);

  if (hoja.getLastRow() <= 1) return [];

  return hoja
    .getRange(2, 1, hoja.getLastRow() - 1, 6)
    .getValues();
}

/************************************************
 * BUSCAR POR PLACA
 ************************************************/
function buscarPlaca(placa) {
  placa = String(placa).trim().toUpperCase();
  const hoja = obtenerHoja(HOJA_DEVOLUCIONES);
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][3]).trim().toUpperCase() === placa) {
      return {
        id: datos[i][0],
        fecha: datos[i][1],
        funcionario: datos[i][2],
        placa: datos[i][3],
        cedula: datos[i][4],
        nombre: datos[i][5]
      };
    }
  }
  return null;
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

    if (filaCabecera === -1) throw new Error("No existe la devolución.");

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
        item.observacion
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
function generarPDFDevolucion(idDevolucion) {

  try {

    const datos = buscarDevolucion(idDevolucion);

    if (!datos || !datos.id) {
      throw new Error("No se encontró la devolución.");
    }

    /************************************************
     * ID DE LAS IMÁGENES EN DRIVE
     ************************************************/
    const ID_ENCABEZADO = "1wPY_QJ4G_W7rz5bdkz7ObGc0L0cN9_ML";
    const ID_PIE = "1m-KztMZ-KSlX-tu4BS61qrRQ-9TX8YpR";

    const encabezado =
      "data:image/jpeg;base64," +
      Utilities.base64Encode(
        DriveApp.getFileById(ID_ENCABEZADO)
        .getBlob()
        .getBytes()
      );

    const pie =
      "data:image/png;base64," +
      Utilities.base64Encode(
        DriveApp.getFileById(ID_PIE)
        .getBlob()
        .getBytes()
      );

    const fecha = Utilities.formatDate(
      new Date(datos.fecha),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

    let html = `
<!DOCTYPE html>
<html>

<head>

<meta charset="UTF-8">

<style>

@page{
size:letter;
margin:18mm;
}

body{

font-family:Arial;
font-size:11pt;
color:#222;
margin:0;
padding:0;

}

.encabezado img{

width:100%;

}

.titulo{

text-align:center;
font-size:18pt;
font-weight:bold;
margin-top:15px;
color:#003A70;

}

.subtitulo{

text-align:center;
font-size:11pt;
margin-bottom:20px;

}

table{

width:100%;
border-collapse:collapse;

}

.info td{

border:1px solid #CFCFCF;
padding:8px;

}

.label{

background:#EFEFEF;
font-weight:bold;
width:22%;

}

.detalle{

margin-top:20px;

}

.detalle th{

background:#003A70;
color:white;
padding:8px;
border:1px solid #DDD;

}

.detalle td{

padding:8px;
border:1px solid #DDD;
vertical-align:top;

}

.detalle tr:nth-child(even){

background:#F8F8F8;

}

.firmas{

margin-top:70px;

}

.linea{

width:220px;
border-top:1px solid #000;
margin:auto;

}

.footer{

margin-top:50px;

}

.footer img{

width:100%;

}

</style>

</head>

<body>

<div class="encabezado">

<img src="${encabezado}">

</div>

<div class="titulo">

COMPROBANTE DE DEVOLUCIÓN DE TRÁMITE

</div>

<div class="subtitulo">

Radicado No. ${datos.id}

</div>

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

<td colspan="3">

${datos.nombre}

</td>

</tr>

</table>

<h3 style="color:#003A70">

Motivos de devolución

</h3>

<table class="detalle">

<tr>

<th width="35%">Motivo</th>

<th>Observación</th>

</tr>
`;

    datos.detalle.forEach(function(fila){

      html += `
<tr>

<td><b>${fila.motivo}</b></td>

<td>${fila.observacion || ""}</td>

</tr>
`;

    });

    html += `

</table>

<table class="firmas">

<tr>

<td align="center">

<div class="linea"></div>

<br>

<b>${datos.funcionario}</b>

<br>

Funcionario Responsable

</td>

<td align="center">

<div class="linea"></div>

<br>

<b>${datos.nombre}</b>

<br>

Ciudadano

</td>

</tr>

</table>

<div class="footer">

<img src="${pie}">

</div>

</body>

</html>
`;

    const pdf = HtmlService
      .createHtmlOutput(html)
      .getAs(MimeType.PDF)
      .setName(
        "Devolucion_" +
        datos.placa +
        "_" +
        datos.id +
        ".pdf"
      );

    return {

      ok:true,

      base64:Utilities.base64Encode(pdf.getBytes()),

      nombreArchivo:pdf.getName()

    };

  } catch(error){

    return{

      ok:false,

      mensaje:error.message

    };

  }

}

