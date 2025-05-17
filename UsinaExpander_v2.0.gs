/**
 * API_KEY=AIzaSyBgmk4IMW5XjsOGB6fP2N59aBjlLucLKwU
 * response json ID=AAKfycbx8SvNNDKnKHQH6Gz8Dr3ATtuVwj0-gFmffJo_nM6Dadm-FwVmnq1ic2AUfTbIMR7BZ
 * correo de servicio = usina-nomina-service@nomina-usina.iam.gserviceaccount.com
 */

// CONSTANTES GLOBALES
const ID_FRANQUICIAS = "1_rNr_zdX4xogtGIZxUOWKyXRAg5PWdeLpkXtlj5yjms";
const ID_AUTO_DERIV = "1_rNr_zdX4xogtGIZxUOWKyXRAg5PWdeLpkXtlj5yjms";
const FRANQUICIAS_VALIDAS = [
  "ATENEA",
  "EROS",
  "FENIX",
  "FLASHBET",
  "GANA24",
  "FORTUNA",
  "PADRINO",
];
const COMANDOS_VALIDOS = ["a1", "a3"];

function doGet(e) {
  try {
    const {
      accion,
      nombre,
      orden,
      comando,
      sheet: sheetName,
      cell: cellRef,
      spreadsheetId,
      fila,
      tipo,
    } = parseParams(e);

    switch (accion) {
      case "getFranquicia":
        return handleGetFranquicia(nombre);
      case "getTelefonos":
        return handleGetTelefonos(nombre);
      // NUEVO: Endpoint para obtener teléfonos con datos de cargas
      case "getTelefonosConCargas":
        return handleGetTelefonosConCargas(nombre);
      case "incrementarOrden":
        return handleIncrementarOrden(nombre, orden);
      case "contarUso":
        return handleContarUso(comando);
      case "registrarUso":
        return handleRegistroUso(sheetName, cellRef);
      // NUEVO: Endpoint para registrar cargas en columna H
      case "registrarCarga":
        return handleRegistroCarga(sheetName, cellRef);
      case "sumarValor":
        return handleSumarValor(sheetName, cellRef, spreadsheetId);
      case "registrarPublicidadExtra":
        return handleRegistrarPublicidadExtra(fila, tipo);
      default:
        return crearRespuestaError("Acción no válida");
    }
  } catch (error) {
    console.error("Error en doGet:", error);
    return crearRespuestaError("Error interno del servidor");
  }
}

function handleRegistrarPublicidadExtra(fila, tipo) {
  if (!fila || !["a1", "a3"].includes(tipo)) {
    return crearRespuestaError("Parámetros inválidos");
  }

  const hoja = SpreadsheetApp.openById(
    "1vq2C76OM7z2CWJPkDwYufPHy2vwfqjtaGXnT7ryGnQw"
  ).getSheetByName("SERVER4");

  const columna = tipo === "a1" ? 9 : 10; // I=9, J=10
  const celda = hoja.getRange(fila, columna);
  const valorActual = Number(celda.getValue()) || 0;
  celda.setValue(valorActual + 1);

  return crearRespuestaJson({ ok: true, fila, tipo });
}

// ===== HANDLERS PRINCIPALES =====
function handleGetFranquicia(nombre) {
  validarParametro(nombre, "nombre");
  return crearRespuestaJson(getDatosFranquicia(nombre));
}

function handleGetTelefonos(nombre) {
  validarParametro(nombre, "nombre");
  return crearRespuestaJson(getTelefonosActivos(nombre));
}

// NUEVO: Handler para obtener teléfonos con datos de cargas
function handleGetTelefonosConCargas(nombre) {
  validarParametro(nombre, "nombre");
  return crearRespuestaJson(getTelefonosActivosConCargas(nombre));
}

function handleIncrementarOrden(nombre, orden) {
  validarParametro(nombre, "nombre");
  validarParametro(orden, "orden");
  return crearRespuestaJson(incrementarConteoTelefono(nombre, orden));
}

function handleContarUso(comando) {
  validarParametro(comando, "comando");
  if (!COMANDOS_VALIDOS.includes(comando)) {
    return crearRespuestaError("Comando no válido");
  }
  return crearRespuestaJson(incrementarContadorComando(comando));
}

function handleSumarValor(sheetName, cell, spreadsheetId) {
  try {
    const spreadsheet = spreadsheetId
      ? SpreadsheetApp.openById(spreadsheetId)
      : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    const range = sheet.getRange(cell);
    const currentValue = range.getValue() || 0;
    range.setValue(currentValue + 1);
    return crearRespuestaOk();
  } catch (e) {
    return crearRespuestaError(e.message);
  }
}

function handleRegistroUso(sheetName, cellRef) {
  validarParametro(sheetName, "sheet");
  validarParametro(cellRef, "cell");

  const spreadsheet = SpreadsheetApp.openById(ID_AUTO_DERIV);
  const sheet = obtenerHoja(spreadsheet, sheetName);
  const range = sheet.getRange(cellRef);

  const currentValue = Number(range.getValue()) || 0;
  const newValue = currentValue + 1;

  range.setValue(newValue);
  registrarEnLogs(spreadsheet, sheetName, cellRef, newValue);

  // Actualizar totales y verificar resultado
  if (!actualizarTotales(spreadsheet, sheetName)) {
    console.warn("No se pudo actualizar totales para:", sheetName);
  }

  SpreadsheetApp.flush();
  return crearRespuestaJson({ ok: true, newValue });
}

// NUEVO: Handler para registrar cargas en columna H
function handleRegistroCarga(sheetName, cellRef) {
  validarParametro(sheetName, "sheet");
  validarParametro(cellRef, "cell");

  // Verificar que la celda sea de columna H
  if (!cellRef.toUpperCase().startsWith("H")) {
    // Convertir referencia de columna G a H
    if (cellRef.toUpperCase().startsWith("G")) {
      const fila = cellRef.substring(1);
      cellRef = "H" + fila;
    } else {
      return crearRespuestaError(
        "La celda debe ser de columna H para registrar cargas"
      );
    }
  }

  const spreadsheet = SpreadsheetApp.openById(ID_AUTO_DERIV);
  const sheet = obtenerHoja(spreadsheet, sheetName);
  const range = sheet.getRange(cellRef);

  const currentValue = Number(range.getValue()) || 0;
  const newValue = currentValue + 1;

  range.setValue(newValue);
  registrarEnLogs(spreadsheet, sheetName, cellRef, newValue);

  SpreadsheetApp.flush();
  return crearRespuestaJson({ ok: true, newValue });
}

// ===== FUNCIONES DE DATOS =====
function getDatosFranquicia(nombre) {
  if (!FRANQUICIAS_VALIDAS.includes(nombre)) {
    return { error: "Franquicia no válida" };
  }

  try {
    const hoja =
      SpreadsheetApp.openById(ID_FRANQUICIAS).getSheetByName("FRANQUICIAS");
    const datos = hoja.getRange("A2:F" + hoja.getLastRow()).getValues();
    const encabezados = ["nombre", "pass", "cvu", "alias", "titular", "link"];

    for (const fila of datos) {
      if (fila[0].toString().toUpperCase().trim() === nombre) {
        return crearObjetoFranquicia(encabezados, fila);
      }
    }
    return { error: "Franquicia no encontrada" };
  } catch (error) {
    console.error("Error en getDatosFranquicia:", error);
    return { error: "Error al obtener datos" };
  }
}

function getTelefonosActivos(nombre) {
  if (!FRANQUICIAS_VALIDAS.includes(nombre)) {
    return { error: "Franquicia no válida" };
  }

  try {
    const hoja = obtenerHoja(SpreadsheetApp.openById(ID_FRANQUICIAS), nombre);
    const datos = hoja.getRange("A2:C" + hoja.getLastRow()).getValues();

    return datos
      .filter((fila) => fila[2]?.toString().toUpperCase() === "ACTIVO")
      .map((fila) => ({
        orden: fila[0]?.toString().padStart(2, "0") ?? "00",
        numero: fila[1]?.toString() ?? "",
      }));
  } catch (error) {
    console.error("Error en getTelefonosActivos:", error);
    return { error: "Error al obtener teléfonos" };
  }
}

// NUEVO: Función para obtener teléfonos activos con datos de enviados y cargas
function getTelefonosActivosConCargas(nombre) {
  if (!FRANQUICIAS_VALIDAS.includes(nombre)) {
    return { error: "Franquicia no válida" };
  }

  try {
    // Obtener datos básicos de teléfonos (A:orden, B:numero, C:estado, D:meta, E:disponibles, F:overgoal)
    const hojaFranquicias = obtenerHoja(
      SpreadsheetApp.openById(ID_FRANQUICIAS),
      nombre
    );
    const lastRow = hojaFranquicias.getLastRow();
    const datosTelefonos = hojaFranquicias
      .getRange("A2:F" + lastRow)
      .getValues();

    // Filtrar solo los activos
    const telefonosActivos = datosTelefonos
      .filter((fila) => fila[2]?.toString().toUpperCase() === "ACTIVO")
      .map((fila) => ({
        orden: fila[0]?.toString().padStart(2, "0") ?? "00",
        numero: fila[1]?.toString() ?? "",
        meta: Number(fila[3]) || 1, // Columna D
        disponibles: Number(fila[4]) || 0, // Columna E
        overgoal:
          typeof fila[5] !== "undefined" &&
          String(fila[5]).toLowerCase() === "true"
            ? true
            : false, // Columna F
        enviados: 0, // Valor por defecto
        cargas: 0, // Valor por defecto
      }));

    if (telefonosActivos.length === 0) {
      return [];
    }

    // Obtener datos de enviados (columna G) y cargas (columna H)
    const hojaDerivaciones = obtenerHoja(
      SpreadsheetApp.openById(ID_AUTO_DERIV),
      nombre
    );
    const datosDerivaciones = hojaDerivaciones
      .getRange("A2:H" + hojaDerivaciones.getLastRow())
      .getValues();

    // Mapear datos de enviados y cargas a los teléfonos activos
    for (const telefono of telefonosActivos) {
      for (const fila of datosDerivaciones) {
        const ordenFila = fila[0]?.toString().padStart(2, "0") ?? "";
        if (ordenFila === telefono.orden) {
          telefono.enviados = Number(fila[6]) || 0; // Columna G (índice 6)
          telefono.cargas = Number(fila[7]) || 0; // Columna H (índice 7)
          break;
        }
      }
    }

    return telefonosActivos;
  } catch (error) {
    console.error("Error en getTelefonosActivosConCargas:", error);
    return { error: "Error al obtener teléfonos con cargas" };
  }
}

// ===== FUNCIONES DE ACTUALIZACIÓN =====
function incrementarConteoTelefono(nombre, orden) {
  try {
    const spreadsheet = SpreadsheetApp.openById(ID_AUTO_DERIV);
    const hoja = obtenerHoja(spreadsheet, nombre);
    const datos = hoja.getRange("A2:A" + hoja.getLastRow()).getValues();

    for (let i = 0; i < datos.length; i++) {
      if (
        datos[i][0]?.toString().trim().padStart(2, "0") ===
        orden.padStart(2, "0")
      ) {
        const celda = hoja.getRange(i + 2, 7); // Columna G
        const nuevoValor = (Number(celda.getValue()) || 0) + 1;

        celda.setValue(nuevoValor);
        registrarEnLogs(spreadsheet, nombre, `G${i + 2}`, nuevoValor);
        actualizarTotales(spreadsheet, nombre);

        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
    return { error: "Orden no encontrada" };
  } catch (error) {
    console.error("Error en incrementarConteoTelefono:", error);
    return { error: "Error al incrementar conteo" };
  }
}

function incrementarContadorComando(comando) {
  try {
    const hoja = obtenerHoja(SpreadsheetApp.openById(ID_AUTO_DERIV), "TOTALES");
    const columna = comando === "a1" ? 1 : 2; // A=1 (a1), B=2 (a3)

    const celda = hoja.getRange(2, columna);
    celda.setValue((Number(celda.getValue()) || 0) + 1);

    SpreadsheetApp.flush();
    return { ok: true };
  } catch (error) {
    console.error("Error en incrementarContadorComando:", error);
    return { error: "Error al incrementar contador" };
  }
}

function sumarValor(sheetName, cell, spreadsheetId) {
  try {
    var spreadsheet = spreadsheetId
      ? SpreadsheetApp.openById(spreadsheetId)
      : SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    var range = sheet.getRange(cell);
    var currentValue = range.getValue() || 0;
    range.setValue(currentValue + 1);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ===== FUNCIONES AUXILIARES =====
function parseParams(e) {
  return {
    accion: e.parameter.accion,
    nombre: e.parameter.nombre,
    orden: e.parameter.orden,
    comando: e.parameter.comando,
    sheet: e.parameter.sheet,
    cell: e.parameter.cell,
    spreadsheetId: e.parameter.spreadsheetId,
    fila: e.parameter.fila,
    tipo: e.parameter.tipo,
  };
}

function validarParametro(valor, nombreParam) {
  if (!valor) throw new Error(`Parámetro '${nombreParam}' requerido`);
}

function obtenerHoja(spreadsheet, nombreHoja) {
  const hoja = spreadsheet.getSheetByName(nombreHoja);
  if (!hoja) throw new Error(`Hoja '${nombreHoja}' no encontrada`);
  return hoja;
}

function crearObjetoFranquicia(encabezados, fila) {
  const json = {};
  encabezados.forEach(
    (col, i) => (json[col] = fila[i]?.toString()?.trim() || "")
  );
  return json;
}

function registrarEnLogs(spreadsheet, franquicia, celda, valor) {
  const logSheet =
    spreadsheet.getSheetByName("REGISTROS") ||
    spreadsheet.insertSheet("REGISTROS");
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Fecha", "Franquicia", "Celda", "Valor"]);
  }
  logSheet.appendRow([new Date(), franquicia, celda, valor]);
}

function actualizarTotales(spreadsheet, franquicia) {
  try {
    const timezone = Session.getScriptTimeZone();
    const ahora = new Date();
    const fechaFormateada = Utilities.formatDate(ahora, timezone, "dd/MM/yyyy");

    const hojaTotales = obtenerHojaTotales(spreadsheet);
    const { fila, existe } = encontrarFilaTotales(
      hojaTotales,
      fechaFormateada,
      franquicia
    );

    if (!existe) {
      // Calcular acumulado inicial
      const acumuladoInicial = calcularAcumulado(hojaTotales, franquicia);

      hojaTotales.getRange(fila, 1).setValue(fechaFormateada); // Fecha
      hojaTotales.getRange(fila, 2).setValue(franquicia); // Franquicia
      hojaTotales.getRange(fila, 3).setValue(0); // Diario (inicia en 0)
      hojaTotales.getRange(fila, 4).setValue(acumuladoInicial); // Acumulado
    }

    // Actualizar valores
    const rango = hojaTotales.getRange(fila, 3, 1, 3); // C, D, E
    const valores = rango.getValues();

    const nuevoDiario = Number(valores[0][0]) + 1;
    const nuevoAcumulado = Number(valores[0][1]) + 1;

    rango.setValues([[nuevoDiario, nuevoAcumulado, ahora]]);

    return true;
  } catch (error) {
    console.error("Error en actualizarTotales:", error);
    return false;
  }
}

function calcularAcumulado(hoja, franquicia) {
  const datos = hoja.getRange("B2:D" + hoja.getLastRow()).getValues();
  let acumulado = 0;

  for (const fila of datos) {
    if (fila[0] === franquicia) {
      acumulado += Number(fila[2]) || 0; // Suma columna Acumulado (D)
    }
  }

  return acumulado;
}

function obtenerHojaTotales(spreadsheet) {
  let hoja = spreadsheet.getSheetByName("TOTALES");
  if (!hoja) {
    hoja = spreadsheet.insertSheet("TOTALES");
    // Encabezados mejorados
    hoja
      .getRange("A1:E1")
      .setValues([
        ["Fecha", "Franquicia", "Diario", "Acumulado", "Última Actualización"],
      ]);
    hoja.getRange("A1:E1").setFontWeight("bold");
    // Formatos
    hoja.getRange("A2:A").setNumberFormat("dd/mm/yyyy");
    hoja.getRange("C2:D").setNumberFormat("#,##0");
    hoja.getRange("E2:E").setNumberFormat("dd/mm/yyyy hh:mm:ss");
    // Congelar la primera fila
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function encontrarFilaTotales(hoja, fechaFormateada, franquicia) {
  const fechas = hoja.getRange("A2:A" + hoja.getLastRow()).getDisplayValues();
  const franquicias = hoja.getRange("B2:B" + hoja.getLastRow()).getValues();

  // Buscar coincidencia exacta de fecha y franquicia
  for (let i = 0; i < fechas.length; i++) {
    if (fechas[i][0] === fechaFormateada && franquicias[i][0] === franquicia) {
      return { fila: i + 2, existe: true };
    }
  }

  // Si no existe, buscar última fila de la misma franquicia para mantener orden
  for (let i = fechas.length - 1; i >= 0; i--) {
    if (franquicias[i][0] === franquicia) {
      return { fila: i + 3, existe: false }; // Insertar después del último de la misma franquicia
    }
  }

  // Si no hay registros de la franquicia
  return { fila: hoja.getLastRow() + 1, existe: false };
}

function crearRespuestaJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function crearRespuestaOk(data = {}) {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, ...data })
  ).setMimeType(ContentService.MimeType.JSON);
}

function crearRespuestaError(mensaje) {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: false, error: mensaje })
  ).setMimeType(ContentService.MimeType.JSON);
}
