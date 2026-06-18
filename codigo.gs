/**
 * MENÚ PERSONALIZADO PARA GESTIÓN DE COLUMNAS
 * Este menú aparecerá en la barra superior de Google Sheets
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Administración')
      .addItem('Ocultar Columna A (Todas las pestañas)', 'ocultarColumnaAEnTodas')
      .addItem('Mostrar Columna A (Todas las pestañas)', 'mostrarColumnaAEnTodas')
      .addSeparator() // Una línea divisoria
      .addItem('Convertir vacíos en 0 (Selección)', 'ponerCerosEnSeleccion')
      .addToUi();
}

function onEdit(e) {
  var hoja = e.source.getActiveSheet();
  var rango = e.range;
  
  // Si el cambio fue en la celda B1 y es de la hoja "Panel"
  if (hoja.getName() == "9-4" && rango.getA1Notation() == "B1") {
    if (rango.getValue() == true) {
      ocultarColumnaAEnTodas();
    } else {
      mostrarColumnaAEnTodas();
    }
  }
}

/**
 * Función para ocultar la Columna A en todas las hojas del libro
 */
function ocultarColumnaAEnTodas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = ss.getSheets();
  
  hojas.forEach(function(hoja) {
    // La columna 1 es la columna 'A'
    hoja.hideColumns(1);
  });
  
  SpreadsheetApp.getUi().alert('Se ha ocultado la primera columna en las ' + hojas.length + ' pestañas.');
}

/**
 * Función para volver a mostrar la Columna A en todas las hojas
 */
function mostrarColumnaAEnTodas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = ss.getSheets();
  
  hojas.forEach(function(hoja) {
    hoja.showColumns(1);
  });
  
  SpreadsheetApp.getUi().alert('La primera columna es visible nuevamente en todas las pestañas.');
}

// --- TUS FUNCIONES ACTUALES (doGet, doPost) CONTINÚAN ABAJO ---
/**
 * Versión adaptada para Libro Único con 5 pestañas (8-3, 8-4, 9-1, 9-2, 9-3)
 */

/**
 * Versión adaptada para Libro Único con 5 pestañas (8-3, 8-4, 9-1, 9-2, 9-3)
 */

function doGet(e) {
  try {
    // Obtenemos el grupo desde el parámetro de la URL (ej: ?grupo=8-3)
    var nombreGrupo = e.parameter.grupo;
    
    if (!nombreGrupo) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "No se especificó el grupo" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(nombreGrupo);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "La pestaña del grupo no existe" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Leemos los nombres de la columna A (desde la fila 2)
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    
    var studentNames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var namesList = studentNames.map(row => row[0]).filter(name => name !== "");
    
    return ContentService.createTextOutput(JSON.stringify(namesList))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Evita colisiones entre envíos simultáneos

  try {
    var data = JSON.parse(e.postData.contents);
    
    // Extraemos los datos enviados desde el juego
    var nombre = data.nombre;
    var puntaje = data.puntaje;
    var juego = data.juego;
    var grupo = data.grupo; 

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(grupo);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        "status": "error", 
        "message": "El grupo '" + grupo + "' no existe en el libro."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- 1. BUSCAR LA FILA DEL ESTUDIANTE (Columna A) ---
    var lastRow = sheet.getLastRow();
    var playerRow = 0;
    
    if (lastRow > 0) {
      var columnANames = sheet.getRange(1, 1, lastRow, 1).getValues().map(row => row[0].toString().trim());
      playerRow = columnANames.indexOf(nombre.toString().trim()) + 1;
    }

    // Si no existe el nombre, lo agregamos al final de la lista
    if (playerRow <= 0) {
      sheet.appendRow([nombre]);
      playerRow = sheet.getLastRow();
    }

    // --- 2. BUSCAR O CREAR LA COLUMNA DEL JUEGO (Fila 1) ---
    var lastColumn = sheet.getLastColumn();
    var gameColumn = 0;
    var headers = [];
    
    if (lastColumn > 0) {
      headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
      gameColumn = headers.indexOf(juego) + 1;
    }

    if (gameColumn <= 0) {
      // Si el juego no tiene columna, creamos una nueva al final
      gameColumn = (lastColumn > 0) ? lastColumn + 1 : 1;
      sheet.getRange(1, gameColumn).setValue(juego);
    }

    // --- 3. ESCRIBIR EL PUNTAJE COMO NÚMERO ---
    // Convertimos el puntaje a número decimal puro
    var valorNumerico = parseFloat(puntaje);
    
    if (!isNaN(valorNumerico)) {
      var celda = sheet.getRange(playerRow, gameColumn);
      celda.setValue(valorNumerico);
      // Aplicamos formato numérico con 2 decimales (Sheets usará la coma si es tu config regional)
      celda.setNumberFormat("#,##0.00");
    } else {
      // Si por alguna razón no es un número, lo guarda como texto para no perder el dato
      sheet.getRange(playerRow, gameColumn).setValue(puntaje);
    }

    return ContentService.createTextOutput(JSON.stringify({ 
        "status": "success", 
        "message": "Puntaje de " + nombre + " guardado como NÚMERO en " + grupo 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
        "status": "error", 
        "message": error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}


function doPost2(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Evita colisiones si varios alumnos envían al mismo tiempo

  try {
    var data = JSON.parse(e.postData.contents);
    
    // Extraemos datos: ahora el frontend envía también el 'grupo'
    var nombre = data.nombre;
    var puntaje = data.puntaje;
    var juego = data.juego;
    var grupo = data.grupo; 

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(grupo);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({"status":"error", "message": "Grupo no encontrado"}));
    }

        // --- EL ÚNICO CAMBIO: Convertir punto a coma y asegurar que sea texto ---
    var puntajeFinal = puntaje.toString().replace(".", ",");

    // --- 1. BUSCAR LA FILA DEL ESTUDIANTE (Columna A) ---
    var columnANames = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(row => row[0]);
    var playerRow = columnANames.indexOf(nombre) + 1;

    // Si no existe el nombre, lo agregamos al final
    if (playerRow === 0) {
      sheet.appendRow([nombre]);
      playerRow = sheet.getLastRow();
    }

    // --- 2. BUSCAR O CREAR LA COLUMNA DEL JUEGO (Fila 1) ---
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var gameColumn = headers.indexOf(juego) + 1;

    if (gameColumn === 0) {
      // Si el juego no tiene columna, creamos una nueva al final
      gameColumn = headers.length + 1;
      sheet.getRange(1, gameColumn).setValue(juego);
    }

    // --- 3. ESCRIBIR EL PUNTAJE EN LA INTERSECCIÓN ---
        // Agregamos el apóstrofe (') para que Google Sheets NO lo convierta en fecha jamás
    sheet.getRange(playerRow, gameColumn).setValue("'" + puntajeFinal);
  
    
    // Opcional: Podrías poner la fecha en una columna fija si quisieras, 
    // pero esta lógica mantiene tu estilo de "Nombre vs Juego"

    return ContentService.createTextOutput(JSON.stringify({ 
        "status": "success", 
        "message": "Puntaje de " + nombre + " actualizado en " + grupo 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }));
  } finally {
    lock.releaseLock();
  }
}



function doPost2(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); 

  try {
    var data = JSON.parse(e.postData.contents);
    
    var nombre = data.nombre;
    var puntaje = data.puntaje; // Aquí llega el 5.3 o 5,3
    var juego = data.juego;
    var grupo = data.grupo; 

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(grupo);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({"status":"error", "message": "Grupo no encontrado"}));
    }

    // --- EL ÚNICO CAMBIO: Convertir punto a coma y asegurar que sea texto ---
    var puntajeFinal = puntaje.toString().replace(".", ",");

    // --- 1. BUSCAR LA FILA DEL ESTUDIANTE ---
    var columnANames = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(row => row[0]);
    var playerRow = columnANames.indexOf(nombre) + 1;

    if (playerRow === 0) {
      sheet.appendRow([nombre]);
      playerRow = sheet.getLastRow();
    }

    // --- 2. BUSCAR O CREAR LA COLUMNA DEL JUEGO ---
    var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    var gameColumn = headers.indexOf(juego) + 1;

    if (gameColumn === 0) {
      gameColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, gameColumn).setValue(juego);
    }

    // --- 3. ESCRIBIR EL PUNTAJE ---
    // Agregamos el apóstrofe (') para que Google Sheets NO lo convierta en fecha jamás
    sheet.getRange(playerRow, gameColumn).setValue("'" + puntajeFinal);
    
    return ContentService.createTextOutput(JSON.stringify({ 
        "status": "success", 
        "message": "Puntaje de " + nombre + " actualizado a " + puntajeFinal 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }));
  } finally {
    lock.releaseLock();
  }
}



