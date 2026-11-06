// ============================================================
//  XV MORE - Apps Script
//  Recibe: confirmaciones (rsvp), canciones (cancion) y fotos (foto)
// ============================================================

var CARPETA_FOTOS = "1ePxLg4jjDmUne9i3A0AA-mNjn-WDZYFb";   // carpeta de Drive para las fotos
var EMAIL_AVISO   = "matieguido@gmail.com,mnosdeo@gmail.com"; // mails que reciben el aviso (separados por coma)

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var tipo = data.tipo || "rsvp";

    if (tipo === "cancion") return guardarCancion(data);
    if (tipo === "foto")    return guardarFoto(data);
    return guardarConfirmacion(data);

  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}

// ---------- CONFIRMACIONES (primera hoja: una fila por persona) ----------
function guardarConfirmacion(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var nombre = (data.nombre || "Sin nombre").toString().trim();
  var asiste = data.asiste || "No especificado";
  var notas  = data.mensaje || "";

  var lista = data.lista || (data.acompanantes ? data.acompanantes.toString().split(",") : []);
  lista = lista.map(function (n) { return n.toString().trim(); }).filter(String);
  if (asiste !== "Sí") lista = [];

  var ahora = new Date();
  // Columnas: Fecha | Nombre | Asiste | Tipo | Confirmado por | Notas
  var filas = [[ahora, nombre, asiste, "Titular", nombre, notas]];
  lista.forEach(function (n) { filas.push([ahora, n, asiste, "Acompañante", nombre, ""]); });
  sheet.getRange(sheet.getLastRow() + 1, 1, filas.length, 6).setValues(filas);

  if (asiste === "Sí") {
    var asunto  = "¡Nueva confirmación de " + nombre + "!";
    var mensaje = nombre + " confirmó asistencia para " + filas.length + " persona(s).";
    if (lista.length) mensaje += "\nAcompañantes: " + lista.join(", ");
    if (notas)        mensaje += "\n" + notas;
    MailApp.sendEmail(EMAIL_AVISO, asunto, mensaje);
  }
  return ok();
}

// ---------- CANCIONES (hoja "Canciones", se crea sola) ----------
function guardarCancion(data) {
  var cancion = (data.cancion || "").toString().trim();
  if (!cancion) return ok();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Canciones");
  if (!sheet) {
    sheet = ss.insertSheet("Canciones");
    sheet.appendRow(["Fecha", "Canción sugerida"]);
  }
  sheet.appendRow([new Date(), cancion]);
  return ok();
}

// ---------- FOTOS (a la carpeta de Drive) ----------
function guardarFoto(data) {
  if (!data.data) return ok();
  var tipos = ["image/jpeg", "image/png", "image/webp"];
  var mime = tipos.indexOf(data.mime) >= 0 ? data.mime : "image/jpeg";
  var autor = (data.autor || "Invitado").toString().replace(/[^\wÁÉÍÓÚÑáéíóúñ ]/g, "").trim().slice(0, 40) || "Invitado";
  var nombreArchivo = Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "yyyyMMdd-HHmmss")
                      + " - " + autor + " - " + Math.floor(Math.random() * 10000) + ".jpg";
  var blob = Utilities.newBlob(Utilities.base64Decode(data.data), mime, nombreArchivo);
  DriveApp.getFolderById(CARPETA_FOTOS).createFile(blob);
  return ok();
}

function ok() {
  return ContentService.createTextOutput("Éxito").setMimeType(ContentService.MimeType.TEXT);
}

// Ejecutá esta función UNA vez desde el editor para dar el permiso de Drive.
function autorizar() {
  DriveApp.getFolderById(CARPETA_FOTOS).getName();
  SpreadsheetApp.getActiveSpreadsheet().getName();
}
