/*function copiarYOrdenar(evento){
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName("Gestor de correos")
  const rangoEditado = evento.range.getA1Notation()
  const info = hoja.getRange(rangoEditado).getValues()
  Logger.log(info)
  
  libro.toast(rangoEditado)
  
}*/

function onEdit() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = libro.getActiveSheet();
  const nombreOrigen = hojaOrigen.getSheetName();
  const hojaDestino = libro.getSheetByName('Seguimiento');
  const celdaActiva=hojaOrigen.getActiveCell();
  const filaActiva=celdaActiva.getRow();
  const columnaActiva=celdaActiva.getColumn();
  const valor=celdaActiva.getValue()
  Logger.log(nombreOrigen)
  Logger.log(columnaActiva)
  Logger.log(valor)

  if(filaActiva>=2 && columnaActiva==16 && valor==="OK" && nombreOrigen=="Gestor de correos"){
    var rangoDatosOrigen = hojaOrigen.getRange(filaActiva,1,1,hojaOrigen.getLastColumn());
   
    var rangoDatosDestino = hojaDestino.getRange(hojaDestino.getLastRow()+1,1);
    
    rangoDatosOrigen.moveTo(rangoDatosDestino);
    hojaOrigen.deleteRow(filaActiva)
    var fechaWhatsApp = hojaDestino.getRange(hojaDestino.getLastRow(),15).getValue()
    if (fechaWhatsApp=="OK"){
    fechaWhatsApp = new Date()
    hojaDestino.getRange(hojaDestino.getLastRow(),15).setValue(fechaWhatsApp)
    }

  }else if(filaActiva>=2 && columnaActiva==16 && valor==="" && nombreOrigen=="Seguimiento"){
    const hojaDestino=libro.getSheetByName('Gestor de correos')
    var rangoDatosOrigen = hojaOrigen.getRange(filaActiva,1,1,hojaOrigen.getLastColumn());
    var rangoDatosDestino = hojaDestino.getRange(hojaDestino.getLastRow()+1,1);
    
    rangoDatosOrigen.moveTo(rangoDatosDestino);
    hojaOrigen.deleteRow(filaActiva)
    var fechaWhatsApp = hojaDestino.getRange(hojaDestino.getLastRow(),15).getValue()
    fechaWhatsApp = ""
    hojaDestino.getRange(hojaDestino.getLastRow(),15).setValue(fechaWhatsApp)
  }

 /* var rangoDatosOrigen = hojaOrigen.getRange(hojaOrigen.getLastRow(),1,1,hojaOrigen.getLastColumn());
  var rangoDatosDestino = hojaDestino.getRange(hojaDestino.getLastRow()+1,1,1,10);
  var rangoDatosDestinoPenultimo = hojaDestino.getRange(hojaDestino.getLastRow(),1,1,10);
    
  rangoDatosOrigen.copyTo(rangoDatosDestino, SpreadsheetApp.CopyPasteType.PASTE_NORMAL)
  }*/
  
}

function copiaDeFacebook() {
  var libro = SpreadsheetApp.getActiveSpreadsheet()
  var hojaOrigen = libro.getSheetByName("BD Facebook a Gestor");
  var hojaDestino = libro.getSheetByName("Gestor de correos");

  var rangoDatosOrigen = hojaOrigen.getRange(2,1,hojaOrigen.getLastRow(), 9);
  var rangoDatosDestino = hojaDestino.getRange(hojaDestino.getLastRow()+1,1);
  rangoDatosOrigen.copyTo(rangoDatosDestino);
}
