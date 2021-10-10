//Variables globales

const ss = SpreadsheetApp.getActiveSpreadsheet();
const hojaOrigen = ss.getSheetByName('inventario y resp');
const hojaDestino = ss.getActiveSheet();
const ultimaFilaOrigen = hojaOrigen.getLastRow();
const rangoOrigen = hojaOrigen.getRange(7,1,ultimaFilaOrigen,18)
const pcsAutorizadas = hojaOrigen.getRange('b2').getValue();

const idPlantilla = '10APQHQDVjb2DtbSoScT81r1ZomIreaDhOdDLNC9Fdbk';
const url = "https://docs.google.com/spreadsheets/d/"+idPlantilla;



const onOpen = () => {

  let menu=SpreadsheetApp.getUi().createMenu('Controles');
  menu.addItem("Pegar PC's",'pegarPCs')
  .addItem("Insertar Fecha y Hora Inicio",'fechaIni')
  .addItem("Insertar Fecha Fin",'fechaFin');
  menu.addToUi();

}


const pegarPCs = () => {

  let uI = SpreadsheetApp.getUi();

  if (hojaDestino.getSheetName()==hojaOrigen.getSheetName()){
    uI.alert('Error! Posicionese en la hoja del mes correspondiente');
    return;
  } else if (hojaOrigen.getRange('b1').getValue()=="") {
    uI.alert('Error! No se cargó el nombre de la empresa');
    return;
  }

  //Crea filtro según criterio de valor escondido BAJA
  rangoOrigen.createFilter();
  let criterio = SpreadsheetApp.newFilterCriteria().setHiddenValues(['BAJA']).build();
  hojaOrigen.getFilter().setColumnFilterCriteria(16,criterio);

  criterio = SpreadsheetApp.newFilterCriteria().build()

  //toma los valores filtrados de PCs
  hojaOrigen.getRange(7,1,100,1).activate();


  //pega los valores filtrados en la hoja destino
  let datosFiltrados = hojaOrigen.getSelection()
    .getActiveRange()
    .copyTo(hojaDestino.getRange('a3'),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); 

  //toma los valores filtrados de Ubicación
  hojaOrigen.getRange(7,9,100,1).activate();

  //pega los valores filtrados de ubicación en la hoja destino
  hojaOrigen.getSelection()
    .getActiveRange()
    .copyTo(hojaDestino.getRange('c3'),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  //toma los valores filtrados de Fecha Actualización
  hojaOrigen.getRange(7,15,100,1).activate();

  //pega los valores filtrados de ubicación en la hoja destino
  hojaOrigen.getSelection().getActiveRange()
    .copyTo(hojaDestino.getRange('b3'),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  //calcula la cantidad de PCs pegadas
  let datosPegados=hojaDestino.getRange('a4:a').getValues();
  let ultimaFilaPegada=datosPegados.filter(String).length;

  //Concicional en caso de que no haya pcs cargadas
  if (ultimaFilaPegada==0){
    hojaOrigen.getFilter().remove();
    hojaOrigen.getRange('a8').activateAsCurrentCell();
    uI.alert('Error','No hay PCs cargadas en el inventario',uI.ButtonSet.OK);
    return;}

  //Copia y pega los valores de las PCs
  hojaOrigen.getRange('r7:s').activate();

  datosFiltrados=hojaOrigen.getSelection()
  .getActiveRange()
  .copyTo(hojaDestino.getRange('Z3'),SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  datosFiltrados=hojaOrigen.getSelection()
  .getActiveRange()
  .copyTo(hojaDestino.getRange('Z3'),SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  //para remover el filtro
  ss.getActiveSheet().getFilter().remove();

  //para copiar y pegar celdas y armar la tabla
  ss.setActiveSheet(hojaDestino);

  //Si hay una sola PC en inventario, evita error en filas cuando son 0
  if (ultimaFilaPegada==1) {
    let rangoTabla=hojaDestino.getRange('a1').offset(4,3,1,21)
    hojaDestino.getRange('d4:y4').copyTo(rangoTabla);
    Logger.log(rangoTabla.getA1Notation())
    let rangoBorrar=hojaDestino.getRange(ultimaFilaPegada+4,2).offset(0,0,100,24);
    rangoBorrar.clear();
    rangoBorrar.clearDataValidations();
    return;}
  else {
  let rangoTabla=hojaDestino.getRange('a1').offset(4,3,ultimaFilaPegada-1,21)
  hojaDestino.getRange('d4:y4').copyTo(rangoTabla);
  Logger.log(rangoTabla.getA1Notation())
  let rangoBorrar=hojaDestino.getRange(ultimaFilaPegada+4,2).offset(0,0,100,24);
  Logger.log(rangoBorrar.getA1Notation());
  rangoBorrar.clear();
  rangoBorrar.clearDataValidations();
  }

  let rangoBorrar=hojaDestino.getRange(pcsAutorizadas+4,2).offset(0,0,100,24);
  Logger.log(rangoBorrar.getA1Notation());
  rangoBorrar.clear();
  rangoBorrar.clearDataValidations();

}


const fechaIni = () => {

  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getActiveSheet();
  let activa = sheet.getActiveCell();
  let nombreEmpresa = sheet.getRange('a1').getValue();
  let fila = activa.getRow();
  let equipo = sheet.getRange(fila,1).getValue();
  let mes = sheet.getRange('c1').getValue();
  Logger.log(fila);

  if (sheet.getRange(fila,21).getValue()=="") {

    let celdaFecha=sheet.getRange(fila,21).setValue(new Date())
    let celdahorario=sheet.getRange(fila,22).setValue(new Date());
    let hojaLog=ss.getSheetByName('Log horarios');
    let ultimaFilaLog = hojaLog.getLastRow() + 1;
    hojaLog.getRange(ultimaFilaLog,1).setValue(nombreEmpresa);
    hojaLog.getRange(ultimaFilaLog,2).setValue(equipo);
    hojaLog.getRange(ultimaFilaLog,3).setValue(mes);
    hojaLog.getRange(ultimaFilaLog,4).setValue(new Date());
    hojaLog.getRange(ultimaFilaLog,5).setValue(new Date());
    hojaLog.getRange(ultimaFilaLog,7).setFormula(sheet.getSheetName()+'!' + sheet.getRange(fila,22).getA1Notation());
    hojaLog.getRange(ultimaFilaLog,8).setFormula(hojaLog.getRange(ultimaFilaLog,5)
    .getA1Notation()+'-'+hojaLog.getRange(ultimaFilaLog,7)
    .getA1Notation()).setNumberFormat("0.000");

  } else { SpreadsheetApp.getUi().alert("La celda ya tiene cargada su fecha de inico") }
}


const fechaFin = () => {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let mes = sheet.getRange('c1').getValue();
  let nombreEmpresa = sheet.getRange('a1').getValue();
  let activa = sheet.getActiveCell();
  let fila = activa.getRow();
  let equipo = sheet.getRange(fila,1).getValue();

  if (sheet.getRange(fila,23).getValue()=="") {

    let celdaFecha = sheet.getRange(fila,23).setValue(new Date());
    let hojaLog = ss.getSheetByName('Log horarios');
    let ultimaFilaLog = hojaLog.getLastRow() +1;
    hojaLog.getRange(ultimaFilaLog,1).setValue(nombreEmpresa);
    hojaLog.getRange(ultimaFilaLog,2).setValue(equipo);
    hojaLog.getRange(ultimaFilaLog,3).setValue(mes);
    hojaLog.getRange(ultimaFilaLog,6).setValue(new Date());
    hojaLog.getRange(ultimaFilaLog,9).setFormula(sheet.getSheetName()+'!' + sheet.getRange(fila,23).getA1Notation());
    hojaLog.getRange(ultimaFilaLog,10).setFormula(hojaLog.getRange(ultimaFilaLog,6)
    .getA1Notation()+'-'+hojaLog.getRange(ultimaFilaLog,9)
    .getA1Notation()).setNumberFormat("0.000");

  } else { SpreadsheetApp.getUi().alert("La celda ya tiene cargada su fecha de fin") }
}


const irAlMesEnCurso = () => {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let idss = ss.getId();

  if (idss==idPlantilla) {
    SpreadsheetApp.getUi().alert('Cree el archivo primero. Usted intenta escribir en la plantilla')
    return;
  }
  let mes = ss.getSheetByName('controles').getRange('ae1').getValue();
  let hojaMes = ss.setActiveSheet(ss.getSheetByName(mes)).activate();

}


const abrirArchivoEmpresa = () => {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let idArchivo = ss.getId();

  if (idArchivo!=idPlantilla) {

  SpreadsheetApp.getUi().alert('Error, debe elegir los archivos desde la Plantilla');
  return;

  } else {

  let id = ss.getSheetByName('menu').getRange('h16').getValue();

  //ABRE EL ARCHIVO CREADO
  let url = "https://docs.google.com/spreadsheets/d/"+id;
  let html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  let userInterface = HtmlService.createHtmlOutput(html);
   SpreadsheetApp.getUi().showModalDialog(userInterface, "Abrir archivo");
  }
}


const abrirPlantilla = () => {

   let idActivo = SpreadsheetApp.getActiveSpreadsheet().getId();

   if (idPlantilla==idActivo) {
     SpreadsheetApp.getUi().alert('Usted ya se encuentra en la Plantilla');
     return;
   }

  //ABRE EL ARCHIVO CREADO
  let html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  let userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Abrir archivo");
}


const irAlInventario = () => {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let idss = ss.getId();

  if (idss==idPlantilla) {
    SpreadsheetApp.getUi().alert('Cree el archivo primero. Usted intenta escribir en la plantilla')
    return;
  }

  let hojaInve = ss.setActiveSheet(ss.getSheetByName('inventario y resp')).activate();

}
