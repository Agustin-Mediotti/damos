//Variables globales

var ss = SpreadsheetApp.getActiveSpreadsheet();
const hojaOrigen=ss.getSheetByName('inventario y resp');
const hojaDestino=ss.getActiveSheet();

var ultimaFilaOrigen=hojaOrigen.getLastRow();
var rangoOrigen=hojaOrigen.getRange(7,1,ultimaFilaOrigen,18)
var pcsAutorizadas=hojaOrigen.getRange('b2').getValue();

function pegarPCs(){
    var uI=SpreadsheetApp.getUi();
    if (hojaDestino.getSheetName()==hojaOrigen.getSheetName()){
      uI.alert('Error! Posicionese en la hoja del mes correspondiente');
      return;
    }
    if (hojaOrigen.getRange('b1').getValue()==""){
      uI.alert('Error! No se cargó el nombre de la empresa');
      return;
    }

//Crea filtro según criterio de valor escondido BAJA
rangoOrigen.createFilter();
var criterio=SpreadsheetApp.newFilterCriteria().setHiddenValues(['BAJA']).build();
hojaOrigen.getFilter().setColumnFilterCriteria(16,criterio);
criterio=SpreadsheetApp.newFilterCriteria().build()

//toma los valores filtrados de PCs
hojaOrigen.getRange(7,1,100,1).activate();


//pega los valores filtrados en la hoja destino
var datosFiltrados=hojaOrigen.getSelection().getActiveRange().copyTo(hojaDestino.getRange('a3'),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); 

//toma los valores filtrados de Ubicación
hojaOrigen.getRange(7,9,100,1).activate();

//pega los valores filtrados de ubicación en la hoja destino
var ubcacionFiltrados=hojaOrigen.getSelection().getActiveRange().copyTo(hojaDestino.getRange('c3'),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

//toma los valores filtrados de Fecha Actualización
hojaOrigen.getRange(7,15,100,1).activate();

//pega los valores filtrados de ubicación en la hoja destino
var ubcacionFiltrados=hojaOrigen.getSelection().getActiveRange().copyTo(hojaDestino.getRange('b3'),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

//calcula la cantidad de PCs pegadas
var datosPegados=hojaDestino.getRange('a4:a').getValues();
var ultimaFilaPegada=datosPegados.filter(String).length;

//Concicional en caso de que no haya pcs cargadas
if (ultimaFilaPegada==0){
  hojaOrigen.getFilter().remove();
  hojaOrigen.getRange('a8').activateAsCurrentCell();
  uI.alert('Error','No hay PCs cargadas en el inventario',uI.ButtonSet.OK);
  return;}

//Copia y pega los valores de las PCs
hojaOrigen.getRange('r7:s').activate();
 var datosFiltrados=hojaOrigen.getSelection().getActiveRange().copyTo(hojaDestino.getRange('Z3'),SpreadsheetApp.CopyPasteType.PASTE_VALUES, false); 
 var datosFiltrados=hojaOrigen.getSelection().getActiveRange().copyTo(hojaDestino.getRange('Z3'),SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); 

//para remover el filtro
ss.getActiveSheet().getFilter().remove();

//para copiar y pegar celdas y armar la tabla

ss.setActiveSheet(hojaDestino);

//Si hay una sola PC en inventario, evita error en filas cuando son 0
if (ultimaFilaPegada==1) {
  var rangoTabla=hojaDestino.getRange('a1').offset(4,3,1,21)
  hojaDestino.getRange('d4:y4').copyTo(rangoTabla);
  Logger.log(rangoTabla.getA1Notation())
  var rangoBorrar=hojaDestino.getRange(ultimaFilaPegada+4,2).offset(0,0,100,24);
  rangoBorrar.clear();
  rangoBorrar.clearDataValidations();
  return;}
else {
var rangoTabla=hojaDestino.getRange('a1').offset(4,3,ultimaFilaPegada-1,21)
hojaDestino.getRange('d4:y4').copyTo(rangoTabla);
Logger.log(rangoTabla.getA1Notation())
var rangoBorrar=hojaDestino.getRange(ultimaFilaPegada+4,2).offset(0,0,100,24);
Logger.log(rangoBorrar.getA1Notation());
rangoBorrar.clear();
rangoBorrar.clearDataValidations();
}

var rangoBorrar=hojaDestino.getRange(pcsAutorizadas+4,2).offset(0,0,100,24);
Logger.log(rangoBorrar.getA1Notation());
rangoBorrar.clear();
rangoBorrar.clearDataValidations();

}


function fechaIni(){
var ss =SpreadsheetApp.getActiveSpreadsheet()
var sheet=ss.getActiveSheet();
var activa=sheet.getActiveCell();
var nombreEmpresa=sheet.getRange('a1').getValue();
var fila=activa.getRow();
var equipo=sheet.getRange(fila,1).getValue();
var mes=sheet.getRange('c1').getValue();
Logger.log(fila);
if (sheet.getRange(fila,21).getValue()=="") {
  var celdaFecha=sheet.getRange(fila,21).setValue(new Date())
  var celdahorario=sheet.getRange(fila,22).setValue(new Date());
  var hojaLog=ss.getSheetByName('Log horarios');
  var ultimaFilaLog=hojaLog.getLastRow();
  hojaLog.getRange(ultimaFilaLog+1,1).setValue(nombreEmpresa);
  hojaLog.getRange(ultimaFilaLog+1,2).setValue(equipo);
  hojaLog.getRange(ultimaFilaLog+1,3).setValue(mes);
  hojaLog.getRange(ultimaFilaLog+1,4).setValue(new Date());
  hojaLog.getRange(ultimaFilaLog+1,5).setValue(new Date());
  hojaLog.getRange(ultimaFilaLog+1,7).setFormula(sheet.getSheetName()+'!' + sheet.getRange(fila,22).getA1Notation());
  hojaLog.getRange(ultimaFilaLog+1,8).setFormula(hojaLog.getRange(ultimaFilaLog+1,5).getA1Notation()+'-'+hojaLog.getRange(ultimaFilaLog+1,7).getA1Notation()).setNumberFormat("0.000");
  
   }
  
  else{ mensajeAlerta();


  }}
 function mensajeAlerta(){
  SpreadsheetApp.getUi().alert("La celda ya tiene cargada su fecha de inico");
}

function fechaFin(){
var ss =SpreadsheetApp.getActiveSpreadsheet();
var sheet=ss.getActiveSheet();
var mes=sheet.getRange('c1').getValue();
var nombreEmpresa=sheet.getRange('a1').getValue();
var activa=sheet.getActiveCell();
var fila=activa.getRow();
var equipo=sheet.getRange(fila,1).getValue();
if (sheet.getRange(fila,23).getValue()=="") {
  var celdaFecha=sheet.getRange(fila,23).setValue(new Date());
  var hojaLog=ss.getSheetByName('Log horarios');
  var ultimaFilaLog=hojaLog.getLastRow();
  hojaLog.getRange(ultimaFilaLog+1,1).setValue(nombreEmpresa);
  hojaLog.getRange(ultimaFilaLog+1,2).setValue(equipo);
  hojaLog.getRange(ultimaFilaLog+1,3).setValue(mes);
  hojaLog.getRange(ultimaFilaLog+1,6).setValue(new Date());
  hojaLog.getRange(ultimaFilaLog+1,9).setFormula(sheet.getSheetName()+'!' + sheet.getRange(fila,23).getA1Notation());
  hojaLog.getRange(ultimaFilaLog+1,10).setFormula(hojaLog.getRange(ultimaFilaLog+1,6).getA1Notation()+'-'+hojaLog.getRange(ultimaFilaLog+1,9).getA1Notation()).setNumberFormat("0.000");
  }
   
  else{ 
    
  mensajeAlerta2();


  }
}
 function mensajeAlerta2(){
  SpreadsheetApp.getUi().alert("La celda ya tiene cargada su fecha de fin");
}


function onOpen(){
   crearMenu();
   
}
function crearMenu(){
var menu=SpreadsheetApp.getUi().createMenu('Controles');
menu.addItem("Pegar PC's",'pegarPCs').addItem("Insertar Fecha y Hora Inicio",'fechaIni').addItem("Insertar Fecha Fin",'fechaFin');
menu.addToUi();

}
function irAlMesEnCurso(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var idss=ss.getId();
  const idPlantilla='10APQHQDVjb2DtbSoScT81r1ZomIreaDhOdDLNC9Fdbk';
  if(idss==idPlantilla) {
    SpreadsheetApp.getUi().alert('Cree el archivo primero. Usted intenta escribir en la plantilla')
    return;
  }
  var mes=ss.getSheetByName('controles').getRange('ae1').getValue();
  var hojaMes=ss.setActiveSheet(ss.getSheetByName(mes)).activate();

}


function abrirArchivoEmpresa(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  const idPlantilla='10APQHQDVjb2DtbSoScT81r1ZomIreaDhOdDLNC9Fdbk';
  var idArchivo=ss.getId();
  if(idArchivo!=idPlantilla){
  SpreadsheetApp.getUi().alert('Error, debe elegir los archivos desde la Plantilla');
  return;}
  else{
  
  var id=ss.getSheetByName('menu').getRange('h16').getValue();

  //ABRE EL ARCHIVO CREADO
  var url = "https://docs.google.com/spreadsheets/d/"+id;
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
   SpreadsheetApp.getUi().showModalDialog(userInterface, "Abrir archivo");
  }
}
function abrirPlantilla(){
   const idPlantilla='10APQHQDVjb2DtbSoScT81r1ZomIreaDhOdDLNC9Fdbk';
   var idActivo=SpreadsheetApp.getActiveSpreadsheet().getId();

   if(idPlantilla==idActivo){
     SpreadsheetApp.getUi().alert('Usted ya se encuentra en la Plantilla');
     return;

   }
  //ABRE EL ARCHIVO CREADO
  var url = "https://docs.google.com/spreadsheets/d/"+idPlantilla;
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
   SpreadsheetApp.getUi().showModalDialog(userInterface, "Abrir archivo");
}

function irAlInventario (){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var idss=ss.getId();
  const idPlantilla='10APQHQDVjb2DtbSoScT81r1ZomIreaDhOdDLNC9Fdbk';
  if(idss==idPlantilla) {
    SpreadsheetApp.getUi().alert('Cree el archivo primero. Usted intenta escribir en la plantilla')
    return;
  }
  var hojaInve=ss.setActiveSheet(ss.getSheetByName('inventario y resp')).activate();

}
