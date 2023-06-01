//variable donde guardaremos todos los estilos>>>>
var estilos_sheet = PropertiesService.getDocumentProperties();

function onOpen() {
  //para crear el menu en la ventana de googlesheet
  SpreadsheetApp.getUi().createMenu('Aulaenlanube')
    //items que componen el menu       //nombre de la funion que se ejecuta en el opcion de menu   
    .addItem('Mostrar barra lateral', 'mostrarBarraLateral')
    //cierra el codigo y crea la barra
    .addToUi();

}

function mostrarBarraLateral() {

  var barra = HtmlService.createTemplateFromFile('BarraLateral').evaluate().setTitle('Barra lateral Aulaenlanube');
  //Mostrar la barra lateral, obtener la interface
  SpreadsheetApp.getUi().showSidebar(barra);


}
/* forma 1
function aplicarEstilo1(){
//Guardamos la hoja actual / SpreadsheetApp.getActive:nos obtenga la hoja activa y getActiveSpreSpreadsheetApp:obtenga cual es la hoja activa(pestaña) dentro de la hoja de calculo
var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

//guardamos las celdas que tengamos seleccionadas / 
var celdas = hojaActual.getActiveRange();
      //Atributos de la celda
  
celdas.setFontColor(estilos_sheet.getProperty('color'))
      .setBackground(estilos_sheet.setProperty('colorFondo'))
      .setFontSize(estilos_sheet.setProperty('size'));

}
*/

// forma 2
function aplicarEstilo(estilo) {
  //primero borramos el estilo de las celdas activas
  borrarEstilos();

  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas = hojaActual.getActiveRange();

  celdas.setFontColor(estilos_sheet.getProperty('colorLetra' + estilo))
    .setBackground(estilos_sheet.getProperty('colorFondo' + estilo))
    .setFontSize(estilos_sheet.getProperty('sizeFuente' + estilo))
    .setValue('Estilo ' + estilo);

  //APLICAR BORDES  


}

function comprobarBordes(borde, estilo) {

  switch (borde) {
    case 'sup': return estilos_sheet.getProperty('BordeSupCO' + estilo) != null;
    case 'izq': return estilos_sheet.getProperty('BordeIzqCO' + estilo) != null;
    case 'inf': return estilos_sheet.getProperty('BordeInfCO' + estilo) != null;
    case 'der': return estilos_sheet.getProperty('BordeDerCO' + estilo) != null;
  }
}

function obtenerEnumBorde(tipoBorde) {

  switch (tipoBorde) {
    case 'DOTTED': return SpreadsheetApp.BorderStyle.DOTTED;
    case 'DASHED': return SpreadsheetApp.BorderStyle.DASHED;
    case 'SOLID': return SpreadsheetApp.BorderStyle.SOLID;
    case 'SOLID_MEDIUM': return SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    case 'SOLID_THICK': return SpreadsheetApp.BorderStyle.SOLID_THICK;
    case 'DOUBLE': return SpreadsheetApp.BorderStyle.DOUBLE;
    default: return null;
  }
}

function guardarEstilo(estilo) {


  //borramos previamente los estilos
  eliminarEstilo(estilo);

  //obtenemos la celda activa
  var celda = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();

  //guardamos los bordes
  guardarBordes(celda, estilo);

  //guardamos colores y tamaño
  estilos_sheet.setProperty('colorLetra' + estilo, celda.getFontColor())
    .setProperty('colorFondo' + estilo, celda.getBackground())
    .setProperty('sizeFuente' + estilo, celda.getFontSize() + '');

  return {
    colorFondo: estilos_sheet.getProperty('colorFondo' + estilo),
    colorLetra: estilos_sheet.getProperty('colorLetra' + estilo),
    BordeSupCO: estilos_sheet.getProperty('BordeSupCO' + estilo),
    BordeSupST: estilos_sheet.getProperty('BordeSupST' + estilo),
    BordeInfCO: estilos_sheet.getProperty('BordeInfCO' + estilo),
    BordeInfST: estilos_sheet.getProperty('BordeInfST' + estilo),
    BordeDerCO: estilos_sheet.getProperty('BordeDerCO' + estilo),
    BordeDerST: estilos_sheet.getProperty('BordeDerST' + estilo),
    BordeIzqCO: estilos_sheet.getProperty('BordeIzqCO' + estilo),
    BordeIzqST: estilos_sheet.getProperty('BordeIzqST' + estilo)
  };
}

function guardarBordes(celda, estilo) {
  //obtenemos los bordes
  var bordes = celda.getBorder();

  if (bordes != null) {
    var borde_sup = bordes.getTop();
    var borde_inf = bordes.getBottom();
    var borde_izq = bordes.getLeft();
    var borde_der = bordes.getRight();

    //borde sup, comprabmos que hay color y un estilo
    if (borde_sup.getColor() != null && borde_sup.getBorderStyle() != null) {
      estilos_sheet.setProperty('BordeSupCO' + estilo, borde_sup.getColor().asRgbColor().asHexString());
      estilos_sheet.setProperty('BordeSupST' + estilo, borde_sup.getBorderStyle());
    }
    //borde inf
    if (borde_inf.getColor() != null && borde_inf.getBorderStyle() != null) {
      estilos_sheet.setProperty('BordeInfCO' + estilo, borde_inf.getColor().asRgbColor().asHexString());
      estilos_sheet.setProperty('BordeInfST' + estilo, borde_inf.getBorderStyle());
    }
    //borde der
    if (borde_der.getColor() != null && borde_der.getBorderStyle() != null) {
      estilos_sheet.setProperty('BordeDerCO' + estilo, borde_der.getColor().asRgbColor().asHexString());
      estilos_sheet.setProperty('BordeDerST' + estilo, borde_der.getBorderStyle());
    }
    //borde izq
    if (borde_izq.getColor() != null && borde_izq.getBorderStyle() != null) {
      estilos_sheet.setProperty('BordeIzqCO' + estilo, borde_izq.getColor().asRgbColor().asHexString());
      estilos_sheet.setProperty('BordeIzqST' + estilo, borde_izq.getBorderStyle());
    }
  }
}

function cargarEstilos(estilo) {

  return estilos_sheet.getProperties();

}

function eliminarEstilo(estilo) {
  //colores
  estilos_sheet.deleteProperty('colorLetra' + estilo);
  estilos_sheet.deleteProperty('colorFondo' + estilo);
  estilos_sheet.deleteProperty('sizeFuente' + estilo);

  //bordes
  estilos_sheet.deleteProperty('BordeSupCO' + estilo);
  estilos_sheet.deleteProperty('BordeSupST' + estilo);
  estilos_sheet.deleteProperty('BordeInfCO' + estilo);
  estilos_sheet.deleteProperty('BordeInfST' + estilo);
  estilos_sheet.deleteProperty('BordeIzqCO' + estilo);
  estilos_sheet.deleteProperty('BordeIzqST' + estilo);
  estilos_sheet.deleteProperty('BordeDerCO' + estilo);
  estilos_sheet.deleteProperty('BordeDerST' + estilo);
}

function borrarEstilos() {

  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({ formatOnly: true });
}


function borrarTodo() {
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear();
}


/*
function aplicarEstilo2(){
//Guardamos la hoja actual / SpreadsheetApp.getActive:nos obtenga la hoja activa y getActiveSpreSpreadsheetApp:obtenga cual es la hoja activa(pestaña) dentro de la hoja de calculo
var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var celdas = hojaActual.getActiveRange();
      //Atributos de la celda
celdas.setBackground("green")
      .setFontColor("white")
      .setFontWeight("bold")
      //El valor que va tener dentro la celda
      .setValue("Estilo 2");
}
*/








