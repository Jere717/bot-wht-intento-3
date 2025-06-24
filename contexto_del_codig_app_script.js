var jo = {};
var response = "No";
var token_session = "";
var random_inicio = "";
var random_fin = "";
var excel_qr = SpreadsheetApp.getActiveSpreadsheet();
var sheet_qr = excel_qr.getSheetByName("Configuracion");
var rows_config = sheet_qr.getRange(1, 1, sheet_qr.getLastRow(), sheet_qr.getLastColumn()).getValues();
var api_interna = "https://comfortable-morganica-tattosedm-99d22847.koyeb.app/getqr";
var api_interna_fb = "";
var options = { 'headers': { "Content-Type": "application/json" }, 'method': "POST" };
var appscript = rows_config[0][1];
var gmt = rows_config[4][1];
var recibemensajegrupos = (rows_config[4][3] ? rows_config[4][3] : "");

function onEdit(e) {
  var activeCell = e.range;
  var val = activeCell.getValue();
  var r = activeCell.getRow();
  var c = activeCell.getColumn();
  if (activeCell.getSheet().getName() === "Configuracion" && r == 8 && c == 2) {
    var modelodefault = "";
    if (val == "BOT API GPT" || val == "BOT Asistente GPT") {
      modelodefault = "gpt-4o";
    } else if (val == "BOT Asistente DEEPSEEK") {
      modelodefault = "deepseek-chat";
    } else if (val == "BOT Asistente GEMINI") {
      modelodefault = "gemini-1.5-flash";
    } else if (val == "BOT Asistente QWEN ALIBABA") {
      modelodefault = "qwen-plus";
    } else if (val == "BOT Asistente GROK") {
      modelodefault = "grok-2-vision-latest";
    } else if (val == "BOT Asistente MISTRAL") {
      modelodefault = "pixtral-large-latest";
    } else if (val == "BOT Asistente Llama DeepInfra") {
      modelodefault = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8";
    }
    sheet_qr.getRange(9, 8).setValue(modelodefault);
  }
}
function almacenararchivo(base64String) {
  try {
    var folderurl_completo = (sheet_qr.getRange(13, 4).getValue() ? sheet_qr.getRange(13, 4).getValue() : sheet_qr.getRange(10, 8).getValue()).split("/");
    var folderurl_id = ((folderurl_completo[folderurl_completo.length - 1]).split("?"))[0];
    var carpeta_global = DriveApp.getFolderById(folderurl_id);
    var decodedBytes = Utilities.base64Decode(base64String);
    var blob = Utilities.newBlob(decodedBytes);
    var file = carpeta_global.createFile(blob.setName("" + new Date().getTime() + ".png"));
    jo.status = '0';
    jo.message = ' ok';
    jo.url = "https://drive.usercontent.google.com/download?id=" + file.getId();
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}
function generartextokey(data, textoemail) {
  textoemail = "" + textoemail;
  // Reemplazos generales
  Object.entries(data).forEach(([key, value]) => {
    if (value && key !== "productos") {
      textoemail = textoemail.replaceAll(`@${key}@`, value);
    }
  });
  // Reemplazos de productos
  const productos = data.productos || [];
  for (let i = 0; i < 15; i++) {
    const producto = productos[i];
    const reemplazos = [
      [`@productos${i + 1}@`, producto ? producto.nombre : ""],
      [`@cantidad${i + 1}@`, producto ? producto.cantidad : ""],
      [`@precio${i + 1}@`, producto ? producto.precio : ""],
      [`@subtotal${i + 1}@`, producto ? producto.subtotal : ""]
    ];
    reemplazos.forEach(([placeholder, valor]) => {
      textoemail = textoemail.replaceAll(placeholder, valor);
    });
  }
  return textoemail;
}
function enviaremail(data, textoemail, correos, titulo, adjunto) {
  try {
    textoemail = generartextokey(data, textoemail);
    correos = generartextokey(data, correos);
    var payloademail = {
      to: correos,
      subject: titulo,
      htmlBody: textoemail
    }
    if (adjunto) {
      var blobl = (DriveApp.getFileById(adjunto)).getBlob();
      payloademail.attachments = blobl;
    }
    MailApp.sendEmail(payloademail);
    jo.status = '0';
    jo.message = "OK se envio";
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return jo;
}
function generarreporte(data, plantilla_url, carpeta_url) {
  var urlpdf = "";
  try {
    var plantillaId = plantilla_url.split("/").slice(-2, -1)[0]; // Obtener ID de plantilla
    var plantilla = DriveApp.getFileById(plantillaId);
    var folderId = carpeta_url.split("/").pop().split("?")[0]; // Obtener ID de carpeta
    var carpetaDestino = DriveApp.getFolderById(folderId);
    var copia = plantilla.makeCopy(carpetaDestino);
    var documento = DocumentApp.openById(copia.getId());
    var body = documento.getBody();
    // Reemplazos generales
    Object.keys(data).forEach(function (key) {
      if (data[key] && key !== "productos") {
        body.replaceText("@" + key + "@", data[key]);
      }
    });
    try {
      var fechasistema = Utilities.formatDate(new Date(), gmt, "dd-MM-yyyy");
      var horasistema = Utilities.formatDate(new Date(), gmt, "HH:mm");
      body.replaceText("@fechasistemas@", fechasistema).replaceText("@horasistemas@", horasistema);
    } catch (eww) {
    }
    // Manejo de productos
    var productos = data.productos || [];
    for (let i = 1; i <= 15; i++) {

      const producto = productos[i - 1];
      body.replaceText("@productos" + i + "@", producto ? producto.nombre : "").replaceText("@cantidad" + i + "@", producto ? producto.cantidad : "").replaceText("@precio" + i + "@", producto ? producto.precio : "").replaceText("@subtotal" + i + "@", producto ? producto.subtotal : "");
    }
    documento.saveAndClose();
    var docblob = documento.getAs('application/pdf');
    docblob.setName(Utilities.formatDate(new Date(), "GMT", "yyyyMMdd HHmmss") + ".pdf");
    var file = carpetaDestino.createFile(docblob);
    Utilities.sleep(1000);
    copia.setTrashed(true);
    urlpdf = file.getId();
  } catch (e) {
//    sheet_qr.getRange(20,22).setValue(e);
    Logger.log(e.toString());
  }
  return urlpdf;
}
function actualizarblacklist(operacion) {
  try {
    var sheet_black = excel_qr.getSheetByName("BlackListBOT");
    sheet_black.appendRow([operacion]);
    SpreadsheetApp.flush();
    jo.blacklist = obtenerblacklist();
    jo.status = '0';
    jo.message = "OK se actualizo la memoria BOT";
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}
function aplicarblacklist(palabras, numero) {
  if ((sheet_qr.getRange(7, 6).getValue() + "") == "SI" && sheet_qr.getRange(7, 8).getValue()) {
    try {
      var arrayPalabras = ("" + sheet_qr.getRange(7, 8).getValue()).split(":::");
      for (var ii = 0; ii < arrayPalabras.length; ii++) {
        if (arrayPalabras[ii] && (palabras + "").toUpperCase().includes(arrayPalabras[ii].toUpperCase())) {
          var sheet = excel_qr.getSheetByName("BlackListBOT");
          sheet.appendRow(["" + numero, new Date()]);
        }
      }
    } catch (e) {
    }
  }
}
function aplicarlimpiezapalabras(palabras) {
  if ((sheet_qr.getRange(5, 6).getValue() + "") == "SI" && sheet_qr.getRange(5, 8).getValue()) {
    try {
      var arrayPalabras = ("" + sheet_qr.getRange(5, 8).getValue()).split(":::");
      for (var ii = 0; ii < arrayPalabras.length; ii++) {
        if (arrayPalabras[ii]) {
          palabras = palabras.replaceAll(arrayPalabras[ii], "");
        }
      }
    } catch (e) {
    }
  }
  return palabras;
}

function generarmensajes(salida, numero_enviar) {
  var mensajes = [];
  aplicarblacklist(salida, numero_enviar);
  var url = ((salida + "").match(/<url>.*?<\/url\>/g));
  if (url && url.length > 0) {
    for (let i = 0; i < url.length; i++) {
      salida = (salida + "").replace(url[i], "");
      mensajes.push({ "tipo": "url", "nombrearchivo": "archivo", "mensaje_salida": (url[i] + "").replace("<url>", "").replace("</url>", "") })
    }
  }
  var mapa = ((salida + "").match(/<mapa>.*?<\/mapa\>/g));
  if (mapa && mapa.length > 0) {
    for (let i = 0; i < mapa.length; i++) {
      salida = (salida + "").replace(mapa[i], "");
      mensajes.push({ "tipo": "location", "nombrearchivo": "archivo", "mensaje_salida": (mapa[i] + "").replace("<mapa>", "").replace("</mapa>", "") })
    }
  }
  var registros_ = (salida.match(/https:\/\/[^\s]+\.(png|jpg|jpeg|pdf|mp3|ogg)/gi));
  if (registros_ && registros_.length > 0) {
    for (var ii = 0; ii < registros_.length; ii++) {
      salida = salida.replace(registros_[ii], "[ver 👇]");
      mensajes.push({ "tipo": "url", "nombrearchivo": "archivo", "mensaje_salida": registros_[ii] })
    }
  }
  mensajes.unshift({ "tipo": "mensaje", "mensaje_salida": (aplicarlimpiezapalabras(salida) + "").trim() })
  return mensajes;
}
function onOpen() {
  createMenus();
}
function createMenus() {
  var menu = SpreadsheetApp.getUi().createMenu("Whatsapp");
  menu.addItem('Obtener TOKEN Session', 'qrwhatsapp');
  menu.addItem('Enviar Mensaje Manual', 'enviarwhatsapp');
  menu.addItem('Validar Numeros', 'validarwhatsapp');
  menu.addItem('Enviar Mensaje Manual con validacion', 'enviarwhatsappvalidacion');
  menu.addItem('Obtener Grupos', 'recuperargrupos');
  menu.addItem('Obtener Contactos', 'recuperarcontactos');
  menu.addItem('Agregar Participantes Grupos', 'agregargrupocontactos');
  menu.addItem('Programar Iniciar Programacion', 'programar');
  menu.addItem('Programar Detener Programacion', 'detenerprogramar');
  menu.addItem('Iniciar BOT', 'enviarasistentebot');
  menu.addItem('Actualizar Memoria BOT', 'actualizarmemoriabot');
  menu.addToUi();

  var menu = SpreadsheetApp.getUi().createMenu("Whatsapp API");
  menu.addItem('Iniciar BOT', 'enviarasistentebotwhatsappapi');
  menu.addItem('Actualizar Memoria BOT', 'actualizarmemoriabotwhatsappapi');
  menu.addToUi();
}
function actualizarmemoriabotwhatsappapi() {
  iniciarbotia("actualizar Memoria BOT","APIWHATSAPP");
}
function enviarasistentebotwhatsappapi() {
  iniciarbotia("ASISTENTE","APIWHATSAPP");
}
function enviarasistentebot() {
  iniciarbotia("ASISTENTE","APINOOFICIALWHATSAPP");
}
function actualizarmemoriabot() {
  iniciarbotia("actualizar Memoria BOT","APINOOFICIALWHATSAPP");
}
function iniciarbotia(categoria,tipored) {
  var tipobot = (sheet_qr.getRange(8, 2).getValue() + "");
  var tipomodelo = "ASISTENTE";
  if(tipored.includes("APIWHATSAPP")){
    if ((sheet_qr.getRange(2, 5).getValue() + "") == "" || (sheet_qr.getRange(3, 5).getValue() + "") == "") {
      Browser.msgBox('INGRESE DATOS API WHATSAPP OFICIAL TOKEN o URL API ', Browser.Buttons.OK);
      return;
    }
  }else{
    if ((sheet_qr.getRange(3, 2).getValue() + "") == "") {
      Browser.msgBox('No ha obtenido el TOKEN SESSION WHATSAPP', Browser.Buttons.OK);
      return;
    }
  }
  if (tipobot == "") {
    Browser.msgBox('favor seleccione el TIPO BOT ', Browser.Buttons.OK);
    return;
  }
  if (tipobot == "BOT Asistente GPT") {
    if ((sheet_qr.getRange(8, 4).getValue() + "") == "" || (sheet_qr.getRange(8, 6).getValue() + "") == "") {
      Browser.msgBox('No ha completado el token o ID Asistente', Browser.Buttons.OK);
      return;
    }
  } else if (tipobot == "BOT AutoResponder") {
    tipomodelo = "AUTORESPONDER";
  } else {
    if ((sheet_qr.getRange(8, 4).getValue() + "") == "" || (sheet_qr.getRange(9, 2).getValue() + "") == "") {
      Browser.msgBox('No ha completado el token o el entrenamiento', Browser.Buttons.OK);
      return;
    }
  }
  tipobot = (categoria=="actualizar Memoria BOT"?"Actualizar Memoria ": "Activar ") + tipobot;
  if (("" + sheet_qr.getRange(9, 8).getValue()) != "") {
    tipobot+= " ( Modelo " + sheet_qr.getRange(9, 8).getValue()+")";
  }
  if (confirmarenvioqr(tipobot)) {
    var payload_ = {"sheet_id": excel_qr.getId(), "error_en_grupos": recibemensajegrupos, "op": "registermessage", "token_qr": token_session, "conversacion_bot": (categoria=="actualizar Memoria BOT" ? [{ "inicio": "memoria" }, { "inicio": "memoria" }] : [{ "inicio": "datos" }, { "inicio": "datos" }]), "app_script": appscript, "tipobot": tipomodelo,"tipored":tipored };
    invocarapi(payload_, api_interna);
  }
}
function enviarapigptfb() {
  iniciarbotiafb("CHATGPTAPI_ASISTENTE", "activar bot Facebook API Chatgpt");
}
function iniciarbotiafb(modelo, texto) {
  if ((sheet_qr.getRange(2, 5).getValue() + "") == "") {
    Browser.msgBox('No ha obtenido el TOKEN SESSION FACEBOOK', Browser.Buttons.OK);
    return;
  }
  if (appscript == "") {
    Browser.msgBox('Debe ingresar la URL APPSCRIPT', Browser.Buttons.OK);
    return;
  }
  if (texto == "activar bot asistente GPT") {
    if ((sheet_qr.getRange(8, 4).getValue() + "") == "" || (sheet_qr.getRange(8, 6).getValue() + "") == "") {
      Browser.msgBox('No ha completado el token o ID Asistente', Browser.Buttons.OK);
      return;
    }
  } else if (texto == "activar bot") {
  } else {
    if ((sheet_qr.getRange(8, 4).getValue() + "") == "" || (sheet_qr.getRange(9, 2).getValue() + "") == "") {
      Browser.msgBox('No ha completado el token o el entrenamiento', Browser.Buttons.OK);
      return;
    }
  }
  if (confirmarenvioqr(texto)) {
    sheet_qr.getRange(8, 2).setValue(modelo);
    SpreadsheetApp.flush();
    invocarapi({ "op": "cargarcache", "app_script": appscript }, api_interna_fb);
  }
}
function gruposcontacto(resultado) {
  try {
    if (resultado.data && resultado.data && resultado.data.op.length > 0) {
      var hoja = excel_qr.getSheetByName("GruposParticipantes");
      const datos = hoja.getDataRange().getValues();
      const indexMap = new Map();
      datos.slice(1).forEach((fila, index) => {
        indexMap.set(index.toString(), index + 2); // +2 to account for header in sheet
      });
      resultado.data.op.forEach(grupo => {
        grupo.registros.forEach(contacto => {
          const filaIndex = indexMap.get(contacto.index);
          if (filaIndex) {
            const resultado = contacto.resultado || "OK SE CREO"; // Define logic to determine result
            hoja.getRange(filaIndex, 3).setValue(resultado);
          }
        });
      });
    }
    jo.status = '0';
    jo.message = ' Se grabo el registro';
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}
function contactos(resultado) {
  try {
    if (resultado.mensajes && resultado.mensajes.length > 0) {
      var sheet = excel_qr.getSheetByName("Contactos");
      for (let i = 0; i < resultado.mensajes.length; i++) {
        sheet.appendRow([resultado.mensajes[i].id_contacto, "'" + resultado.mensajes[i].nombre_contacto]);
      }
    }
    jo.status = '0';
    jo.message = ' Se grabo el registro';
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}
function grupos(resultado) {
  try {
    if (resultado.mensajes && resultado.mensajes.length > 0) {
      var sheet = excel_qr.getSheetByName("Grupos");
      for (let i = 0; i < resultado.mensajes.length; i++) {
        sheet.appendRow([resultado.mensajes[i].id_grupo, resultado.mensajes[i].nombre_grupo]);
      }
    }
    jo.status = '0';
    jo.message = ' Se grabo el registro';
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}
function agregargrupocontactos() {
  if (confirmarenvioqr("agregar grupos y participantes")) {
    const hoja = excel_qr.getSheetByName('GruposParticipantes');
    const datos = hoja.getDataRange().getValues();
    const gruposMap = {};
    datos.slice(1).forEach((fila, index) => { // Empezar después de la fila de encabezado
      const [numero, nombreGrupo] = fila;
      if (!gruposMap[nombreGrupo]) {
        gruposMap[nombreGrupo] = {
          nombregrupo: nombreGrupo,
          registros: []
        };
      }
      gruposMap[nombreGrupo].registros.push({
        contacto: numero.toString(),
        index: index.toString() // Índice relativo al conjunto de datos original
      });
    });
    const gruposArray = Object.values(gruposMap);
    invocarapi({ "app_script": appscript, "op": "registermessage", "token_qr": token_session, "listener": "true", "grupocontactos": gruposArray }, api_interna);
  }
}
function recuperarcontactos() {
  if (confirmarenvioqr("recuperar contactos")) {
    invocarapi({ "app_script": appscript, "op": "registermessage", "token_qr": token_session, "listener": "true", "contactos": [{ "mensaje": "contacto" }] }, api_interna);
  }
}
function recuperargrupos() {
  if (confirmarenvioqr("recuperar grupos")) {
    invocarapi({ "app_script": appscript, "op": "registermessage", "token_qr": token_session, "listener": "true", "grupos": [{ "mensaje": "grupo" }] }, api_interna);
  }
}
function nuevoregistro(resultado) {
  try {
    var sheet = excel_qr.getSheetByName("Solicitudes");
    var fila = [new Date(), resultado.idpersona, "LEAD", resultado.mensaje];
    sheet.appendRow(fila);
    jo.status = '0';
    jo.message = "OK";
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}
function crearContacto(nombre, numero) {
  try {
    var contactResource = { "names": [{ "familyName": nombre, }], "phoneNumbers": [{ 'value': numero }] }
    var contactResourceName = People.People.createContact(contactResource)["resourceName"];
    var groupName = "ContactosBOT";
    var groups = People.ContactGroups.list()["contactGroups"];
    var group = groups.find(group => group["name"] === groupName);
    if (!group) {
      var groupResource = { contactGroup: { name: groupName } }
      group = People.ContactGroups.create(groupResource);
    }
    var groupResourceName = group["resourceName"];
    var membersResource = { "resourceNamesToAdd": [contactResourceName] }
    People.ContactGroups.Members.modify(membersResource, groupResourceName);
  } catch (err) {
    console.log('Failed to get the connection with an error %s', err.message);
  }
}

function registrarsolicitudes(resultado) {
  try {
    var aplicarenviarstock = false;
    var aplicarenviarblacklist = false;
    var sheet = excel_qr.getSheetByName("Solicitudes");
    if (rows_config[19][1] == "SI") {
      var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      var index = rows.findLastIndex((item) => item[1] + "" === "" + resultado.idpersona && item[2] == "LEAD");
      if (index > -1) {
        sheet.deleteRow(index + 2);
        SpreadsheetApp.flush();
        sheet = excel_qr.getSheetByName("Solicitudes");
      }
    }
    var fila = [new Date(), resultado.idpersona, "SOLICITADO"];
    // se limpia campos sin uso
    Object.keys(resultado).forEach(function (key) {
      if (resultado[key] && !(key == "op" || key == "mensaje_salida" || key == "fechasistema" || key == "horasistema" || key == "idpersona" || key == "estado" || key == "identificador")) {
        fila.push(resultado[key]);
      }
    });
    sheet.appendRow(fila);
    jo.status = '0';
    var salida_respuesta = "OK se registro ";
    var aplicarobtenerpostsolicitud = false;   
    var aplicarobtenerpostreporte = true;   
    
    var existepost = obtenerpostsolicitud();
    if(existepost){
      for (var ii = 0; ii < existepost.length; ii++) {
        if((existepost[ii][0]+"") && resultado.mensaje_salida.toUpperCase().includes((existepost[ii][0]+"").toUpperCase())){
          aplicarobtenerpostsolicitud=true;
          if(existepost[ii][5]!="SI"){
            aplicarobtenerpostreporte=false;
          }
          break;
        }  
      }
    }
//    sheet_qr.getRange(20,21).setValue("zzz"+aplicarobtenerpostreporte);
    var factura_id = "";
    if(aplicarobtenerpostreporte){
      if ((rows_config[9][9] && rows_config[9][7]) || (rows_config[12][1] && rows_config[12][3])) {
        factura_id = generarreporte(resultado, (rows_config[12][1] ? rows_config[12][1] : rows_config[9][9]), (rows_config[12][3] ? rows_config[12][3] : rows_config[9][7]));
        if (factura_id) {
          var factura = "https://drive.usercontent.google.com/download?id=" + factura_id;
          salida_respuesta += " y su PDF es el siguiente : " + factura + ".espdf";
          jo.archivo = (factura ? factura + ".espdf" : "");
        } else {
          salida_respuesta += " y su PDF no se pudo generar ";
        }
      }  
    }
    if(aplicarobtenerpostsolicitud){
      for (var ii = 0; ii < existepost.length; ii++) {
        var existepalabra =false;
        if(resultado.mensaje_salida.toUpperCase().includes((existepost[ii][0]+"").toUpperCase())){
          existepalabra=true;
        }  
        if(existepalabra){
          if(existepost[ii][1]=="CORREO" &&  existepost[ii][2]  &&  existepost[ii][3]) {
            var enviar = enviaremail(resultado, existepost[ii][3], existepost[ii][2], "CORREO SOLICITUD", factura_id);
            if (enviar.status == "0") {
              salida_respuesta += " y se envio el email ";
            }
          }
          if(existepost[ii][1]=="WHATSAPP" &&  existepost[ii][2] &&  existepost[ii][3]) {
              var enviar = generartextokey(resultado, existepost[ii][3]);
              if (enviar) {
                jo.numeroreenviar = existepost[ii][2];
                jo.mensajereenviar = enviar;
              }
          }
          if (existepost[ii][4] == "SI") {
            actualizarblacklist(resultado.idpersona);
            aplicarenviarblacklist = true;
          }
        }
      }
    }
    
    if(!aplicarobtenerpostsolicitud){
      if (rows_config[13][1] && rows_config[13][3] && rows_config[13][5]) {
        var enviar = enviaremail(resultado, rows_config[13][1], rows_config[13][5], rows_config[13][3], factura_id);
        if (enviar.status == "0") {
          salida_respuesta += " y se envio el email ";
        }
      }
      if (rows_config[14][1] && rows_config[14][3]) {
        var enviar = generartextokey(resultado, rows_config[14][1]);
        if (enviar) {
          jo.numeroreenviar = rows_config[14][3];
          jo.mensajereenviar = enviar;
        }
      }
      if (rows_config[15][1] == "Apagar BOT al cliente") {
        actualizarblacklist(resultado.idpersona);
        aplicarenviarblacklist = true;
      }
    } 
    // EN CASO DESEA ALMACENAR CONTACTOS
    if (rows_config[15][3] == "SI") {
      crearContacto(resultado.idpersona, (resultado.nombre ? resultado.nombre : resultado.idpersona));
      salida_respuesta += " y se creo contacto BD ";
    }    
    jo.message = salida_respuesta;
    // EN CASO DESEA DESCONTAR STOCK
    if (rows_config[15][5] == "SI") {
      try {
        if (resultado.productos && resultado.productos.length > 0) {
          var sheet_inventario = excel_qr.getSheetByName('Inventario');
          var rows = sheet_inventario.getRange(2, 1, sheet_inventario.getLastRow() - 1, sheet_inventario.getLastColumn()).getValues();
          for (var ii = 0; ii < resultado.productos.length; ii++) {
            var evento = rows.findIndex((item) => ((item[0] + "").toUpperCase() == (resultado.productos[ii].producto + "").toUpperCase() || (item[0] + "").toUpperCase() == (resultado.productos[ii].nombre + "").toUpperCase()));
            if (evento > -1) {
              var cantidad = parseInt(resultado.productos[ii].cantidad);
              sheet_inventario.getRange(evento + 2, 2).setValue((sheet_inventario.getRange(evento + 2, 2).getValue() - cantidad));
              aplicarenviarstock = true;
            }
          }
        }
      } catch (ees) {
      }
    }
    //aplica registrar el caledario
    if(resultado.op=="registrarcita"){
        jo.message = "Se registro la cita";
        var events = CalendarApp.getDefaultCalendar();
        var fechaInicio = Utilities.parseDate(resultado.fecha_hora_cita, rows_config[4][1], "dd/MM/yyyy HH:mm");
        var fechaFin = new Date(fechaInicio.getTime() + 30 * 60000);
        events.createEvent("CITA:" + resultado.nombre, fechaInicio, fechaFin, {
            location: 'remote',
            description: 'Comentario:' + resultado.resumen_cita,
            guests: resultado.correo
        });
    }
    //se envian tags de configuracion
    if (aplicarenviarstock) {
      jo.configuracion = obtenerconfiguracion();
    }
    if (aplicarenviarblacklist) {
      jo.blacklist = obtenerblacklist();
    }
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}

function validanumero(resultado) {
  try {
    if (resultado.validar_numero && resultado.validar_numero.length > 0) {
      var sheet = excel_qr.getSheetByName("Validacion");
      for (let i = 0; i < resultado.validar_numero.length; i++) {
        sheet.getRange((2 + (parseInt(resultado.validar_numero[i].posicion))), 2).setValue(resultado.validar_numero[i].estado);
      }
    }
    jo.status = '0';
    jo.message = ' Se grabo el registrp';
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}
function validarwhatsapp() {
  if (confirmarenvioqr("validar numeros")) {
    if (excel_qr.getSheetByName("Validacion")) {
      var dataArray = [];
      var sheet = excel_qr.getSheetByName("Validacion");
      var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      for (var i = 0, l = rows.length; i < l; i++) {
        var numero = rows[i][0];
        var arrayNumero = ("" + numero).split(";");
        for (var ii = 0; ii < arrayNumero.length; ii++) {
          var record = {};
          record['numero'] = arrayNumero[ii];
          record['posicion'] = "" + i;
          dataArray.push(record);
        }
      }
      invocarapi({ "app_script": appscript, "op": "registermessage", "token_qr": token_session, "listener": "true", "validar_numero": dataArray }, api_interna);
    }
  }
}
function qrwhatsapp() {
  var response = "No"
  try {
    var response = Browser.msgBox('Seguro que quiere generar QR ahora ?', Browser.Buttons.YES_NO);
  } catch (e) {
    Browser.msgBox('La acción no se ha realizado', Browser.Buttons.OK);
  }
  if (response == "yes") {
    enviar();
  }
}
function enviar() {
  sheet_qr.getRange(2, 2).setValue("NO CONECTADO");
  SpreadsheetApp.flush();
  invocarapi({ "op": "iniciarqr", "app_script": appscript, "sheet_id": excel_qr.getId(), "fechahora": Utilities.formatDate(new Date(), "GMT-5", "yyMMddHHmmss") }, api_interna);
}
function generar(qr) {
  try {
    sheet_qr.getRange(2, 2).setValue(encodeURIComponent(qr.qr));
    if (qr.numero && qr.qr == "CONECTADO") {
      sheet_qr.getRange(3, 2).setValue(qr.session);
      sheet_qr.getRange(3, 3).setValue(qr.numero);
    }
    jo.status = '0';
    jo.message = ' Se grabo el registro';
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}

function enviarwhatsappvalidacion() {
  if (confirmarenvioqr("enviar los mensajes con validacion")) {
    aplicarenviarwhatsapp(true);
  }
}
function enviarwhatsapp() {
  if (confirmarenvioqr("enviar los mensajes")) {
    aplicarenviarwhatsapp(false);
  }
}
function aplicarenviarwhatsapp(aplicavalidar) {
  if (excel_qr.getSheetByName("MensajeManual")) {
    var dataArray = [];
    var sheet = excel_qr.getSheetByName("MensajeManual");
    try {
      random_inicio = (sheet_qr.getRange(7, 2).getValue() && parseInt(sheet_qr.getRange(7, 2).getValue()) > 0 ? sheet_qr.getRange(7, 2).getValue() : "");
      random_fin = (sheet_qr.getRange(7, 4).getValue() && parseInt(sheet_qr.getRange(7, 4).getValue()) > 0 ? sheet_qr.getRange(7, 4).getValue() : "");
    } catch (e) {
    }
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    for (var i = 0, l = rows.length; i < l; i++) {
      try {
        if (rows[i][3] && (rows[i][3] + "").toUpperCase() == "ENVIADO") {
          continue;
        }
      } catch (e) {
      }
      var numero = rows[i][0];
      var mensaje = rows[i][1];
      var url = rows[i][2];
      var arrayNumero = ("" + numero).split(";");
      for (var ii = 0; ii < arrayNumero.length; ii++) {
        var intervalo_mensaje = "";
        if (random_inicio && random_fin) {
          intervalo_mensaje = generarramdon(parseInt(random_inicio), parseInt(random_fin));
        }
        if (mensaje) {
          dataArray.push({ "aplicavalidacion": (aplicavalidar == true ? "SI" : ""), "numero": arrayNumero[ii], "mensaje": mensaje, "intervalo_mensaje": intervalo_mensaje, "posicion": "" + i });
        }
        if (url) {
          dataArray.push({ "aplicavalidacion": (aplicavalidar == true ? "SI" : ""), "numero": arrayNumero[ii], "intervalo_mensaje": intervalo_mensaje, "url": url, "posicion": "" + i });
        }
      }
    }
    invocarapi({ "op": "registermessage", "token_qr": token_session, "listener": true, "mensajes": dataArray, "app_script": appscript }, api_interna);
  }
}
function invocarapi(payload, api_enviar) {
  try {
    options.payload = JSON.stringify(payload);
    var response = UrlFetchApp.fetch(api_enviar, options);
    var json = JSON.parse(response.getContentText());
    if (json.status == "0") {
      Browser.msgBox(json.message ? json.message : 'Se envio la solicitud a la API', Browser.Buttons.OK);
    } else {
      Browser.msgBox('Error al iniciar : ' + json.message, Browser.Buttons.OK);
    }
  } catch (e) {
    Browser.msgBox('Se notificaron los mensajes asincrono ', Browser.Buttons.OK);
  }
}
function confirmarenvioqr(mensajes) {
  try {
    token_session = sheet_qr.getRange(3, 2).getValue();
    var response = Browser.msgBox('Seguro que quiere ' + mensajes + ' del token de session sera : ' + token_session + ' ?', Browser.Buttons.YES_NO);
  } catch (e) {
    Browser.msgBox('La acción no se ha realizado', Browser.Buttons.OK);
  }
  if (response == "yes") {
    return true;
  }
  return false;
}
function soloenviar() {
  if (confirmarenvioqr("mensajes ahora hoja programacion")) {
    var respuesta = solonotificarwhatsapp();
    Browser.msgBox(respuesta.message, Browser.Buttons.OK);
  }
}

function resultado(resultado) {
  try {
    var sheet = excel_qr.getSheetByName("MensajeManual");
    sheet_qr.getRange(6, 2).setValue("Se procesaron :" + new Date() + " :: TOTAL " + resultado.mensajes.length);
    for (let i = 0; i < resultado.mensajes.length; i++) {
      sheet.getRange((2 + (parseInt(resultado.mensajes[i].posicion))), 4).setValue(resultado.mensajes[i].estado);
    }
    jo.status = '0';
    jo.message = ' Se grabo el registrp';
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}

function solonotificarwhatsapp() {
  var excel = SpreadsheetApp.getActiveSpreadsheet();
  if (excel.getSheetByName("Programados")) {
    var dataArray = [];
    var sheet = excel.getSheetByName("Programados");
    var sheet_qr = excel.getSheetByName("Configuracion");
    var gmt = sheet_qr.getRange(5, 2).getValue();
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var fechasistema = Utilities.formatDate(new Date(), gmt, "yyyyMMdd");
    var horasistema = parseInt(Utilities.formatDate(new Date(), gmt, "HH"));
    var token_session = sheet_qr.getRange(3, 2).getValue();;
    var appscript = sheet_qr.getRange(1, 2).getValue();
    var random_inicio = "";
    var random_fin = "";
    try {
      random_inicio = (sheet_qr.getRange(7, 2).getValue() && parseInt(sheet_qr.getRange(7, 2).getValue()) > 0 ? sheet_qr.getRange(7, 2).getValue() : "");
      random_fin = (sheet_qr.getRange(7, 4).getValue() && parseInt(sheet_qr.getRange(7, 4).getValue()) > 0 ? sheet_qr.getRange(7, 4).getValue() : "");
    } catch (e) {
    }
    for (var i = 0; i < rows.length; i++) {
      // ELIMINAR REGISTROS NO VALIDOS
      Logger.log("rows[i][3]" + rows[i][3] + "ACTIVO" + rows[i][0] + "--" + (rows[i][1] + "" + rows[i][2]));
      if (!(rows[i][3] == "ACTIVO" && rows[i][0] && (rows[i][1] || rows[i][2]))) {
        continue;
      }
      Logger.log("FECHA APROGRAMACION" + fechasistema + "--" + horasistema + "sss" + rows[i][4] + "xxxx" + rows[i][5]);
      // VALIDAR FECHA PROGRAMACION ENVIOS
      if ((rows[i][4]) && !((rows[i][4] + "") == ("" + fechasistema))) {
        continue;
      }
      Logger.log("FECHA SISTEMA");
      if ((rows[i][5]) && !(("" + rows[i][5]) == ("" + horasistema))) {
        continue;
      }
      Logger.log("HORA SISTEMA");
      var numero = rows[i][0];
      var mensaje = rows[i][1];
      var url = rows[i][2];
      var arrayNumero = ("" + numero).split(";");
      for (var ii = 0; ii < arrayNumero.length; ii++) {
        var intervalo_mensaje = "";
        if (random_inicio && random_fin) {
          intervalo_mensaje = generarramdon(parseInt(random_inicio), parseInt(random_fin));
        }
        if (mensaje) {
          dataArray.push({ "numero": arrayNumero[ii], "mensaje": mensaje, "intervalo_mensaje": intervalo_mensaje, "posicion": "" + i });
        }
        if (url) {
          dataArray.push({ "numero": arrayNumero[ii], "intervalo_mensaje": intervalo_mensaje, "url": url, "posicion": "" + i });
        }
      }
    }
    if (dataArray.length == 0) {
      sheet_qr.getRange(6, 2).setValue("No existen registros a enviar " + new Date());
      jo.status = '1';
      jo.message = "No existen registros a enviar " + new Date();
      return jo;
    }
    var payload = { "config": { "operacion": "resultadoprogramar" }, "op": "registermessage", "app_script": appscript, "token_qr": token_session, "listener": "true", "mensajes": dataArray };
    options.payload = JSON.stringify(payload);
    try {
      var response = UrlFetchApp.fetch(api_interna, options);
      var json = JSON.parse(response.getContentText());
      if (json.status == "0") {
        sheet_qr.getRange(6, 2).setValue("Se notificaron cantidad de registros " + dataArray.length + " a las " + new Date());
        jo.status = '0';
        jo.message = "Se notificaron cantidad de registros " + dataArray.length + " a las " + new Date();
      } else {
        sheet_qr.getRange(6, 2).setValue("Error de la API  " + json.message + " registros notificados " + dataArray.length + " a las " + new Date());
        jo.status = '1';
        jo.message = "Error de la API  " + json.message + " registros notificados " + dataArray.length + " a las " + new Date();
      }
    } catch (e) {
      jo.status = '-1';
      jo.message = "Error de la API  " + e.toString() + " registros notificados " + dataArray.length + " a las " + new Date();
      sheet_qr.getRange(6, 2).setValue("Error de la API  " + e.toString() + " registros notificados " + dataArray.length + " a las " + new Date());
    }
  }
  return jo;
}

function detenerprogramar() {
  if (confirmarenvioqr("detener tareas")) {
    //ELIMINAMOS LA PROGRAMACION
    eliminareventos();
    Browser.msgBox('La acción ha sido realizada', Browser.Buttons.OK);
  }
}
function eliminareventos() {
  //ELIMINAMOS LA PROGRAMACION
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getTriggerSource() == ScriptApp.TriggerSource.CLOCK && triggers[i].getHandlerFunction() == "solonotificarwhatsapp") {
      ScriptApp.deleteTrigger(triggers[i]);
    };
  };
}
function programar() {
  console.log("Funcion inicio programar : la fecha y hora: " + new Date());
  var minutos = 60; //sheet_qr.getRange(4, 2).getValue();
  if (confirmarenvioqr("programar tarea cada  " + minutos + " minutos")) {
    eliminareventos();
    ScriptApp.newTrigger("solonotificarwhatsapp").timeBased().everyHours(1).create();
    Browser.msgBox('La acción ha sido realizada', Browser.Buttons.OK);
  }
}

function resultadoprogramar(resultado) {
  try {
    if (resultado.mensajes && resultado.mensajes.length > 0) {
      var sheet = excel_qr.getSheetByName("Programados");
      var sheet_trasacciones = excel_qr.getSheetByName("Transacciones");
      var gmt = sheet_qr.getRange(5, 2).getValue();
      for (let i = 0; i < resultado.mensajes.length; i++) {
        try {
          var posicion_mov = parseInt(resultado.mensajes[i].posicion);
          var numero = sheet.getRange((2 + posicion_mov), 1).getValue();
          var mensaje = sheet.getRange((2 + posicion_mov), 2).getValue();
          var url = sheet.getRange((2 + posicion_mov), 3).getValue();
          var cantidad = ("" + sheet.getRange((2 + posicion_mov), 7).getValue());
          sheet_trasacciones.appendRow([numero, mensaje, url, resultado.mensajes[i].estado, new Date()]);
          if (resultado.mensajes[i].estado == "Enviado") {
            if (cantidad) {
              var fecha = "" + sheet.getRange((2 + posicion_mov), 5).getValue();
              var fechadesde = Utilities.parseDate(fecha, gmt, "yyyyMMdd");
              var otraFecha = fechadesde.setDate(fechadesde.getDate() + parseInt(cantidad));
              sheet.getRange((2 + posicion_mov), 5).setValue(Utilities.formatDate(new Date(otraFecha), gmt, "yyyyMMdd"));
            } else {
              sheet.getRange((2 + posicion_mov), 4).setValue("INACTIVO");
            }
          }
        } catch (errror) {
        }
      }
    }
    jo.status = '0';
    jo.message = ' Se grabo el registrp';
  } catch (e) {
    jo.status = '-1';
    jo.message = e.toString();
  }
  return JSON.stringify(jo);
}

function obtenerevento(numero_enviar, mensaje_buscar) {
  var conversacion_ingresado = {};
  try {
    var cache = CacheService.getScriptCache();
    var conversaciones = [];
    if (cache.get("conversaciones" + numero_enviar) != null) {
      conversaciones = JSON.parse(cache.get("conversaciones" + numero_enviar));
    }
    var sheet = excel_qr.getSheetByName("Conversacion");
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var evento = rows.find((item) => ((item[1] + "").toUpperCase()).split(";").includes(mensaje_buscar.toUpperCase()));
    var evento_error = rows.find((item) => item[0] === "Error");
    var evento_start = rows.find((item) => item[0] === "Start");
    var id_fila = -1;
    //OBTENER LA POSICION
    var evento_start_inicio = conversaciones.find((item) => item.evento === "Start");
    if (evento_start_inicio) {
      id_fila = evento_start_inicio.id_fila;
    }
    if (evento === undefined) {
      var filtroconversacion = conversaciones.filter((item) => item.numero == numero_enviar);
      evento = evento_error;
      if (filtroconversacion.length > 0) {
        var evento_retorno = filtroconversacion[filtroconversacion.length - 1].retornar;
        var evento_retornar = rows.find((item) => evento_retorno != "" && item[0] === evento_retorno);
        if (evento_retornar !== undefined) {
          evento = evento_retornar;
        } else {
          // CASO SEAN ARREGLOS ::
          var evento_retorna_arreglo_menu = evento_retorno.split(";");
          if (evento_retorna_arreglo_menu.length > 1) {
            for (let jl = 0; jl < evento_retorna_arreglo_menu.length; jl++) {
              var evento_msj = evento_retorna_arreglo_menu[jl].split(",");
              if ((evento_msj[1] + "").toUpperCase() === ("" + mensaje_buscar).toUpperCase()) {
                var evento_retornar_menu = rows.find((item) => item[0] === evento_msj[0]);
                if (evento_retornar_menu !== undefined) {
                  evento = evento_retornar_menu;
                  break;
                }
              }
            }
          }
        }
      }
      if (evento[0] == "Error" && evento_start && evento_start[1].includes("%%%")) {
        // CONVERSACION NUMERO CELULAR
        evento = evento_start;
        if (filtroconversacion.length == 0 || (filtroconversacion.length == 1 && filtroconversacion[0].evento.includes("Close"))) {
          evento = evento_start;
        }
      }
    }
    conversacion_ingresado = { conversaciones: conversaciones, status: "0", numero: numero_enviar, mensaje_entrada: mensaje_buscar, evento: evento[0], retornar: evento[3], mensaje_salida: evento[2], id_fila: id_fila };
  } catch (e) {
    conversacion_ingresado = { status: "-1", message: e.toString() };
  }
  return conversacion_ingresado;
}
function enviar_bot(numero_enviar, mensaje_buscar, base64, tipo) {
  var conversacion_ingresado = {};
  var flag_close = false;
  try {
    conversacion_ingresado = obtenerevento(numero_enviar, mensaje_buscar);
    var conversaciones = [];
    if (conversacion_ingresado.status != "0") {
      return JSON.stringify({ status: "-1", message: "No existe el evento" })
    } else {
      conversaciones = conversacion_ingresado.conversaciones;
      conversacion_ingresado.conversaciones = [];
    }
    if (((conversacion_ingresado.evento) + "").toUpperCase().includes("fin")) {
      flag_close = true;
    }
    conversacion_ingresado.mensajes = generarmensajes(conversacion_ingresado.mensaje_salida, numero_enviar);
    var conversacion_ingresado_temp = JSON.parse(JSON.stringify(conversacion_ingresado));
    conversaciones.push(conversacion_ingresado_temp);
    if ((conversacion_ingresado.evento).includes("Start") || flag_close) {
      conversaciones = [];
      if ((conversacion_ingresado.evento).includes("Start")) {
        conversaciones.push(conversacion_ingresado_temp);
      }
    }
    var cache = CacheService.getScriptCache();
    cache.put('conversaciones' + numero_enviar, JSON.stringify(conversaciones), 600);
  } catch (e) {
    conversacion_ingresado.status = "0";
    conversacion_ingresado.tipo = "mensaje";
    conversacion_ingresado.mensaje_salida = "" + e.toString();
  }
  return JSON.stringify(conversacion_ingresado);
}
function obtenerwhitelist() {
  try {
    var sheet_blacklist = excel_qr.getSheetByName("WhiteListBOT");
    var itemsblacklist = sheet_blacklist.getRange(2, 1, sheet_blacklist.getLastRow(), sheet_blacklist.getLastColumn()).getValues();
    return [...new Set(itemsblacklist.map(item => item[0]))].toString();
  } catch (e) {
  }
  return "";
}
function obtenerblacklist() {
  try {
    var sheet_blacklist = excel_qr.getSheetByName("BlackListBOT");
    var itemsblacklist = sheet_blacklist.getRange(2, 1, sheet_blacklist.getLastRow(), sheet_blacklist.getLastColumn()).getValues();
    return [...new Set(itemsblacklist.map(item => item[0]))].toString();
  } catch (e) {
  }
  return "";
}
function obtenerpostsolicitud() {
  try {
    var sheet_blacklist = excel_qr.getSheetByName("PostSolicitud");
    var itemsblacklist = sheet_blacklist.getRange(2, 1, sheet_blacklist.getLastRow(), sheet_blacklist.getLastColumn()).getValues();
    return itemsblacklist;
  } catch (e) {
  }
  return "";
}

function obtenerdatosinventarios() {
  try {
    var sheet = excel_qr.getSheetByName('Inventario');
    var lastColumn = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headersRange = sheet.getRange(1, 1, 1, lastColumn);
    var headers = headersRange.getValues()[0].map(header => header.replace(/ /g, '_'));
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    var data = dataRange.getValues();
    var inventory = [];
    data.forEach(function (row) {
      var item = {};
      headers.forEach(function (header, index) {
        item[header] = row[index];
      });
      inventory.push(item);
    });
    var jsonResult = JSON.stringify(inventory, null, 2);
    return "\nEstos son los productos que vendo, presentados en formato JSON:\n" + jsonResult + "\nPor favor, procesa esta información según sea necesario.\n";
  } catch (e) {
    return "\nLo siento no hay productos\n";
  }
}
function obtenerconfiguracion() {
  try {
    if (rows_config[8][1] && (rows_config[8][1] + "").includes("@productos@")) {
      var datoss = obtenerdatosinventarios();
      rows_config[8][1] = (rows_config[8][1] + "").replace("@productos@", datoss);
      if (rows_config[8][5] && (rows_config[8][5] + "").includes("@productos@")) {
        rows_config[8][5] = (rows_config[8][5] + "").replace("@productos@", datoss);
      }
    }
    return rows_config;
  } catch (e) {
  }
  return "";
}

function doPost(e) {
  var operacion = JSON.parse(e.postData.contents)
  var respuesta = "";
  if (operacion.op == "qr") {
    respuesta = generar(operacion);
  } else if (operacion.op == "resultadoprogramar") {
    respuesta = resultadoprogramar(operacion);
  } else if (operacion.op == "resultado") {
    respuesta = resultado(operacion);
  } else if (operacion.op == "grupos") {
    respuesta = grupos(operacion);
  } else if (operacion.op == "contactos") {
    respuesta = contactos(operacion);
  } else if (operacion.op == "subirimagen") {
    respuesta = almacenararchivo(operacion.data);
  } else if (operacion.op == "obtenersheet") {
    jo.configuracion = obtenerconfiguracion();
    if (jo.configuracion) {
      jo.status = '0';
      jo.message = ' Se grabo el registro';
      jo.blacklist = obtenerblacklist();
      jo.whitelist = obtenerwhitelist();
    } else {
      jo.status = '-1';
      jo.message = e.toString();
    }
    respuesta = JSON.stringify(jo);
  } else if (operacion.op == "resultado_gruposcontacto") {
    respuesta = gruposcontacto(operacion);
  } else if (operacion.op == "save_validanumero") {
    respuesta = validanumero(operacion);
  } else if (operacion.op == "registrarsolicitudes") {
    //sheet_qr.getRange(20,20).setValue(JSON.stringify(operacion))
    respuesta = registrarsolicitudes(operacion);
  } else if (operacion.op == "nuevoregistro") {
    respuesta = nuevoregistro(operacion);
  } else if (operacion.op == "registrarcita") {
    respuesta = registrarsolicitudes(operacion);
  } else if (operacion.op == "actualizarblacklist") {
    respuesta = actualizarblacklist(operacion.mensaje, api_interna);
  } else if (operacion.op == "fechasistema") {
    respuesta = fechasistema(operacion);
  } else if (operacion.op == "find_conversacion") {
    var numero_enviar = (operacion.numero).substring(0, (operacion.numero).lastIndexOf("@"));
    var mensaje_buscar = operacion.mensaje;
    var aplicarblacklist = false;
    if ((rows_config[6][5] + "") == "SI") {
      try {
        var sheet = excel_qr.getSheetByName("BlackListBOT");
        var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        var evento = rows.find((item) => ((item[0] + "") == numero_enviar + ""));
        if (evento) {
          aplicarblacklist = true;
        }
      } catch (e) {
      }
    }
    if (aplicarblacklist == false) {
      var base64 = "";
      var tipo = "";
      if (operacion.mensaje == "documento_send" && (operacion.documento.mimetype + "").includes("audio/ogg")) {
        base64 = operacion.documento.data;
        tipo = "audio";
        mensaje_buscar = "";
      } else if (operacion.mensaje == "documento_send" && ((operacion.documento.mimetype + "").includes("image/jpeg") || (operacion.documento.mimetype + "").includes("image/jpg") || (operacion.documento.mimetype + "").includes("image/png"))) {
        base64 = operacion.documento.data;
        tipo = "image";
        mensaje_buscar = "";
      }
      if ((rows_config[7][1] + "") == "BOT AutoResponder") {
        respuesta = enviar_bot(numero_enviar, mensaje_buscar, base64, "");
      }
    } else {
      respuesta = JSON.stringify({ "status": "-1", "message": "NO APLICA" })
    }
  }
  return ContentService.createTextOutput(respuesta).setMimeType(ContentService.MimeType.JSON);
}

function obtenerQRyGuardarConReintentos(reintentos) {
  var url = "https://comfortable-morganica-tattosedm-99d22847.koyeb.app/getqr";
  var options = {
    'method': 'POST',
    'headers': { 'Content-Type': 'application/json' },
    'muteHttpExceptions': true,
    'payload': JSON.stringify({ sheet_id: SpreadsheetApp.getActiveSpreadsheet().getId() })
  };
  var response = UrlFetchApp.fetch(url, options);
  var json = {};
  try {
    json = JSON.parse(response.getContentText());
  } catch (e) {
    Browser.msgBox("Respuesta inválida del backend: " + response.getContentText());
    return;
  }
  if (json.qr) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuracion");
    // Guarda el QR en B2 (columna 2)
    sheet.getRange(2, 2).setValue(json.qr);
    // Inserta la imagen en C2 (columna 3, fila 2)
    try {
      sheet.insertImage(decodeURIComponent(json.qr), 3, 2);
    } catch (e) {
      // Si falla, solo deja el texto
    }
    Browser.msgBox("QR recibido y guardado.");
  } else if (json.status == "-1" && reintentos > 0) {
    Utilities.sleep(2500); // Espera 2.5 segundos
    obtenerQRyGuardarConReintentos(reintentos - 1);
  } else {
    Browser.msgBox("Estado: " + json.message);
  }
}