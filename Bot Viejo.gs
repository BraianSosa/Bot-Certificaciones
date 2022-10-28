/*// Code by Taiel Martinez 2020
// El uso de este script para fines comerciales esta prohibido.
// Registro: https://github.com/TaielMartinez/
function CertificacionEntidadesStart() {
    // conf //
    const develop = false;
    const production = true;
    const conf_letra_valores = "B";
    
    const letra_numero = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
                          'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ',
                          'BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ',
                          'CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ',
                          'DA','DB','DC','DD','DE','DF','DG','DH','DI','DJ','DK','DL','DM','DN','DO','DP','DQ','DR','DS','DT','DU','DV','DW','DX','DY','DZ'];
    var date = Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy")
    
    console.log(" --------- start --------- ");
    
    
    ///////////////////////////////////////////////////////////////////////////// DECLARACION DE VARIABLES /////////////////////////////////////////////////////////////////////////////
    
    
    var cantidad_columnas_copiadas = 0;
    var cantidad_mails_enviados = 0;
    
    
    var respuestasNombreHoja = "Respuestas de formulario";
    var respuestasLetraFinal = "CN";
    var respuestasNumeroInicio = 2;
    var respuestasColumnaID = 0;
    var respuestasColumnaTipoEntidad = 16;
    
    var certificacionesGeneralHoja = "Certificaciones General";
    var certificacionesGeneralID = 0;
    var certificacionesGeneralNombreEntidad = 2;
    var certificacionesGeneralNombreResponsable = 11;
    var certificacionesGeneralTelefonoResponsable = 12;
    var certificacionesGeneralMail = 14;
    var certificacionesGeneralAutomatico = 1;
    var certificacionesGeneralCantidadColumnasPersonalizadas = 7;
    var certificacionesGeneralCantidadColumnasDelFormulario = 93;
    
    
    var conf_sheet_mail = getSheet("Config Mail", "A", 1, "F", 23).getValues();
    var hoja_respuestas_name = certificacionesGeneralHoja;
    var mail_adress = 21;
    var replace_motivo = 6;
    //var replace_asunto = config(20);
    var replace_organizacion = 9;
    var replace_responsable = 18;
    var replace_tipo = 23;
    var veredicto = 2;
    var date_enviado = 5;
    var mail_enviado = 4;
    var texto_enviar_mail = "PENDIENTE";
    
    
    var columna_mails = {
      "Club de Esports": 4,
      "Medio de Esports": 7,
      "Liga de Esports": 10,
      "Comunidad": 13,
      "Organizador de Esports": 16
    }
    
    var mail_message = {
      "Aprobado A": 5,
      "Aprobado B": 10,
      "Aprobado C": 15,
      "Rechazado": 20
    }
    
    var subject_aprobado_A = conf_sheet_mail[5][4];
    var message_aprobado_A = conf_sheet_mail[6][4];
    
    var subject_aprobado_B = conf_sheet_mail[10][4];
    var message_aprobado_B = conf_sheet_mail[11][4];
    
    var subject_aprobado_C = conf_sheet_mail[15][4];
    var message_aprobado_C = conf_sheet_mail[16][4];
    
    var subject_rechazado = conf_sheet_mail[20][4];
    var message_rechazado = conf_sheet_mail[21][4];
    
    // los numeros representan el numero de columna que tene que ser SI para que puedan entrar a esa categoria
    // 78, // registro de jugadores
    const validacionEntidad = {
      "Club de Esports":{
        "C": [
          85, // logo
          84 // +1000 seguidores 
        ],
        "B": [
          83, // +2 redes sociales
          77, // reglamento interno
          72, // registro de marca
          71, // capacidad de facturar
          73, // estructura administrativa parcial
          75 // estructura deportiva parcial
        ],
        "A": [
          79, // plan de redes sociales
          80, // contrato estructura administrativa
          81, // contrato estructura competitiva
          70, // poseo personería jurídica
          74, // estructura administrativa completa
          76 // estructura deportiva completa
        ]
      },
      "Medio de Esports":{
        "C": [
          68, // +1000 seguidores
          66, // Redes sociales 1
          69  // Logo
          
        ],
        "B": [
          61, // Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).
          62  // Poseo Reglamento Interno. (Código de conducta).
  
        ],
        "A": [
          60, //Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.).
          57, //Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)
          63, //Poseo más de 1 año de antigüedad.
          59, // Poseo registro de marca.
          67  //Poseo 2 o más redes sociales.
          
        ]
      },
      "Liga de Esports":{
        "C": [
          55, // +1000 seguidores
          53, // Redes sociales 1
          56  // Logo
          
        ],
        "B": [
          47, // Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).
          44, // Poseo capacidad de facturación (Monotributo, Responsable Inscripto).
          48, // Poseo Reglamento Interno. (Código de conducta).
          54, // Poseo 2 o más redes sociales.
          49  // Poseo Reglamentos por cada disciplina.
  
        ],
        "A": [
          46, // Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.).
          50, // Poseo más de 1 año de antigüedad.
          43, // Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)
          45  // Poseo registro de marca.
          
        ]
      },
      "Comunidad":{
        "C": [
          28, // +1000 seguidores
          26, // Redes sociales 1
          29  // Logo
          
        ],
        "B": [
          21, // Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).
          22, // Poseo Reglamento Interno. (Código de conducta).
          27  // Poseo 2 o más redes sociales.
          
        ],
        "A": [
          17, // Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)
          19, // Poseo registro de marca.
          25, // Poseo contratos con Estructura Administrativa.
          24  // Poseo Plan de Redes Sociales.
          
        ]
      },
      "Organizador de Esports":{
        "C": [
          41, // +1000 seguidores
          39, // Redes sociales 1
          42  // Logo
          
          
        ],
        "B": [
          34, // Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).
          31, // Poseo capacidad de facturación (Monotributo, Responsable Inscripto).
          35, // Poseo Reglamento Interno. (Código de conducta).
          40  // Poseo 2 o más redes sociales.
          
          
        ],
        "A": [
          33, // Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.).
          36, // Poseo más de 1 año de antigüedad.
          30  // Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)
          
          
        ]
      }
    }
    
    const validacionNombre = {
      47: 'Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).',
      44: 'Poseo capacidad de facturación (Monotributo, Responsable Inscripto).',
      48: 'Poseo Reglamento Interno. (Código de conducta).',
      54: 'Poseo 2 o más redes sociales.',
      49: 'Poseo Reglamentos por cada disciplina.',
      46: 'Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.).',
      50: 'Poseo más de 1 año de antigüedad.',
      43: 'Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)',
      45: 'Poseo registro de marca.',
      28: '+1000 seguidores',
      26: 'Redes sociales 1',
      29: 'Logo',
      21: 'Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).',
      22: 'Poseo Reglamento Interno. (Código de conducta).',
      27: 'Poseo 2 o más redes sociales.',
      17: 'Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)',
      19: 'Poseo registro de marca.',
      25: 'Poseo contratos con Estructura Administrativa.',
      24: 'Poseo Plan de Redes Sociales.',
      41: '+1000 seguidores',
      39: 'Redes sociales 1',
      42: 'Logo',
      34: 'Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).',
      31: 'Poseo capacidad de facturación (Monotributo, Responsable Inscripto).',
      35: 'Poseo Reglamento Interno. (Código de conducta).',
      40: 'Poseo 2 o más redes sociales.',
      33: 'Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.).',
      36: 'Poseo más de 1 año de antigüedad.',
      55: '+1000 seguidores',
      53: 'Redes sociales 1',
      56: 'Logo',
      30: 'Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)',
      60: 'Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.)',
      61: 'Poseo Estructura Administrativa Parcial (Periodista, CM, etc.)',
      62: 'Poseo Reglamento Interno. (Código de conducta)',
      57: 'Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)',
      63: 'Poseo más de 1 año de antigüedad',
      59: 'Poseo registro de marca',
      67: 'Poseo 2 o más redes sociales.',
      68: '+1000 seguidores',
      66: 'Redes sociales 1',
      69: 'Logo',
      70: 'poseo Personería Jurídica',
      71: 'poseo capacidad de facturación',
      72: 'poseo registro de marca del Club de Esports.',
      73: 'poseo Estructura Administrativa Completa',
      74: 'poseo Estructura Administrativa Parcial',
      75: 'poseo Estructura Competitiva Completa',
      76: 'poseo Estructura Competitiva Parcial',
      77: 'poseo Reglamento Interno',
      78: 'poseo Registro de Jugadores',
      79: 'poseo Plan de Redes Sociales',
      80: 'poseo contratos con Estructura Administrativa',
      81: 'poseo contratos con Estructura Competitiva',
      82: 'poseo 1 red social',
      83: 'poseo 2 o más redes sociales',
      84: 'poseo más de 1000 seguidores en una de mis redes sociales',
      85: 'poseo Logo Propio en PNG/PDF con fondo transparente'
    }
    
    console.log("variables declaradas");
    
    var id_utilizadas = getSheet(certificacionesGeneralHoja, "A", 1, "A").getValues();
    
    // Para los Logs
    var nombre_hoja_logs = "LogError";
    var log_cant = getSheet(nombre_hoja_logs, "A", 1, "A").getValues().length;
    //var logsSheet = getSheet(nombre_hoja_logs, "A", 1, "C").getValues();
    var hojaLogs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre_hoja_logs);
    if(!hojaLogs) console.log("Error 'hojaLogs' no existe")
    var log_fila = log_cant;
    ///////////////////////////////////////////////////////////////////////////// declaracion de variables /////////////////////////////////////////////////////////////////////////////
    
    
    
    ///////////////////////////////////////////////////////////////////////////// COPIAR RESPUESTA A NUEVA HOJA /////////////////////////////////////////////////////////////////////////////
    
    var respuestasSheet = getSheet(respuestasNombreHoja, "A", 1, respuestasLetraFinal).getValues();
    var hojaFormulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(respuestasNombreHoja);
    var hojaCertificacionesGeneral = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(certificacionesGeneralHoja);
    
    var newLineCertificacionesGeneral = id_utilizadas.length;
    //console.log("newLineCertificacionesGeneral "+ newLineCertificacionesGeneral)
    //console.log("respuestas largo " + respuestasSheet.length)
    for (let i = respuestasNumeroInicio; i < respuestasSheet.length; i++) {
      let row = respuestasSheet[i-1];

      //console.log(row[1]); console.log(i);
      var row_id = row[respuestasColumnaID]
      
      //console.log("ID= "+row_id)
      //console.log("line: " + i)
      //console.log("nombre: " + row[2])
      //if(row_id == "") console.log("vacio")
      //if(row_id == undefined) console.log("undefined")
      //if(row_id.length < 5) console.log("corto")
      if(row[respuestasColumnaID] == ""){
        var newID = "CE-"+id_generate();
        console.log("ID"+row[respuestasColumnaID])
        newLineCertificacionesGeneral++;
        cantidad_columnas_copiadas = cantidad_columnas_copiadas + 1;
        for (let c = 0; c < certificacionesGeneralCantidadColumnasDelFormulario; c++) {
          hojaCertificacionesGeneral.getRange(letra_numero[c + certificacionesGeneralCantidadColumnasPersonalizadas]+newLineCertificacionesGeneral).setValue(row[c]);
        }
        hojaCertificacionesGeneral.getRange(letra_numero[certificacionesGeneralID]+newLineCertificacionesGeneral).setValue(newID);
        hojaCertificacionesGeneral.getRange(letra_numero[certificacionesGeneralAutomatico]+newLineCertificacionesGeneral).setValue(determinarTipoDeCertificacion(row));
        hojaFormulario.getRange(letra_numero[respuestasColumnaID]+i).setValue(newID);
      }
    }
    
    ///////////////////////////////////////////////////////////////////////////// copiar respuestas a nueva hoja ////////////////////////////////////////////////////////////////////////////
    
    
    
    
    
    
    guardarLog("Start", "Script ejecutado", date);
    
    
    
    
    
    
    
    ///////////////////////////////////////////////////////////////////////////// ENVIAR MENSAJES /////////////////////////////////////////////////////////////////////////////
    
    
    
    var dataRange = getSheet(hoja_respuestas_name, "A", 2, "Y");
    var hojaFormulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja_respuestas_name);
    var data = dataRange.getValues();
    var index = 2;
    for (var i in data) {
      var row = data[i];
      if(row[mail_enviado] == texto_enviar_mail && (row[veredicto] == "Aprobado A" || row[veredicto] == "Aprobado B" || row[veredicto] == "Aprobado C" || row[veredicto] == "Rechazado")){
        
        //try {
        if(false){ // consologear variables
          console.log(`veredicto: ${veredicto} - row[veredicto]: ${row[veredicto]} - mail_message[row[veredicto]]: ${mail_message[row[veredicto]]}`);
          console.log(`replace_tipo: ${replace_tipo} - columna_mails: ${columna_mails[row[replace_tipo]]}`);
          console.log(`subject: ${conf_sheet_mail[mail_message[row[veredicto]]][columna_mails[row[replace_tipo]]]}`);
          console.log("mail_message[row[parseInt(veredicto) + 1]] " + mail_message[row[parseInt(veredicto) + 1]])
          console.log("columna_mails[row[replace_tipo]] " + columna_mails[row[replace_tipo]])
        }
        
        try{
          var subject = conf_sheet_mail[mail_message[row[veredicto]]][columna_mails[row[replace_tipo]]];
          var message = conf_sheet_mail[mail_message[row[veredicto]]+1][columna_mails[row[replace_tipo]]];
        } catch (error) {
          guardarLog('Mail', 'error en subject o message', error)
          console.error("-- error en subject o message --");
          console.log("veredicto: " + veredicto);
          console.log("row[veredicto]: " + row[veredicto]);
          console.log("mail_message[row[veredicto]]: " + mail_message[row[veredicto]]);
          console.log("replace_tipo: " + replace_tipo);
          console.log("row[replace_tipo]: " + row[replace_tipo]);
          console.log("columna_mails[row[replace_tipo]]: " + Number(columna_mails[row[replace_tipo]]));
          console.log("conf_sheet_mail[Number(mail_message[row[veredicto]])][Number(columna_mails[row[replace_tipo]])]: " + conf_sheet_mail[mail_message[row[veredicto]]][columna_mails[row[replace_tipo]]]);
          console.log("-----------------");
          console.log("veredicto + 1: " + (Number(veredicto) + 1));
          console.log("row[veredicto + 1]: " + row[veredicto + 1]);
          console.log("mail_message[row[parseInt(veredicto) + 1]]: " + mail_message[row[Number(veredicto) + 1]]);
          console.log("mail_message[row[veredicto]]: " + mail_message[row[veredicto]]);
          console.log("row[replace_tipo]: " + row[replace_tipo]);
          console.log("columna_mails[row[replace_tipo]]: " + columna_mails[row[replace_tipo]]);
        }
        
        try {
          subject = subject.toString().replace("<nombre>", row[replace_organizacion]).replace("<tipo>", row[replace_tipo]);
          message = message.toString().replace("<nombre>", row[replace_organizacion]).replace("<tipo>", row[replace_tipo]);
          
          var emailAddress = "tmartinez@deva.org.ar";   //develop
          if(develop == false && production == true){
            console.log("Mail produccion: " + row[mail_adress]);
            emailAddress = row[mail_adress];
          } else {
            console.log("Mail develop = " + develop + " production = " + production);
          }
          
          
          //console.log(message);
          if(emailAddress && message && subject){
            cantidad_mails_enviados = cantidad_mails_enviados + 1;
            console.log("------------------ Send Mail ------------------");
            //console.log("email: " + emailAddress);
            //console.log("asunto: " + subject);
            //console.log("mensaje: " + message);
            MailApp.sendEmail(emailAddress, subject, message);
            
            Utilities.sleep(5000);
            var threads = GmailApp.search("is:unread subject:\"delivery status notification\" " + emailAddress);
            if (threads.length > 0) {
              console.log("Failed: Delivery failed. " + threads[0].getPermalink() + "  || emailAddress: "+ emailAddress)
              guardarLog("Mail", "Mail no enviado", emailAddress);
            }
            
            console.log(letra_numero[mail_enviado]+index);
            hojaFormulario.getRange(letra_numero[mail_enviado]+index).setValue("ENVIADO");
            hojaFormulario.getRange(letra_numero[date_enviado]+index).setValue(date);
            console.log("=============================================================");
          }
        } catch (error) {
          console.error("-- error en subject o message --");
          console.log("veredicto: " + veredicto);
          console.log("row[veredicto]: " + row[veredicto]);
          console.log("mail_message[row[veredicto]]: " + mail_message[row[veredicto]]);
          console.log("replace_tipo: " + replace_tipo);
          console.log("row[replace_tipo]: " + row[replace_tipo]);
          console.log("columna_mails[row[replace_tipo]]: " + columna_mails[row[replace_tipo]]);
          console.log("conf_sheet_mail[mail_message[row[veredicto]]][columna_mails[row[replace_tipo]]]: " + conf_sheet_mail[mail_message[row[veredicto]]][columna_mails[row[replace_tipo]]]);
          console.log("-----------------");
          console.log("veredicto + 1: " + veredicto + 1);
          console.log("row[veredicto + 1]: " + row[veredicto + 1]);
          console.log("mail_message[row[parseInt(veredicto) + 1]]: " + mail_message[row[parseInt(veredicto) + 1]]);
          console.log("mail_message[row[veredicto]]: " + mail_message[row[veredicto]]);
          console.log("row[replace_tipo]: " + row[replace_tipo]);
          console.log("columna_mails[row[replace_tipo]]: " + columna_mails[row[replace_tipo]]);
        }
      }
      index++;
    }
    
    ///////////////////////////////////////////////////////////////////////////// enviar mensajes /////////////////////////////////////////////////////////////////////////////
    
      
    
    
    

    
    
    
    ///////////////////////////////////////////////////////////////////////////// TOOLS /////////////////////////////////////////////////////////////////////////////
    
    
    function getSheet(hoja, inicio_letra, incio_numero, final_letra, final_numero){
      console.log("getSheet - start");
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja);
      var _sheet = "";
      if(!final_numero){
        final_numero = sheet.getLastRow();
      }

      if(hoja){
        _sheet = "'"+hoja+"'!";
      }
      console.log("hoja: "+hoja+" inicio_letra: "+inicio_letra+" incio_numero: "+incio_numero+" final_letra: "+final_letra+" final_numero: "+final_numero);
      console.log("getSheet - end");
      return (sheet.getRange(_sheet+inicio_letra+incio_numero+":"+final_letra+final_numero));
      
    }
    
    function remplazar_mensaje(texto, row){
      try{
        return texto.toString().replace("<nombre>", row[replace_organizacion]).replace("<tipo>", row[replace_tipo])
      } catch (error) {
        console.error(error);
      }
    }
    
    function id_generate(){
      console.log("id_generate - start");
      var id_repetida = true;
      while (id_repetida){
        id_repetida = false;
        var result           = '';
        var characters       = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        var charactersLength = characters.length;
        for ( var i = 0; i < 7; i++ ) {
          result += characters.charAt(Math.floor(Math.random() * charactersLength));
        }
        console.log("id creada " + result);
        for ( var i = 0; i < id_utilizadas.length; i++ ) {
          if(id_utilizadas[i][0].includes(result))
            id_repetida = true;
        }
      }
      console.log("id_generate - end");
      return result;
    }
    
    function determinarTipoDeCertificacion(row){
      console.log("determinarTipoDeCertificacion - start");
      var result = 'Rechazar';
      var motivos = '';
      var rechazado = false;
      //try {
      let types = Object.keys(validacionEntidad[row[respuestasColumnaTipoEntidad]]);
      //console.log(types);
      //console.log(types.length);
      for (let c = types.length; c > 0; c--) {
        if(!rechazado){
          var type = validacionEntidad[row[respuestasColumnaTipoEntidad]];
          let letra = letra_numero[c-1];
          //console.log("letra "+letra);
          type = type[letra];
          var validType = true;
          for (let i = 0; i < type.length; i++) {
            //console.log("linea: " + type[i]);
            //console.log(row[type[i]]);
            if(row[type[i]] != "SI"){
              validType = false;
              motivos = motivos + " - No " + validacionNombre[type[i]];
              rechazado = true;
              //console.log("================false================")
            } else {
              //console.log("true")
            }
          }
          if(validType) {
            console.log("determinarTipoDeCertificacion - end");
            result = letra;
          }
        }
      } 
      
      
      // ///////////////////// COMPROBAR REPETIDOS ///////////////////////////
      
      /*
      let repetido = ''; 
      console.log("COMPROBAR REPETIDOS "+hojaCertificacionesGeneral.length)
      for (let i = 0; i < hojaCertificacionesGeneral.length; i++) {
      console.log("COMPROBAR REPETIDOS")
      let nombreEntidad = hojaCertificacionesGeneral[i][certificacionesGeneralNombreEntidad + certificacionesGeneralCantidadColumnasPersonalizadas];
      nombreEntidad = nombreEntidad.toLowerCase()
      //console.log('nombreEntidad: '+ nombreEntidad)
      if(row[certificacionesGeneralNombreEntidad] == nombreEntidad) {
      repetido = repetido + "REPETIDO > Nombre de entidad ("+hojaCertificacionesGeneral[i][certificacionesGeneralID]+") -"
      }
      
      let mail = hojaCertificacionesGeneral[i][certificacionesGeneralMail + certificacionesGeneralCantidadColumnasPersonalizadas];
      mail = mail.toLowerCase()
      //console.log('mail: '+mail)
      if(row[certificacionesGeneralMail] == mail) {
      repetido = repetido + "REPETIDO > Mail ("+hojaCertificacionesGeneral[i][certificacionesGeneralID]+") -"
      }
      
      let telefonoResponsable = hojaCertificacionesGeneral[i][certificacionesGeneralTelefonoResponsable + certificacionesGeneralCantidadColumnasPersonalizadas];
      telefonoResponsable = telefonoResponsable.toLowerCase()
      telefonoResponsable = telefonoResponsable.replace("+", "");
      //console.log('telefonoResponsable: '+telefonoResponsable)
      if(row[certificacionesGeneralTelefonoResponsable] == telefonoResponsable) {
      repetido = repetido + "REPETIDO > Telefono del responsable ("+hojaCertificacionesGeneral[i][certificacionesGeneralID]+") -"
      }
      
      let nombreResponsable = hojaCertificacionesGeneral[i][certificacionesGeneralNombreResponsable + certificacionesGeneralCantidadColumnasPersonalizadas];
      nombreResponsable = nombreResponsable.toLowerCase()
      //console.log('nombreResponsable: '+nombreResponsable)
      if(row[certificacionesGeneralNombreResponsable] == nombreResponsable) {
      repetido = repetido + "REPETIDO > Nombre del responsable ("+hojaCertificacionesGeneral[i][certificacionesGeneralID]+") -"
      }
      }
      //poner fin de comentario
      
      return result + " -" + motivos;
      //}
      //catch(err) {
      //  console.log(err)
      //}
    }
    
    
    
    function guardarLog(tipo, error, variable){
      console.log("guardarLog - start");
      if(!hojaLogs) console.log("Error 'hojaLogs' no existe en funcion")
      log_fila = log_fila + 1;
      console.log("log_fila: "+log_fila)
      
      hojaLogs.getRange("A"+log_fila).setValue(tipo);
      hojaLogs.getRange("B"+log_fila).setValue(error);
      hojaLogs.getRange("C"+log_fila).setValue(variable);
      
      console.log("LOG: "+tipo+" | "+error+" | "+variable)
      console.log("guardarLog - end");
    }
    
    ///////////////////////////////////////////////////////////////////////////// tools /////////////////////////////////////////////////////////////////////////////
    
    console.log("-------------------------------------------------------------");
    console.log("Cantidad de columnas copiadas: " + cantidad_columnas_copiadas);
    console.log("Cantidad de mails enviados: " + cantidad_mails_enviados);
    console.log("-------------------------------------------------------------");
    
  }
*/



