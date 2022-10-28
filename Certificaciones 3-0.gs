/*
* Sistema de generación de certificados | Copyright DEVA 2020
* Versión: v3.0.0 release
* Changelog: verificar en hoja changelog
* Desarrollado por la Dirección de Sistemas
*/
/**** Declaración de variables****/
//VARIABLES DE HOJAS
 const hoja1 = "Respuestas de formulario";
 const hoja2 = "Certificaciones General";
 const hoja3 = "Entity";
 var reason = '';

console.log(" INICIA EL PROGRAMA ");

/**** FUNCION MAIN****/
//CARGA DE DATOS
function loadData() {
const lastrow   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja1).getLastRow()
let data        = getSheet(hoja1, 'A', lastrow, 'CL', lastrow, 'withValues');
let current_row = data[0]
setPrefix(lastrow);
let verdict_var   = verdict(current_row, lastrow);
let entity_name = current_row[2];
console.log(`${entity_name} veredicto:  ${verdict_var}`)
copyData(lastrow);
console.log(` Datos de entidad ${entity_name} copiada en Hoja: Certificaciones General `);
console.log(` FIN `);
}


/*
* La función getSheet obtiene la hoja mediante el párametro sheetName
*luego obtiene el rango mediante los demás párametros
*y por ultimo se le asigna alguno de los sig parametros: "withValues" que trae el rango con los valores , "onlyRange" que trae unicamente el rango y "onlySheet" que solo trae el nombre de la hoja
*/
function getSheet(sheetName, sheetFirstLetter, sheetFirstNumber, sheetSecondLetter, sheetSecondNumber, sheetOption){
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if(!sheetSecondNumber || sheetSecondNumber === ""){
    sheetSecondNumber = sheet.getLastRow();
    }
    let isData;
    let rangeDefinition;
    switch(sheetOption){
      case "withValues":
        rangeDefinition = `${sheetFirstLetter}${sheetFirstNumber}:${sheetSecondLetter}${sheetSecondNumber}`;
        let rangeData = sheet.getRange(rangeDefinition);
        isData = rangeData.getValues();
        break;
      case "onlyRange":
        rangeDefinition = `${sheetFirstLetter}${sheetFirstNumber}:${sheetSecondLetter}${sheetSecondNumber}`;
        isData = sheet.getRange(rangeDefinition);
        break;
      case "onlySheet":
        isData = sheet;
        break;
      case "onlyLastRow":
        isData = sheet.getLastRow();
        break;
      default: 
        console.log("Something Happens");
        console.log("-> Check getSheet function <-");
    }
    return isData;
  }
//ID BY: Taiel Martinez
 function id(){
      console.log("Se crea la ID");   
        let char      = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
        var result    = '';
        // Se crea una ID con la longitud igual al numero que indiquemos en el For de abajo.
        for ( var i = 0; i < 7; i++ ) {
        //Si no se concatena Result no vamos a lograr crear una ID de mas de un caracter.
          result   += char.charAt(Math.floor(Math.random() * char.length));
        }
      console.log("La ID es: " + result);     
      console.log("Funcion ID Completa");
      return result;
    }
//Set prefix by: Exec
function setPrefix(lastrow){

  /*
  * Obtiene los valores con los rangos Q2:Q
  */
  let getEntityType = getSheet(hoja1, "A", 2, "Q", lastrow, 'withValues');
  let obtainedResult = getEntityType;

  let entityId = 0;
  let entityType = 16;

  /*
  * Declaramos una variables llamada newPrefix pero no la inicializamos
  */
  let newPrefix;

  for(let i = 0; i < obtainedResult.length; i++){

  let entityData = obtainedResult[i];
  
  /*
  * Declaramos idCol, esta variable se posiciona en la columna
  */
  let idCol = i + 2;
  let toStringValue = entityData[entityType].toString();
  let obtainedId = id();
  switch(toStringValue){
    case 'Club de Esports':
      newPrefix = `CE-${obtainedId}`;
      console.log(`Obtained result: ${newPrefix} for Club de Esports`);
      break;
    case 'Comunidad':
      newPrefix = `CM-${obtainedId}`;
      console.log(`Obtained result: ${newPrefix} for Comunidad`);
      break;
    case 'Organizador de Esports':
      newPrefix = `OE-${obtainedId}`;
      console.log(`Obtained result: ${newPrefix} for Organizador de Esports`);
      break;
    case 'Liga de Esports':
      newPrefix = `CM-${obtainedId}`;
      console.log(`Obtained result: ${newPrefix} for Liga de Esports`);
      break;
    case 'Medio de Esports':
      newPrefix = `CM-${obtainedId}`;
      console.log(`Obtained result: ${newPrefix} for Medio de Esports`);
      break;
    default: console.log("Something happen");
  }
        if(!entityData[entityId] || entityData[entityId] == ""){
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja1).getRange(idCol, 1).setValue(newPrefix) // Setear valores en Columna de ID (A) en certis gral.
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja2).getRange(idCol, 1).setValue(newPrefix) // Setear valores en Columna de ID (A) en certis gral.
        } else {
          console.log("Unable to copy id, because this already exists");
        }
  }
}


 //VEREDICTO BY: Braian Sosa & Sebastian Szostak
function verdict(current_row, lastrow){
    //DECLARACION DE CONSTANTES DE COLUMNAS
  const CDE = {  //Club De E-Sports
          C: [
            85, // logo //Columna CH
            84 // +1000 seguidores  //Columna CG
          ],
          B: [
            83, // +2 redes sociales // Columna CF
            77, // reglamento interno // Columna BZ
            72, // registro de marca // Columna BU
            71, // capacidad de facturar // Columna BT
            73, // estructura administrativa parcial // Columna BV
            75 // estructura deportiva completa // Columna BX
          ],
          A: [
            79, // plan de redes sociales //Columna CB
            80, // contrato estructura administrativa // Columna CC
            81, // contrato estructura competitiva // Columna CD
            70, // poseo personería jurídica // Columna BS
            74, // estructura administrativa Parcial //// Columna BW
            76 // estructura competitiva parcial  // Columna BY
          ]
      }
  const MDE = { //Medio De E-Sports
          C: [
            69,  // logo //Columna BR
            68,  // +1000 seguidores  // Columna BQ
            66   // Redes sociales 1 // Columna BO
          ],
          B: [
            62,  // Poseo Reglamento Interno. (Código de conducta). //Columna BK
            61 // Poseo Estructura Administrativa Parcial (Periodista, CM, etc.). //Columna BJ
  
          ],
          A: [
            60,  //Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.). //Columna BI
            57,  //Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.) //Columna BF
            63,  //Poseo más de 1 año de antigüedad. //Columna BL
            59,  // Poseo registro de marca. // Columna BH
            67   //Poseo 2 o más redes sociales. //Columna BP
          ]
      }
  const LDE = { //Liga De E-Sports
          C: [
            56,  // Logo //Columna BE
            55, // +1000 seguidores //Columna BD
            53 // Redes sociales 1 //Columna BB
  
          ],
          B: [
            47, // Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).  //Columna AV
            44, // Poseo capacidad de facturación (Monotributo, Responsable Inscripto). //Columna AS
            48, // Poseo Reglamento Interno. (Código de conducta). //Columna AW
            54, // Poseo 2 o más redes sociales. //Columna BC
            49  // Poseo Reglamentos por cada disciplina. //Columna AX
          ],
          A: [
            50, // Poseo más de 1 año de antigüedad. //Columna AY
            46, // Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.). //Columna AU
            45, // Poseo registro de marca. //Columna AT
            43 // Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.) //Columna AR
          ]
      }
  const ODE = { //Organizador De E-Sports
          C: [
            42,  // Logo //Columna AQ
            41, // +1000 seguidores //Columna AP
            39 // Redes sociales 1 // Columna AN
          ],
          B: [
            40,  // Poseo 2 o más redes sociales. //Columna AO
            35, // Poseo Reglamento Interno. (Código de conducta). //Columna AJ
            34, // Poseo Estructura Administrativa Parcial (Periodista, CM, etc.). //Columna AI
            31 // Poseo capacidad de facturación (Monotributo, Responsable Inscripto). //Columna AF
          ],
          A: [
            36, // Poseo más de 1 año de antigüedad. //Columna AK
            33, // Poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.). //Columna AH
            30  // Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.) //Columna AE
          ]
      }
  const COMU = { //Comunidad De E-Sports
          C: [
            29,  // Logo 
            28, // +1000 seguidores
            26 // Redes sociales 1
          ],
          B: [
            27,  // Poseo 2 o más redes sociales.
            22, // Poseo Reglamento Interno. (Código de conducta).
            21 // Poseo Estructura Administrativa Parcial (Periodista, CM, etc.).
          ],
          A: [
            25, // Poseo contratos con Estructura Administrativa.
            24, // Poseo Plan de Redes Sociales.
            19, // Poseo registro de marca.
            17 // Poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)
          ]
      }
  const valid = {
      47: 'poseo Estructura Administrativa Parcial (Periodista, CM, etc.).',
      44: 'poseo capacidad de facturación (Monotributo, Responsable Inscripto).',
      48: 'poseo Reglamento Interno. (Código de conducta).',
      54: 'poseo 2 o más redes sociales.',
      49: 'poseo Reglamentos por cada disciplina.',
      46: 'poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.).',
      50: 'poseo más de 1 año de antigüedad.',
      43: 'poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)',
      45: 'poseo registro de marca.',
      28: 'poseo +1000 seguidores',
      26: 'poseo Redes sociales 1',
      29: 'poseo Logo',
      21: 'poseo Estructura Administrativa Parcial (Periodista, CM, etc.).',
      22: 'poseo Reglamento Interno. (Código de conducta).',
      27: 'poseo 2 o más redes sociales.',
      17: 'poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)',
      19: 'poseo registro de marca.',
      25: 'poseo contratos con Estructura Administrativa.',
      24: 'poseo Plan de Redes Sociales.',
      41: 'poseo +1000 seguidores',
      39: 'poseo Redes sociales 1',
      42: 'poseo Logo',
      34: 'poseo Estructura Administrativa Parcial (Periodista, CM, etc.).',
      31: 'poseo capacidad de facturación (Monotributo, Responsable Inscripto).',
      35: 'poseo Reglamento Interno. (Código de conducta).',
      40: 'poseo 2 o más redes sociales.',
      33: 'poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.).',
      36: 'poseo más de 1 año de antigüedad.',
      55: 'poseo +1000 seguidores',
      53: 'poseo Redes sociales 1',
      56: 'poseo Logo',
      30: 'poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)',
      60: 'poseo Estructura Administrativa Completa (Director, periodista, editor, CM, etc.)',
      61: 'poseo Estructura Administrativa Parcial (Periodista, CM, etc.)',
      62: 'poseo Reglamento Interno. (Código de conducta)',
      57: 'poseo Personería Jurídica (SRL,SAS,Asociación Civil, etc.)',
      63: 'poseo más de 1 año de antigüedad',
      59: 'poseo registro de marca',
      67: 'poseo 2 o más redes sociales.',
      68: 'poseo +1000 seguidores',
      66: 'poseo Redes sociales 1',
      69: 'poseo Logo',
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
  
  //FIN DECLARACION DE CONSTANTES DE COLUMNAS.
  let  verdict //Variable de retorno
  let  entity_type = current_row[16] //Posicion del array donde esta la columna del tipo de entidad en la hoja de formulario
  let  category; //Booleano para verificar si es o no la categoria
  let  date = Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy");
  switch(entity_type){
    case 'Club de Esports':
           category = compare(CDE.C, lastrow, valid)
           if(category == false){
           verdict = 'Rechazado';
           break;
           } 
           else{
           verdict = 'Aprobado C';
           }
           category = compare(CDE.B, lastrow, valid)
           if(category){
           verdict = 'Aprobado B';
           }
           else{
           break;
           }
           category = compare(CDE.A, lastrow, valid)
           if(category){
           verdict = 'Aprobado A';
           break;
           }
           else{
           break;
           }
    case 'Liga de Esports':
           category = compare(LDE.C, lastrow, valid)
           if(category == false){
           verdict = 'Rechazado';
           break;
           } 
           else{
           verdict = 'Aprobado C';
           }
           category = compare(LDE.B, lastrow, valid)
           if(category){
           verdict = 'Aprobado B';
           }
           else{
           break;
           }
           category = compare(LDE.A, lastrow, valid)
           if(category){
           verdict = 'Aprobado A';
           break;
           }
           else{
           break;
           }
    case 'Medio de Esports':
           category = compare(MDE.C, lastrow, valid)
           if(category == false){
           verdict = 'Rechazado';
           break;
           } 
           else{
           verdict = 'Aprobado C';
           }
           category = compare(MDE.B, lastrow, valid)
           if(category){
           verdict = 'Aprobado B';
           }
           else{
           break;
           }
           category = compare(MDE.A, lastrow, valid)
           if(category){
           verdict = 'Aprobado A';
           break;
           }
           else{
           break;
           }
    case 'Organizador de Esports':
           category = compare(ODE.C, lastrow, valid)
           if(category == false){
           verdict = 'Rechazado';
           break;
           } 
           else{
           verdict = 'Aprobado C';
           }
           category = compare(ODE.B, lastrow, valid)
           if(category){
           verdict = 'Aprobado B';
           }
           else{
           break;
           }
           category = compare(ODE.A, lastrow, valid)
           if(category){
           verdict = 'Aprobado A';
           break;
           }
           else{
           break;
           }
    case 'Comunidad':
           category = compare(COMU.C, lastrow, valid)
           if(category == false){
           verdict = 'Rechazado';
           break;
           } 
           else{
           verdict = 'Aprobado C';
           }
           category = compare(COMU.B, lastrow, valid)
           if(category){
           verdict = 'Aprobado B';
           }
           else{
           break;
           }
           category = compare(COMU.A, lastrow, valid)
           if(category){
           verdict = 'Aprobado A';
           break;
           }
           else{
           break;
           }
    default: 
    console.error('Error en veredicto (Switch Case)')
  }
SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja2).getRange(lastrow, 2).setValue(verdict+reason) //Veredicto reemplazado en columna B (automatico)
//SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja2).getRange(lastrow, 3).setValue(verdict) //Veredicto reemplazado en columna C (Veredicto)
SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja2).getRange(lastrow, 5).setValue(date) //Fecha del veredicto definida en columna E
console.log('----Valores reemplazados en Certificaciones General -----')
return verdict
}
//Funcion de comparacion de los requisitos de cada categoria By: Braian Sosa & Seba_insertarapellido, Tincho Miguens
function compare(object_propiety, lastrow, validation){
    let cell
    let state = true;
    let columns = object_propiety
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja1);
    for(let x = 0; x<columns.length; ++x){
        cell = sheet.getRange(lastrow,columns[x]+1).getValues();
        let varcompare = cell.toString();
        let validc = columns[x];
        if(varcompare.toUpperCase() =='NO' || cell[0] == ''){
           state = false
           reason = `${reason}; No ${validation[validc]}`;
        }
    }
  return state; //Si es true, entonces es la categoria, se utilizaria en el switch case para determinar categoria y dar un break. Si retorna false rompe el bucle y retorna false.
}

//ESCRITURA DE DATOS BY:  Tincho Miguens
function copyData(lastrow){
let data = getSheet(hoja1, 'C', lastrow, 'CK', lastrow, 'withValues');
let current_row_paste = data[0]
    for(let x = 0; x < current_row_paste.length; x++){
      let k = x+11
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja2).getRange(lastrow, k).setValue(current_row_paste[x]);
    }
}

function preventDuplicate(){

  let idFormColumn = getSheet(hoja1, 'A',2,'A','','withValues');
  let idGeneralColum = getSheet(hoja2,'A',2,'A','','withValues');

  let duplicate = false;


  idFormColumn.forEach((formValues) => {
    idGeneralColum.forEach((generalValues) =>{
      let formString = formValues.toString();
      let generalString = generalValues.toString();
      if(formString === generalString){
        console.log('Found duplicate id');
        duplicate = true;
      }
    })
  })
  return duplicate;
}


/********** Envío de emails  **********/

function sendEmail(){

  /*
  * Obtiene los valores de la hoja de "Certificaciones Generales"
  */
  let sheetData = getSheet(hoja2,"A",2,"CT","","withValues");

  /*
  * Obtiene solo la hoja de "Certificaciones Generales"
  */
  let sheetOnlySheet = getSheet(hoja2,"A",2,"D","","onlySheet");

  /*
  * Declaramos la variable messageTemplate para el archivo HTML a utilizar
  */
  let messageTemplate = HtmlService.createTemplateFromFile('messageTemplate');

  /*
  * Definimos la variable date para escribir el día y horario en el que se envío em email
  */
  const date = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy - HH:mm:ss");

  /*
  * Declaramos variables con las posiciones de las columnas
  */ 
  // Columna de veredicto (B2)
  let verdict = 2;
  // Columna de entrevista (D2)
  let interviewDate = 3;
  // Columna de estado de email (F2)
  let emailStatus = 5;
  // Columna de envío de email (G2)
  let emailSendDate = 6;
  // Columna de nombre de la entidad (K2)
  let entityName = 10;
  // Columna de email de la entidad (W2)
  let entityEmail = 22;
  // Columna de tipo de entidad (Y2)
  let entityType = 24;

  // Almacenamos los datos obtenidos de sheetData en una variable
  let data = sheetData;

  for(let i = 0; i < data.length; i++ ){
      
      let values = data[i];

      /*
      * Declaramos una constante para los tipos de aprobación, su subject y su mensaje
      * En caso de no poseer un mensajes a modificar, se le debe asignar <null> a los campos <p>
      */
      const messageType = 
      [
        {
        type: "Entrevista",
        subject: `DEVA - ENTREVISTA - ${values[entityType]} - ${values[entityName]}`,
        message: {
          p1: null
        }
        },
        {
        type: "Aprobado A",
        subject: `DEVA - CERTIFICACIÓN APROBADA - ${values[entityType]} - ${values[entityName]}`,
        message: {
          p1: "Estructura mínima: Staff Administrativo/Staff Deportivo. (obligatorio).",
          p2: "Personería Jurídica. (obligatorio).",
          p3: "Registro Marca a nombre de la persona jurídica declarada. (obligatorio).",
          p4: "Contratos con Staff y Jugadores. (obligatorio).",
          p5: "Reglamento Interno.(Obligatorio).",
          p6: "Plan de Redes sociales: (obligatorio).",
          p7: "Logo propio en PNG/PDF fondo transparente con el nombre de la entidad",
          p8: "Desde este momento cuentan con 10 días hábiles para adjuntar la documentación correspondiente o solicitada."
          }
        },
        {
        type: "Aprobado B",
        subject: `DEVA - CERTIFICACIÓN APROBADA - ${values[entityType]} - ${values[entityName]}`,
        message: {
          p1: "Estructura mínima: Staff Administrativo/Staff Deportivo. (obligatorio).",
          p2: "Capacidad de facturar. (obligatorio). // Monotributo.",
          p3: "Registro de marca.",
          p4: "Reglamento Interno.(obligatorio).",
          p5: "Logo propio en PNG/PDF fondo transparente con el nombre de la entidad",
          p6: "Desde este momento cuentan con 10 días hábiles para adjuntar la documentación correspondiente o solicitada."
        }
        },
        { 
        type: "Aprobado C",
        subject: `DEVA - CERTIFICACIÓN APROBADA - ${values[entityType]} - ${values[entityName]}`,
        message: {
          p1: "Logo de tu entidad en formato .PNG/PDF, con fondo transparente, en la mejor resolución posible y con el nombre de tu entidad",
          p2: "Pueden ingresar a la siguiente página http://www.remove.bg para quitarle el fondo,",
          p3: "Desde este momento cuentan con 10 días hábiles para adjuntar la documentación correspondiente o solicitada."
        }
        },
          { 
        type: "Rechazado",
        subject: `DEVA - CERTIFICACIÓN RECHAZADA - ${values[entityType]} - ${values[entityName]}`,
        message: {
          p1: null
        }
        }
      ];

        /*
        * Recorremos el array con forEach
        */
        messageType.forEach((msgReplace) => {
          
          /*
          * Validamos que el veredicto coincida con nuestro array
          */
          if(values[verdict] === msgReplace.type){
            
            /*
            * Escribimos los datos en nuestra plantilla de HTML
            */
            messageTemplate.entityName = values[entityName];
            messageTemplate.type = msgReplace.type;
            messageTemplate.entityType = values[entityType];
            messageTemplate.interviewDate = values[interviewDate];
            messageTemplate.parrafo1 = msgReplace.message.p1;
            messageTemplate.parrafo2 = msgReplace.message.p2;
            messageTemplate.parrafo3 = msgReplace.message.p3;
            messageTemplate.parrafo4 = msgReplace.message.p4;
            messageTemplate.parrafo5 = msgReplace.message.p5;
            messageTemplate.parrafo6 = msgReplace.message.p6;
            messageTemplate.parrafo7 = msgReplace.message.p7;
            messageTemplate.parrafo8 = msgReplace.message.p8;

            /*
            * Validamos que el estado del email sea PENDIENTE
            */
            if(values[emailStatus] === "PENDIENTE"){
                try{
                  MailApp.sendEmail({
                          to: values[entityEmail],
                          subject: msgReplace.subject,
                          htmlBody: messageTemplate.evaluate().getContent(),
                  });
                /*
                * Actualizamos los valores en el sheet para evitar duplicaciones de envío
                */
               sheetOnlySheet.getRange(i + 2, emailStatus + 1).setValue("ENVIADO");
               sheetOnlySheet.getRange(i + 2, emailSendDate + 1).setValue(date);
          } catch(error){
            console.log("Program finish with errors");
            console.log("Email error ->" + error)
          }
        }
      }
    })
  };  
}