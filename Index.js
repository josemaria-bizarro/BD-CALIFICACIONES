// *************************************************************
//tabla de calificaciones

let bd=SpreadsheetApp.getActiveSpreadsheet()
let sheetBD=bd.getSheetByName("CALIFICACIONES")
let orriaBd=bd.getSheetByName("PORTADA")
let datosbd =sheetBD.getDataRange().getDisplayValues()
var ufilaIx = sheetBD.getLastRow()
var ufilaIndex=sheetBD.getRange(ufilaIx,1).getDisplayValue()


// **************************************************************
//concentrado de calificaciones

//diseño let bdDg=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1ihnDwqApaat4E6h2FdgVJibE2nVMqQ-y3oBkt93v1P0/')

//gastronomia let bdDg=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1BaU8lycurZpycvxOR-I2DZbNuTTo7SjOfP8CuSnKEfE/')

//comunicacion let bdDg=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/10FuRW_o7wh0G2tY_tiHUmXZ9ctRitTo0ga12CsVoRC0/')

//SAETI
  let bdDg=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1cyLa8JqTBLlwYBzYfy2x6y6eBHG3cwuO0t0wmUvpvB0/')

//LAE 
//let bdDg=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/17S7h8sUE7vxMpxNpkZeMbIuxS_ZIYulAZH9NtlZ2CGE/')


// ************************************************************************
// HOJA DATOS HISTORICOS

let sheetdbDg=bdDg.getSheetByName('PASO2')
//let datuakdbDg=sheetdbDg.getDataRange().getDisplayValues()
var lrow = sheetdbDg.getLastRow()-1
let datuakRango=sheetdbDg.getRange(2,1,lrow,23) // PASO2 datos
let datuakdbDg=datuakRango.getDisplayValues()



// catalogo alumnos
let bdikasleak=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1kzmgvrbkQ7ehUslRxR4ghkpSeI2yT99M3h-Dyo-EEVY/')
let sheetbdikasleak=bdikasleak.getSheetByName('ACTIVOS_FORMATEADO')
let datuakbdikasleak=sheetbdikasleak.getDataRange().getDisplayValues()

//catalogo ceec
let katalogo=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1T5KQCvXiYsGewSoc_A8qugUwSs-ft1NNds4oaNgQxkw/')
let sheetKatalogo=katalogo.getSheetByName('TABLA ASIGNATURAS')
let datuakKatalogo=sheetKatalogo.getDataRange().getDisplayValues()
let sheetPersonal=katalogo.getSheetByName('PERSONAL')
let datuakbdirakasle=sheetPersonal.getDataRange().getDisplayValues()
let sheetAukera=katalogo.getSheetByName('OPCIONES EDUCATIVAS')
let datuaAukera=sheetAukera.getDataRange().getDisplayValues()
let sheetAldia=katalogo.getSheetByName('PERIODOS EDUCATIVOS')
let datuaAldia=sheetAldia.getDataRange().getDisplayValues()
let sheetOpEdu=katalogo.getSheetByName('OPCIONES EDUCATIVAS')
let datuaOpEdu=sheetOpEdu.getDataRange().getDisplayValues()


// carpeta origen para mover listas

let jatorri=DriveApp.getFolderById('10kLiAvotzT2kgrouht4gAfyuEv7BY_D8')

// carpeta destino y para proceso de listas
//let norako=DriveApp.getFolderById('1deSpXPymOeU2j_lpWs9DCDJPA3Duy0Fl')  //LAD *************************
//let norako=DriveApp.getFolderById('10kLiAvotzT2kgrouht4gAfyuEv7BY_D8')

var datuakUrl =orriaBd.getRange("C1").getValue();    //URL de carpeta a procesar
var encuentra =datuakUrl.lastIndexOf("/")+1
var datuakId = datuakUrl.substr(encuentra,50)
let norako=DriveApp.getFolderById(datuakId)

// log de errores
let logDeErrores =SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rEXfk56euyX74jq2kRwZzsE86xEu_om9mNUIHWYna6Q/')
let sheetlog=logDeErrores.getSheetByName("log")
let ufilaLog =sheetlog.getLastRow()+1
let logerror =[]

//campos varios
let fechoy=(new Date())
let sheetDatos=""
let idAsigna=""
let varduplica=0
let data=[]





//*******************************************************************************\\

// MENU DE OPCIONES
function onOpen()
 {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('OPCIONES')
   
   menu
      .addItem("Carga listas", 'cargaDB')
      
  menu.addToUi();
}


// DESPLIEGA MENSAJE DE ERROR
function mensajeError(msg)
{
  var html=HtmlService.createHtmlOutput(msg)
  .setWidth(400)
  .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html,'ATENCIÓN!')
}

// ******************************************************************************************
// PROCESO DE CARGA DE LISTAS DE ASISTENCIA EN TABLA DE CALIFICACIONES
// ******************************************************************************************
function cargaDB() 
{

// extrae archivoS de la carpeta origen
  var filesIterator = norako.getFiles();   
  
  var file;
  var fileType;
  var ssID;
  var combinedData = [];
 
  
// procesa archivo por archivo
  while(filesIterator.hasNext())
  {
     file = filesIterator.next();
     fileType = file.getMimeType();

    
    if(fileType === "application/vnd.google-apps.spreadsheet")
    {
        ssID = file.getId();
                                                   //  verifica que no se haya cargado esta lista
        varduplica=0                           
        duplica(ssID)
                                                
        if (varduplica==0)                           // si no es duplicado continua
        {                                                   
          tomaDatos(ssID,data)
        }
       
    } //if ends here
  }//while ends here
  // escribe log de errores
  if(logerror.length>0)
  {
    sheetlog.getRange(ufilaLog,1,logerror.length,4).setValues(logerror)
    var msg ="Proceso concluido con incidencias, revisar log"
    mensajeError(msg)
  }

  var askenilara= sheetBD.getLastRow()+1
 
 if (data.length>0)
 { 
  sheetBD.getRange(askenilara,1,data.length,14).setValues(data)
 }
}



//function getDataFromSpreadsheet(ssID)
function tomaDatos(ssID,data)
{
 
  var ss = SpreadsheetApp.openById(ssID);
  let urlss =ss.getUrl()                                  //URL DE ARCHIVO DE LISTA 
  var ws = ss.getSheets()[0];
  let fechPeriodo=ws.getRange("BA3").getDisplayValue()      //FECHA PERIODO
  let wsParcial=ws.getRange("AH1").getDisplayValue()      
  var nomDocente =ws.getRange("F2").getDisplayValue()                 
  var wsPersonal = sheetPersonal.getDataRange().getDisplayValues()  //BUSCA CLAVE DE PROFESOR 
  var personalIlara = wsPersonal.filter(ilara=>(ilara[0]==nomDocente))
  let cveDocente                                                      //CLAVE DE DOCENTE
  if (personalIlara==0)
  {
    var msg ="error en id Asignatura Lista de asistencia"
    logerror.push([ssID,"error en docente",msg,fechoy])
    cveDocente="error "+ nomDocente
  } 
  else
  {
    cveDocente=personalIlara[0][1]
  }
  
  let creditos=0                                                    //CREDITOS DE MATERIAS

  var credTabla= datuakKatalogo.filter(ilara=> ilara[0] == idAsigna)
  if (credTabla.length > 0)
    {
      creditos = credTabla[0][17] 
    }
  else 
    {
      creditos =0
    }
  


  let finPeriod                                                   //FECHA DE FIN DE PERIODO
  
  var fecFinPeriodo =datuaAldia.filter(ilara =>ilara[0]==fechPeriodo)
 
  
  if(fecFinPeriodo.length>0)
  {
     finPeriod = fecFinPeriodo[0][5]
  }
  else
  {
    finPeriod= fechPeriodo
  }

 let fecha=sheetDatos.getRange("AH4").getDisplayValue()  //FECHA DE FINAL DEL PARCIAL

 let parcial=0                                           //PARCIAL EN NUMERO

 switch(wsParcial)
 {
  case "PRIMERO":
      parcial=1;
      break;
  case "SEGUNDO":
      parcial=2;
      break;
  case "TERCERO":
      parcial=3;
      break;
  case "FINAL":
      parcial=4;
      break;
  default:
      parcial=4
      break;
 }


                                                  // *********************************************************
                                                  // empieza a trabajar con los alumnos de la lista

 var datosLista = ws.getRange("C10:CG" + ws.getLastRow()).getValues();
 let tipo=""
 let calif=0

 datosLista.forEach(ilara=>                 //RECORRE LOS DATOS DE ALUMNOS EN LA LISTA
 {
    if (ilara[1]!=="")
    {

      var nombre =ilara[0]
                                                                       // BUSCA DATOS DE ALUMNO
      var ikasleDatua = datuakbdikasleak.filter(fila => fila[0] == nombre);

      
      if (ikasleDatua.length > 0)                                                 //Si alumno no en catalogo = baja
      {
        
        // obtiene datos de alumno

          var correoIns = ikasleDatua[0][1];  //cuenta institucional
          var taldeI = ikasleDatua[0][5]      // grupo activo
          calif = ilara[80]          // calificacion
          if (calif=="")
          {
            calif=5
          }
          var opcEduIk = ikasleDatua[0][9]    //opc edu del alumno
          var opcEduBd = datuaOpEdu.filter(fila => fila[0] == opcEduIk)   //busca opc edu en catalogo
          if (opcEduBd.length > 0)
          {
            var opcEdu = opcEduBd[0][8]
          }
          else 
          {
            var opcEdu = opcEduIk
          }
                      
            var existe  =datosbd.filter(fila => fila[1] == correoIns && fila[6]==idAsigna && fila[9]==fecha)
            var globalA =datosbd.filter(fila => fila[1] == correoIns && fila[6]==idAsigna && fila[9]!=fecha)

            console.log(fecha)
            
            // let existe =datosbd.filter(fila => fila[6] == fecha)
                
              if (existe.length>0)
                {
                  var msg =correoIns+" "+idAsigna+" "+fecha
                  logerror.push([ssID,"duplicado",msg,fechoy])
                }
              else
                {
                  
                  if(globalA.length>0)
                    {
                      tipo="gbl"
                    }
                  else
                    {
                      tipo="cur"
                    } 
                    ufilaIndex++
                    data.push([ufilaIndex,correoIns,taldeI,opcEdu,fechPeriodo,parcial,idAsigna,calif,tipo,fecha,cveDocente,urlss,creditos,finPeriod])
                }
      }
    }
 })
 
 return data;

}



// **************************************************************************************************
// VERIFICA SI LOS DATOS SON DUPLICADOS o es una corrección de calificación o es una adición a la BD

function duplica(ssID)
{
                                                          // Lista de asistencia con la que va a trabajar
   let ssdatos=SpreadsheetApp.openById(ssID)
   sheetDatos=ssdatos.getSheetByName("ORIGEN")
   let ufila=sheetDatos.getRange("B10:B").getLastRow()
   let datos=sheetDatos.getRange(10,3,ufila,83).getDisplayValues() //Cuerpo de la lista alumnos,calificaciones

   // CVE ASIGNATURA UNIFICADA 
   idAsigna=sheetDatos.getRange("D8").getDisplayValue() //si es de 5 caracteres, busca el equivalente 
  
   if (idAsigna.length!==4)
    {                         //proceso clave de asignatura antigua

        // DATOS tabla asignatura

        var parteCve= idAsigna.slice(0,3)
      
        switch(parteCve)
        {
          case "BTG":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[11]);
              break;
          case "DPR":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[11]);
              break;
          case "BTC":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[12]);
              break;
          case "CDI":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[12]);
              break;
          case "BTD":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[13]);
              break;
          case "BDW":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[13]);
              break;
          case "BTA":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[15]);
              break;
          case "LAD":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[16]);
              break;
          case "DAD":
              var cveUni= datuakKatalogo.filter(ilara=> idAsigna == ilara[16]);
              break;
          default:
            var cveUni=[];
            
        } //fin switch

        if (cveUni.length==0)
          {
            // escribe log de errores
            var msg ="error en id Asignatura Lista de asistencia"
            logerror.push([ssID,"valida id asignatura",msg,fechoy])
          
            // despliega mensaje avisando que hay un error
            mensajeError(msg)
    
          }
        else
          {
            idAsigna = cveUni[0][0]    
          }
        
    } // fin if
   

   let fecha=sheetDatos.getRange("AH4").getDisplayValue()

   
// datosbd = DB calificaciones Verifica que no exista la URL
   var urlOrigen =ssdatos.getUrl()
   let existeUrl =datosbd.filter(fila => fila[11] == urlOrigen) 

    //let existe =datosbd.filter(fila => fila[0] == nombre && fila[3]==idAsigna && fila[6]==fecha)
   // let existe =datosbd.filter(fila => fila[6] == fecha)
     
    if (existeUrl.length > 0)  //verifica si la url ya existe en DB
      {
        datos.forEach(fila=>
        {   //valida que existan los alumnos en la lista

          if (fila[0]!=="")
          {
            let existeikasle=datosbd.filter(ila=>fila[2]==ila[1] && urlOrigen==ila[11] )

            if (existeikasle.length>0)
            {
                //  REVISA QUE LA CALIFICACION SEA DIFERENTE 
                var ilaraAct = (existeikasle[0][0])
               
                var ilaraIx=0;
                //busca en que posicion se encuentra el registro
                for(var ix=0;ix<ufilaIx;ix++)
                {
                  if (datosbd[ix][0]==ilaraAct)
                    {
                      ilaraIx=ix+1
                      ix=ufilaIx
                    }
                }      
                      
               var califant= existeikasle[0][7] 

               let existeCalif=datosbd.filter(ila=>fila[2]==ila[1] && urlOrigen==ila[11] && fila[80]==ila[7] )
               
               if(existeCalif.length ==0) // si calificacion no es igual a la BD
               {
                //modifica la calificacion
                 sheetBD.getRange(ilaraIx,8).setValue(fila[80])
                 varduplica =1
                 
               }
            }
          }
        
        })
        
      }
      else   //  NO EXISTE LA URL EN BD
        {
          var ok=""
         
        }
}  // FIN FUNCION DUPLICA



//************************************************************

// Mueve archivos de una carpeta a otra
function mugitu()
{
  var file;
  var origen = jatorri.getFiles()
  while(origen.hasNext())
  {
    file=origen.next();
    file.moveTo(norako);
  }
}

// sustituye el grupo con subgrupo por el grupo en bd alumno
function taldea()
{
  var arr=[]
 datosbd.forEach(ilara=>
    {
      var taldeAct = ilara[1]
      
      try
      {
          var ikasle=datuakbdikasleak.filter(ila=> ilara[0]==ila[1])
          
          var talde=ikasle[0][5]
          arr.push([taldeAct,talde])
      }
      catch(e)
      {
          var msg ="clave de grupo "+(ilara[0])
          logerror.push([ssID,"error en clave de grupo",msg,fechoy])
      }
    }) 
    sheetBD.getRange(2,15,arr.length,2).setValues(arr)
}



//sustituye id prof por correo
function idprof()
{
  let arrayDatos=[]
// actuaiza asignatura en historico
datuakdbDg.forEach(ilara=>
{
  var clave=datuakbdirakasle.filter(ila => ilara[19]==ila[0])
  
  try
  {
    var idAsignaBerri= clave[0][1]
 
     arrayDatos.push([ilara[19],idAsignaBerri])
  
  }
  catch(err)
  {
       var msg ="clave de docente "+(ilara[19])
        logerror.push([ssID,"error en clave de docente",msg,fechoy])
     arrayDatos.push([ilara[19],"error"])
  }
  
  
})

let datuakIrteera=sheetdbDg.getRange(2,24,arrayDatos.length,2)
datuakIrteera.setValues(arrayDatos)
}


//sustituye id asignatura antiguo por nuevo
function asignatura()
{
  let arrayDatos=[]
// actuaiza asignatura en historico
datuakdbDg.forEach(ilara=>
{
  var clave=datuakKatalogo.filter(ila => ilara[6]==ila[16])
  
  try
  {
    var idAsignaBerri= clave[0][0]
   
  
  arrayDatos.push([ilara[6],idAsignaBerri])
  
  }
  catch(err)
  {
    var msg=err.message
   
     //logerror.push([ilara[0],"carga de BD calificaciones",msg])
     arrayDatos.push([ilara[6],"error"])
  }
  
  
})

let datuakIrteera=sheetdbDg.getRange(2,22,arrayDatos.length,2)
datuakIrteera.setValues(arrayDatos)
}



function aukera()
{
  
  
  datos.forEach(row =>
  {
    
      let nombre =row[2]
      let idAsigna=sheetDatos.getRange("D8").getDisplayValue()
      let fecha=sheetDatos.getRange("AH4").getDisplayValue()


    let existe =datosbd.filter(fila => fila[0] == nombre && fila[3]==idAsigna && fila[6]==fecha)
   // let existe =datosbd.filter(fila => fila[6] == fecha)
       console.log(fecha+""+nombre+""+idAsigna)
      
    if (existe.length>0)
        {console.log("duplicado")
        
        }else
        {
          console.log("NO DUPLICADO")
         
        }
  })
}




// CARGA BDs ANTERIORES

function carga()
{
//let logerror =[]
let arrayDatos =[]

//obtiene correo de catalog alumnos y obtiene datos de historico
datuakdbDg.forEach(ilara=>
{
  try
  {
  var aukCve=datuaAukera.filter(cve=> ilara[2]==cve[0])
  var aukeraCve=aukCve[0][8]
  var esta=datuakbdikasleak.filter(ila => ilara[0]==ila[0])
  
  }
  catch(e)
  {
    console.log(ilara)
  }
  try
  {
    var posta= esta[0][1]
   
  arrayDatos.push([posta,ilara[16],aukeraCve,ilara[4],ilara[5],ilara[22],ilara[7],ilara[10],ilara[11],ilara[12],ilara[13],ilara[14],ilara[18]])
  
  }
  catch(err)
  {
    var msg=err.message
     logerror.push([ilara[0],"carga de BD calificaciones",msg,fechoy])
    
  }
 
            

})


var ufila= sheetBD.getLastRow()+1
var totfila =arrayDatos.length


sheetBD.getRange(ufila,1,totfila,13).setValues(arrayDatos)
sheetlog.getRange(ufilaLog,1,logerror.length,4).setValues(logerror)

}

//ESTA ES LA FUNCION PARA OBTENER LAS CALIFICACIONES DE UNA CARPETA DEL PERIODO
