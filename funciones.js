const bdCeec = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1T5KQCvXiYsGewSoc_A8qugUwSs-ft1NNds4oaNgQxkw/") //catalogo ceec bueno
let cveOpcEdu =0;
let idOpEd="";
let periodoW="";
let docentew="";
let correodocente="";
let nomDocenteMin="";
let aulasW=[];
let nombreLegal="";
let nombreInterno="";
let ponderaExamen=0;
let ponderaEvaluacion=0;
let numPeriodo="";
let asignaOrria="";
let asignaPaso="";
let opcEduca="";
let wparcial="";
let wname="";
let wkaldea="";
let txanda="";
let archSalida="";
let alumnosLista=[];
let wsalida=[];
let cveAula="";
let fecIniPer="";
let fecFinPer="";
let docenteNom="";
let nomClase="";



function doGet()
{
    const html = HtmlService.createTemplateFromFile('Main');
    const irteeran = html.evaluate();
    return irteeran;
}

function include(filename)
{
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function obtieneOpcEdu()                // OBTIENE DATOS DE OPC EDU DE DB
{
    const opcEduOrria=bdCeec.getSheetByName("OPCIONES EDUCATIVAS");
    datuaOpcEdu=opcEduOrria.getDataRange().getDisplayValues();
    datuaOpcEdu.shift();
    return datuaOpcEdu;
}

function obtieneCicloEsc(wsOpcEdu)          //OBTIENE DATOS DE CICLO ESC DE DB
{
    const cicloEscOrria=bdCeec.getSheetByName("PERIODOS EDUCATIVOS");
    var azkenIlara= cicloEscOrria.getLastRow();
    var cicloPaso=cicloEscOrria.getRange(3,1,azkenIlara,3).getDisplayValues();
    const datuaCicloEsc=cicloPaso.filter(ilara=>ilara[1]!=""&&ilara[2]==wsOpcEdu);
    opcEduW=wsOpcEdu;
    if (datuaCicloEsc.length>0)
    {
        return datuaCicloEsc;
    }
    else    
    {
        console.log("filtro vacio ciclo escolar");
    }
    
}

function obtieneAsignatura(wsOpcEdu)                    // OBTIENE DATOS DE ASIGATURA DE DB
{
    asignaOrria=bdCeec.getSheetByName("TABLA ASIGNATURAS");
    asignaPaso=asignaOrria.getDataRange().getDisplayValues();
    asignaPaso.shift()
    datuaAsigna=asignaPaso.filter(ilara=>ilara[3]==wsOpcEdu);
    if(datuaAsigna.length>0)
    {
        return datuaAsigna;
        
    }
    else
    {
        console.log("filtro asignatura vacio");
    }
}

function obtieneDocente()                           //OBTIENE DATOS DE DOCENTE DE DB
{
    const irakasleOrria=bdCeec.getSheetByName("PERSONAL");
    const irakaslePaso=irakasleOrria.getDataRange().getDisplayValues();
    irakaslePaso.shift();
    const datuaDocente=irakaslePaso.filter(ilara=> ilara[2]=="ACTIVO"&&ilara[5]=="DOCENTE");
    if (datuaDocente.length>0)
    {
        return datuaDocente;
       
    }
    else 
    {
        console.log("filtro de Docente vacío");
    }
}

function obtieneGrupo(wsOpcEdu)                                     //OBTIENE DATOS DE GRUPOS DE DB
{
    const taldeaOrria=bdCeec.getSheetByName("GRUPOS");
    const taldeaPaso=taldeaOrria.getDataRange().getDisplayValues();
    taldeaPaso.shift();
    const datuaGrupoPaso=taldeaPaso.filter(ilara=>ilara[0]==wsOpcEdu&&ilara[7]!="NO")
    if (datuaGrupoPaso.length>0)
    {
        let datuaGrupo=[];
        var taldeAnt=0;
        for(var i=0;i<datuaGrupoPaso.length;i++)
        {
            //if (datuaGrupoPaso[i][1]!=taldeAnt)
            //{
                var grupoS=datuaGrupoPaso[i][1]+datuaGrupoPaso[i][2]
                datuaGrupo.push([grupoS]);
            //    taldeAnt=datuaGrupoPaso[i][1];
            //}
        }
        return datuaGrupo
    }
    else
    {
        console.log("filtro de Grupo Vacío")
    }
    
}

function grabaRegistroAula(formSalida)                      //GENERA DATOS DE PARA NOMBRE LISTA ************
{
    let werror=0;
    

                                                                // OBTIENE CLAVE DE OPCION EDUCATIVA
    var datuakOpcedu=bdOpcEdu.getDataRange().getDisplayValues();
    var datuakFiltro=datuakOpcedu.filter(ilara=>ilara[0]==formSalida[0])
    if(datuakFiltro.length>0)
    {
        idOpEd=datuakFiltro[0][2];                           //CLAVE DE OPC EDU
        cveOpcEdu=datuakFiltro[0][8];                        //NUMERO DE OPC EDU
        
    }
    else
    {
        console.log("error al filtrar Opc Edu");
  
      werror=1;
      return werror;
    }
                                                                //OBTIENE DATOS PERIODO
    var periodoOrria=bdPeriodo.getDataRange().getDisplayValues();
    var datuaPeriodo=periodoOrria.filter(ilara=> ilara[0]==formSalida[2]&& ilara[2]==formSalida[0]);
    if (datuaPeriodo.length>0)
    {
        
        fecIniPer=datuaPeriodo[0][6];
        fecFinPer=datuaPeriodo[0][7];
        periodoW=formSalida[2]; 
        werror= seleccionaAulas();
        if (werror==1)
        {
        return werror;
        }
    }
    else 
    {
        console.log("filtro de Periodo vacío formsalida: "+formSalida[0]);
        werror=1
        return werror;
    }

    werror= procesoAula();
        if (werror==1)
        {
        return werror;
        }
}


function seleccionaAulas(formSalida)                                 //PROCESO DE IMPRESION***************************
{
/*
OBTIENE LAS CLASES EXISTENTES
FILTRA POR 
    OPC EDUCATIVA - TRES PRIMERAS LETRAS DE CAMPO - description == idOpEd
    ACTIVOS EN CAMPO - courseState == "ACTIVOS"
    PERIODO EN CAMPO - Room == periodoW
    aulasW. push() //contiene los datos de las aulas seleccionadas
*/

werror=0;
opcEduca=formSalida[0];
var cursosW=Classroom.Courses.list().courses; //obtiene todos los cursos

var periodoW=formSalida[2];
wparcial=formSalida[1];

datuaOpcEdu=obtieneOpcEdu()
// OBTIENE LA CLAVE DE LA OPC EDU
var datuaOpcEduFil =datuaOpcEdu.filter(ilara=>ilara[0]==formSalida[0]);

if (datuaOpcEduFil.length >0)
{
    idOpEd = datuaOpcEduFil[0][2];
}
else
{
    console.log("error en clave de opc edu")
    werror=1
    return werror;
}


try{                    //filtra los cursos por activos, ciclo y clave de opc edu
       // var cursosFil=cursosW.filter(ilara=>ilara.courseState == "ACTIVE" && ilara.room==periodoW&&ilara.description.substring(0,3)==idOpEd)
        var cursosFil=cursosW.filter(ilara=>ilara.courseState == "ACTIVE" && ilara.room==periodoW)
        //var cursosFil=cursosW.filter(ilara=>ilara.section=="ENE - MAY 2024" )
        if (cursosFil.length>0)
        //if(ilara.description.substring(0,3)==idOpEd)
        {   /*DOCENTE
            OBTIENE LA CLAVE DEL DOCENTE DEL CAMPO - name CUATRO PRIMERAS POSICIONES
            GUARDA LA CLAVE DEL DOCENTE PARA USARLA EN ALIAS DEL NOMBRE DEL ARCHIVO
            OBTIENE DE CATALOGO EL NOMBRE EN MINUSCULAS USANDO LA CLAVE DEL DOCENTE */
            
            cursosFil.forEach(ilara=>
            {
            wname=ilara.name;
            wkaldea=ilara.section;
            
            docentew = ilara.name.substring(0,4);
            
            bdDocenteW=bdDocente.getDataRange().getDisplayValues();
            bdDocenteWF=bdDocenteW.filter(ilara1=>ilara1[4]==docentew&&ilara1[2]=="ACTIVO")
            
            if(bdDocenteWF.length>0)
            {
                nomDocenteMin=bdDocenteWF[0][0];
                correodocente=bdDocenteWF[0][1];
                aulasW.push([ilara.name,ilara.id,docentew,nomDocenteMin,wkaldea,correodocente])
                
            }
            else
            {
                console.log("error en clave docente")
                werror=1
                return
            }
            
            })
            
        }
        else
        {
            console.log("ERROR cursosFil.description.substring(0,3) "+ cursosFil.description);
            console.log("Periodo: "+periodoW);
            console.log("Opc Edu: "+idOpEd);
            console.log(cursosW);
        }
    }
    catch(err)
    {
        console.log("Error al obtener todos los cursos... Error: "+err+"  Cursosw: "+cursosFil)
        werror=1;
        return werror;
    }
//})
//console.log(aulasW)
werror = procesoAula(aulasW,formSalida);
return werror;
}

function procesoAula(aulasW,formSalida)
{
/*
    PROCESO POR AULA
      OBTIENE LOS SIGUIENTES DATOS:
        ASIGNATURA
            OBTIENE LA CLAVE DE LA ASIGNATURA DEL CAMPO - name ENTRE LOS PARENTESIS. CUATRO POSICIONES
            GUARDA CLAVE ASIGATURA PARA USAR EN ARCHIVO SALIDA
            CON LA CLAVE BUSCA EN CATALOGO 
                LA DESCRIPCION DE LA ASIGNATURA LEGAL E INTERNA
                EL PORCENTAJE PONDERADO
                TIPO DE PERIODO (CUATRI, SEM, TRI)
            NUMERO DE PARCIAL EN LETRA PARA ARCHIVO DE SALIDA
        GRUPOS
            OBTIENE DE CAMPO - section LOS GRUPOS PARA USARLOS EN ARCHIVO DE SALIDA TANTO EN NOMBRE COMO EN LISTA
            BUSCA TURNO EN CATALOGO
                SI BTA = ONLINE
                FILTRAR CATALOGO
                    OPC EDU
                    BUSCAR EN REGISTROS OBTENIDOS PERIODO GUARDADNO SIEMPRE EL ANTERIOR JUNTO CON EL TURNO
                        SI ENCUENTRA = OBTIENE TURNO
                        SI NO ENCUENTRA, USA EL ULTIMO PERIODO Y OBTIENE TURNO
        ARMA NOMBRE DE ARCHIVO DE SALIDA
            APJU-Lengua Adicional al Español I.-BTA1O23301B -PRIMERO
    OBTIENE ALUMNOS
            INVITADOS
            ALUMNOS
            OBTIENE:
                NOMBRE
                GRUPO
                INSTITUCIONAL
            ORDENAR AZ
            SI RECURSADOR
                ESCRIBE EN TABLA RECURSADORES
*/
var lastindex=0;
werror=0;
warro=1


aulasW.forEach(ilara=>          
    //AULAS SEGUN OPC EDU Y PERIODO aulasW.push([ilara.name,ilara.id,docentew,nomDocenteMin,wkaldea])
    {
        /*
        ASIGNATURA
            GUARDA CLAVE ASIGATURA PARA USAR EN ARCHIVO SALIDA= wcveAula
            CON LA CLAVE wcveAula BUSCA EN CATALOGO LA DESCRIPCION DE LA ASIGNATURA LEGAL wlegal E INTERNA winterna
                EL PORCENTAJE PONDERADO wponderaEx wponderaCon
                TIPO DE PERIODO (CUATRI, SEM, TRI) wperiodo
                NUMERO DE PARCIAL EN LETRA PARA ARCHIVO DE SALIDA wparcial */        

        var windex =ilara[0].lastIndexOf("(")+1 ;
        var wcveAula = ilara[0].substring(windex,windex+4); //OBTIENE LA CLAVE DE LA ASIGNATURA DEL CAMPO - name ENTRE LOS PARENTESIS. CUATRO POSICIONES
        var widAula=ilara[1];                             //ID DE AULA
        var wCvedocente=ilara[2];                         //CVE DOCENTE
        var wdocente=ilara[3];                            //NOMBRE DOCENTE
        var wOpEd=idOpEd;                                 //cve opc educa

        
        //Obtiene datos de asignatura
        var asigna =bdAsigna.getDataRange().getDisplayValues();
        asigna.shift();
        var asignaF = asigna.filter(ilara=>ilara[0]==wcveAula&&ilara[3]==opcEduca)
        if (asignaF.length>0)
        {
         var wlegal =asignaF[0][1];             //NOMBRE LEGAL
         var winterna = asignaF[0][2];          //NOMBRE INTERNO
         var wperiodo = asignaF[0][8];          //PERIODO
         var wponderaEx = asignaF[0][9];        //PONDERACION EXAMEN
         var wponderaCon = asignaF[0][10];      //PONDERACION CON
                  
        }   
        else
        {
            console.log("error en asignatura:  "+wcveAula+"  "+opcEduca);
            werror=1;
        }             

        /*
        GRUPOS
            OBTIENE DE CAMPO - section LOS GRUPOS PARA USARLOS EN ARCHIVO DE SALIDA TANTO EN NOMBRE COMO EN LISTA
            
            BUSCA TURNO EN CATALOGO
                SI BTA = ONLINE
                FILTRAR CATALOGO
                    OPC EDU
                    BUSCAR EN REGISTROS OBTENIDOS PERIODO GUARDADNO SIEMPRE EL ANTERIOR JUNTO CON EL TURNO
                        SI ENCUENTRA = OBTIENE TURNO
                        SI NO ENCUENTRA, USA EL ULTIMO PERIODO Y OBTIENE TURNO */

        wkaldea=ilara[4];   //GRUPOS      
        var windex =wkaldea.lastIndexOf(" ");
        if (windex<0)
        {
            windex=4;
        }

        if (windex>4)
        {
            windex=4;
        }
            
        var kaldeaTurnoW=wkaldea.substring(0,windex); //primer grupo del string para buscar turno
        
        if (kaldeaTurnoW=="ESPE")
        {
            if(wOpEd==4)
            {
                txanda="ONLINE";
            }
            else
            {
                txanda="MATUTINO";
            }
        }
        else
        {
            var wkal =kaldeaTurnoW.replaceAll("-","");  
            var wkal2 =wkal.replaceAll(" ","");  //grupo sin espacios ni guiones
            var wkal3 =wkal2.replaceAll("A",""); //remueve literales
            var wkal4 =wkal3.replaceAll("B",""); //remueve literales
            var wkal5 =wkal4.replaceAll("C",""); //remueve literales
            var wkal6 =wkal5.replaceAll("D",""); //remueve literales
            var kaldeaTurno =wkal6.replaceAll("E",""); //remueve literales
            
            //aqui me quede
            // va a buscar el turno
            var wtxanda=bdgrupo.getDataRange().getDisplayValues();
            wtxanda.shift();
            var wtxandaF=wtxanda.filter(ilara=> ilara[1]==kaldeaTurno);
            if (wtxandaF.length>0)
            {
                txanda=wtxandaF[0][4];
                var txandaIni=txanda.substring(0,1);
                
            }
            else
            {
                console.log("error al buscar grupo / turno "+kaldeaTurno);
                werror=1;
            }
        }
        /** OBTIENE EL PERIODO SEGUN LA ASIGNATURA         */
        var datAsigna=bdAsigna.getDataRange().getDisplayValues();
        var datAsignaF=datAsigna.filter(ilara=> ilara[0]==wcveAula);
        if(datAsignaF.length>0)
        {
            var peridoNum=datAsignaF[0][20];

        }
        else
        {
            werror=1;
            console.log("Error al buscar el periodo de la asignatura");
            return werror;
        }
        /*ARMA NOMBRE DE ARCHIVO DE SALIDA
            APJU-Lengua Adicional al Español I.-BTA1O23301B(OPCEDU QUATRI TURNO YY GPOS) -PRIMERO */
            var wkaldeaM=wkaldea.replaceAll("-","");
            var wkaldeaN=wkaldeaM.replaceAll(" ","");
            var YYw=new Date().getFullYear();
            var YYn=Number(YYw);
            var YY=YYn-2000;
            
            
        //                  ARMA EL NOMBRE DEL ARCHIVO DE SALIDA.
        if(wkaldeaN=="ESPECIAL")
        {
            archSalida=wCvedocente+"-"+wlegal+"-"+wkaldeaN+"-"+"ESPECIAL";
        }
        else
        {
                archSalida=wCvedocente+"-"+wlegal+"-"+idOpEd+peridoNum+txandaIni+YY+wkaldeaN+"-"+wparcial;
                 
        }
        
        /*OBTIENE ALUMNOS DEL AULA: INVITADOS Y ACTIVOS
        
        
        OBTIENE ALUMNOS
            INVITADOS
            ALUMNOS
            OBTIENE:
                NOMBRE
                GRUPO
                INSTITUCIONAL
            ORDENAR AZ
            SI RECURSADOR
                ESCRIBE EN TABLA RECURSADORES 
            widAula 
                */
        var optionalArgs = 
            {
              courseId:widAula
            };
        try
        {
            var alumnosAula=Classroom.Courses.Students.list(widAula); //OBTIENE ALUMNOS DEL AULA
            
            alumnosAula.students.forEach(ilara=>
            {
                var alumid=Classroom.UserProfiles.get(ilara.userId);
                alumnosLista.push(alumid.emailAddress)  
            })
        
        
        }
        catch(err)
        {
            console.log("no han entrado alumnos en aula: "+widAula)
        }

        try
        {
            var alumnosInvit=Classroom.Invitations.list(optionalArgs).invitations;     //OBTIENE INVITADOS DE AULA
            var alumnosInvitF=alumnosInvit.filter(ilara=> ilara.role=="STUDENT")
            if (alumnosInvitF.length>0)
            {
                alumnosInvitF.forEach(row=>
                {
                    var cveinvitado=row.userId;
                    var alumid=Classroom.UserProfiles.get(cveinvitado);
                    alumnosLista.push(alumid.emailAddress)
                })
            }
            else
            {
                werror=1;
                console.log("error al buscar invitados")
            }
        }
        catch(err)
        {
            console.log("no hay invitados en aula: "+widAula)
        }

        var periodoOrria=bdPeriodo.getDataRange().getDisplayValues();
        var datuaPeriodo=periodoOrria.filter(ilara=> ilara[0]==formSalida[2]&& ilara[2]==formSalida[0]);
        if (datuaPeriodo.length>0)
        {
            
            fecIniPer=datuaPeriodo[0][6];
            fecFinPer=datuaPeriodo[0][7];
            periodoW=formSalida[2]; 
        }
        else 
        {
            console.log("filtro de Periodo vacío formsalida: "+formSalida[0]);
            werror=1
            return werror;
        }
        
         //CVE ASIGNATURA,OPC EDU,NOMBRE LEGAL,NOMBRE INTERNO,CUATRI,PONDERACION EXAMEN,PONDERACION CON,NOMBRE ARCHIVO,DOCENTE,CORREO DOCENTE,PERIODO,FECHA INICIAL,FECHA FINAL,PARCIAL,GRUPOS,TURNO,LISTA ALUMNOS
        werror=creaLista(wcveAula,idOpEd,wlegal,winterna,wperiodo,wponderaEx,wponderaCon,archSalida,wdocente,correodocente,periodoW,fecIniPer,fecFinPer,wparcial,wkaldeaM,txanda,alumnosLista);  
        
        alumnosLista=[];  
        /*
        PREPARA DATOS PARA ARMAR LISTA Y CREA LISTA 
        }*/
    })
  

return werror;
}

/***************************************************************************************************************** */
/*                           PROCESO DE ARMADO DE LISTA DE IMPRESION                                               */
/***************************************************************************************************************** */

//CVE ASIGNATURA,OPC EDU,NOMBRE LEGAL,NOMBRE INTERNO,CUATRI,PONDERACION EXAMEN,PONDERACION CONT,NOMBRE ARCHIVO,DOCENTE,CORREO DOCENTE,PERIODO,FECHA INICIAL,FECHA FINAL,PARCIAL,GRUPOS,TURNO,LISTA ALUMNOS
function creaLista(wcveAula,idOpEd,wlegal,winterna,wperiodo,wponderaEx,wponderaCon,archSalida,wdocente,correodocente,periodoW,fecIniPer,fecFinPer,wparcial,wkaldeaM,txanda,alumnosLista)
{
werror=0;
/*
OBTENER CARPETAS DE DESTINO (CONCENTRADOS)IDFolderDestino
OBTENER PLANTILLA SEGUN OPC EDU
*/
var opceduOrriaW=bdOpcEdu.getDataRange().getDisplayValues()
var opceduIlara=opceduOrriaW.filter(ilara=> ilara[2]==idOpEd)
if (opceduIlara.length>0)
{
    var IDFolderDestino=opceduIlara[0][14];  //url folder destino Contenedor
    
    var url_plantilla=opceduIlara[0][10];      //url plantilla OrigenX
    var encuentra = url_plantilla.lastIndexOf("/")+1
    var IDPlantilla =url_plantilla.substr(encuentra,50)

}
else
{
    console.log("Error al acceder a datos de OPC EDU");
    werror=1;
    return werror;
}


// FECHA DE ELABORACIÓN *********************
var CurDate = Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy"); 
//var CurDateB = Utilities.formatDate(new Date(), "GMT-6", "yyyy");         //FECHA CORTA





//CREA NUEVO DOCUMENTO ****************************
const SheetFile = DriveApp.getFileById(IDPlantilla);
const FolderConte = DriveApp.getFolderById(IDFolderDestino);
const ListCopia = SheetFile.makeCopy(FolderConte);
var Idlis = ListCopia.getId();

//ABRE EL NUEVO DOCUMENTO Y JALA LA HOJA *******************
var Lista = SpreadsheetApp.openById(Idlis);
var Hoja = Lista.getSheetByName("ORIGEN");
var PLANT ="CAFETALES";

// ********************************************************************************
//   ARMA LISTA DE ASISTENCIA
                                    /**** DATOS DE LA HOJA DE CALCULO   
                 CVE ASIGNATURA,OPC EDU,NOMBRE LEGAL,NOMBRE INTERNO,CUATRI,PONDERACION EXAMEN,PONDERACION EVA,NOMBRE ARCHIVO,DOCENTE,PERIODO,FECHA INICIAL,FECHA FINAL,PARCIAL,GRUPOS,TURNO,LISTA ALUMNOS
function creaLista(wcveAula,idOpEd,wlegal,winterna,wperiodo,wponderaEx,wponderaCon,archSalida,wdocente,periodoW,fecIniPer,fecFinPer,wparcial,wkaldeaM,txanda,alumnosLista)                 
                                    */
                                    
Hoja.getRange("B6").setValue(wlegal);  //ASIGNATURA LEGAL
Hoja.getRange("B8").setValue(winterna);  //ASIGNATURA INTERNA
Hoja.getRange("D8").setValue(wcveAula);  //CVE ASIGNATURA
Hoja.getRange("F2").setValue(wdocente);  // DOCENTE
Hoja.getRange("AH1").setValue(wparcial); //PERIODO
Hoja.getRange("BA1").setValue(PLANT); //PLANTEL
Hoja.getRange("BU1").setValue(txanda); //TURNO
Hoja.getRange("AH3").setValue(fecIniPer); //INICIO PERIODO
Hoja.getRange("AH4").setValue(fecFinPer); //FIN PERIODO
Hoja.getRange("BA3").setValue(periodoW); //CICLO ESCOLAR 
Hoja.getRange("AV4").setValue(wkaldeaM); //GRUPOS
Hoja.getRange("BL3").setValue(wperiodo); //CUATRIMESTE
Hoja.getRange("BY5").setValue(wponderaCon); //PONDERACIÓN EVALUACION
Hoja.getRange("CA5").setValue(wponderaEx); //PONDERACIÓN EXAMEN

Hoja.getRange("CA1").setValue("Creación " + CurDate); //FECHA DE CREACIÓN


var Alumnos1 =[];
let Alumnos2 =[]
var datuaIkasle= bdikasleOrria.getDataRange().getDisplayValues();
datuaIkasle.shift();

              //PREPARAR HASTA 30 ALUMNOS ***********************

    for (var i=0; i<alumnosLista.length; i++)
    {
        //OBTIENE DATOS DE CATALOGO ALUMNOS
        var ikasleF= datuaIkasle.filter(ilara=>alumnosLista[i]==ilara[1]);
        if (ikasleF.length>0)
        {
            var wgpo=ikasleF[0][5]+ikasleF[0][6];
            Alumnos1.push([ikasleF[0][0],wgpo,alumnosLista[i]]) //NOMBRE+ GRUPO+CORREO INSTITUCIONAL
            //ORDENA DATOS ALUMNOS POR NOMBRE
            Alumnos1.sort();
        }
    }
 
    try
    {
        
        //************************************************************ */

        //DETERMINAR SI LA CANTIDAD DE ALUMNOS REBASA LOS 30 ELEMENTOS PARA HACER UNA SEGUNDA HOJA ***********
        if(alumnosLista.length>30)
        {
            /*SEGUNDA PAGINA ****************
            determinar la cantidad de alumnos en parte1 y en parte 2 */
            var TotalAlum1 = 30
            var TotalAlum2 = Math.round(alumnosLista.length-30)    
        }
        else
        {
            var TotalAlum1 = alumnosLista.length;
            var TotalAlum2 = 0;
            Hoja.getRange(10,3,TotalAlum1,3).setValues(Alumnos1);//GRABA HASTA 30 ALUMNOS CORREO INSTITUCIONAL
        }

        
       
        
        if (TotalAlum2 > 0)   //GRABA ALUMNOS SI SON MAS QUE 30
        {
            Alumnos2 =[]
            for (var i=0; i<TotalAlum1; i++)
            {
                
                Alumnos2.push(Alumnos1[i])
            }
            
            Hoja.getRange(10,3,TotalAlum1,3).setValues(Alumnos2);

            Alumnos2 =[]
            for (var i=30; i<alumnosLista.length; i++)
            {
                
                Alumnos2.push(Alumnos1[i])
            }
            Hoja.getRange(42,3,TotalAlum2,3).setValues(Alumnos2);
            
        }
    } // fin TRY
    catch(err)
    {
        var errMsg=('NO HAY ALUMNOS SELECCIONADOS...',err.message)
        console.log(errMsg);
        /*
        var html= HtmlService.createHtmlOutput(errMsg)
            .setWidth(500)
            .setHeight(100);
            //SpreadsheetApp.getUi()
            .showModalDialog(html,'ERROR');
            */
    }
    

    //ESCRIBE FORMULAS
    //Evidencias 
    Hoja.getRange("BX10").setFormula('=IF(C10:C>0,IF(COUNT(F10:BW10)=0,"-",COUNT({G10,I10,K10,M10,O10,Q10,S10,U10,W10,Y10,AA10,AC10,AE10,AG10,AI10,AK10,AM10,AO10,AQ10,AS10,AU10,AW10,AY10,BA10,BC10,BE10,BG10,BI10,BK10,BM10,BO10,BQ10,BS10,BU10,BW10})),"")');
    var FillEvidencias = Hoja.getRange(10,76,TotalAlum1)//RANGO PARA RELLENAR MAS ADELANTE
    
    //Evaluacion continua        
    Hoja.getRange("BY10").setFormula('=IF(COUNT(F10:BW10)=0,"",FIXED(SUM({G10,I10,K10,M10,O10,Q10,S10,U10,W10,Y10,AA10,AC10,AE10,AG10,AI10,AK10,AM10,AO10,AQ10,AS10,AU10,AW10,AY10,BA10,BC10,BE10,BG10,BI10,BK10,BM10,BO10,BQ10,BS10,BU10,BW10})/$BY$4,1))');
    var FillEvaCont = Hoja.getRange(10,77,TotalAlum1)//RANGO PARA RELLENAR MAS ADELANTE
    
    //Numero de faltas
    Hoja.getRange("CF10").setFormula('=IF((COUNTA(F10:BW10))=0,"",FIXED(COUNTIF({F10,H10,J10,L10,N10,P10,R10,T10,V10,X10,Z10,AB10,AD10,AF10,AH10,AJ10,AL10,AN10,AP10,AR10,AT10,AV10,AX10,AZ10,BB10,BD10,BF10,BH10,BJ10,BL10,BN10,BP10,BR10,BT10,BV10},"=/"),))');
    var FillNoFaltas = Hoja.getRange(10,84,TotalAlum1)//RANGO PARA RELLENAR MAS ADELANTE
    
                                                //RELLENA CON FORMULAS
    Hoja.getRange("BX10").copyTo(FillEvidencias);
    Hoja.getRange("BY10").copyTo(FillEvaCont);
    Hoja.getRange("CF10").copyTo(FillNoFaltas);
    
    if (TotalAlum2>0)                                //RELLENA FORMULAS EN SEGUNDA HOJA
    {
        var FillEvidencias1 = Hoja.getRange(10,76,TotalAlum2)
        var FillEvaCont1 = Hoja.getRange(10,77,TotalAlum2)
        var FillNoFaltas1 = Hoja.getRange(10,84,TotalAlum2)
        Hoja.getRange("BX42").copyTo(FillEvidencias1);
        Hoja.getRange("BY42").copyTo(FillEvaCont1);
        Hoja.getRange("CF42").copyTo(FillNoFaltas1);
    }
    

    try
    {
    Lista.addEditor(correodocente);  //ASIGNA PRIVILEGIOS DE EDITOR A CORREO INSTITUCIONAL DOCENTE
    }
    catch(err)
    {
        var errMsg=('Error al asignar profesor como editor %s',err.message)
        console.log(errMsg);
        /*
        var html= HtmlService.createHtmlOutput(errMsg)
            .setWidth(500)
            .setHeight(100);
        SpreadsheetApp.getUi()
            .showModalDialog(html,'FALLA');
            */
    }
                                        //PREPARA AREAS PARA BLOQUEO DE SEGURIDAD
    var Bloqueo1 = Hoja.getRange('A1:E');
    var Bloqueo2 = Hoja.getRange('F1:CG5');
    var Bloqueo3 = Hoja.getRange('BX5:BZ');
    var Bloqueo4 = Hoja.getRange('CA5:CG9');
    var Bloqueo5 = Hoja.getRange('CB10:CC');
    var Bloqueo6 = Hoja.getRange('CE10:CG');

    var Protection1 = Bloqueo1.protect();
    var Protection2 = Bloqueo2.protect();
    var Protection3 = Bloqueo3.protect();
    var Protection4 = Bloqueo4.protect();
    var Protection5 = Bloqueo5.protect();
    var Protection6 = Bloqueo6.protect();

    Protection1.removeEditor(correodocente);         //RETIRA PRIVILEGIOS DE EDICION AL DOCENTE
    Protection2.removeEditor(correodocente);
    Protection3.removeEditor(correodocente);
    Protection4.removeEditor(correodocente);
    Protection5.removeEditor(correodocente);
    Protection6.removeEditor(correodocente);
    

DriveApp.getFileById(Idlis).setName(archSalida)  //RENOMBRA ARCHIVO DE SALIDA
//ASIGNA PERMISOS DE ACCESO A DOCENTE
DriveApp.getFileById(Idlis).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
// ********************************************************************************
return werror
}

function siguiente()
{
    var indexAnt=0
    const datuakIkasgela= ikasgelaOrria.getDataRange().getValues();
    var azkenIlara = ikasgelaOrria.getLastRow()+1;
    var ai = ikasgelaOrria.getLastRow();
    if (ai==1)
    {
        indexAnt=0
    }
    else
    {
        indexAnt=ikasgelaOrria.getRange(ai,1).getValue();
    }
    var index=indexAnt+1
    var aa=new Date().getFullYear();
    cveAula=aa+cveOpcEdu+index;
    var cveAsigna=formSalida[2].slice(0,4);
    var long=formSalida[2].length-4;
    var nomAsigna=formSalida[2].slice(5,long).toUpperCase();
    
    var nomAsigAula=formSalida[9].toUpperCase();

    nomClase=cveDocente+" - "+nomAsigAula+"("+cveAsigna+")" //RIAG - AUDITORÍA ADMINISTRATIVA.-(LA42)

    var seccion = formSalida[4]+" "+formSalida[5]+" "+formSalida[6]+" "+formSalida[7]+" "+formSalida[8];   //SECCION(GRUPOS PARTICIPANTES)	
    
    var descrClase=idOpEd+" | "+nomDocente+" | "+formSalida[1] //CVE OP EDU | DOCENTE | PERIDO

    wsalida.push([index,                     //indice
                cveAula,                    //clave unica de aula  año opc edu index
                formSalida[0],              //op edu
                idOpEd,                     // cve op edu
                formSalida[1],              //periodo
                cveAsigna,                  //cve asignatura
                cveDocente,                 //cve docente
                docenteNom,                 //nombre docente
                formSalida[4],              //grupo1
                formSalida[5],              //grupo2
                formSalida[6],              //grupo3
                formSalida[7],              //grupo4
                formSalida[8],              //grupo5
                "",                         //grupo6
                "",                         //grupo7
                nomClase,                   //NOMBRE CLASE RIAG - AUDITORÍA ADMINISTRATIVA.-(LA42)
                seccion,                    //SECCION(GRUPOS PARTICIPANTES)	
                formSalida[0],              //MATERIA(NOMBRE OPCION EDU)	
                formSalida[1],              //AULA(PERIODO)	
                descrClase                  //DESCR CLASE(CVE OP EDU | DOCENTE | PERIDO
])
  
    try
    {
      //obtiene rango a grabar 
       //escribe datos modificados
    ikasgelaOrria.getRange(azkenIlara,1,1,20).setValues(wsalida)
    bdCeec.waitForAllDataExecutionsCompletion(2)
    
    }
    catch(err)
    {
      console.log("error al grabar datos en tabla Aulas "+err);
  
      werror=1;
    }
  
  return werror;
    
}