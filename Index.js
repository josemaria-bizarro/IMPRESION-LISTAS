const bd_id='1T5KQCvXiYsGewSoc_A8qugUwSs-ft1NNds4oaNgQxkw';
const bd = SpreadsheetApp.openById(bd_id);
const bdAsigna=bd.getSheetByName('TABLA ASIGNATURAS');
const bdPeriodo=bd.getSheetByName('PERIODOS EDUCATIVOS');
const bdDocente=bd.getSheetByName('PERSONAL');
const bdOpcEdu=bd.getSheetByName('OPCIONES EDUCATIVAS');
const bdgrupo=bd.getSheetByName('GRUPOS');
const bdikasleId='1kzmgvrbkQ7ehUslRxR4ghkpSeI2yT99M3h-Dyo-EEVY';
const bdikasle=SpreadsheetApp.openById(bdikasleId);
const bdikasleOrria=bdikasle.getSheetByName('ACTIVOS_FORMATEADO');
let datuaAsigna;
