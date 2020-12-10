/*Agrega los títulos, Validación de datos, Fórmulas, Formatos y Protección
Se requiere correr una única vez al crear una nueva hoja de respuestas*/
function formatearHoja() {
  const S = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const S2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Turnos Habilitados');
  
  S.setName('Solicitudes');
  //  S.getRange('D1').setValue("Nombre Escrito");
  //  S.getRange('G1').setValue("Materia Escrita");
  S.getRange('K1').setValue("Cantidad");
  S.getRange('M1').setValue("Disponibilidad");
  S.getRange('N1').setValue("Estado");
  S.getRange('O1').setValue("Confirmación Estado");
  S.getRange('P1').setValue("Asistencia");
  S.getRange('Q1').setValue("Confirmación Asistencia");
  S.getRange('R1').setValue("idSolicitud");
  S.hideColumns(18);
  S.getRange('S1').setValue("idReserva");
  S.getRange('T1').setValue("Correo Personal");
  S.getRange('V1').setValue("Filas Nulas");
  S.getRange('W1').setValue("Deshabilitados");
  S.getRange('X1').setValue("Prioritarios");
  S.getRange('Y1').setValue("Errores De Fecha");
  S.getRange('Z1').setValue("Errores De Cantidad");
  S.getRange('AA1').setValue("Solicitudes Duplicadas");
  
  S.getRange('M2:M').setDataValidation(SpreadsheetApp.newDataValidation()
                                       .setAllowInvalid(false)
                                       .requireValueInList(['Turno Ocupado', 'Turno Disponible'], true)
                                       .build());
  S.getRange('N2:N').setDataValidation(SpreadsheetApp.newDataValidation()
                                       .setAllowInvalid(false)
                                       .requireValueInList(['Aceptar', 'Rechazar', 'Cancelar', 'Cancelar X Usuario'], true)
                                       .build());
  S.getRange('O2:O').setDataValidation(SpreadsheetApp.newDataValidation()
                                       .setAllowInvalid(false)
                                       .requireValueInList(['Aceptado', 'Rechazado', 'Cancelado X Admin', 'Cancelado X Usuario'], true)
                                       .build());
  S.getRange('P2:P').setDataValidation(SpreadsheetApp.newDataValidation()
                                       .setAllowInvalid(false)
                                       .requireValueInList(['Asistió', 'Ausente Con Aviso', 'Ausente Sin Aviso', 'Rechazado', 'Cancelado X Admin', 'Cancelado X Usuario'], true)
                                       .build());
  S.getRange('Q2:Q').setDataValidation(SpreadsheetApp.newDataValidation()
                                       .setAllowInvalid(false)
                                       .requireValueInList(['Encuesta Enviada', 'ACA', 'ASA', 'Rechazado', 'CXA', 'CXU'], true)
                                       .build());
  
  S.getRange('T2').insertCheckboxes();
  S.getRange('U2').setValue("2-");
  S.getRange('V2').setFormula('=AND(ISBLANK($O2)=FALSE;ISBLANK($Q2)=FALSE;$I2<TODAY())');
  S.getRange('W2').setFormula('=IF(COUNTIF(\'Turnos Habilitados\'!$O$3:$T$23;$H2&WEEKDAY($I2)&$J2)=1;TRUE)');
  S.getRange('X2').setFormula('=AND(ISBLANK($O2);OR($I2=TODAY();$I2=TODAY()+1))');
  S.getRange('Y2').setFormula('=AND(ISBLANK($O2),OR(WEEKDAY($I2)=1,AND(WEEKDAY($I2)=7,$J2="Noche"),$I2<TODAY(),YEAR($I2)>YEAR(NOW())))');
  S.getRange('Z2').setFormula('=AND(ISBLANK($O2);OR(AND($H2="Medrano-Azul";$K2>50);AND($H2="Medrano-Multimedia";$K2>15);AND($H2="Medrano-Rojo";$K2>44);AND($H2="Medrano-Verde";$K2>22);AND($H2="Medrano-WorkGroup Lab 1";$K2>30);AND($H2="Campus-Campus";$K2>30);AND($H2="Campus-WorkGroup Lab 2";$K2>24)))');
  S.getRange('AA2').setFormula('=AND(ISBLANK($O2);SUMPRODUCT(--(($H$2:H&$I$2:I&$J$2:J)=($H2&$I2&$J2)))>1)');
  
  S.getRange('I2').setNumberFormat('dddd"  "d"/"mm"/"yy');
  S.getRange(1, 1, S.getMaxRows(), S.getMaxColumns()).setFontFamily('Arial').setFontSize(10);
  S.setColumnWidth(1, 122);
  S.setColumnWidth(2, 200);
  S.setColumnWidth(6, 250);
  S.setColumnWidth(8, 169);
  S.setColumnWidth(9, 123);
  S.setColumnWidth(10, 55);
  S.setColumnWidth(11, 60);
  S.setColumnWidth(13, 127);
  S.setColumnWidth(14, 84);
  S.setColumnWidth(15, 131);
  S.setColumnWidth(18, 75);
  S.setColumnWidth(19, 104);
  S.autoResizeColumns(20, 7);
  
  var rules = S.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
             .setRanges([S.getRange('I2:I')]) //Fecha
             .whenFormulaSatisfied('=IF(Y2:Y=TRUE,TRUE)') //Errores De Fecha
             .setBackground('#FF9900')
             .setFontColor('#FFFFFF')
             .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
             .setRanges([S.getRange('I2:I')]) //Fecha
             .whenFormulaSatisfied('=IF(X2:X=TRUE,TRUE)') //Prioritarios
             .setBackground('#FFFFFF')
             .setFontColor('#FF0000')
             .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
             .setRanges([S.getRange('J2:J')]) //Turno
             .whenFormulaSatisfied('=IF(W2:W=TRUE,TRUE)') //Deshabilitados
             .setBackground('#A64D79')
             .setFontColor('#FFFFFF')
             .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
             .setRanges([S.getRange('K2:K')]) //Cantidad
             .whenFormulaSatisfied('=IF(Z2:Z=TRUE,TRUE)') //Errores De Cantidad
             .setBackground('#FF3A61')
             .setFontColor('#FFFFFF')
             .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
             .setRanges([S.getRange('M2:M')]) //Disponibilidad
             .whenFormulaSatisfied('=AND(ISBLANK($O2);$M2="Turno Ocupado")') //Turno Ocupado
             .setBackground('#FF0000')
             .setFontColor('#FFFFFF')
             .build());
  
  S.setConditionalFormatRules(rules);
  
  S.protect()
  .setUnprotectedRanges([S.getRange('N2'), S.getRange('P2'), S.getRange('T2')])
  .removeEditors(['insertar lista de editores sin permisos de administrador']);
  
  S2.protect()
  //  .setUnprotectedRanges([S2.getRange('E3:I23'), S2.getRange('J3:J4'), S2.getRange('J6:J7'), S2.getRange('J9:J10'), 
  //                         S2.getRange('J12:J13'), S2.getRange('J15:J16'), S2.getRange('J18:J19'), S2.getRange('J21:J22')])
  .removeEditors(['insertar lista de editores sin permisos de administrador']);
  S2.getRange('A1:Z100').protect()
  .setWarningOnly(true);
  
  /*Se requiere quitar de esta lista quienes no tengan permisos de administrador de editar toda la hoja
  De compartir con nuevos se tiene que quitar manualmente de la protección desde una cuenta administrador*/
};