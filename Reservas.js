//Variables globales constantes en todo el código, cambiar siempre estas de ser necesario y así evitar códigos harcodeados
const HOJA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Solicitudes');
const HOJA2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Turnos Habilitados');
const ULTIMA_FILA = HOJA.getLastRow(), ULTIMA_COLUMNA = HOJA.getLastColumn();
const MES = new Array('Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre');
const DIA_DEL_ANIO = new Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334);
const DIA_SEMANA = new Array('Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado');
const COLORES = new Array('#ffadad', '#adfff3', '#fffdad', '#adffb9', '#ffade4', '#ffc6ad', '#ebffad', '#adccff', '#ffadbd', '#adffda', '#ffecad', '#ddffad', '#d5adff', '#f9adff', '#adffbe', '#caadff');

const COLUMNA_TIMESTAMP = 1, COLUMNA_CORREO = 2, COLUMNA_PROFESOR = 3, COLUMNA_NOMBRE = 4, COLUMNA_ESPECIALIDAD = 5, COLUMNA_MATERIA = 6,
  COLUMNA_MATERIA_ESCRITA = 7, COLUMNA_LABORATORIO = 8, COLUMNA_FECHA = 9, COLUMNA_TURNO = 10, COLUMNA_CANTIDAD = 11, COLUMNA_PETICION = 12,
  COLUMNA_DISPONIBILIDAD = 13, COLUMNA_ESTADO = 14, COLUMNA_CONFIRM_ESTADO = 15, COLUMNA_ASISTENCIA = 16, COLUMNA_COMFIRM_ASISTENCIA = 17,
  COLUMNA_ID_SOLICITUD = 18, COLUMNA_ID_RESERVA = 19, COLUMNA_CORREO_PERSONAL = 20,
  I = 21, COLUMNA_FILAS_NULAS = 22, COLUMNA_DESHABILITACIONES = 23, COLUMNA_PRIORITARIOS = 24, COLUMNA_ERROR_FECHA = 25,
  COLUMNA_ERROR_CANTIDAD = 26, COLUMNA_SOLICITUD_DUPLICADAS = 27;

const MOTIVO_DESHABILITADO = 'Deshabilitado', MOTIVO_EFECHA = 'Error Fecha', MOTIVO_FERIADO = 'Feriado', MOTIVO_OCUPADO = 'Ocupado',
  MOTIVO_CANTIDAD = 'Error Cantidad', MOTIVO_GESTION = 'En Gestion', MOTIVO_ACEPTAR = 'Aceptado', MOTIVO_RECHAZADO_X_ERROR = 'Rechazado X Error',
  MOTIVO_CANCELADO_X_ERROR = 'Cancelado X Error', MOTIVO_RECHAZO = 'Rechazado', MOTIVO_CANCELADO_X_ADMIN = 'Cancelado X Admin',
  MOTIVO_CANCELADO_X_USUARIO = 'Cancelado X Usuario', MOTIVO_ENCUESTA = 'Llenar Encuesta', MOTIVO_DISCULPA_X_VENCIMIENTO = 'Disculpa X Solicitud Vencida';

const OCUPADO = 'Turno Ocupado', DISPONIBLE = 'Turno Disponible';

const ESTADO_ACEPTAR = 'Aceptar', ESTADO_RECHAZAR = 'Rechazar', ESTADO_CANCELAR = 'Cancelar', ESTADO_CANCELAR_X_USUARIO = 'Cancelar X Usuario',
  CORREO_PERSONAL_OK = true, CORREO_PERSONAL_NO_OK = false;

const CONFIRM_ACEPTAR = 'Aceptado', CONFIRM_RECHAZAR = 'Rechazado', CONFIRM_CANCELAR = 'Cancelado X Admin', CONFIRM_CANCELAR_X_USUARIO = 'Cancelado X Usuario';

const ESTADO_ASISTIO = 'Asistió', ESTADO_AUSENTE_CON_AV = 'Ausente Con Aviso', ESTADO_AUSENTE_SIN_AV = 'Ausente Sin Aviso',
  ESTADO_ASIS_RECHAZADO = 'Rechazado', ESTADO_ASIS_CANCELADO_X_ADMIN = 'Cancelado X Admin', ESTADO_ASIS_CANCELADO_X_USUARIO = 'Cancelado X Usuario';

const CONFIRM_ASISTIO = 'Encuesta Enviada', CONFIRM_AUSENTE_CON_AV = 'ACA', CONFIRM_AUSENTE_SIN_AV = 'ASA', CONFIRM_ASIS_RECHAZADO = 'Rechazado',
  CONFIRM_ASIS_CANCELADO_X_ADMIN = 'CXA', CONFIRM_ASIS_CANCELADO_X_USUARIO = 'CXU';

const MENSAJE_ERROR_AL_BORRAR = 'No borre datos de la hoja, se recopilan para generar estadísticas.', TITULO_ERROR_AL_BORRAR = 'ATENCIÓN';

function alAbrir() {
  HOJA.setActiveSelection(HOJA.getRange(ULTIMA_FILA, COLUMNA_CORREO));
};

/*Guarda una cadena solo de las filas necesarias para no recorrer toda la hoja, sino solo esas
Requiere tener vigente un activador de tiempo (recomendado todos los días después de medianoche), solo en la cuenta principal*/
function filasInnecesarias() { //AGREGAR ACTIVADOR TEMPORAL
  let cadena = HOJA.getRange(2, I).getValue();
  let f = cadena.split('-');
  gestionesVencidas(f);
  let filas = '';
  for (let i = f[0]; i <= ULTIMA_FILA; i++) {
    if (!HOJA.getRange(i, COLUMNA_FILAS_NULAS).getValue()) {
      filas += i + '-';
    }
  }
  HOJA.getRange(2, I).setValue(filas);
  habilitarEdicion(filas.split('-'));
};

//Chequea las solicitudes ya vencidas y las gestiona enviando correo de disculpa
function gestionesVencidas(filas) {
  var ayer = new Date();
  ayer.setDate(ayer.getDate() - 1);
  for (let i = 0; i < filas.length - 1; i++) {
    if (HOJA.getRange(filas[i], COLUMNA_FECHA).getValue() < ayer && HOJA.getRange(filas[i], COLUMNA_CONFIRM_ESTADO).getValue() == '') {
      enviarCorreoAuto(agregarDatos(filas[i]), MOTIVO_DISCULPA_X_VENCIMIENTO);
      HOJA.getRange(filas[i], COLUMNA_ESTADO).setValue(ESTADO_RECHAZAR);
      HOJA.getRange(filas[i], COLUMNA_CONFIRM_ESTADO).setValue(CONFIRM_RECHAZAR);
      HOJA.getRange(filas[i], COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_RECHAZADO);
      HOJA.getRange(filas[i], COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_ASIS_RECHAZADO);
      formatearFila(filas[i], 3);
    }
  }
}

/*Agrega fórmulas, chequea errores para rechazos automáticos, disponibilidad, solicitudes duplicadas y habilita edición
Requiere tener vigente el activador "Al enviar formulario", solo en la cuenta principal (Averiguar que sucede si en más de una cuenta)*/
function enEnvioFormulario(e) {
  var fila = e.range.getRow();
  HOJA.getRange(fila, COLUMNA_FECHA).setNumberFormat('dddd"  "d"/"mm"/"yy');
  HOJA.getRange(fila, COLUMNA_CORREO_PERSONAL).insertCheckboxes();
  HOJA.getRange(fila - 1, I + 1, 1, ULTIMA_COLUMNA - I).copyTo(HOJA.getRange(fila, I + 1, 1, ULTIMA_COLUMNA - I));
  let auto = false;
  var feriados = [];
  var datos = agregarDatos(fila);
  if (HOJA.getRange(fila, COLUMNA_DESHABILITACIONES).getValue()) {
    auto = enviarCorreoAuto(datos, MOTIVO_DESHABILITADO, fila);
  } else if (HOJA.getRange(fila, COLUMNA_ERROR_FECHA).getValue()) {
    auto = enviarCorreoAuto(datos, MOTIVO_EFECHA, fila);
  } else {
    try {
      //Usar aca un calendario personal si se quieren modificar, agregar o elegir los días feriados.
      var calendarioFeriados = CalendarApp.getCalendarById('es.ar#holiday@group.v.calendar.google.com');
      if (calendarioFeriados == null) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Chequear rechazo por feriado manualmente.', 'Error en asiganción a calendario de feriados.', -1);
        Logger.log('Línea 93 \nError en asiganción a calendario de feriados. ' +
          '\ncalendarioFeriados: ' + calendarioFeriados + '\nLab: ' + datos.lab + '\nFecha: ' + datos.fechaInicio + '\nTurno: ' + datos.turno);
      } else {
        feriados = calendarioFeriados.getEventsForDay(datos.fechaInicio);
      }
      if (feriados.length > 0 && calendarioFeriados != null) {
        auto = enviarCorreoAuto(datos, MOTIVO_FERIADO, fila);
      } else {
        if (datos.calendario == null) {
          SpreadsheetApp.getActiveSpreadsheet().toast('Chequear disponibilidad manualmente.', 'Error en asiganción a calendario del laboratorio.', -1);
          Logger.log('Línea 103 \nError en asiganción a calendario del laboratorio.' +
            'datos.calendario: ' + datos.calendario + '\nLab: ' + datos.lab + '\nFecha: ' + datos.fechaInicio + '\nTurno: ' + datos.turno);
        } else {
          var reservasDelDia = datos.calendario.getEvents(datos.fechaInicio, datos.fechaFin);
          if (reservasDelDia.length > 0) {
            HOJA.getRange(fila, COLUMNA_DISPONIBILIDAD).setValue(OCUPADO);
            auto = enviarCorreoAuto(datos, MOTIVO_OCUPADO, fila);
          } else {
            HOJA.getRange(fila, COLUMNA_DISPONIBILIDAD).setValue(DISPONIBLE);
          }
        }
      }
    } catch (error) {
      SpreadsheetApp.getActiveSpreadsheet().toast('No se chequeó el calendario. Chequear disponibilidad de feriados y turno manualmente.',
        'Error de CalendarApp en gestión al recibir formulario.', -1);
      Logger.log('Error de CalendarApp en gestión al recibir formulario. ' + error.name + '\n' + error.message);
    }
  }
  if (auto) {
    HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_RECHAZAR);
    HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).setValue(CONFIRM_RECHAZAR);
    HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_RECHAZADO);
    HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_ASIS_RECHAZADO);
    formatearFila(fila, 3);
  } else {
    agregarSolicitud(datos, fila);
    var cadena = HOJA.getRange(2, I).getValue() + fila + '-';
    HOJA.getRange(2, I).setValue(cadena);
    habilitarEdicion(cadena.split('-'));
    if (HOJA.getRange(fila, COLUMNA_SOLICITUD_DUPLICADAS).getValue()) {
      solicitudesDuplicadas(cadena.split('-').map(Number), fila);
    }
    enviarCorreoAuto(datos, MOTIVO_GESTION, fila);
  }
};

//Desprotege y habilita para edición de a cada nueva fila nueva agregada, para evitar cambios en el resto de la hoja debajo
function habilitarEdicion(filas) {
  var desprotecciones = [];
  for (let i = 0; i < filas.length - 1; i++) {
    desprotecciones.push(HOJA.getRange(filas[i], COLUMNA_ESTADO));
    desprotecciones.push(HOJA.getRange(filas[i], COLUMNA_ASISTENCIA));
    desprotecciones.push(HOJA.getRange(filas[i], COLUMNA_CORREO_PERSONAL));
  }
  var proteccion = HOJA.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  proteccion.setUnprotectedRanges(desprotecciones);
};

//Agrega la solicitud al calendario correspondiente
function agregarSolicitud(datos, fila) {
  if (datos.calendarioSolicitud == null) {
    SpreadsheetApp.getActiveSpreadsheet().toast('De requerir agregar solicitud al calendario manualmente.', 'Error en asiganción a calendario de solicitudes.', -1);
    Logger.log('Línea 155 \nError en asiganción a calendario de solicitudes.' +
      'datos.calendarioSolicitud: ' + datos.calendarioSolicitud + '\nLab: ' + datos.lab + '\nFecha: ' + datos.fechaInicio + '\nTurno: ' + datos.turno);
  } else {
    var reserva = datos.calendarioSolicitud.createEvent(datos.nombre, datos.fechaInicio, datos.fechaFin);
    var idSolicitud = reserva.getId();
    HOJA.getRange(fila, COLUMNA_ID_SOLICITUD).setValue(idSolicitud);
  }
};

//Resalta duplicados de solicitudes en distintos colores. Se eligen en el vector COLORES[]
function solicitudesDuplicadas(filas, fila) {
  var duplicados = [];
  for (let i = 0; i < filas.length - 1; i++) {
    if (HOJA.getRange(filas[i], COLUMNA_SOLICITUD_DUPLICADAS).getValue()) {
      duplicados.push(filas[i]);
    }
  }
  if (fila == 3) {
    var c = 0;
  } else {
    var c = HOJA.getRange(3, I).getValue();
  }
  var fila1 = HOJA.getRange(fila, COLUMNA_LABORATORIO).getValue() + HOJA.getRange(fila, COLUMNA_FECHA).getValue() + HOJA.getRange(fila, COLUMNA_TURNO).getValue();
  for (let i = 0; i < duplicados.length; i++) {
    var filaX = HOJA.getRange(duplicados[i], COLUMNA_LABORATORIO).getValue() + HOJA.getRange(duplicados[i], COLUMNA_FECHA).getValue() + HOJA.getRange(duplicados[i], COLUMNA_TURNO).getValue();
    if (filaX == fila1) {
      HOJA.getRange(fila, 1, 1, I - 1).setBackground(COLORES[c]);
      HOJA.getRange(duplicados[i], 1, 1, I - 1).setBackground(COLORES[c]);
    }
  }
  c++;
  if (c >= COLORES.length) {
    c = 0;
  }
  HOJA.getRange(3, I).setValue(c);
};

//Correr únicamente si falla en correr enEnvioFormulario (Ver error de activador). Formatea cada fila perjudicada por la falla y vuelve la hoja a su estado funcional.
function arregloEnEnvioFormulario() {
  var cadena = HOJA.getRange(2, I).getValue();
  var filas = cadena.split('-').map(Number);
  var filasNuevas = [], feriados = [];
  for (let i = filas[0]; i <= ULTIMA_FILA; i++) {
    if (HOJA.getRange(i, COLUMNA_ESTADO).getValue() == '') {
      filasNuevas.push(i);
    }
  }
  HOJA.getRange(filasNuevas[0] - 1, I + 1, 1, ULTIMA_COLUMNA - I).copyTo(HOJA.getRange(filasNuevas[0], I + 1, ULTIMA_FILA - filasNuevas[0], ULTIMA_COLUMNA - I));
  let auto = false;
  for (let i = 0; i < filasNuevas.length; i++) {
    HOJA.getRange(filasNuevas[i], COLUMNA_FECHA).setNumberFormat('dddd"  "d"/"mm"/"yy');
    HOJA.getRange(filasNuevas[i], COLUMNA_CORREO_PERSONAL).insertCheckboxes();
    var datos = agregarDatos(filasNuevas[i]);
    if (HOJA.getRange(filasNuevas[i], COLUMNA_DESHABILITACIONES).getValue()) {
      auto = enviarCorreoAuto(datos, MOTIVO_DESHABILITADO);
    } else if (HOJA.getRange(filasNuevas[i], COLUMNA_ERROR_FECHA).getValue()) {
      auto = enviarCorreoAuto(datos, MOTIVO_EFECHA);
    } else {
      try {
        //Usar aca un calendario personal si se quieren modificar, agregar o elegir los días feriados.
        var calendarioFeriados = CalendarApp.getCalendarById('es.ar#holiday@group.v.calendar.google.com');
        if (calendarioFeriados == null) {
          SpreadsheetApp.getActiveSpreadsheet().toast('Chequear rechazo por feriado manualmente.', 'Error en asiganción a calendario de feriados.', -1);
          Logger.log('Línea 218 \nError en asiganción a calendario de feriados.' +
            '\ncalendarioFeriados: ' + calendarioFeriados + '\nLab: ' + datos.lab + '\nFecha: ' + datos.fechaInicio + '\nTurno: ' + datos.turno);
        } else {
          feriados = calendarioFeriados.getEventsForDay(datos.fechaInicio);
        }
        if (feriados.length > 0 && calendarioFeriados != null) {
          auto = enviarCorreoAuto(datos, MOTIVO_FERIADO);
        } else {
          if (datos.calendario == null) {
            SpreadsheetApp.getActiveSpreadsheet().toast('Chequear disponibilidad manualmente. Puede requerir arreglo.', 'Error en asiganción a calendario del laboratorio.', -1);
            Logger.log('Línea 228 \nError en asiganción a calendario del laboratorio.\n' +
              'datos.calendario: ' + datos.calendario + '\nLab: ' + datos.lab + '\nFecha: ' + datos.fechaInicio + '\nTurno: ' + datos.turno);
          } else {
            var reservasDelDia = datos.calendario.getEvents(datos.fechaInicio, datos.fechaFin);
            if (reservasDelDia.length > 0) {
              HOJA.getRange(filasNuevas[i], COLUMNA_DISPONIBILIDAD).setValue(OCUPADO);
              auto = enviarCorreoAuto(datos, MOTIVO_OCUPADO);
            } else {
              HOJA.getRange(filasNuevas[i], COLUMNA_DISPONIBILIDAD).setValue(DISPONIBLE);
            }
          }
        }
      } catch (error) {
        SpreadsheetApp.getActiveSpreadsheet().toast('No se chequeó el calendario. Chequear disponibilidad de feriados y turno manualmente.',
          'Error de CalendarApp en gestión manual de arregloEnEnvioFormulario.', -1);
        Logger.log('Error de CalendarApp en gestión manual de arregloEnEnvioFormulario.' + error.name + '\n' + error.message);
      }
    }
    if (auto) {
      HOJA.getRange(filasNuevas[i], COLUMNA_ESTADO).setValue(ESTADO_RECHAZAR);
      HOJA.getRange(filasNuevas[i], COLUMNA_CONFIRM_ESTADO).setValue(CONFIRM_RECHAZAR);
      HOJA.getRange(filasNuevas[i], COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_RECHAZADO);
      HOJA.getRange(filasNuevas[i], COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_ASIS_RECHAZADO);
      formatearFila(filasNuevas[i], 3);
    } else {
      agregarSolicitud(datos);
    }
  }

  filasInnecesarias();
  var cadena = HOJA.getRange(2, I).getValue();
  habilitarEdicion(cadena.split('-'));
  var filas = cadena.split('-').map(Number);
  var duplicados = [];
  for (let i = 0; i < filas.length - 1; i++) {
    if (HOJA.getRange(filas[i], COLUMNA_SOLICITUD_DUPLICADAS).getValue()) {
      duplicados.push(filas[i]);
    }
  }
  var c = 0;
  for (let i = 0; i < duplicados.length; i++) {
    var fila1 = HOJA.getRange(duplicados[i], COLUMNA_LABORATORIO).getValue() + HOJA.getRange(duplicados[i], COLUMNA_FECHA).getValue() + HOJA.getRange(duplicados[i], COLUMNA_TURNO).getValue();
    if (c >= COLORES.length) {
      c = 0;
    }
    for (let j = (1 + i); j < duplicados.length; j++) {
      var filaX = HOJA.getRange(duplicados[j], COLUMNA_LABORATORIO).getValue() + HOJA.getRange(duplicados[j], COLUMNA_FECHA).getValue() + HOJA.getRange(duplicados[j], COLUMNA_TURNO).getValue();
      if (fila1 == filaX && HOJA.getRange(duplicados[j], 1).getBackground() == '#ffffff') {
        HOJA.getRange(duplicados[i], 1, 1, I - 1).setBackground(COLORES[c]);
        HOJA.getRange(duplicados[j], 1, 1, I - 1).setBackground(COLORES[c]);
      }
    }
    c++;
  }
};

/*Chequea cambios al editarse la hoja en las columnas de Estado y Asistencia y la hoja de Deshabilitaciones,
 y corre todas las funciones necesarias para cada caso.
Requiere tener vigente y funcionando el Activador en "enEdición" en la cuenta que modifique la hoja*/
function enEdicion(e) {
  var IU = SpreadsheetApp.getUi();
  var fila = e.range.getRow();
  if (e.source.getSheetName() == HOJA.getName()) {
    switch (e.range.getColumn()) {
      case COLUMNA_ESTADO:
        if (e.range.getValue() == ESTADO_ACEPTAR && HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue() != CONFIRM_ACEPTAR) {
          aceptarReserva(fila);
          cambiosDisponibilidad(fila, OCUPADO);
        } else if (e.range.getValue() == ESTADO_RECHAZAR && HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue() != CONFIRM_RECHAZAR) {
          rechazarReserva(fila);
        } else if (e.range.getValue() == ESTADO_CANCELAR && HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue() != CONFIRM_CANCELAR) {
          cancelarReserva(fila, true);
          cambiosDisponibilidad(fila, DISPONIBLE);
        } else if (e.range.getValue() == ESTADO_CANCELAR_X_USUARIO && HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue() != CONFIRM_CANCELAR_X_USUARIO) {
          cancelarReserva(fila, false);
          cambiosDisponibilidad(fila, DISPONIBLE);
        } else if (e.range.getValue() == '' && HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue() != '') {
          deshacerEstado(fila);
          IU.alert(TITULO_ERROR_AL_BORRAR, MENSAJE_ERROR_AL_BORRAR, IU.ButtonSet.OK);
        }
        break;
      case COLUMNA_ASISTENCIA:
        if (e.range.getValue() == ESTADO_ASISTIO && HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).getValue() != CONFIRM_ASISTIO) {
          presentismo(fila, ESTADO_ASISTIO);
        } else if (e.range.getValue() == ESTADO_AUSENTE_CON_AV && HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).getValue() != CONFIRM_AUSENTE_CON_AV) {
          presentismo(fila, ESTADO_AUSENTE_CON_AV);
          cambiosDisponibilidad(fila, DISPONIBLE);
        } else if (e.range.getValue() == ESTADO_AUSENTE_SIN_AV && HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).getValue() != CONFIRM_AUSENTE_SIN_AV) {
          presentismo(fila, ESTADO_AUSENTE_SIN_AV);
        } else if (e.range.getValue() == ESTADO_ASIS_CANCELADO_X_ADMIN && HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).getValue() != CONFIRM_ASIS_CANCELADO_X_ADMIN) {
          presentismo(fila, ESTADO_ASIS_CANCELADO_X_ADMIN);
        } else if (e.range.getValue() == ESTADO_ASIS_RECHAZADO && HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).getValue() != CONFIRM_ASIS_RECHAZADO) {
          presentismo(fila, ESTADO_ASIS_RECHAZADO);
        } else if (e.range.getValue() == '' && HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).getValue() != '') {
          deshacerSeleccionAsistencia(fila);
          IU.alert(TITULO_ERROR_AL_BORRAR, MENSAJE_ERROR_AL_BORRAR, IU.ButtonSet.OK);
        }
        break;
      case COLUMNA_CORREO_PERSONAL:
        if (e.range.isChecked() == null) {
          HOJA.getRange(fila, COLUMNA_CORREO_PERSONAL).insertCheckboxes();
          IU.alert(TITULO_ERROR_AL_BORRAR, MENSAJE_ERROR_AL_BORRAR, IU.ButtonSet.OK);
        } else if (e.range.isChecked()) {
          if (!ventana(datos, CORREO_PERSONAL_OK)) {
            HOJA.getRange(fila, COLUMNA_CORREO_PERSONAL).insertCheckboxes();
          }
        } else if (!(e.range.isChecked())) {
          if (!ventana(datos, CORREO_PERSONAL_NO_OK)) {
            HOJA.getRange(fila, COLUMNA_CORREO_PERSONAL).insertCheckboxes().check();
          }
        }
        break;
    }
  } else if (e.source.getSheetName() == HOJA2.getName()) {
    if (e.range.isChecked() == null) {
      if (HOJA2.getRange(fila + 27, e.range.getColumn()).getValue() == '') {
        e.range.insertCheckboxes().check();
      } else {
        e.range.insertCheckboxes();
      }
      IU.alert(TITULO_ERROR_AL_BORRAR, MENSAJE_ERROR_AL_BORRAR, IU.ButtonSet.OK);
    } else if (!(e.range.isChecked())) {
      deshabilitarTurno(deshabilitaciones(e), e);
    } else if (e.range.isChecked()) {
      habilitarTurno(deshabilitaciones(e), e);
    }
  }
};

function aceptarReserva(fila) {
  var IU = SpreadsheetApp.getUi();
  var datos = agregarDatos(fila);
  if (datos.calendario == null) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No es posible agregar la reserva. Arreglo necesario.', 'Error en asiganción a calendario del laboratorio.', -1);
    Logger.log('Línea 362 \nError en asiganción a calendario del laboratorio.' +
      'datos.calendario: ' + datos.calendario + '\nLab: ' + datos.lab + '\nFecha: ' + datos.fechaInicio + '\nTurno: ' + datos.turno);
  } else {
    var reservasDelDia = datos.calendario.getEvents(datos.fechaInicio, datos.fechaFin);
    if (reservasDelDia[0] == null) {
      if (ventana(datos, ESTADO_ACEPTAR)) {
        reservarLaboratorio(datos, fila);
        switch (HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue()) {
          case '':
            enviarCorreo(datos, MOTIVO_ACEPTAR, fila);
            break;
          case CONFIRM_RECHAZAR:
            enviarCorreo(datos, MOTIVO_RECHAZADO_X_ERROR, fila);
            break;
          case CONFIRM_CANCELAR:
          case CONFIRM_CANCELAR_X_USUARIO:
            enviarCorreo(datos, MOTIVO_CANCELADO_X_ERROR, fila);
            break;
        }
        if (datos.idSolicitud != '') {
          borrarSolicitud(datos, fila);
        }
        HOJA.getRange(fila, COLUMNA_DISPONIBILIDAD).setValue(OCUPADO);
        HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).setValue(CONFIRM_ACEPTAR);
        HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue('');
        HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue('');
        formatearFila(fila, 1);
      } else {
        deshacerEstado(fila);
      }
    } else {
      deshacerEstado(fila);
      IU.alert('ERROR', 'Turno ocupado.', IU.ButtonSet.OK);
    }
  }
};

function rechazarReserva(fila) {
  var IU = SpreadsheetApp.getUi();
  var estado = HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue();
  switch (estado) {
    case '':
      var datos = agregarDatos(fila);
      if (ventana(datos, ESTADO_RECHAZAR)) {
        borrarReserva(datos);
        enviarCorreo(datos, MOTIVO_RECHAZO, fila);
        if (datos.idSolicitud != '') {
          borrarSolicitud(datos, fila);
        }
        HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).setValue(CONFIRM_RECHAZAR);
        HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_RECHAZADO);
        HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_ASIS_RECHAZADO);
        formatearFila(fila, 2);
      } else {
        HOJA.getRange(fila, COLUMNA_ESTADO).setValue('');
      }
      break;
    case CONFIRM_ACEPTAR:
      HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_ACEPTAR);
      IU.alert('ERROR', 'No se puede rechazar una solicitud aceptada.', IU.ButtonSet.OK);
      break;
    case CONFIRM_CANCELAR:
      HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_CANCELAR);
      IU.alert('ERROR', 'No se puede rechazar una solicitud cancelada.', IU.ButtonSet.OK);
      break;
    case CONFIRM_CANCELAR_X_USUARIO:
      HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_CANCELAR_X_USUARIO);
      IU.alert('ERROR', 'No se puede rechazar una solicitud cancelada.', IU.ButtonSet.OK);
      break;
  }
};

function cancelarReserva(fila, xAdmin) {
  var IU = SpreadsheetApp.getUi();
  if (HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue() != CONFIRM_RECHAZAR) {
    var datos = agregarDatos(fila);
    if (xAdmin) {
      if (ventana(datos, ESTADO_CANCELAR)) {
        borrarReserva(datos);
        enviarCorreo(datos, MOTIVO_CANCELADO_X_ADMIN, fila);
        HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).setValue(CONFIRM_CANCELAR);
        HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_CANCELADO_X_ADMIN);
        HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_ASIS_CANCELADO_X_ADMIN);
        if (datos.idSolicitud != '') {
          borrarSolicitud(datos, fila);
        }
        formatearFila(fila, 2);
      } else {
        deshacerEstado(fila);
      }
    } else {
      if (ventana(datos, ESTADO_CANCELAR_X_USUARIO)) {
        borrarReserva(datos);
        enviarCorreo(datos, MOTIVO_CANCELADO_X_USUARIO, fila);
        HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).setValue(CONFIRM_CANCELAR_X_USUARIO);
        HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue('');
        HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue('');
        if (datos.idSolicitud != '') {
          borrarSolicitud(datos, fila);
        }
        formatearFila(fila, 1);
      } else {
        deshacerEstado(fila);
      }
    }
  } else {
    HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_RECHAZAR);
    IU.alert('ERROR', 'No se puede cancelar una solicitud rechazada.', IU.ButtonSet.OK)
  }
};

//Deshace la selección del estado principal
function deshacerEstado(fila) {
  var cambioAnterior = HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue();
  switch (cambioAnterior) {
    case CONFIRM_ACEPTAR:
      HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_ACEPTAR);
      break;
    case CONFIRM_RECHAZAR:
      HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_RECHAZAR);
      break;
    case CONFIRM_CANCELAR:
      HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_CANCELAR);
      break;
    case CONFIRM_CANCELAR_X_USUARIO:
      HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_CANCELAR_X_USUARIO);
      break;
    default:
      HOJA.getRange(fila, COLUMNA_ESTADO).setValue('');
      break;
  }
};

//Elimina la solicitud del calendario de solicitudes correspondiente
function borrarSolicitud(datos, fila) {
  var solicitud = datos.calendarioSolicitud.getEventById(datos.idSolicitud);
  if (solicitud != null) {
    solicitud.deleteEvent();
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('No es necesario para que funcione el sistema.', 'Error al borrar solicitud del calendario.', -1);
    Logger.log('Línea 502 \nError al borrar solicitud del calendario.' +
      'datos.idSolicitud: ' + datos.idSolicitud + '\nLab: ' + datos.lab + '\nFecha: ' + datos.fechaInicio + '\nTurno: ' + datos.turno);
  }
  HOJA.getRange(fila, COLUMNA_ID_SOLICITUD).setValue('');
};

//Elimina la reserva del calendario de reservas correspondiente
function borrarReserva(datos) {
  var reservas = datos.calendario.getEvents(datos.fechaInicio, datos.fechaFin);
  var reserva = reservas[0];
  if (reserva != null) {
    reserva.deleteEvent();
  }
};

//Cambia la disponibilidad de los turnos al aceptar o cancelar
function cambiosDisponibilidad(fila, disponibilidad) {
  var cadena = HOJA.getRange(2, I).getValue();
  var filas = cadena.split('-').map(Number);
  var duplicados = [];
  for (let i = 0; i < filas.length - 1; i++) {
    if (HOJA.getRange(filas[i], COLUMNA_SOLICITUD_DUPLICADAS).getValue()) {
      duplicados.push(filas[i]);
    }
  }
  var fila1 = HOJA.getRange(fila, COLUMNA_LABORATORIO).getValue() + HOJA.getRange(fila, COLUMNA_FECHA).getValue() + HOJA.getRange(fila, COLUMNA_TURNO).getValue();
  for (let i = 0; i < duplicados.length; i++) {
    var filaX = HOJA.getRange(duplicados[i], COLUMNA_LABORATORIO).getValue() + HOJA.getRange(duplicados[i], COLUMNA_FECHA).getValue() + HOJA.getRange(duplicados[i], COLUMNA_TURNO).getValue();
    if (filaX == fila1) {
      HOJA.getRange(duplicados[k], COLUMNA_DISPONIBILIDAD).setValue(disponibilidad);
    }
  }
};

//Maneja y cambia los estados de asistencia
function presentismo(fila, estadoAsistencia) {
  var IU = SpreadsheetApp.getUi();
  let datos = agregarDatos(fila);
  var estadoPrincipal = HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).getValue();
  switch (estadoPrincipal) {
    case CONFIRM_RECHAZAR:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_RECHAZADO);
      IU.alert('ERROR', 'No se puede ingresar asistencia de una solicitud rechazada.', IU.ButtonSet.OK);
      break;
    case CONFIRM_CANCELAR:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_CANCELADO_X_ADMIN);
      IU.alert('ERROR', 'No se puede ingresar asistencia de una solicitud cancelada por la administración.', IU.ButtonSet.OK);
      break;
    case CONFIRM_CANCELAR_X_USUARIO:
      switch (estadoAsistencia) {
        case ESTADO_AUSENTE_CON_AV:
          if (ventana(datos, ESTADO_AUSENTE_CON_AV)) {
            HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_AUSENTE_CON_AV);
            formatearFila(fila, 2);
          } else {
            deshacerSeleccionAsistencia(fila);
          }
          break;
        case ESTADO_ASIS_CANCELADO_X_USUARIO:
          if (ventana(datos, ESTADO_ASIS_CANCELADO_X_USUARIO)) {
            HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_ASIS_CANCELADO_X_USUARIO);
            formatearFila(fila, 2);
          } else {
            deshacerSeleccionAsistencia(fila);
          }
          break;
        default:
          deshacerSeleccionAsistencia(fila);
          IU.alert('ERROR', 'Ingreso Inválido. Solo se puede ingresar "' + ESTADO_AUSENTE_CON_AV + '" o "' + ESTADO_ASIS_CANCELADO_X_USUARIO +
            '" en una solicitud cancelada a pedido del usuario.', IU.ButtonSet.OK);
          break;
      }
      break;
    case CONFIRM_ACEPTAR:
      switch (estadoAsistencia) {
        case ESTADO_ASISTIO:
          if (ventana(datos, ESTADO_ASISTIO)) {
            enviarCorreo(datos, MOTIVO_ENCUESTA, fila);
            HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_ASISTIO);
            formatearFila(fila, 2);
          } else {
            deshacerSeleccionAsistencia(fila);
          }
          break;
        case ESTADO_AUSENTE_CON_AV:
          if (ventana(datos, ESTADO_AUSENTE_CON_AV)) {
            borrarReserva(datos);
            HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_AUSENTE_CON_AV);
            HOJA.getRange(fila, COLUMNA_ESTADO).setValue(ESTADO_CANCELAR_X_USUARIO);
            HOJA.getRange(fila, COLUMNA_CONFIRM_ESTADO).setValue(CONFIRM_CANCELAR_X_USUARIO);
            formatearFila(fila, 2);
          } else {
            deshacerSeleccionAsistencia(fila);
          }
          break;
        case ESTADO_AUSENTE_SIN_AV:
          if (ventana(datos, ESTADO_AUSENTE_SIN_AV)) {
            HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).setValue(CONFIRM_AUSENTE_SIN_AV);
            formatearFila(fila, 2);
          } else {
            deshacerSeleccionAsistencia(fila);
          }
          break;
        default:
          deshacerSeleccionAsistencia(fila);
          IU.alert('ERROR', 'Ingreso Inválido. Solo se puede ingresar "' + ESTADO_ASISTIO + '", "' + ESTADO_AUSENTE_CON_AV +
            '" o "' + ESTADO_AUSENTE_SIN_AV + '" en una solicitud aceptada.', IU.ButtonSet.OK);
          break;
      }
      break;
    default:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue('');
      IU.alert('ERROR', 'No se puede ingresar asistencia en una solicitud no gestionada.', IU.ButtonSet.OK);
      break;
  }
};

//Aplica formato a la fila como color, letra, relleno, etc.
function formatearFila(fila, formato) {
  switch (formato) {
    case 1:
      HOJA.getRange(fila, 1, 1, I - 1).setFontColor('blue');
      HOJA.getRange(fila, 1, 1, I - 1).setBackground('white');
      break;
    case 2:
      HOJA.getRange(fila, 1, 1, I - 1).setFontColor('black');
      HOJA.getRange(fila, 1, 1, I - 1).setBackground('#eeeeee');
      break;
    case 3:
      HOJA.getRange(fila, 1, 1, I - 1).setFontColor('black');
      HOJA.getRange(fila, 1, 1, I - 1).setBackground('#dddddd');
      break;
  }
};

//Deshace la selección del estado de asistencia
function deshacerSeleccionAsistencia(fila) {
  var cambioAnterior = HOJA.getRange(fila, COLUMNA_COMFIRM_ASISTENCIA).getValue();
  switch (cambioAnterior) {
    case CONFIRM_ASISTIO:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASISTIO);
      break;
    case CONFIRM_AUSENTE_SIN_AV:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_AUSENTE_SIN_AV);
      break;
    case CONFIRM_AUSENTE_CON_AV:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_AUSENTE_CON_AV);
      break;
    case CONFIRM_ASIS_RECHAZADO:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_RECHAZADO);
      break;
    case CONFIRM_ASIS_CANCELADO_X_ADMIN:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_CANCELADO_X_ADMIN);
      break;
    case CONFIRM_ASIS_CANCELADO_X_USUARIO:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue(ESTADO_ASIS_CANCELADO_X_USUARIO);
      break;
    default:
      HOJA.getRange(fila, COLUMNA_ASISTENCIA).setValue('');
      break;
  }
};

//Crea una ventana de confirmación
function ventana(datos, estado) {
  var IU = SpreadsheetApp.getUi();
  var titulo = '';
  var mensaje = '';
  var hoja2 = false;
  switch (estado) {
    case ESTADO_ACEPTAR:
      titulo = '¿Aceptar Solicitud?';
      break;
    case ESTADO_RECHAZAR:
      titulo = '¿Rechazar Solicitud?';
      break;
    case ESTADO_CANCELAR:
      titulo = '¿Cancelar Solicitud?';
      break;
    case ESTADO_CANCELAR_X_USUARIO:
      titulo = '¿Cancelar Solicitud a Pedido del Usuario?';
      break;
    case ESTADO_ASISTIO:
      titulo = '¿Confirmar Asistencia?';
      break;
    case ESTADO_AUSENTE_CON_AV:
      titulo = '¿Inasistencia Con Aviso?';
      break;
    case ESTADO_AUSENTE_SIN_AV:
      titulo = '¿Inasistencia Sin Aviso?';
      break;
    case ESTADO_ASIS_CANCELADO_X_USUARIO:
      titulo = '¿Cancelado Por Usuario?';
      break;
    case CORREO_PERSONAL_OK:
      titulo = '¿Desea Redactar Correos Personales?';
      mensaje = 'En los casos donde se enviaría un correo automático, en lugar de eso, se crearía un borrador con estos datos en la cuenta del administrador ' +
        'para editar personalmente. \n\nIMPORTANTE: El sistema continuaría gestionando la solicitud sin notificar al usuario hasta uno no redactar y enviar cada ' +
        'correo manualmente.';
      break;
    case CORREO_PERSONAL_NO_OK:
      titulo = '¿Desea Que Se Vuelvan A Enviar Correos Automáticos?';
      mensaje = 'Se enviarían los correos automáticos ya predeterminados en cada caso, negando la opción de redactarlos personalmente.';
      break;
    case 'Deshabilitar':
      titulo = '¿Seguro Que Desea Deshabilitar?';
      mensaje = 'Se impedirían todas las reservas en el turno, lugar y día seleccionado, empezando desde está semana hasta el fin del año.\n\n';
      hoja2 = true;
      break;
    case 'Habilitar':
      titulo = '¿Seguro Que Desea Habilitar?';
      mensaje = 'Se volverían a tomar reservas en el turno, lugar y día seleccionado, desde la semana en el que fue deshabilitado hasta el fin del año.\n\n';
      hoja2 = true;
      break;
  }
  if (!hoja2) {
    var f = datos.fechaInicio;
    var fecha = DIA_SEMANA[f.getDay()] + ' ' + f.getDate() + ' de ' + MES[f.getMonth()];
    return IU.Button.YES == IU.alert(titulo, 'Nombre: ' + datos.nombre + '\nMateria: ' + datos.materia + '\nLaboratorio: ' + datos.lab +
      '\nFecha: ' + fecha + '\nTurno: ' + datos.turno + '\n\n' + mensaje, IU.ButtonSet.YES_NO);
  } else {
    return IU.Button.YES == IU.alert(titulo, 'Laboratorio: ' + datos.lab + '\nDía: ' + DIA_SEMANA[datos.dia] + '\nTurno: ' +
      datos.turno + '\n\n' + mensaje, IU.ButtonSet.YES_NO);
  }
};

//Crea la reserva y asigna en el calendario y recurso correspondiente de la fila recibida
function reservarLaboratorio(datos, fila) {
  var fecha = HOJA.getRange(fila, COLUMNA_TIMESTAMP).getValue();
  //Crea el Id {
  var dias = (DIA_DEL_ANIO[fecha.getMonth()] + fecha.getDate());
  var numero = parseInt((fecha.getFullYear() % 10) + ('0000' + ((dias - 1) * 24 + fecha.getHours())).slice(-4) + ('0000' + (fecha.getMinutes() * 60 + fecha.getSeconds())).slice(-4) + (fila % 10));
  var id = numero.toString(20).toUpperCase();
  // }
  HOJA.getRange(fila, COLUMNA_ID_RESERVA).setValue(id);
  datos.calendario.createEvent(datos.nombre, datos.fechaInicio, datos.fechaFin,
    {
      description: 'Número de reserva: ' + datos.idReserva + '\nMateria: ' + datos.materia + '\nCantidad de alumnos: ' +
        datos.cantPersonas + '\nPetición: ' + datos.peticiones, guests: datos.recurso + ', ' + datos.email/*, sendInvites: true*/
    });
};

function enviarCorreoAuto(datos, motivo, fila) {
  var encabezadoHtml = '<center><table border=1 style=background-color:#c4302b; width=80% height = 50><div >' +
    '<a href = https://www.frba.utn.edu.ar/que-significa-el-logo-de-la-utn-ba/ >' +
    '<img alt= utn src=https://www.frba.utn.edu.ar/wp-content/uploads/2016/08/logo-utn.ba-horizontal-e1471367724904.jpg width=80 height=35 > </a></div> </table></center>';
  var cuerpoAsunto = 'Gestión De Solicitud De Reserva UTN Laboratorio', cuerpoCorreo = '';
  var pieHtml = "<center><footer><p > <strong> <h4 style= color:#c4302b;> Ing. Ramiro Garbarini </h4> ·  Jefe de Laboratorios · Laboratorio de Sistemas de Información </strong> </p> </footer></center>";

  switch (motivo) {
    case MOTIVO_DESHABILITADO://Significa que el lab está deshabilitado, las razones por que no se saben
      cuerpoCorreo = '';
      /*
      Disculpe, su solicitud de reserva no puede ser gestionada debido a que el laboratorio en la fecha y turno elegidos se encuentra deshabilitado.
 
      {Datos}
 
      Puede verificar los turnos ocupados y también deshabilitados adhiriéndose a los calendarios de cada laboratorio. 
      Se encuentran en la misma página donde solicitar reservas así también con instrucciones [Opcional].
      */
      break;
    case MOTIVO_EFECHA: //Fecha del pasado
      if (datos.fecha < Date) {
        cuerpoCorreo = '';
        /*
        Disculpe, su solicitud de reserva no puede ser gestionada debido a que la fecha elegida ya es pasada.
 
        {Datos}
 
        Recuerde que solo se admiten fechas vigentes del mismo año.
        Fechas pasadas, días sábados en turno noche y domingos serán rechazados.
        */
      } else if (datos.fecha.getDay() == 0) { //Domingo
        cuerpoCorreo = '';
        /*
        Disculpe, su solicitud de reserva no puede ser gestionada debido a que la fecha elegida es un día domingo.
 
        {Datos}
 
        Recuerde que solo se admiten fechas vigentes del mismo año.
        Fechas pasadas, días sábados en turno noche y domingos serán rechazados.
        */
      } else if (datos.fecha.getDay() == 6 && datos.turno == 'Noche') { //Sábado a la noche
        cuerpoCorreo = '';
        /*
        Disculpe, su solicitud de reserva no puede ser gestionada debido a que la fecha elegida es sábado y el turno noche.
 
        {Datos}
 
        Recuerde que solo se admiten fechas vigentes del mismo año.
        Fechas pasadas, días sábados en turno noche y domingos serán rechazados.
        */
      } else if (datos.fecha.getFullYear() > Date.getFullYear()) { //Fecha del año siguiente
        cuerpoCorreo = '';
        /*
        Disculpe, su solicitud de reserva no puede ser gestionada debido a que la fecha elegida pertenece al año siguiente.
 
        {Datos}
 
        Recuerde que solo se admiten fechas vigentes del mismo año.
        Fechas pasadas, días sábados en turno noche y domingos serán rechazados.
        */
      }
      break;
    case MOTIVO_FERIADO:
      cuerpoCorreo = '';
      /*
      Disculpe, su solicitud de reserva no puede ser gestionada debido a que la fecha elegida es feriado.
 
      {Datos}
 
      Puede verificar los días feriados en el calendario académico aquí {Insertar enlace}. 
      */
      break;
    case MOTIVO_OCUPADO:
      cuerpoCorreo = '';
      /*
      Disculpe, su solicitud de reserva fue rechazada debido a que el laboratorio se encuentra ya ocupado en el turno requerido.
 
      {Datos}
 
      Puede verificar los turnos ocupados y también deshabilitados adhiriéndose a los calendarios de cada laboratorio.
      Se encuentran en la misma página donde solicitar reservas así también instrucciones [Opcional].
      */
      break;
    case MOTIVO_GESTION:
      /*
      Su solicitud está en gestión.
 
      {Datos}
 
      En los próximos días recibirá un correo informando su estado dependiendo de qué tan cercana sea la fecha solicitada a la actual y a la prioridad otorgada por la administración.
      Recuerde que esta solicitud no garantiza una reserva.
      */
      break;
  }
  try {
    MailApp.sendEmail(datos.email, cuerpoAsunto, {
      noReply: true, //Para evitar que se responda a estos correos
      htmlBody: encabezadoHtml + cuerpoCorreo + pieHtml
    });
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No se envío el correo. Requiere correr arregloEnEnvioFormulario o enviarCorreoAuto con los parametros en el registro.',
      'Error de GmailApp en gestión al recibir formulario.', -1);
    Logger.log('Error de GmailApp en gestión al recibir formulario. ' + error.name + '\n' + error.message + '\nParametros para ejecución manual: \nagregarDatos(Fila):' + fila + '\nMotivo: ' + motivo);
  } finally {
    return true;
  }
};

function enviarCorreo(datos, motivo, fila) {
  var encabezadoHtml = '<center><table border=1 style=background-color:#c4302b; width=80% height = 50><div >' +
    '<a href = https://www.frba.utn.edu.ar/que-significa-el-logo-de-la-utn-ba/ >' +
    '<img alt= utn src=https://www.frba.utn.edu.ar/wp-content/uploads/2016/08/logo-utn.ba-horizontal-e1471367724904.jpg width=80 height=35 > </a></div> </table></center>';
  var cuerpoAsunto = 'Respuesta Solicitud De Reserva UTN Laboratorio', cuerpoCorreo = '';
  //Agregar Muchas gracias o Atte al pie
  var pieHtml = '<center><footer><p > <strong> <h4 style= color:#c4302b;> Ing. Ramiro Garbarini </h4> ·  Jefe de Laboratorios · Laboratorio de Sistemas de Información </strong> </p> </footer></center>';

  var f = datos.fechaInicio;
  var fecha = DIA_SEMANA[f.getDay()] + " " + f.getDate() + " de " + MES[f.getMonth()];

  switch (motivo) {
    case MOTIVO_ACEPTAR:
      cuerpoCorreo = '';
      /*
      Su solicitud de reserva fue aceptada.
 
      {Datos}
 
      Recuerde siempre, de no poder asistir, comunicarse aquí {insertar info de contacto} para cancelar su reserva. 
      Nos sería de gran ayuda para seguir brindando nuestros servicios.
      */
      break;
    case MOTIVO_RECHAZADO_X_ERROR:
      cuerpoCorreo = '';
      /*
      Su solicitud de reserva fue aceptada.
 
      Debido a un inconveniente recibió por error un correo rechazando su solicitud. 
      Le pedimos que lo desestime.
 
      {Datos}
 
      Recuerde siempre, de no poder asistir, comunicarse aquí {insertar info de contacto} para cancelar su reserva. 
      Nos sería de gran ayuda para seguir brindando nuestros servicios.
      */
      break;
    case MOTIVO_CANCELADO_X_ERROR:
      cuerpoCorreo = '';
      /*
      Su solicitud de reserva fue aceptada.
 
      Debido a un inconveniente recibió por error un correo cancelando su solicitud. 
      Le pedimos que lo desestime.
 
      {Datos}
 
      Recuerde siempre, de no poder asistir, comunicarse aquí {insertar info de contacto} para cancelar su reserva. 
      Nos sería de gran ayuda para seguir brindando nuestros servicios.
        */
      break;
    case MOTIVO_RECHAZO:
      cuerpoCorreo = '';
      /*
      Lo sentimos, su solicitud de reserva fue rechazada a criterio de la administración, debido a una gran demanda de solicitudes en el turno, fecha y laboratorio seleccionado.
 
      {Datos}
      */
      break;
    case MOTIVO_CANCELAR:
      cuerpoCorreo = '';
      /*
      Lo sentimos, su solicitud de reserva fue cancelada, debido a motivos internos de la facultad en ese día, turno y laboratorio.
 
      Le pedimos disculpas por los inconvenientes ocasionados.
 
      {Datos}
        */
      break;
    case MOTIVO_CANCELADO_X_USUARIO:
      cuerpoCorreo = '';
    /*
    Su solicitud fue cancelada a su pedido.
 
    {Datos}
 
    De ser esto un error contáctese con alguien que le importe {insertar contacto de alguien que le importe} para gestionar el estado de su solicitud.
      */
    case MOTIVO_DISCULPA_X_VENCIMIENTO:
      cuerpoCorreo = '';
      /*
      Lo sentimos, su solicitud de reserva fue cancelada.
 
      Debido a problemas en el sistema no concretamos gestionar su solicitud a tiempo. Nos disculpamos por los inconvenientes ocasionados.
 
      {Datos}
      */
      break;
    case ASUNTO_ENCUESTA:
      cuerpoAsunto = 'Encuesta De Satisfacción UTN';
      cuerpoCorreo = '';
      /*
      Esperamos que su reserva del día {día y fecha del mes}
      turno {Turno} en el laboratorio {nombre Lab} haya sido de agrado.
 
      Por favor si es tan amable de ingresar al siguiente enlace y llenar la siguiente encuesta:
      Ingrese aquí {https://forms.gle/eTEGsberRkRhMFLK8}
      
      Nos ayudaría a mejorar el servicio.
      */
      break;
  }
  try {
    if (asunto != ASUNTO_ENCUESTA && datos.correoPersonal) {
      GmailApp.createDraft(datos.email, cuerpoAsunto, {//Si no funciona puede ser mejor enviar un correo a quien deba enviarlo y este lo modifica y lo reenvía
        name: Session.getActiveUser().getEmail(), //Ver como se ve esto o agregar un enviado por
        htmlBody: encabezadoHtml + cuerpoCorreo + pieHtml //Ver si es posible dejarlo listo para editar y meterle texto dentro
      });
    } else {
      MailApp.sendEmail(datos.email, cuerpoAsunto, {
        name: Session.getActiveUser().getEmail(), //Ver como se ve esto o agregar un enviado por
        htmlBody: encabezadoHtml + cuerpoCorreo + pieHtml
      });
    }
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No se envío el correo. Requiere correr enviarCorreo con los parametros en el registro.',
      'Error de GmailApp en gestión al Editar.', -1);
    Logger.log('Error de GmailApp en gestión al Editar. ' + error.name + '\n' + error.message + '\nParametros para ejecución manual: \nagregarDatos(Fila):' + fila + '\nMotivo: ' + motivo);
  }
};

//Crea y devuelve un objeto con todos los parametros de datos de la fila deseada
function agregarDatos(fila) {
  var nombreEscrito, materiaEscrita;
  if (HOJA.getRange(fila, COLUMNA_NOMBRE).getValue() != '') {
    nombreEscrito = HOJA.getRange(fila, COLUMNA_NOMBRE).getValue();
  } else {
    nombreEscrito = HOJA.getRange(fila, COLUMNA_PROFESOR).getValue();
  }
  if (HOJA.getRange(fila, COLUMNA_MATERIA_ESCRITA).getValue() != '') {
    materiaEscrita = HOJA.getRange(fila, COLUMNA_MATERIA_ESCRITA).getValue();
  } else {
    materiaEscrita = HOJA.getRange(fila, COLUMNA_MATERIA).getValue();
  }
  var cadena = HOJA.getRange(fila, COLUMNA_LABORATORIO).getValue();
  var sedeLabo = cadena.split('-');
  try {
    var datos = {
      email: HOJA.getRange(fila, COLUMNA_CORREO).getValue(),
      nombre: nombreEscrito,
      materia: materiaEscrita,
      sede: sedeLabo[0],
      lab: sedeLabo[1],
      fechaInicio: new Date(HOJA.getRange(fila, COLUMNA_FECHA).getValue()),
      fechaFin: new Date(HOJA.getRange(fila, COLUMNA_FECHA).getValue()),
      turno: HOJA.getRange(fila, COLUMNA_TURNO).getValue(),
      cantPersonas: HOJA.getRange(fila, COLUMNA_CANTIDAD).getValue(),
      peticiones: HOJA.getRange(fila, COLUMNA_PETICION).getValue(),
      calendario: CalendarApp.getCalendarById('gaston@ecologix.com.ar'),
      calendarioSolicitud: CalendarApp.getCalendarById('gaston@ecologix.com.ar'),
      recurso: 'nulo',
      idSolicitud: HOJA.getRange(fila, COLUMNA_ID_SOLICITUD).getValue(),
      idReserva: HOJA.getRange(fila, COLUMNA_ID_RESERVA).getValue(),
      correoPersonal: HOJA.getRange(fila, COLUMNA_CORREO_PERSONAL).getValue()
    };
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Requiere consultar el registro.',
      'Error de CalendarApp o formato fecha Date en obtención de datos.', -1);
    Logger.log('Error de CalendarApp o formato fecha Date en obtención de datos. ' + error.name + '\n' + error.message);
  }
  try {
    switch (datos.turno) {
      case 'Mañana':
        datos.fechaInicio.setHours(8);
        datos.fechaInicio.setMinutes(30);
        datos.fechaFin.setHours(13);
        datos.fechaFin.setMinutes(15);
        break;
      case 'Tarde':
        datos.fechaInicio.setHours(13);
        datos.fechaInicio.setMinutes(15);
        datos.fechaFin.setHours(18);
        datos.fechaFin.setMinutes(00);
        break;
      case 'Noche':
        datos.fechaInicio.setHours(18);
        datos.fechaInicio.setMinutes(00);
        datos.fechaFin.setHours(22);
        datos.fechaFin.setMinutes(15);
        break;
      default:
        Logger.log(datos.turno, 'Error en cadena Turno en asiganciones de horario');
        break;
    }
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Requiere consultar el registro por seteo incorrecto de variables Date.',
      'Error de seteo de horario de fecha en obtención de datos.', -1);
    Logger.log('Error de seteo de horario de fecha en obtención de datos. ' + error.name + '\n' + error.message +
      '\ndatos.fechaInicio: ' + datos.fechaInicio + '\ndatos.fechaFin: ' + datos.fechaFin);
  }
  /*Se asignan los Ids del calendario y recurso correspondiente, Desde:
  calendario/configuarción/Integrar el calendario/Id de calendario
  Y Administración/Edificios y recursos/Abrir(Gestión de rescursos)/{Recurso}/Correo electrónico del recurso */
  try {
    switch (datos.lab) {
      case 'Azul':
        datos.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.calendarioSolicitud = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.recurso = 'Insertar id de recurso de Google';
        break;
      case 'Naranja':
        datos.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.calendarioSolicitud = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.recurso = 'Insertar id de recurso de Google';
        break;
      case 'Rojo':
        datos.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.calendarioSolicitud = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.recurso = 'Insertar id de recurso de Google';
        break;
      case 'Verde':
        datos.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.calendarioSolicitud = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.recurso = 'Insertar id de recurso de Google';
        break;
      case 'WorkGroup Lab 1':
        datos.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.calendarioSolicitud = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.recurso = 'Insertar id de recurso de Google';
        break;
      case 'Campus':
        datos.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.calendarioSolicitud = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.recurso = 'Insertar id de recurso de Google';
        break;
      case 'WorkGroup Lab 2':
        datos.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.calendarioSolicitud = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        datos.recurso = 'Insertar id de recurso de Google';
        break;
      default:
        Logger.log(datos.lab, 'Error en cadena Lab en asignación a calendario');
        break;
    }
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No se puede gestionar. Esperar y volver a intentar.',
      'Error de CalendarApp al asignar calendarios en obtención de datos.', -1);
    Logger.log('Error de CalendarApp al asignar calendarios en obtención de datos. ' + error.name + '\n' + error.message);
  }
  return datos;
};

//Crea un objeto con datos necesarios para habilitar o deshabilitar un turno específico
function deshabilitaciones(e) {
  var labo;
  if (HOJA2.getRange(e.range.getRow(), 2).getValue() == '') {
    if (HOJA2.getRange(e.range.getRow() - 1, 2).getValue() == '') {
      labo = HOJA2.getRange(e.range.getRow() - 2, 2).getValue();
    } else {
      labo = HOJA2.getRange(e.range.getRow() - 1, 2).getValue();
    }
  } else {
    labo = HOJA2.getRange(e.range.getRow(), 2).getValue();
  }
  try {
    var evento = {
      lab: labo,
      turno: HOJA2.getRange(e.range.getRow(), 4).getValue(),
      dia: HOJA2.getRange(2, e.range.getColumn()).getValue(),
      fechaInicio: new Date(),
      fechaFin: new Date(),
      calendario: CalendarApp.getCalendarById('gaston@ecologix.com.ar'),
      id: HOJA2.getRange(e.range.getRow() + 27, e.range.getColumn()).getValue()
    };
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Requiere consultar el registro.',
      'Error de CalendarApp o formato fecha Date en obtención de datos.', -1);
    Logger.log('Error de CalendarApp o formato fecha Date en obtención de datos. ' + error.name + '\n' + error.message);
  }
  try {
    diaHoy = new Date().getDay();
    diferencia = evento.dia - diaHoy;
    evento.fechaInicio.setDate(evento.fechaInicio.getDate() + diferencia);
    evento.fechaFin = evento.fechaInicio;
    switch (evento.turno) {
      case 'Mañana':
        evento.fechaInicio.setHours(8);
        evento.fechaInicio.setMinutes(30);
        evento.fechaFin.setHours(13);
        evento.fechaFin.setMinutes(15);
        break;
      case 'Tarde':
        evento.fechaInicio.setHours(13);
        evento.fechaInicio.setMinutes(15);
        evento.fechaFin.setHours(18);
        evento.fechaFin.setMinutes(00);
        break;
      case 'Noche':
        evento.fechaInicio.setHours(18);
        evento.fechaInicio.setMinutes(00);
        evento.fechaFin.setHours(22);
        evento.fechaFin.setMinutes(15);
        break;
      default:
        Logger.log(evento.turno, 'Error en cadena Turno en asiganciones de horario');
        break;
    }
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Requiere consultar el registro por seteo incorrecto de variables Date.',
      'Error de seteo de horario de fecha en obtención de datos.', -1);
    Logger.log('Error de seteo de horario de fecha en obtención de datos. ' + error.name + '\n' + error.message +
      '\nevento.fechaInicio: ' + evento.fechaInicio + '\nevento.fechaFin: ' + evento.fechaFin);
  }
  try {
    switch (evento.lab) {
      case 'Azul':
        evento.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        break;
      case 'Naranja':
        evento.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        break;
      case 'Rojo':
        evento.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        break;
      case 'Verde':
        evento.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        break;
      case 'WorkGroup Lab 1':
        evento.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        break;
      case 'Campus':
        evento.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        break;
      case 'WorkGroup Lab 2':
        evento.calendario = CalendarApp.getCalendarById('Insertar id de calendario de Google');
        break;
      default:
        Logger.log(evento.lab, 'Linea 1123 \nError en cadena Lab en asignación a calendario');
        break;
    }
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No se puede gestionar. Esperar y volver a intentar.',
      'Error de CalendarApp al asignar calendarios en obtención de datos.', -1);
    Logger.log('Error de CalendarApp al asignar calendarios en obtención de datos. ' + error.name + '\n' + error.message);
    evento.calendario = null;
  }
  return evento;
};

function deshabilitarTurno(evento, e) {
  if (evento.calendario == null) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No es posible deshabilitar el turno. Arreglo necesario.', 'Error en asiganción a calendario del laboratorio.', -1);
    Logger.log('Línea 1132 \nError en asiganción a calendario del laboratorio.' +
      'evento.calendario: ' + evento.calendario + '\nLab: ' + evento.lab + '\nFecha: ' + evento.fechaInicio + '\nTurno: ' + evento.turno);
    e.range.insertCheckboxes().check();
  } else {
    if (ventana(evento, 'Deshabilitar')) {
      var deshabilitacion = evento.calendario.createEventSeries('Turno Deshabilitado', evento.fechaInicio, evento.fechaFin, CalendarApp.newRecurrence()
        .addWeeklyRule().until(new Date(evento.fechaInicio.getFullYear() + 1, 0, 1)))
        .setVisibility(CalendarApp.Visibility.PUBLIC);
      HOJA2.getRange(e.range.getRow() + 27, e.range.getColumn()).setValue(deshabilitacion.getId());
      SpreadsheetApp.getActiveSpreadsheet().toast('Laboratorio: ' + evento.lab + '\nDía: ' + DIA_SEMANA[evento.dia] + '\nTurno: ' +
        evento.turno, 'Turno Deshabilitado', 10);
    } else {
      e.range.insertCheckboxes().check();
    }
  }
};

function habilitarTurno(evento, e) {
  if (evento.calendario == null) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No es posible habilitar el turno. Arreglo necesario.', 'Error en asiganción a calendario del laboratorio.', -1);
    Logger.log('Línea 1152 \nError en asiganción a calendario del laboratorio.' +
      'evento.calendario: ' + evento.calendario + '\nLab: ' + evento.lab + '\nFecha: ' + evento.fechaInicio + '\nTurno: ' + evento.turno);
    e.range.insertCheckboxes().uncheck();
  } else {
    if (ventana(evento, 'Habilitar')) {
      var deshabilitacion = evento.calendario.getEventSeriesById(evento.id);
      if (deshabilitacion != null) {
        deshabilitacion.deleteEventSeries();
      }
      HOJA2.getRange(e.range.getRow() + 27, e.range.getColumn()).setValue('');
      SpreadsheetApp.getActiveSpreadsheet().toast('Laboratorio: ' + evento.lab + '\nDía: ' + DIA_SEMANA[evento.dia] + '\nTurno: ' +
        evento.turno, 'Turno Habilitado', 10);
    } else {
      e.range.insertCheckboxes();
    }
  }
};