// ============================================================
//  AUTOMATIZACIÓN DE GESTIÓN DE INCIDENCIAS MÉDICAS
//  Google Apps Script
// ============================================================
//
//  CICLO COMPLETO (todo dentro de un único hilo de Gmail):
//   1. Recibimos un correo con una incidencia (p. ej. de un
//      sistema de tickets / autoservicio de clientes).
//   2. Lo REENVIAMOS al proveedor correspondiente en el mismo
//      hilo (forward), determinado automáticamente por la
//      especialidad/servicio y/o el médico indicado.
//   3a. El proveedor responde con "Caso cerrado" → reenviamos
//       la resolución al solicitante original (mismo hilo).
//   3b. El proveedor responde SIN "Caso cerrado" → reenviamos
//       la respuesta a un área interna para revisión manual.
//
//  ACTIVADORES (Editor de Apps Script > Activadores):
//   • procesarIncidencias              → cada 15 min
//   • procesarRespuestasProveedores    → cada 15 min
//   • gestionarRecordatorios           → cada día (09:00)
//   • resumenSemanal (opcional)        → cada viernes
//
//  PROTECCIONES IMPLEMENTADAS:
//   • LockService en procesarIncidencias y procesarRespuestasProveedores.
//   • Etiquetado del hilo ANTES del procesado (no se reprocesa nunca).
//   • Deduplicación por SR (identificador de caso) en registrarEnSheet.
//   • Captura multilínea de campos del cuerpo del correo.
//   • Match de médicos por palabras completas (tokens).
//   • Distinción solicitante / proveedor en las respuestas.
//   • Filtrado de auto-respuestas (out-of-office, auto-reply).
//   • Recordatorios SLA no duplicados (PropertiesService).
//   • Log persistente en una pestaña "Log" del Sheet de proveedores.
//
// ============================================================
// ============================================================
//  SECCIÓN CONFIG — RELLENA AQUÍ TUS DATOS ANTES DE EJECUTAR
// ============================================================
const CONFIG = {
  // IDs de los Google Sheets (los sacas de la URL del propio Sheet:
  // https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit)
  SHEET_ID_PROVEEDOR1:      "TU_SHEET_ID_PROVEEDOR_1",
  SHEET_ID_PROVEEDOR2:           "TU_SHEET_ID_PROVEEDOR_2",
  // Google Sheet con UNA pestaña por cada proveedor "resto".
  // Pestañas esperadas: PROVEEDOR3 | PROVEEDOR4 | PROVEEDOR5 | PROVEEDOR6 | PROVEEDOR7 | Log
  SHEET_ID_OTROS_PROVEEDORES: "TU_SHEET_ID_OTROS_PROVEEDORES",
  // Google Sheet unificado del cuadro médico (listado de médicos por proveedor).
  // Columnas: PROVEEDOR | NOMBRE | APELLIDO1 | APELLIDO2 | ESPECIALIDAD
  SHEET_ID_CUADRO_MEDICO:     "TU_SHEET_ID_CUADRO_MEDICO",
  HOJA_CUADRO_MEDICO:         "Hoja1",  // si tu pestaña se llama distinto, cámbialo aquí
  HOJA_INCIDENCIAS_PROVEEDOR1: "Incidencias",
  HOJA_INCIDENCIAS_PROVEEDOR2:      "Segumiento incidencias",
  // ── Filtro temporal de incidencias entrantes ──
  // El script solo procesa correos cuya fecha esté en los últimos N días
  // (incluyendo hoy). Con DIAS_VENTANA = 1 se procesa ayer + hoy.
  DIAS_VENTANA: 1,
  // Pestaña de log persistente dentro de SHEET_ID_OTROS_PROVEEDORES
  HOJA_LOG: "Log",
  // Pestaña donde se registran las incidencias clasificadas como TÉCNICA
  // (NO se reenvían a proveedor; requieren acción manual del equipo interno).
  HOJA_TECNICAS: "TECNICAS",
  // Mapa proveedor → nombre de la pestaña dentro de SHEET_ID_OTROS_PROVEEDORES
  HOJAS_OTROS_PROVEEDORES: {
    PROVEEDOR3:        "PROVEEDOR3",
    PROVEEDOR4:    "PROVEEDOR4",
    PROVEEDOR5: "PROVEEDOR5",
    PROVEEDOR6:       "PROVEEDOR6",
    PROVEEDOR7:           "PROVEEDOR7",
  },
  // Buzón de Gmail que recibe las incidencias entrantes (el que ejecuta el script)
  CORREO_ENTRADA: "tu_correo@tudominio.com",
  ASUNTO_FILTRO:  "Nueva gestión de Autoservicio Caso",
  // Destinatarios por proveedor: a quién se reenvía cada tipo de incidencia
  PROVEEDOR2:        { to: "contacto_proveedor2@proveedor.com", cc: "" },
  PROVEEDOR1:   { to: "contacto_proveedor1@proveedor.com", cc: "" },
  PROVEEDOR3:        { to: "contacto_proveedor3@proveedor.com", cc: "" },
  PROVEEDOR4:    { to: "contacto_proveedor4@proveedor.com", cc: "" },
  PROVEEDOR5: { to: "contacto_proveedor5@proveedor.com", cc: "" },
  PROVEEDOR6:       { to: "contacto_proveedor6@proveedor.com", cc: "" },
  PROVEEDOR7:           { to: "contacto_proveedor7@proveedor.com", cc: "" },
  // Las alertas (TÉCNICA, ROUTING, REVISIÓN PROVEEDOR, errores y resumen
  // semanal) siempre se envían al propio CORREO_ENTRADA para centralizarlas
  // donde llegan también las incidencias. Aquí solo configuramos el cc
  // opcional. El "to" de cualquier alerta es CONFIG.CORREO_ENTRADA.
  ALERTA: {
    cc: ""
  },
  // ── ETIQUETAS DE GMAIL ──
  LABEL_PROCESADO:         "Incidencia_Procesada",
  LABEL_CIERRE_NOTIFICADO: "Caso_Cerrado_Notificado",
  LABEL_REVISION_MANUAL:   "Revision_Manual_Proveedor",
  // Etiqueta que se aplica AUTOMÁTICAMENTE a cada correo de alerta interna
  // ([ACCIÓN MANUAL - TÉCNICA / ROUTING / REVISIÓN PROVEEDOR]). Sirve para
  // que procesarIncidencias NO los reprocese como si fueran incidencias nuevas.
  LABEL_ALERTA:            "Alerta_Sistema",
  // ── Texto que dispara el cierre ──
  TEXTO_CIERRE: "Caso cerrado",
  // ── SLA ──
  RECORDATORIOS_DIAS: [3, 5, 7, 14],
  RECORDATORIO_PERIODICO_DIAS: 7,
};
// Lista de proveedores que se registran en SHEET_ID_OTROS_PROVEEDORES (16 cols)
const PROVEEDORES_OTROS = ["PROVEEDOR3", "PROVEEDOR4", "PROVEEDOR5", "PROVEEDOR6", "PROVEEDOR7"];
// Lista de NOMBRES DE CAMPO conocidos en el cuerpo del correo,
// usados para que extraerCampo capture varias líneas hasta el siguiente campo.
const CAMPOS_CONOCIDOS = [
  // SR / Caso
  "Caso",
  // Identificación del solicitante
  "ID solicitante",
  "DNI",
  // Nombre del paciente
  "Nombre",
  "Apellido 1",
  "Apellido 2",
  "Apellidos",
  // Email
  "Correo del usuario",
  "Correo",
  "Email",
  "E-mail",
  // Especialidad
  "Especialidad",
  // Médico
  "Nombre del médico",
  "Doctor",
  "Doctora",
  "Médico",
  // Día de la cita
  "Día de la video consulta",
  "Día de la videoconsulta",
  "Fecha de la cita",
  "Fecha cita",
  "Fecha",
  // Hora de la cita
  "Hora de la video consulta",
  "Hora de la videoconsulta",
  "Hora de inicio",
  "Hora inicio",
  // Texto / descripción
  "Texto del mensaje",
  "Incidencia",
  "Datos faltantes",
  "Motivo",
  // Servicio
  "Servicio",
  "Tipo de servicio",
];
// ============================================================
//  MAPA SERVICIO → PROVEEDOR(ES)
// ============================================================
const NOMBRES_PROVEEDORES = {
  PROVEEDOR3:         "Proveedor 3",
  PROVEEDOR2:         "Proveedor 2",
  PROVEEDOR1:    "Proveedor 1",
  PROVEEDOR4:     "Proveedor 4",
  PROVEEDOR5:  "Proveedor 5",
  PROVEEDOR6:        "Proveedor 6",
  PROVEEDOR7:            "Proveedor 7",
};
const MAPA_SERVICIOS = {
  // Proveedor 3
  "medicina general":              ["PROVEEDOR3"],
  "interpretacion de analiticas":  ["PROVEEDOR3"],
  "interpretacion analiticas":     ["PROVEEDOR3"],
  "analiticas":                    ["PROVEEDOR3"],
  "post operatoria":               ["PROVEEDOR3"],
  "postoperatoria":                ["PROVEEDOR3"],
  // Proveedor 1 (exclusivo)
  "neurologia": ["PROVEEDOR1"],
  "neumologia": ["PROVEEDOR1"],
  // Proveedor 2 + Proveedor 1
  "endocrinologia": ["PROVEEDOR2", "PROVEEDOR1"],
  "dermatologia":   ["PROVEEDOR2", "PROVEEDOR1"],
  "traumatologia":  ["PROVEEDOR2", "PROVEEDOR1"],
  "digestivo":      ["PROVEEDOR2", "PROVEEDOR1"],
  "cardiologia":    ["PROVEEDOR2", "PROVEEDOR1"],
  "alergologia":    ["PROVEEDOR2", "PROVEEDOR1"],
  "ginecologia":    ["PROVEEDOR2", "PROVEEDOR1"],
  "urologia":       ["PROVEEDOR2", "PROVEEDOR1"],
  "psicologia":     ["PROVEEDOR2", "PROVEEDOR1"],
  "pediatria":      ["PROVEEDOR2", "PROVEEDOR1"],
  // Proveedor 4
  "urgencias":               ["PROVEEDOR4"],
  "videoconsulta inmediata": ["PROVEEDOR4"],
  "inmediata":               ["PROVEEDOR4"],
  // Proveedor 5
  "chat de orientacion":     ["PROVEEDOR5"],
  "orientacion diagnostica": ["PROVEEDOR5"],
  "chat":                    ["PROVEEDOR5"],
  // Proveedor 6
  "evaluador de sintomas": ["PROVEEDOR6"],
  "evaluador sintomas":    ["PROVEEDOR6"],
  "sintomas":              ["PROVEEDOR6"],
  // PROVEEDOR7
  "fisioterapia digital": ["PROVEEDOR7"],
  "fisioterapia":         ["PROVEEDOR7"],
  "proveedor7":                 ["PROVEEDOR7"],
};
// ============================================================
//  PASO 1 — PROCESAR INCIDENCIAS ENTRANTES (cada 15 min)
// ============================================================
function procesarIncidencias() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log("[INFO] procesarIncidencias: otra ejecución en curso, salimos.");
    return;
  }
  try {
    const label = obtenerOCrearLabel(CONFIG.LABEL_PROCESADO);
    // Filtro temporal: solo correos a partir de hace DIAS_VENTANA días.
    // Gmail interpreta after:YYYY/MM/DD como "ese día o posterior".
    const dias = (typeof CONFIG.DIAS_VENTANA === "number") ? CONFIG.DIAS_VENTANA : 1;
    const desde = new Date();
    desde.setDate(desde.getDate() - dias);
    const desdeStr = Utilities.formatDate(desde, Session.getScriptTimeZone(), "yyyy/MM/dd");
    // Doble protección contra reprocesado de alertas:
    //   • -label:Alerta_Sistema  → ignora cualquier hilo ya marcado como alerta
    //   • -subject:"[ACCIÓN MANUAL"  → cinturón extra por si la etiqueta no
    //     llegó a aplicarse (Gmail tarda en indexar)
    const query =
      `to:${CONFIG.CORREO_ENTRADA} ` +
      `subject:"${CONFIG.ASUNTO_FILTRO}" ` +
      `-subject:"[ACCIÓN MANUAL" ` +
      `-label:${CONFIG.LABEL_ALERTA} ` +
      `after:${desdeStr} ` +
      `-label:${CONFIG.LABEL_PROCESADO}`;
    const hilos = GmailApp.search(query, 0, 50);
    if (hilos.length === 0) return;
    const wbP1 = SpreadsheetApp.openById(CONFIG.SHEET_ID_PROVEEDOR1);
    const wbP2 = SpreadsheetApp.openById(CONFIG.SHEET_ID_PROVEEDOR2);
    const wbOtros = SpreadsheetApp.openById(CONFIG.SHEET_ID_OTROS_PROVEEDORES);
    const hojaIncP1 = wbP1.getSheetByName(CONFIG.HOJA_INCIDENCIAS_PROVEEDOR1);
    const hojaIncP2 = wbP2.getSheetByName(CONFIG.HOJA_INCIDENCIAS_PROVEEDOR2);
    // ── Carga del cuadro médico (soporta .xlsx y Google Sheet nativo) ──
    // Si es .xlsx, convertimos a Google Sheet temporal, leemos y borramos.
    // Resultado cacheado durante 1h para no convertir en cada ejecución.
    const cuadro = cargarMedicosCuadroUnificadoConCache();
    const medicosP1 = cuadro.medicosP1;
    const medicosP2 = cuadro.medicosP2;
    hilos.forEach(hilo => {
      const labelsHilo = hilo.getLabels().map(l => l.getName());
      if (labelsHilo.indexOf(CONFIG.LABEL_PROCESADO) >= 0) {
        Logger.log(`[SKIP] Hilo ${hilo.getId()} ya tenía etiqueta de procesado.`);
        return;
      }
      // Etiquetar PRIMERO para idempotencia
      hilo.addLabel(label);
      try {
        const mensaje = hilo.getMessages()[0];
        const asunto  = mensaje.getSubject();
        const cuerpo  = mensaje.getPlainBody();
        const fecha   = mensaje.getDate();
        const datos = extraerDatosCorreo(asunto, cuerpo, fecha);
        // ── INCIDENCIA TÉCNICA ──
        // NO se reenvía a proveedor. Se registra en la pestaña TECNICAS y
        // se manda alerta al equipo interno con [ACCIÓN MANUAL] en el asunto.
        if (datos.tipoIncidencia === "TÉCNICA") {
          const hojaTec = wbOtros.getSheetByName(CONFIG.HOJA_TECNICAS);
          if (!hojaTec) {
            _logSheet("WARN", "procesarIncidencias",
              `No existe la pestaña "${CONFIG.HOJA_TECNICAS}". Crea las cabeceras con crearCabecerasOtrosProveedores().`);
          } else {
            registrarEnSheet(hojaTec, datos, "TECNICAS");
          }
          enviarAlertaTecnica(datos, hilo, mensaje);
          _logSheet("OK", "procesarIncidencias",
            `Caso ${datos.sr} → TÉCNICA / ${datos.subcategoria || "GENERAL"} → ACCIÓN MANUAL`);
          return;
        }
        const proveedor = determinarProveedor(datos, medicosP1, medicosP2);
        switch (proveedor) {
          case "PROVEEDOR1":
            registrarEnSheet(hojaIncP1, datos, "PROVEEDOR1");
            enviarAlProveedor("PROVEEDOR1", hilo, mensaje, datos);
            break;
          case "PROVEEDOR2":
            registrarEnSheet(hojaIncP2, datos, "PROVEEDOR2");
            enviarAlProveedor("PROVEEDOR2", hilo, mensaje, datos);
            break;
          case "PROVEEDOR3":
          case "PROVEEDOR4":
          case "PROVEEDOR5":
          case "PROVEEDOR6":
          case "PROVEEDOR7": {
            const nombreHoja = CONFIG.HOJAS_OTROS_PROVEEDORES[proveedor];
            const hojaOtro = wbOtros.getSheetByName(nombreHoja);
            if (!hojaOtro) {
              _logSheet("WARN", "procesarIncidencias", `No existe la pestaña "${nombreHoja}".`);
            } else {
              registrarEnSheet(hojaOtro, datos, proveedor);
            }
            enviarAlProveedor(proveedor, hilo, mensaje, datos);
            break;
          }
          default:
            enviarAlertaManual(datos, hilo, mensaje, "MEDICO_NO_ENCONTRADO");
        }
        _logSheet("OK", "procesarIncidencias",
          `Caso ${datos.sr} → ${proveedor || "ALERTA MANUAL"}`);
      } catch (e) {
        _logSheet("ERROR", "procesarIncidencias",
          `Hilo ${hilo.getId()}: ${e.message}`);
        try {
          GmailApp.sendEmail(
            CONFIG.CORREO_ENTRADA,
            `⚠️ Error procesando incidencia (hilo ${hilo.getId()})`,
            `El hilo se marcó como procesado para evitar duplicados, pero la ejecución falló:\n\n${e.message}\n\n${e.stack || ""}`,
            { cc: CONFIG.ALERTA.cc, name: "Sistema Automatización Salud Digital" }
          );
        } catch (_) { /* no-op */ }
      }
    });
  } finally {
    lock.releaseLock();
  }
}
// ============================================================
//  PASO 3 — PROCESAR RESPUESTAS DE PROVEEDORES (cada 15 min)
// ============================================================
function procesarRespuestasProveedores() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log("[INFO] procesarRespuestasProveedores: otra ejecución en curso, salimos.");
    return;
  }
  try {
    const labelCierre = obtenerOCrearLabel(CONFIG.LABEL_CIERRE_NOTIFICADO);
    const labelManual = obtenerOCrearLabel(CONFIG.LABEL_REVISION_MANUAL);
    const query =
      `label:${CONFIG.LABEL_PROCESADO} ` +
      `-label:${CONFIG.LABEL_CIERRE_NOTIFICADO} ` +
      `-label:${CONFIG.LABEL_REVISION_MANUAL}`;
    const hilos = GmailApp.search(query, 0, 50);
    if (hilos.length === 0) return;
    const correoEntrada = CONFIG.CORREO_ENTRADA.toLowerCase();
    hilos.forEach(hilo => {
      try {
        const mensajes = hilo.getMessages();
        if (mensajes.length < 2) return; // todavía no hay respuesta
        const primerMensaje = mensajes[0];
        // ── B2: distinguir solicitante vs proveedor ──
        const emailSolicitante = extraerEmail(primerMensaje.getFrom()).toLowerCase();
        let respuestaProveedor = null;
        for (let i = mensajes.length - 1; i > 0; i--) {
          const m = mensajes[i];
          const fromEmail = extraerEmail(m.getFrom() || "").toLowerCase();
          // Ignorar nuestros propios mensajes y los del solicitante original
          if (!fromEmail) continue;
          if (fromEmail === correoEntrada) continue;
          if (fromEmail === emailSolicitante) continue;
          respuestaProveedor = m;
          break;
        }
        if (!respuestaProveedor) return;
        // ── R3: filtrar auto-respuestas (OOO / autoreply) ──
        if (_esAutoRespuesta(respuestaProveedor)) {
          Logger.log(`[SKIP] Auto-respuesta ignorada en hilo ${hilo.getId()}`);
          return;
        }
        const cuerpoRespuesta = respuestaProveedor.getPlainBody() || "";
        const asuntoRespuesta = respuestaProveedor.getSubject() || "";
        const sr =
          extraerSRdeAsunto(asuntoRespuesta) ||
          extraerSRdeAsunto(primerMensaje.getSubject()) ||
          "N/A";
        // ── Detección robusta del cierre ──
        // Normaliza acentos, puntuación, NBSP, espacios múltiples...
        const contieneCierre = _contieneFraseNormalizada(cuerpoRespuesta, CONFIG.TEXTO_CIERRE);
        // Log de diagnóstico: primeros 300 caracteres normalizados
        _logSheet(contieneCierre ? "DEBUG_CIERRE" : "DEBUG_NO_CIERRE",
          "procesarRespuestas",
          `${sr} · cuerpoNorm="${_normalizarFrase(cuerpoRespuesta).substring(0, 300)}"`);
        if (contieneCierre) {
          if (!emailSolicitante) {
            _logSheet("WARN", "procesarRespuestas",
              `Sin email del solicitante para ${sr}`);
            return;
          }
          const cuerpoTexto =
            `Hola,\n\n` +
            `Os trasladamos la resolución del caso. Más abajo tenéis el historial completo del hilo con la respuesta del proveedor.\n\n` +
            `Un saludo.`;
          respuestaProveedor.forward(emailSolicitante, {
            name: "Salud Digital Mapfre",
            htmlBody: _construirHtmlConHilo(hilo, cuerpoTexto)
          });
          cerrarCasoEnSheet(sr);
          hilo.addLabel(labelCierre);
          _logSheet("OK", "procesarRespuestas",
            `Resolución de ${sr} → ${emailSolicitante}`);
        } else {
          const cuerpoTexto =
            `⚠️ REVISIÓN MANUAL — Respuesta del proveedor\n\n` +
            `El proveedor ha respondido al caso ${sr} pero la respuesta NO contiene la confirmación "Caso cerrado".\n\n` +
            `Por favor, revisad si el caso está realmente resuelto, requiere acciones adicionales o un nuevo contacto con el proveedor antes de trasladar la resolución al cliente. Más abajo tenéis el historial completo del hilo.\n\n` +
            `Un saludo,\nSistema Automatización Salud Digital.`;
          // Correo nuevo + etiqueta Alerta_Sistema para que no se reprocese.
          _enviarAlertaInterna(
            _asuntoAccionManual(respuestaProveedor.getSubject(), "REVISIÓN PROVEEDOR"),
            cuerpoTexto,
            _construirHtmlConHilo(hilo, cuerpoTexto)
          );
          hilo.addLabel(labelManual);
          _logSheet("REV_MANUAL", "procesarRespuestas",
            `${sr} → área médica`);
        }
      } catch (e) {
        _logSheet("ERROR", "procesarRespuestas",
          `Hilo ${hilo.getId()}: ${e.message}`);
      }
    });
  } finally {
    lock.releaseLock();
  }
}
// ============================================================
//  ENVÍO AL PROVEEDOR (forward → mismo hilo de Gmail)
// ============================================================
function enviarAlProveedor(proveedor, hilo, mensaje, datos) {
  const cfg = CONFIG[proveedor];
  if (!cfg || !cfg.to) {
    _logSheet("WARN", "enviarAlProveedor",
      `Sin destinatario para ${proveedor}. Caso ${datos.sr} no enviado.`);
    return;
  }
  const nombreProveedor = NOMBRES_PROVEEDORES[proveedor] || proveedor;
  const cuerpoTexto =
    `Hola ${nombreProveedor},\n\n` +
    `Os derivamos la siguiente incidencia recibida desde el área de clientes. ` +
    `Os agradeceríamos que la reviséis y nos confirméis estado/resolución a la mayor brevedad.\n\n` +
    `RESUMEN DE LA INCIDENCIA\n` +
    `• Caso (SR):      ${datos.sr || "—"}\n` +
    `• Paciente:       ${[datos.nombre, datos.apellido1, datos.apellido2].filter(Boolean).join(" ") || "—"}\n` +
    `• DNI:            ${datos.dni || "—"}\n` +
    `• Email paciente: ${datos.emailPaciente || "—"}\n` +
    `• Especialidad:   ${datos.especialidad || "—"}\n` +
    `• Médico:         ${datos.medico || "—"}\n` +
    `• Cita:           ${[datos.diaCita, datos.horaCita].filter(Boolean).join(" ") || "—"}\n` +
    `• Tipo:           ${datos.tipoIncidencia || "—"}\n\n` +
    `DESCRIPCIÓN DEL CLIENTE\n${datos.textoIncidencia || "(sin texto)"}\n\n` +
    `Quedamos a la espera de vuestra respuesta.\n\n` +
    `Un saludo,\nEquipo de Salud Digital.`;
  mensaje.forward(cfg.to, {
    cc: cfg.cc,
    name: "Salud Digital Mapfre",
    htmlBody: _construirHtmlConHilo(hilo, cuerpoTexto)
  });
}
// ============================================================
//  ENVÍO DE ALERTA MANUAL (cuando no se puede enrutar al proveedor)
// ============================================================
function enviarAlertaManual(datos, hilo, mensaje, motivo) {
  const motivoTexto = motivo === "MEDICO_NO_ENCONTRADO"
    ? `No se ha podido determinar el proveedor a partir de la especialidad/servicio ` +
      `("${datos.especialidad}") y el médico indicado ("${datos.medico}") tampoco aparece ` +
      `en los listados de proveedores.`
    : `Caso que requiere atención manual.`;
  const cuerpoTexto =
    `⚠️ ACCIÓN MANUAL REQUERIDA — Routing no resuelto\n\n` +
    `El sistema de automatización ha detectado un caso que requiere intervención humana.\n\n` +
    `DETALLES DEL CASO:\n` +
    `• ID Caso (SR): ${datos.sr}\n` +
    `• Paciente: ${datos.nombre} ${datos.apellido1} ${datos.apellido2} (DNI: ${datos.dni})\n` +
    `• Médico reportado: ${datos.medico}\n` +
    `• Especialidad: ${datos.especialidad}\n` +
    `• Tipo incidencia detectado: ${datos.tipoIncidencia} / ${datos.subcategoria || "GENERAL"}\n\n` +
    `MOTIVO DE LA ALERTA:\n${motivoTexto}\n\n` +
    `ACCIONES REQUERIDAS:\n` +
    `1. Revisar el hilo completo (incluido más abajo).\n` +
    `2. Si es un médico nuevo, actualizar el Listado de Médicos en el Google Sheet correspondiente.\n` +
    `3. Gestionar manualmente la incidencia con el proveedor correcto.`;
  // Correo nuevo (hilo independiente) con etiqueta Alerta_Sistema aplicada.
  _enviarAlertaInterna(
    _asuntoAccionManual(mensaje.getSubject(), "ROUTING"),
    cuerpoTexto,
    _construirHtmlConHilo(hilo, cuerpoTexto)
  );
}
// Reenvía la incidencia al buzón interno cuando ha sido clasificada como TÉCNICA.
// No se envía a ningún proveedor.
function enviarAlertaTecnica(datos, hilo, mensaje) {
  const cuerpoTexto =
    `⚠️ ACCIÓN MANUAL REQUERIDA — Incidencia TÉCNICA\n\n` +
    `El sistema de automatización ha clasificado esta incidencia como TÉCNICA. ` +
    `NO se ha reenviado a ningún proveedor médico — requiere gestión por el equipo interno.\n\n` +
    `DETALLES DEL CASO:\n` +
    `• ID Caso (SR): ${datos.sr || "—"}\n` +
    `• Paciente: ${[datos.nombre, datos.apellido1, datos.apellido2].filter(Boolean).join(" ") || "—"} (DNI: ${datos.dni || "—"})\n` +
    `• Email paciente: ${datos.emailPaciente || "—"}\n` +
    `• Especialidad/Servicio reportado: ${datos.especialidad || "—"} / ${datos.tipoServicio || "—"}\n` +
    `• Tipo: ${datos.tipoIncidencia} / ${datos.subcategoria || "GENERAL"}\n\n` +
    `DESCRIPCIÓN DEL CLIENTE\n${datos.textoIncidencia || "(sin texto)"}\n\n` +
    `ACCIONES POSIBLES:\n` +
    `1. Si es una incidencia técnica/administrativa real (app, login, cobertura, etc.) → ` +
    `redirigir al equipo correspondiente (IT, atención al cliente, administración).\n` +
    `2. Si está mal clasificada y en realidad es asistencial → reenviar manualmente al proveedor médico que toque.\n` +
    `3. El caso queda registrado en la pestaña "${CONFIG.HOJA_TECNICAS}" del Sheet de OTROS_PROVEEDORES.`;
  // Correo nuevo + etiqueta Alerta_Sistema para que no se reprocese.
  _enviarAlertaInterna(
    _asuntoAccionManual(mensaje.getSubject(), `TÉCNICA / ${datos.subcategoria || "GENERAL"}`),
    cuerpoTexto,
    _construirHtmlConHilo(hilo, cuerpoTexto)
  );
}
// ============================================================
//  ENVÍO DE ALERTA INTERNA + ETIQUETADO AUTOMÁTICO
// ============================================================
//
//  Envía un correo nuevo a CORREO_ENTRADA (hilo independiente del de la
//  incidencia, para que el asunto [ACCIÓN MANUAL] se vea sin abrir el
//  correo) y le aplica la etiqueta LABEL_ALERTA. Así procesarIncidencias
//  NO lo confunde luego con una incidencia nueva, aunque el asunto
//  contenga "Nueva gestión de Autoservicio Caso".
//
function _enviarAlertaInterna(asunto, cuerpoTexto, htmlBody) {
  GmailApp.sendEmail(
    CONFIG.CORREO_ENTRADA,
    asunto,
    cuerpoTexto,
    {
      cc: CONFIG.ALERTA.cc,
      name: "Sistema Automatización Salud Digital",
      htmlBody: htmlBody
    }
  );
  // Etiquetar el hilo recién creado.
  try {
    const label = obtenerOCrearLabel(CONFIG.LABEL_ALERTA);
    // Gmail puede tardar 1-2 s en indexar. Esperamos brevemente.
    Utilities.sleep(1500);
    const asuntoQuery = asunto.replace(/"/g, '\\"');
    const hilos = GmailApp.search(
      `subject:"${asuntoQuery}" newer_than:1d to:${CONFIG.CORREO_ENTRADA}`,
      0, 5
    );
    if (hilos.length === 0) {
      Logger.log(`[WARN] No se encontró el hilo recién enviado para etiquetar: "${asunto}"`);
    } else {
      hilos.forEach(h => h.addLabel(label));
    }
  } catch (e) {
    Logger.log(`[WARN] No se pudo aplicar la etiqueta de alerta: ${e.message}`);
  }
}
// Helper de mantenimiento: aplica LABEL_ALERTA a TODAS las alertas
// ya existentes en la bandeja (las anteriores a este cambio) para
// que dejen de ser candidatas a reprocesado. Ejecútalo una vez
// manualmente desde el editor.
function etiquetarAlertasExistentes() {
  const label = obtenerOCrearLabel(CONFIG.LABEL_ALERTA);
  const hilos = GmailApp.search(
    `subject:"[ACCIÓN MANUAL" -label:${CONFIG.LABEL_ALERTA}`,
    0, 200
  );
  hilos.forEach(h => h.addLabel(label));
  Logger.log(`[INFO] Etiquetadas ${hilos.length} alertas existentes con "${CONFIG.LABEL_ALERTA}".`);
}
// Construye un asunto con prefijo [ACCIÓN MANUAL - <tipo>]. Quita prefijos previos
// y los "Fwd:" / "RE:" iniciales para que no se acumulen al re-procesar el hilo.
function _asuntoAccionManual(asuntoOriginal, tipo) {
  const tipoUp = (tipo || "").toString().toUpperCase().trim();
  const prefijo = tipoUp ? `[ACCIÓN MANUAL - ${tipoUp}]` : `[ACCIÓN MANUAL]`;
  let limpio = (asuntoOriginal || "").toString();
  // Quita acciones manuales previas
  limpio = limpio.replace(/^(\[\s*ACCI[ÓO]N\s+MANUAL[^\]]*\]\s*)+/gi, "");
  // Quita Fwd:/RE: iniciales
  limpio = limpio.replace(/^(\s*(fwd|re|fw|rv)\s*:\s*)+/i, "");
  return `${prefijo} ${limpio}`.trim();
}
// ============================================================
//  CIERRE DEL CASO EN EL GOOGLE SHEET
// ============================================================
function cerrarCasoEnSheet(sr) {
  if (!sr || sr === "N/A") return;
  const ahora = new Date();
  const sheets = [
    {
      hoja: SpreadsheetApp.openById(CONFIG.SHEET_ID_PROVEEDOR1)
              .getSheetByName(CONFIG.HOJA_INCIDENCIAS_PROVEEDOR1),
      proveedor: "PROVEEDOR1"
    },
    {
      hoja: SpreadsheetApp.openById(CONFIG.SHEET_ID_PROVEEDOR2)
              .getSheetByName(CONFIG.HOJA_INCIDENCIAS_PROVEEDOR2),
      proveedor: "PROVEEDOR2"
    }
  ];
  const wbOtros = SpreadsheetApp.openById(CONFIG.SHEET_ID_OTROS_PROVEEDORES);
  PROVEEDORES_OTROS.forEach(prov => {
    const nombreHoja = CONFIG.HOJAS_OTROS_PROVEEDORES[prov];
    const hoja = wbOtros.getSheetByName(nombreHoja);
    if (hoja) sheets.push({ hoja, proveedor: prov });
  });
  sheets.forEach(({ hoja, proveedor }) => {
    if (!hoja) return;
    const datos    = hoja.getDataRange().getValues();
    if (datos.length < 1) return;
    const cabecera = datos[0];
    const idxSR         = cabecera.indexOf("SR");
    const idxResolucion = cabecera.indexOf("FECHA RESOLUCION");
    const idxEstado     = cabecera.indexOf("ESTADO");
    if (idxSR < 0 || idxResolucion < 0) return;
    for (let i = 1; i < datos.length; i++) {
      if ((datos[i][idxSR] || "").toString().toUpperCase() === sr.toUpperCase()) {
        if (!datos[i][idxResolucion]) {
          hoja.getRange(i + 1, idxResolucion + 1).setValue(ahora);
        }
        if (idxEstado >= 0 && (proveedor === "PROVEEDOR2" || PROVEEDORES_OTROS.indexOf(proveedor) >= 0)) {
          hoja.getRange(i + 1, idxEstado + 1).setValue("RESUELTO");
        }
        Logger.log(`[OK] Caso ${sr} cerrado en Sheet ${proveedor}`);
        return;
      }
    }
  });
  // Limpiar marcadores de recordatorios para evitar acumular Properties
  _limpiarRecordatoriosSR(sr);
}
// ============================================================
//  EXTRACCIÓN DE DATOS DEL CORREO
// ============================================================
function extraerDatosCorreo(asunto, cuerpo, fecha) {
  const clasif = clasificarTipoIncidencia(cuerpo);
  // Helper interno: prueba varios nombres de campo y devuelve el primer match.
  const primero = (...candidatos) => {
    for (let i = 0; i < candidatos.length; i++) {
      const v = extraerCampo(cuerpo, candidatos[i]);
      if (v) return v;
    }
    return "";
  };
  return {
    sr:            extraerCampo(cuerpo, "Caso") || extraerSRdeAsunto(asunto) || "N/A",
    dni:           primero("ID solicitante", "DNI"),
    nombre:        extraerCampo(cuerpo, "Nombre"),
    apellido1:     primero("Apellido 1", "Apellidos"),
    apellido2:     extraerCampo(cuerpo, "Apellido 2"),
    emailPaciente: primero("Correo del usuario", "Correo", "Email", "E-mail"),
    especialidad:  extraerCampo(cuerpo, "Especialidad"),
    medico:        primero("Nombre del médico", "Doctor", "Doctora", "Médico"),
    diaCita:       primero(
                      "Día de la video consulta",
                      "Día de la videoconsulta",
                      "Fecha de la cita",
                      "Fecha cita",
                      "Fecha"
                   ),
    horaCita:      primero(
                      "Hora de la video consulta",
                      "Hora de la videoconsulta",
                      "Hora de inicio",
                      "Hora inicio"
                   ),
    textoIncidencia: primero("Texto del mensaje", "Datos faltantes", "Incidencia", "Motivo"),
    tipoServicio:  primero("Servicio", "Tipo de servicio"),
    tipoIncidencia: clasif.tipo,
    subcategoria:  clasif.subcategoria,
    fechaCorreo:   fecha,
    asunto:        asunto,
    cuerpoOriginal: cuerpo,
  };
}
// ── B3: extracción multilínea con stop en el siguiente campo conocido ──
function extraerCampo(cuerpo, campo) {
  if (!cuerpo || !campo) return "";
  const escapaRegex = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const campoEsc = escapaRegex(campo);
  // Set de campos "siguientes" candidatos: cualquiera distinto del actual
  const siguientes = CAMPOS_CONOCIDOS
    .filter(c => c.toLowerCase() !== campo.toLowerCase())
    .map(escapaRegex)
    .join("|");
  // Captura: [-•]?<campo>:<contenido multilínea no goloso> hasta nuevo campo o fin
  const regex = new RegExp(
    `[-•]?\\s*${campoEsc}\\s*:\\s*([\\s\\S]*?)(?=\\n\\s*[-•]?\\s*(?:${siguientes})\\s*:|$)`,
    "i"
  );
  const match = cuerpo.match(regex);
  return match ? match[1].trim() : "";
}
// ── B1: ahora devuelve "" si no encuentra SR (no "N/A") para que || funcione ──
function extraerSRdeAsunto(asunto) {
  if (!asunto) return "";
  const match = asunto.match(/SR\d+/i);
  return match ? match[0].toUpperCase() : "";
}
// ============================================================
//  CATÁLOGO DE KEYWORDS / FRASES POR (CATEGORÍA, SUBCATEGORÍA)
// ============================================================
//
//  • Las keywords se escriben SIN acentos (la normalización las quita).
//  • El matching se hace con BORDES DE PALABRA: se enmarcan tanto el
//    cuerpo como cada keyword con espacios, así "login" no matchea
//    "alógino" y "no puedo entrar" se busca como frase completa.
//  • Ganadora = categoría con al menos 1 match; dentro de la ganadora,
//    subcategoría con MÁS matches. Empate → la primera por orden.
//  • Si no hay ningún match → ASISTENCIAL / GENERAL.
//
const CATEGORIZACION_INCIDENCIAS = {
  TECNICA: {
    LOGIN_ACCESO: [
      "login", "logout", "password", "contrasena", "credencial", "credenciales",
      "no puedo entrar", "no me deja entrar", "no me reconoce",
      "olvide mi contrasena", "no recuerdo mi contrasena",
      "iniciar sesion", "cerrar sesion", "no consigo iniciar sesion",
      "fallo de autenticacion", "error de autenticacion",
      "no puedo acceder a la app", "no puedo acceder a mi cuenta",
    ],
    APP_FALLO: [
      "crashea", "crashee", "crash",
      "la app no funciona", "la app no carga", "la app no abre",
      "se cierra la app", "se cierra sola", "se cierra solo",
      "la aplicacion no funciona", "no abre la app", "no carga la app",
      "pantalla en blanco", "pantalla en negro",
      "se queda colgado", "se queda colgada", "se queda cargando",
      "se bloquea", "boton no responde", "la app va lenta",
    ],
    COMUNICACIONES: [
      "no me llega el sms", "no me llega el codigo",
      "no recibo el correo", "no recibo el email", "no recibo el sms",
      "no me llega la notificacion", "no me llega el aviso",
      "no me ha llegado el correo", "no me ha llegado el sms",
      "no me ha llegado el codigo", "no me llega el email",
    ],
    CONEXION_VIDEOCONSULTA: [
      "no se conecta", "no me conecta", "no hay conexion", "sin conexion",
      "se corta la videoconsulta", "se corta la videollamada", "se corta la llamada",
      "videoconsulta no funciona", "videollamada no funciona",
      "fallo de conexion", "sin senal", "no consigo conectar",
      "audio no funciona", "video no funciona", "no se oye", "no se ve",
      "mala calidad de imagen", "mala calidad de sonido",
    ],
    ERROR_TECNICO: [
      "error tecnico", "mensaje de error", "codigo de error",
      "error al cargar", "error al iniciar", "error en la app",
      "no responde el sistema", "error 500", "error 404",
      "no funciona correctamente", "fallo de sistema",
    ],
    COBERTURA_POLIZA: [
      "copago", "cobertura", "poliza", "facturacion", "factura",
      "no esta cubierto", "no me cubre",
      "fuera de cobertura", "fuera de poliza", "limite de poliza",
      "exceso de poliza", "no me llega la factura",
      "no me reembolsan", "no me han reembolsado",
      "cuestion de cobertura", "tema administrativo",
    ],
  },
  ASISTENCIAL: {
    NO_PRESENTADO: [
      "no se conecto", "no aparecio", "no me llamo", "no atendio",
      "nadie respondio", "nadie me contacto",
      "el medico no se conecto", "el medico no aparecio",
      "el medico no se presento", "no se presento", "no asistio",
      "la doctora no se conecto", "el doctor no se conecto",
    ],
    DEMORA: [
      "tardo mucho", "espere mucho", "espere mas de",
      "tiempo de espera", "demasiado tiempo", "muy lento",
      "lleva sin contestar", "horas de espera", "minutos de espera",
    ],
    CITA: [
      "cancelar cita", "reprogramar cita", "cambiar la cita",
      "modificar la cita", "cancelacion de cita",
      "no puedo agendar", "no puedo reservar", "cita cancelada",
      "anular cita", "anular la cita", "asignar cita",
    ],
    TRATAMIENTO_INFORME: [
      "receta", "diagnostico", "tratamiento", "informe", "prescripcion",
      "no me dieron el informe", "no he recibido el informe",
      "no me llego la receta", "informe medico",
      "discrepancia clinica", "segunda opinion",
    ],
    ATENCION_MEDICA: [
      "mala atencion", "trato inadecuado", "trato incorrecto",
      "mal trato", "queja sobre la consulta",
      "atencion deficiente", "atencion incorrecta",
    ],
  },
};
// Clasificación principal: devuelve { tipo, subcategoria }.
function clasificarTipoIncidencia(cuerpo) {
  const cuerpoNorm = _normalizarFrase(cuerpo);
  if (!cuerpoNorm) return { tipo: "ASISTENCIAL", subcategoria: "GENERAL" };
  // Enmarcamos con espacios para hacer matching de palabra completa
  const cuerpoBordes = ` ${cuerpoNorm} `;
  const conteo = {};
  for (const categoria in CATEGORIZACION_INCIDENCIAS) {
    const subs = CATEGORIZACION_INCIDENCIAS[categoria];
    conteo[categoria] = {};
    for (const sub in subs) {
      const lista = subs[sub];
      let count = 0;
      for (let i = 0; i < lista.length; i++) {
        const kwNorm = _normalizarFrase(lista[i]);
        if (!kwNorm) continue;
        if (cuerpoBordes.indexOf(` ${kwNorm} `) >= 0) count++;
      }
      if (count > 0) conteo[categoria][sub] = count;
    }
  }
  const hayTecnica     = Object.keys(conteo.TECNICA     || {}).length > 0;
  const hayAsistencial = Object.keys(conteo.ASISTENCIAL || {}).length > 0;
  // Si hay match en TÉCNICA, gana TÉCNICA (es lo más específico).
  if (hayTecnica) {
    return { tipo: "TÉCNICA", subcategoria: _subcategoriaGanadora(conteo.TECNICA) };
  }
  if (hayAsistencial) {
    return { tipo: "ASISTENCIAL", subcategoria: _subcategoriaGanadora(conteo.ASISTENCIAL) };
  }
  return { tipo: "ASISTENCIAL", subcategoria: "GENERAL" };
}
function _subcategoriaGanadora(conteoPorSub) {
  let mejorSub = "";
  let mejorCount = -1;
  for (const sub in conteoPorSub) {
    if (conteoPorSub[sub] > mejorCount) {
      mejorCount = conteoPorSub[sub];
      mejorSub = sub;
    }
  }
  return mejorSub;
}
// ============================================================
//  ROUTING POR SERVICIO → PROVEEDOR
// ============================================================
function normalizarTexto(t) {
  return (t || "").toString().toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ").trim();
}
function proveedoresPorServicio(datos) {
  const texto = normalizarTexto(
    [datos.especialidad, datos.tipoServicio, datos.asunto].filter(Boolean).join(" | ")
  );
  if (!texto) return [];
  const encontrados = new Set();
  for (const clave in MAPA_SERVICIOS) {
    if (texto.includes(clave)) {
      MAPA_SERVICIOS[clave].forEach(p => encontrados.add(p));
    }
  }
  return [...encontrados];
}
function determinarProveedor(datos, medicosP1, medicosP2) {
  const candidatos = proveedoresPorServicio(datos);
  if (candidatos.length === 1) return candidatos[0];
  if (candidatos.length > 1) {
    if (candidatos.includes("PROVEEDOR1") && buscarMedico(datos.medico, medicosP1)) return "PROVEEDOR1";
    if (candidatos.includes("PROVEEDOR2")      && buscarMedico(datos.medico, medicosP2)) return "PROVEEDOR2";
    return null;
  }
  if (buscarMedico(datos.medico, medicosP1)) return "PROVEEDOR1";
  if (buscarMedico(datos.medico, medicosP2)) return "PROVEEDOR2";
  return null;
}
// ============================================================
//  BÚSQUEDA DE MÉDICOS EN LOS LISTADOS
// ============================================================
// ── Carga unificada del cuadro médico con caché y soporte .xlsx ──
//
// La fuente (CONFIG.SHEET_ID_CUADRO_MEDICO) puede ser:
//   • Google Sheet nativa → se lee directamente.
//   • .xlsx subido a Drive → se convierte a Google Sheet TEMPORAL,
//     se lee y se elimina la copia. Requiere "Drive API" como
//     servicio avanzado habilitado en el editor de Apps Script.
//
// La tabla tiene columnas:
//   A: PROVEEDOR ("PROVEEDOR2" | "PROVEEDOR1")
//   B: NOMBRE
//   C: APELLIDO1
//   D: APELLIDO2
//   E: ESPECIALIDAD (informativa, no se usa aquí)
//
// El resultado se cachea 1 hora (CacheService) para no convertir el
// .xlsx en cada ejecución del trigger.
const CACHE_KEY_CUADRO = "cuadro_medico_v1";
const CACHE_TTL_SEGUNDOS = 3600;  // 1 hora
function cargarMedicosCuadroUnificadoConCache() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_CUADRO);
  if (cached) {
    try {
      const obj = JSON.parse(cached);
      Logger.log(`[INFO] Cuadro médico desde caché: ${obj.medicosP2.length} PROVEEDOR2 + ${obj.medicosP1.length} PROVEEDOR1`);
      return obj;
    } catch (e) { /* cache corrupto, recargamos */ }
  }
  const { hoja, tempId } = _obtenerHojaCuadroMedico();
  try {
    const resultado = cargarMedicosCuadroUnificado(hoja);
    try {
      cache.put(CACHE_KEY_CUADRO, JSON.stringify(resultado), CACHE_TTL_SEGUNDOS);
    } catch (e) {
      Logger.log(`[WARN] No se pudo cachear el cuadro médico: ${e.message}`);
    }
    return resultado;
  } finally {
    // Borrar la copia temporal de la conversión .xlsx → Google Sheet
    if (tempId) {
      try {
        DriveApp.getFileById(tempId).setTrashed(true);
      } catch (e) {
        Logger.log(`[WARN] No se pudo borrar la copia temporal del cuadro (${tempId}): ${e.message}`);
      }
    }
  }
}
// Obtiene la hoja del cuadro médico. Si el fichero origen es un .xlsx,
// crea una copia temporal en formato Google Sheet (a borrar después).
// Devuelve { hoja, tempId }. tempId es null si no se creó copia.
function _obtenerHojaCuadroMedico() {
  const fileId = CONFIG.SHEET_ID_CUADRO_MEDICO;
  const file = DriveApp.getFileById(fileId);
  const mime = file.getMimeType();
  // 1) Google Sheet nativa → abrir directamente.
  if (mime === MimeType.GOOGLE_SHEETS) {
    const ss = SpreadsheetApp.openById(fileId);
    const hoja = ss.getSheetByName(CONFIG.HOJA_CUADRO_MEDICO) || ss.getSheets()[0];
    return { hoja: hoja, tempId: null };
  }
  // 2) .xlsx (o similar) → convertir a Google Sheet temporal vía Drive Advanced API.
  if (typeof Drive === "undefined" || !Drive.Files || typeof Drive.Files.insert !== "function") {
    throw new Error(
      `El cuadro médico es un .xlsx (mime "${mime}") y necesita el servicio avanzado "Drive API" ` +
      `habilitado en el editor de Apps Script. ` +
      `Pasos: Editor → Servicios (icono "+") → busca "Drive API" → Añadir.`
    );
  }
  const resource = {
    title: "_cuadro_medico_temporal_" + new Date().getTime(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  const blob = file.getBlob();
  const converted = Drive.Files.insert(resource, blob, { convert: true });
  Logger.log(`[INFO] Cuadro médico .xlsx convertido a Sheet temporal (${converted.id}).`);
  const ss = SpreadsheetApp.openById(converted.id);
  const hoja = ss.getSheetByName(CONFIG.HOJA_CUADRO_MEDICO) || ss.getSheets()[0];
  return { hoja: hoja, tempId: converted.id };
}
// Procesa la hoja (ya en formato Google Sheet) y devuelve los arrays.
function cargarMedicosCuadroUnificado(hoja) {
  if (!hoja) {
    Logger.log("[WARN] cargarMedicosCuadroUnificado: hoja nula.");
    return { medicosP2: [], medicosP1: [] };
  }
  const datos = hoja.getDataRange().getValues();
  const medicosP2 = [];
  const medicosP1 = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila || fila.length === 0) continue;
    const proveedorRaw = (fila[0] || "").toString().toUpperCase().trim();
    if (!proveedorRaw) continue;
    const nombreCompleto = [fila[1], fila[2], fila[3]].filter(Boolean).join(" ");
    const normalizado = normalizarNombre(nombreCompleto);
    if (!normalizado) continue;
    if (proveedorRaw.indexOf("PROVEEDOR2") >= 0) {
      medicosP2.push(normalizado);
    } else if (proveedorRaw.indexOf("PROVEEDOR1") >= 0) {
      medicosP1.push(normalizado);
    } else {
      Logger.log(`[WARN] Proveedor desconocido en cuadro médico fila ${i + 1}: "${proveedorRaw}"`);
    }
  }
  Logger.log(`[INFO] Cuadro médico cargado: ${medicosP2.length} PROVEEDOR2 + ${medicosP1.length} PROVEEDOR1`);
  return { medicosP2, medicosP1 };
}
// Función auxiliar opcional: forzar la recarga del cuadro médico
// limpiando la caché. Útil si actualizas el .xlsx y quieres que el
// cambio se aplique inmediatamente en lugar de esperar 1h al TTL.
function invalidarCacheCuadroMedico() {
  CacheService.getScriptCache().remove(CACHE_KEY_CUADRO);
  Logger.log("[INFO] Caché del cuadro médico invalidada.");
}
function normalizarNombre(nombre) {
  return (nombre || "").toString().toUpperCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\b(DR|DRA|DON|DOÑA)\.?\s*/gi, "")
    .replace(/[^A-Z0-9\s]/g, " ")
    .replace(/\s+/g, " ").trim();
}
// ── B4: match por tokens completos (palabras enteras) ──
// Reglas:
//   • Igualdad exacta normalizada → match.
//   • Si el buscado tiene ≥2 tokens y todos sus tokens están en el listado → match.
//   • Si el listado tiene ≥2 tokens y todos sus tokens están en el buscado → match.
//   • Tokens de 1 letra se ignoran (iniciales).
function buscarMedico(nombreCorreo, listadoNormalizado) {
  if (!nombreCorreo) return false;
  const buscado = normalizarNombre(nombreCorreo);
  if (!buscado) return false;
  const tokensBuscado = buscado.split(/\s+/).filter(t => t.length >= 2);
  if (tokensBuscado.length === 0) return false;
  return listadoNormalizado.some(nombre => {
    if (!nombre) return false;
    if (nombre === buscado) return true;
    const tokensListado = nombre.split(/\s+/).filter(t => t.length >= 2);
    if (tokensListado.length === 0) return false;
    if (tokensBuscado.length >= 2 && tokensBuscado.every(t => tokensListado.indexOf(t) >= 0)) {
      return true;
    }
    if (tokensListado.length >= 2 && tokensListado.every(t => tokensBuscado.indexOf(t) >= 0)) {
      return true;
    }
    return false;
  });
}
// ============================================================
//  REGISTRO EN GOOGLE SHEETS
// ============================================================
function _existeSRenSheet(hoja, sr) {
  if (!sr || sr === "N/A") return false;
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return false;
  const cabecera = datos[0];
  const idxSR = cabecera.indexOf("SR");
  if (idxSR < 0) return false;
  const srUp = sr.toString().toUpperCase();
  for (let i = 1; i < datos.length; i++) {
    if ((datos[i][idxSR] || "").toString().toUpperCase() === srUp) return true;
  }
  return false;
}
function registrarEnSheet(hoja, datos, proveedor) {
  const ahora = new Date();
  if (_existeSRenSheet(hoja, datos.sr)) {
    Logger.log(`[SKIP] Caso ${datos.sr} ya estaba registrado en ${proveedor}. No se duplica.`);
    return;
  }
  if (proveedor === "PROVEEDOR1") {
    hoja.appendRow([
      "PENDIENTE", "ORIGEN", datos.fechaCorreo, datos.asunto,
      datos.sr, datos.dni, datos.textoIncidencia, datos.tipoIncidencia,
      datos.especialidad, ahora,
    ]);
    return;
  }
  if (proveedor === "PROVEEDOR2") {
    hoja.appendRow([
      "PENDIENTE", "ORIGEN", datos.fechaCorreo, datos.asunto,
      datos.sr, datos.dni, datos.tipoIncidencia, datos.textoIncidencia,
      ahora,
    ]);
    return;
  }
  if (PROVEEDORES_OTROS.indexOf(proveedor) >= 0 || proveedor === "TECNICAS") {
    hoja.appendRow([
      "PENDIENTE",            // ESTADO
      "ORIGEN",               // ORIGEN
      datos.fechaCorreo,      // FECHA CORREO
      datos.asunto,           // ASUNTO CORREO
      datos.sr,               // SR
      datos.dni,              // DNI
      datos.tipoIncidencia,   // TIPO INCIDENCIA (TÉCNICA/ASISTENCIAL)
      datos.textoIncidencia,  // OBSERVACIONES
      ahora,                  // FECHA DE RECEPCION
      "",                     // FECHA RESOLUCION
      "",                     // Responsable proveedor
      "",                     // Responsable interno
      "",                     // Responsable área médica
      datos.especialidad,     // Especialidad
      "",                     // Solución
      "",                     // FECHA CONTESTACIÓN
      datos.subcategoria || "", // Col 17 extra: SUBCATEGORIA
    ]);
    return;
  }
  Logger.log(`[WARN] registrarEnSheet: proveedor desconocido "${proveedor}". No se registra.`);
}
// ============================================================
//  RECORDATORIOS SLA
// ============================================================
function gestionarRecordatorios() {
  const ahora = new Date();
  _procesarRecordatoriosSheet(
    SpreadsheetApp.openById(CONFIG.SHEET_ID_PROVEEDOR1)
      .getSheetByName(CONFIG.HOJA_INCIDENCIAS_PROVEEDOR1),
    "PROVEEDOR1", ahora
  );
  _procesarRecordatoriosSheet(
    SpreadsheetApp.openById(CONFIG.SHEET_ID_PROVEEDOR2)
      .getSheetByName(CONFIG.HOJA_INCIDENCIAS_PROVEEDOR2),
    "PROVEEDOR2", ahora
  );
  const wbOtros = SpreadsheetApp.openById(CONFIG.SHEET_ID_OTROS_PROVEEDORES);
  PROVEEDORES_OTROS.forEach(prov => {
    const nombreHoja = CONFIG.HOJAS_OTROS_PROVEEDORES[prov];
    const hoja = wbOtros.getSheetByName(nombreHoja);
    if (hoja) _procesarRecordatoriosSheet(hoja, prov, ahora);
  });
}
function _procesarRecordatoriosSheet(hoja, proveedor, ahora) {
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return;
  const cabecera = datos[0];
  const tieneEstado = (proveedor === "PROVEEDOR2") || (PROVEEDORES_OTROS.indexOf(proveedor) >= 0);
  const idxEstado     = tieneEstado ? cabecera.indexOf("ESTADO") : -1;
  const idxFechaRec   = cabecera.indexOf("FECHA DE RECEPCION");
  const idxSR         = cabecera.indexOf("SR");
  const idxDNI        = cabecera.indexOf("DNI");
  const idxAsunto     = cabecera.indexOf("ASUNTO CORREO");
  const idxResolucion = cabecera.indexOf("FECHA RESOLUCION");
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const fechaResolucion = fila[idxResolucion];
    if (fechaResolucion && fechaResolucion !== "") continue;
    if (idxEstado >= 0) {
      const estado = (fila[idxEstado] || "").toString().toUpperCase();
      if (estado === "RESOLUCION" || estado === "RESUELTO" || estado === "CERRADO") continue;
    }
    const fechaRecepcion = fila[idxFechaRec];
    if (!fechaRecepcion || !(fechaRecepcion instanceof Date)) continue;
    const dias = Math.floor((ahora - fechaRecepcion) / (1000 * 60 * 60 * 24));
    const sr     = (fila[idxSR]     || "N/A").toString();
    const dni    = fila[idxDNI]    || "N/A";
    const asunto = fila[idxAsunto] || `Seguimiento incidencia ${sr}`;
    if (!_tocaRecordatorio(dias)) continue;
    // ── R2: anti-duplicado de recordatorios ──
    if (_yaSeEnvioRecordatorio(sr, dias)) {
      Logger.log(`[SKIP] Recordatorio ya enviado para ${sr} (${dias} días)`);
      continue;
    }
    _enviarRecordatorio(proveedor, sr, dni, asunto, dias);
    _marcarRecordatorioEnviado(sr, dias);
    _logSheet("RECORD", "gestionarRecordatorios",
      `Caso ${sr} → ${proveedor} (${dias} días)`);
  }
}
function _tocaRecordatorio(dias) {
  if (CONFIG.RECORDATORIOS_DIAS.includes(dias)) return true;
  const ultimoFijo = Math.max(...CONFIG.RECORDATORIOS_DIAS);
  if (dias > ultimoFijo && (dias - ultimoFijo) % CONFIG.RECORDATORIO_PERIODICO_DIAS === 0) return true;
  return false;
}
function _enviarRecordatorio(proveedor, sr, dni, asuntoOriginal, dias) {
  const cfg = CONFIG[proveedor];
  if (!cfg || !cfg.to) {
    Logger.log(`[WARN] Sin destinatario para recordatorio ${proveedor} / ${sr}`);
    return;
  }
  const nombreProveedor = NOMBRES_PROVEEDORES[proveedor] || proveedor;
  let asunto, cuerpo;
  if (dias <= 3) {
    asunto = `RE: ${asuntoOriginal} (RECORDATORIO ${dias * 24}H)`;
    cuerpo =
      `Hola ${nombreProveedor},\n\n` +
      `Os escribimos en relación a la incidencia del paciente con DNI ${dni} ` +
      `(Caso ${sr}) enviada hace ${dias} día(s).\n\n` +
      `¿Podríais confirmarnos el estado de la resolución?\n\n` +
      `Un saludo,\nEquipo de Salud Digital.`;
  } else if (dias <= 7) {
    asunto = `URGENTE: Seguimiento Incidencia ${sr} - Sin resolución tras ${dias} días`;
    cuerpo =
      `Hola ${nombreProveedor},\n\n` +
      `No hemos recibido actualización sobre el caso ${sr}.\n\n` +
      `Es necesario disponer de respuesta lo antes posible para evitar una reclamación formal.\n\n` +
      `Atentamente,\nEquipo de Salud Digital.`;
  } else {
    asunto = `RECLAMACIÓN PENDIENTE: Caso ${sr} - Fuera de SLA (${dias} días)`;
    cuerpo =
      `Hola ${nombreProveedor},\n\n` +
      `El caso ${sr} continúa abierto sin resolución desde hace ${dias} días.\n\n` +
      `Rogamos máxima prioridad para cerrar este caso.\n\n` +
      `Un saludo,\nEquipo de Salud Digital.`;
  }
  GmailApp.sendEmail(cfg.to, asunto, cuerpo, {
    cc: cfg.cc,
    name: "Salud Digital Mapfre"
  });
}
// ── R2: tracking de recordatorios via PropertiesService ──
function _claveRecordatorio(sr, dias) {
  return `reminder:${(sr || "").toString().toUpperCase()}:${dias}`;
}
function _yaSeEnvioRecordatorio(sr, dias) {
  if (!sr || sr === "N/A") return false;
  return PropertiesService.getScriptProperties().getProperty(_claveRecordatorio(sr, dias)) === "1";
}
function _marcarRecordatorioEnviado(sr, dias) {
  if (!sr || sr === "N/A") return;
  PropertiesService.getScriptProperties().setProperty(_claveRecordatorio(sr, dias), "1");
}
function _limpiarRecordatoriosSR(sr) {
  if (!sr || sr === "N/A") return;
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const prefijo = `reminder:${sr.toString().toUpperCase()}:`;
  Object.keys(all).forEach(key => {
    if (key.indexOf(prefijo) === 0) props.deleteProperty(key);
  });
}
// ============================================================
//  RESUMEN SEMANAL
// ============================================================
function resumenSemanal() {
  const ahora = new Date();
  const hace7 = new Date(ahora.getTime() - 7 * 24 * 60 * 60 * 1000);
  const resumenMC = _generarResumenSheet(
    SpreadsheetApp.openById(CONFIG.SHEET_ID_PROVEEDOR1)
      .getSheetByName(CONFIG.HOJA_INCIDENCIAS_PROVEEDOR1),
    hace7, ahora
  );
  const resumenDL = _generarResumenSheet(
    SpreadsheetApp.openById(CONFIG.SHEET_ID_PROVEEDOR2)
      .getSheetByName(CONFIG.HOJA_INCIDENCIAS_PROVEEDOR2),
    hace7, ahora
  );
  const wbOtros = SpreadsheetApp.openById(CONFIG.SHEET_ID_OTROS_PROVEEDORES);
  const resumenOtros = {};
  PROVEEDORES_OTROS.forEach(prov => {
    const nombreHoja = CONFIG.HOJAS_OTROS_PROVEEDORES[prov];
    const hoja = wbOtros.getSheetByName(nombreHoja);
    resumenOtros[prov] = hoja
      ? _generarResumenSheet(hoja, hace7, ahora)
      : { total: 0, resueltas: 0, pendientes: 0, tiempoMedio: "N/A" };
  });
  let totalGlobal = resumenMC.total + resumenDL.total;
  let pendGlobal  = resumenMC.pendientes + resumenDL.pendientes;
  PROVEEDORES_OTROS.forEach(prov => {
    totalGlobal += resumenOtros[prov].total;
    pendGlobal  += resumenOtros[prov].pendientes;
  });
  let cuerpo =
    `📊 RESUMEN SEMANAL DE INCIDENCIAS\n` +
    `Período: ${_formatFecha(hace7)} al ${_formatFecha(ahora)}\n\n` +
    `PROVEEDOR1\n• Total: ${resumenMC.total}\n• Resueltas: ${resumenMC.resueltas}\n` +
    `• Pendientes: ${resumenMC.pendientes}\n• Tiempo medio: ${resumenMC.tiempoMedio} días\n\n` +
    `PROVEEDOR2\n• Total: ${resumenDL.total}\n• Resueltas: ${resumenDL.resueltas}\n` +
    `• Pendientes: ${resumenDL.pendientes}\n• Tiempo medio: ${resumenDL.tiempoMedio} días\n\n`;
  PROVEEDORES_OTROS.forEach(prov => {
    const r = resumenOtros[prov];
    cuerpo +=
      `${NOMBRES_PROVEEDORES[prov] || prov}\n• Total: ${r.total}\n• Resueltas: ${r.resueltas}\n` +
      `• Pendientes: ${r.pendientes}\n• Tiempo medio: ${r.tiempoMedio} días\n\n`;
  });
  cuerpo += `GLOBAL\n• Total semana: ${totalGlobal}\n• Pendientes: ${pendGlobal}`;
  GmailApp.sendEmail(
    CONFIG.CORREO_ENTRADA,
    `📊 Resumen Semanal Incidencias — ${_formatFecha(ahora)}`,
    cuerpo,
    { cc: CONFIG.ALERTA.cc, name: "Sistema Automatización Salud Digital" }
  );
  Logger.log("[RESUMEN SEMANAL] Enviado.");
}
function _generarResumenSheet(hoja, desde, hasta) {
  const datos    = hoja.getDataRange().getValues();
  if (datos.length < 2) return { total: 0, resueltas: 0, pendientes: 0, tiempoMedio: "N/A" };
  const cabecera = datos[0];
  const idxFechaRec = cabecera.indexOf("FECHA DE RECEPCION");
  const idxFechaRes = cabecera.indexOf("FECHA RESOLUCION");
  if (idxFechaRec < 0) return { total: 0, resueltas: 0, pendientes: 0, tiempoMedio: "N/A" };
  let total = 0, resueltas = 0, pendientes = 0, sumaDias = 0;
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const fechaRec = fila[idxFechaRec];
    if (!fechaRec || !(fechaRec instanceof Date)) continue;
    if (fechaRec < desde || fechaRec > hasta) continue;
    total++;
    const fechaRes = idxFechaRes >= 0 ? fila[idxFechaRes] : null;
    if (fechaRes && fechaRes instanceof Date) {
      resueltas++;
      sumaDias += Math.floor((fechaRes - fechaRec) / (1000 * 60 * 60 * 24));
    } else {
      pendientes++;
    }
  }
  return {
    total, resueltas, pendientes,
    tiempoMedio: resueltas > 0 ? (sumaDias / resueltas).toFixed(1) : "N/A"
  };
}
// ============================================================
//  R3 — DETECCIÓN DE AUTO-RESPUESTAS
// ============================================================
//
// Heurísticas combinadas:
//   • Asunto que empieza con patrones típicos de OOO/autoreply.
//   • Cabeceras MIME del raw: Auto-Submitted, X-Autoreply, etc.
//
function _esAutoRespuesta(mensaje) {
  try {
    const asunto = (mensaje.getSubject() || "").toLowerCase();
    const patronesAsunto = [
      "out of office",
      "out-of-office",
      "automatic reply",
      "auto-reply",
      "autoreply",
      "auto reply",
      "fuera de la oficina",
      "respuesta automática",
      "respuesta automatica",
      "ausente",
      "no estoy disponible",
      "vacaciones",
      "absent",
    ];
    if (patronesAsunto.some(p => asunto.includes(p))) return true;
    // Inspección de cabeceras del raw (puede no estar siempre disponible)
    const raw = mensaje.getRawContent() || "";
    const rawLow = raw.toLowerCase();
    const cabecerasIndicadoras = [
      "auto-submitted: auto-replied",
      "auto-submitted: auto-generated",
      "x-autoreply:",
      "x-autorespond:",
      "x-auto-response-suppress:",
      "precedence: auto_reply",
      "precedence: bulk",
    ];
    if (cabecerasIndicadoras.some(c => rawLow.indexOf(c) >= 0)) return true;
  } catch (e) {
    Logger.log(`[WARN] _esAutoRespuesta: ${e.message}`);
  }
  return false;
}
// ============================================================
//  D6 — LOG PERSISTENTE EN PESTAÑA "Log"
// ============================================================
function _logSheet(nivel, contexto, mensaje) {
  Logger.log(`[${nivel}] ${contexto}: ${mensaje}`);
  try {
    const wb = SpreadsheetApp.openById(CONFIG.SHEET_ID_OTROS_PROVEEDORES);
    let hoja = wb.getSheetByName(CONFIG.HOJA_LOG);
    if (!hoja) {
      hoja = wb.insertSheet(CONFIG.HOJA_LOG);
      hoja.appendRow(["Fecha", "Nivel", "Contexto", "Mensaje"]);
      hoja.setFrozenRows(1);
      hoja.getRange(1, 1, 1, 4).setFontWeight("bold");
    }
    hoja.appendRow([new Date(), nivel, contexto, mensaje]);
  } catch (e) {
    Logger.log(`[META] No se pudo escribir en Log: ${e.message}`);
  }
}
// ============================================================
//  UTILIDADES
// ============================================================
function obtenerOCrearLabel(nombreLabel) {
  let label = GmailApp.getUserLabelByName(nombreLabel);
  if (!label) {
    label = GmailApp.createLabel(nombreLabel);
    Logger.log(`[INFO] Etiqueta "${nombreLabel}" creada.`);
  }
  return label;
}
function _formatFecha(fecha) {
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
}
function extraerEmail(remitente) {
  if (!remitente) return "";
  const match = remitente.match(/<([^>]+)>/);
  return (match ? match[1] : remitente).trim().toLowerCase();
}
// ============================================================
//  NORMALIZACIÓN ROBUSTA PARA DETECTAR FRASES
// ============================================================
//
//  Convierte el texto a minúsculas, quita acentos, sustituye
//  cualquier carácter no alfanumérico (espacios, NBSP, puntuación,
//  saltos de línea, espacios tipográficos, etc.) por un único
//  espacio y trimea. Así "Càso  cerrado." y "caso cerrado"
//  acaban siendo "caso cerrado".
//
function _normalizarFrase(texto) {
  if (!texto) return "";
  return texto
    .toString()
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[̀-ͯ]/g, "")     // quita marcas de acento
    .replace(/[^a-z0-9]+/g, " ")         // colapsa cualquier no-alfanumérico
    .trim();
}
function _contieneFraseNormalizada(texto, frase) {
  const normTexto = _normalizarFrase(texto);
  const normFrase = _normalizarFrase(frase);
  if (!normTexto || !normFrase) return false;
  return normTexto.indexOf(normFrase) >= 0;
}
function textoAHtml(texto) {
  return (texto || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>");
}
// ============================================================
//  CONSTRUCTOR DE CUERPO HTML CON HISTORIAL DEL HILO
// ============================================================
//
//  Devuelve un htmlBody con:
//    1. Nuestro mensaje nuevo (cuerpoTextoNuevo, texto plano → HTML).
//    2. Separador "Historial del hilo".
//    3. Todos los mensajes anteriores del hilo en orden cronológico
//       inverso (estilo Gmail), cada uno con su cabecera De/Para/CC/
//       Fecha/Asunto y el cuerpo en texto plano.
//
//  De esta forma, cuando reenviamos al proveedor, al solicitante o al
//  área médica, todos pueden ver el contexto completo.
//
function _construirHtmlConHilo(hilo, cuerpoTextoNuevo) {
  let html = textoAHtml(cuerpoTextoNuevo);
  html +=
    `<br><br>` +
    `<hr style="border:none; border-top:1px solid #ccc; margin:16px 0;">` +
    `<div style="color:#555; font-size:0.95em; margin-bottom:8px;">` +
    `<b>--- Historial del hilo ---</b></div>`;
  let mensajes = [];
  try {
    mensajes = hilo ? hilo.getMessages() : [];
  } catch (e) {
    Logger.log(`[WARN] _construirHtmlConHilo: ${e.message}`);
  }
  for (let i = mensajes.length - 1; i >= 0; i--) {
    const m = mensajes[i];
    let from = "—", to = "—", cc = "", asunto = "", cuerpoPlano = "", fechaStr = "";
    try {
      from   = m.getFrom() || "—";
      to     = m.getTo()   || "—";
      cc     = m.getCc()   || "";
      asunto = m.getSubject() || "";
      // Cortamos solo el texto NUEVO escrito por la persona, sin la cita >>> ni
      // las atribuciones tipo "El X escribió:" que Gmail mete automáticamente.
      cuerpoPlano = _extraerCuerpoNuevo(m.getPlainBody() || "");
      const fecha = m.getDate();
      fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(),
                                       "EEE, dd MMM yyyy 'a las' HH:mm");
    } catch (e) {
      Logger.log(`[WARN] _construirHtmlConHilo msg ${i}: ${e.message}`);
    }
    html +=
      `<div style="border-left:3px solid #ccc; margin:12px 0; padding:8px 12px; ` +
      `background:#f9f9f9; color:#333;">` +
      `<div style="font-size:0.85em; color:#555; margin-bottom:8px;">` +
      `<b>De:</b> ${textoAHtml(from)}<br>` +
      `<b>Para:</b> ${textoAHtml(to)}<br>` +
      (cc ? `<b>CC:</b> ${textoAHtml(cc)}<br>` : "") +
      `<b>Fecha:</b> ${textoAHtml(fechaStr)}<br>` +
      `<b>Asunto:</b> ${textoAHtml(asunto)}` +
      `</div>` +
      `<div style="white-space:pre-wrap; font-family:Arial,Helvetica,sans-serif;">` +
      `${textoAHtml(cuerpoPlano)}</div>` +
      `</div>`;
  }
  return html;
}
// ── Devuelve SOLO el texto nuevo escrito por la persona en ese mensaje,
//    descartando la cita automática del mensaje anterior (>, "El X escribió:",
//    separadores, etc.). Es lo que evita que el historial salga con todo
//    prefijado por ">".
function _extraerCuerpoNuevo(plainBody) {
  if (!plainBody) return "";
  const lineas = plainBody.split(/\r?\n/);
  const out = [];
  for (let i = 0; i < lineas.length; i++) {
    const linea = lineas[i];
    const trim = linea.trim();
    // 1) Línea citada con >
    if (trim.indexOf(">") === 0) break;
    // 2) Separadores horizontales típicos de cita (------, ______, ======)
    if (/^[-_=*]{5,}$/.test(trim)) break;
    // 3) Cabeceras de cita ("De: Alguien <correo@x>", "From: ...")
    if (/^(de|from)\s*:\s*.+<[^>]+@/i.test(trim)) break;
    // 4) Atribución estilo Gmail: "El …, X escribió:" / "On …, X wrote:"
    //    El "escribió:" / "wrote:" puede caer hasta 3 líneas después.
    if (/^(el|on)\b/i.test(trim)) {
      const ventana = [trim, lineas[i + 1] || "", lineas[i + 2] || "", lineas[i + 3] || ""].join(" ");
      if (/(escribi[oó]|wrote)\s*:?/i.test(ventana) && /\d/.test(ventana)) {
        break;
      }
    }
    out.push(linea);
  }
  // Trimea líneas en blanco al principio y al final
  return out.join("\n").replace(/^\s+|\s+$/g, "");
}
// ============================================================
//  HELPER OPCIONAL — CREAR CABECERAS EN LAS PESTAÑAS
// ============================================================
function crearCabecerasOtrosProveedores() {
  const CABECERAS = [
    "ESTADO",
    "ORIGEN",
    "FECHA CORREO",
    "ASUNTO CORREO",
    "SR",
    "DNI",
    "TIPO INCIDENCIA (TÉCNICA/ASISTENCIAL)",
    "OBSERVACIONES",
    "FECHA DE RECEPCION",
    "FECHA RESOLUCION",
    "Responsable proveedor",
    "Responsable interno",
    "Responsable área médica",
    "Especialidad",
    "Solución",
    "FECHA CONTESTACIÓN",
    "SUBCATEGORIA",
  ];
  const wb = SpreadsheetApp.openById(CONFIG.SHEET_ID_OTROS_PROVEEDORES);
  // Pestañas de proveedores OTROS + pestaña TECNICAS (mismo formato 16 cols)
  const nombresPestanas = PROVEEDORES_OTROS.map(p => CONFIG.HOJAS_OTROS_PROVEEDORES[p]);
  nombresPestanas.push(CONFIG.HOJA_TECNICAS);
  nombresPestanas.forEach(nombre => {
    let hoja = wb.getSheetByName(nombre);
    if (!hoja) {
      hoja = wb.insertSheet(nombre);
      Logger.log(`[INFO] Pestaña "${nombre}" creada.`);
    }
    hoja.getRange(1, 1, 1, CABECERAS.length)
        .setValues([CABECERAS])
        .setFontWeight("bold");
    hoja.setFrozenRows(1);
  });
  // Crear también la pestaña Log si no existe
  let hojaLog = wb.getSheetByName(CONFIG.HOJA_LOG);
  if (!hojaLog) {
    hojaLog = wb.insertSheet(CONFIG.HOJA_LOG);
    hojaLog.appendRow(["Fecha", "Nivel", "Contexto", "Mensaje"]);
    hojaLog.setFrozenRows(1);
    hojaLog.getRange(1, 1, 1, 4).setFontWeight("bold");
    Logger.log(`[INFO] Pestaña "${CONFIG.HOJA_LOG}" creada.`);
  }
  Logger.log("[OK] Cabeceras inicializadas en SHEET_ID_OTROS_PROVEEDORES.");
}
