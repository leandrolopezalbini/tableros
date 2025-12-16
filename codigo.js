// Codigo.gs
// ================== CONSTANTES GLOBALES ==================

// Nombres de Hojas
const HOJA = {
  TABLEROS: "Tableros",
  HISTORIAL: "Historial",
  MAPS: "Maps",
  PARADAS_SEGURAS: "Paradas_Seguras",
};

// Estados y Valores permitidos
const ESTADO = {
  PEDIDO: 'PEDIDO',
  INSTALADO: 'INSTALADO',
  CONECTADA: 'CONECTADA',
  PENDIENTE: 'PENDIENTE',
  ACCESO: 'ACCESO',
  CREACION: 'Creación',
};

// Nombres de Campos (Headers) en HOJA.TABLEROS
const CAMPOS = {
  NOMBRE: 'Nombre',
  ESTADO: 'Estado',
  ENERGIA: 'Energia',
  CONECTIVIDAD: 'Conectividad',
  PROVEEDOR: 'Proveedor', // Usado para conectividad/Switch
  TIPO: 'Tipo', // Usado para conectividad/Switch
  DIRECCION: 'Direccion',
  OBSERVACIONES: 'Observaciones',
  MARCADO: 'Marcado',
  // Campos de MAPS
  LATITUD: 'Latitud',
  LONGITUD: 'Longitud',
  SWITCH: 'Switch',
  FECHA_SWITCH: 'Fecha Switch',
};

// Prefijo para la generación de IDs
const PREFIJO_ID = "NSDS-";

// Permisos: editar según tu organización
const PERMISSIONS = {
  PEDIDO: ['mauri@mm.com', 'lean@mm.com', 'valen@mm.com'],
  INSTALACION: ['marce@mm.com', 'valen@mm.com'],
  ENERGIA: ['valentin2000cm@gmail.com', 'german@mm.com'],
  CONECTIVIDAD: ['valentin2000cm@gmail.com', 'marce@mm.com']
};
const FULL_ACCESS = ['leandrolopezalbini@gmail.com', 'valentin2000cm@gmail.com', 'mauricioteodorotabarez@gmail.com'];


// ================== MENU / DOGET ==================

/** Crea el menú personalizado al abrir la hoja. */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Tableros')
    .addItem('Panel Tableros', 'abrirPanelTablerosModal')
    .addToUi();
}

/** Mantiene esta función para cargar la WebApp dentro del Sheet */
function abrirPanelTablerosModal() {
  try {
    registrarAccesoUsuario("Panel Principal Tableros (SPA)");
    const html = HtmlService.createHtmlOutputFromFile("PanelTableros")
      .setWidth(1000)
      .setHeight(750);
    SpreadsheetApp.getUi().showModalDialog(html, "Gestión de Tableros");
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    const msg = `Error al inicializar la aplicación. Detalle: ${e.message}`;
    Logger.log("Error CRÍTICO en abrirPanelTablerosModal: " + e);
    ui.alert('Error de Inicialización', msg, ui.ButtonSet.OK);
  }
}

/** Mantiene doGet para acceder a la WebApp por URL (Despliegue) */
function doGet() {
  return HtmlService.createTemplateFromFile('PanelTableros')
    .evaluate()
    .setTitle("Panel de Gestión de Tableros")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Retorna la URL base de la WebApp. */
function getWebAppBaseUrl() {
  return ScriptApp.getWebAppUrl();
}


// ================== UTILIDADES HOJAS Y DATOS ==================

/** Obtiene una hoja por su nombre, lanza error si no existe. */
function obtenerHoja(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) throw new Error(`La hoja "${nombreHoja}" no se encontró.`);
  return hoja;
}

/** Obtiene la primera fila de la hoja como array de encabezados. */
function getHeaders(hoja) {
  const lastCol = hoja.getLastColumn();
  const lastRow = hoja.getLastRow();
  if (lastRow === 0 || lastCol === 0) return [];
  const vals = hoja.getRange(1, 1, 1, lastCol).getValues();
  return (vals && vals[0]) ? vals[0].map(h => (h || '').toString().trim()) : [];
}

/** Normaliza el texto del encabezado para búsqueda. */
function normalizeHeader(h) {
  return h
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

/** * Busca el índice de columna (0-based) para el nombre dado o alias. */
function getHeaderIndex(headers, name) {
  const targets = Array.isArray(name) ? name.map(normalizeHeader) : [normalizeHeader(name)];
  for (let i = 0; i < headers.length; i++) {
    if (targets.indexOf(normalizeHeader(headers[i])) !== -1) {
      return i;
    }
  }
  // No lanzar error para que _getDatosTableros pueda mapear todos los campos
  // throw new Error("No se encontró el encabezado requerido: " + name);
  return -1; // Retorna -1 si no lo encuentra, en lugar de lanzar error
}

/** * Asegura que los encabezados mínimos existan en la hoja HOJA.TABLEROS. */
function ensureTablerosHeaders() {
  const hoja = obtenerHoja(HOJA.TABLEROS);
  const headers = getHeaders(hoja);

  const required = Object.values(CAMPOS).filter(c =>
    c !== CAMPOS.LATITUD && c !== CAMPOS.LONGITUD
  );

  if (!headers || headers.length === 0 || headers.join('').trim() === '') {
    hoja.clear();
    hoja.appendRow(required);
    return;
  }

  const missing = required.filter(h => headers.indexOf(h) === -1);

  if (missing.length > 0) {
    let lastCol = hoja.getLastColumn();
    missing.forEach(m => {
      hoja.insertColumnAfter(lastCol);
      lastCol = hoja.getLastColumn();
      hoja.getRange(1, lastCol).setValue(m);
    });
  }
}

/** * Busca la fila index (1-based) de la primera fila que tenga colValue en columna headerName. */
function findRowByColValue(hoja, headerName, colValue) {
  const headers = getHeaders(hoja);
  let idx;
  try {
    idx = getHeaderIndex(headers, headerName);
  } catch (e) {
    Logger.log(`Error al buscar encabezado ${headerName}: ${e.message}`);
    return -1;
  }

  if (idx === -1) return -1; // Maneja el caso de que getHeaderIndex retorne -1

  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return -1;

  const vals = hoja.getRange(2, idx + 1, lastRow - 1, 1).getValues();
  const normalizedValue = colValue.toString().trim();

  for (let i = 0; i < vals.length; i++) {
    if ((vals[i][0] || '').toString().trim() === normalizedValue) {
      return i + 2; // Retorna índice de fila (1-based)
    }
  }
  return -1;
}

// ================== OPTIMIZACIÓN DE DATOS ==================

/**
 * Función central para cargar los datos de HOJA.TABLEROS y mapear los índices.
 * OPTIMIZACIÓN: Minimiza las llamadas a la API (solo una llamada a getValues()).
 * @returns {{headers: string[], data: Array<Array<any>>, indices: Object}}
 */
function _getDatosTableros() {
  ensureTablerosHeaders(); // Asegura la estructura
  const hoja = obtenerHoja(HOJA.TABLEROS);
  const vals = hoja.getDataRange().getValues();

  if (!vals || vals.length < 2) {
    return { headers: [], data: [], indices: {} };
  }

  const headers = vals[0];
  const data = vals.slice(1); // Datos sin encabezados

  // Mapeo de índices para fácil acceso (se hace una sola vez)
  const indices = {};

  // Si getHeaderIndex devuelve -1, el índice es ignorado por la función que lo usa.
  indices[CAMPOS.NOMBRE] = getHeaderIndex(headers, CAMPOS.NOMBRE);
  indices[CAMPOS.ESTADO] = getHeaderIndex(headers, CAMPOS.ESTADO);
  indices[CAMPOS.ENERGIA] = getHeaderIndex(headers, [CAMPOS.ENERGIA, 'Energía']);
  indices[CAMPOS.CONECTIVIDAD] = getHeaderIndex(headers, CAMPOS.CONECTIVIDAD);
  indices[CAMPOS.DIRECCION] = getHeaderIndex(headers, CAMPOS.DIRECCION);
  indices[CAMPOS.MARCADO] = getHeaderIndex(headers, CAMPOS.MARCADO);
  indices[CAMPOS.PROVEEDOR] = getHeaderIndex(headers, CAMPOS.PROVEEDOR);
  indices[CAMPOS.TIPO] = getHeaderIndex(headers, CAMPOS.TIPO);
  indices[CAMPOS.SWITCH] = getHeaderIndex(headers, CAMPOS.SWITCH);
  indices[CAMPOS.FECHA_SWITCH] = getHeaderIndex(headers, CAMPOS.FECHA_SWITCH);

  // Verificación de índices críticos:
  if (indices[CAMPOS.NOMBRE] === -1 || indices[CAMPOS.DIRECCION] === -1 || indices[CAMPOS.ESTADO] === -1) {
    throw new Error("Faltan columnas esenciales (Nombre, Direccion o Estado) en la hoja 'Tableros'.");
  }

  return { headers, data, indices };
}

// ================== AUTH / PERMISOS ==================

/** Obtiene el email del usuario activo. */
function usuarioActivo() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  } catch (e) {
    return 'Desconocido';
  }
}

/** Verifica si el usuario tiene permiso para una acción. */
function usuarioTienePermiso(accion, email) {
  if (!email) email = usuarioActivo();
  if (!email) return false;
  if (FULL_ACCESS.indexOf(email) !== -1) return true;
  const lista = PERMISSIONS[accion] || [];
  return lista.indexOf(email) !== -1;
}

// ================== BITÁCORA / HISTORIAL / onEdit ==================

/** Registra accesos generales al panel. */
function registrarAccesoUsuario(seccion) {
  try {
    const hoja = obtenerHoja(HOJA.HISTORIAL);
    const usuario = Session.getActiveUser().getEmail() || "Desconocido";
    const detalle = `Ingreso a sección: ${seccion}`;
    hoja.appendRow(["-", ESTADO.ACCESO, "-", new Date(), detalle, usuario]);
  } catch (err) {
    Logger.log("Error al registrar acceso: " + err);
  }
}

/** * Registra un evento en la hoja de historial. */
function registrarHistorial(tipo, idTablero, nuevoValor, detalle, usuarioEmail, campo) {
  try {
    const sheet = obtenerHoja(HOJA.HISTORIAL);
    const email = usuarioEmail || usuarioActivo() || 'Desconocido';
    const fecha = new Date();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([CAMPOS.NOMBRE, 'Tipo', CAMPOS.ESTADO, 'FechaHora', 'Detalle', 'Email']);
    }
    sheet.appendRow([idTablero, tipo, nuevoValor, fecha, detalle, email]);

  } catch (e) {
    Logger.log("Error registrarHistorial: " + e);
  }
}

/** Registra modificaciones manuales en la hoja HOJA.TABLEROS. */
function onEdit(e) {
  // 1. Validaciones iniciales
  if (!e.range || e.range.getRow() === 1 || !e.user || !e.value) return;
  const hojaEditada = e.range.getSheet();
  if (hojaEditada.getName() !== HOJA.TABLEROS) return;

  const fila = e.range.getRow();
  const columna = e.range.getColumn();
  const valorAnterior = e.oldValue;
  const nuevoValor = e.value;

  if (valorAnterior === nuevoValor) return;

  const encabezados = getHeaders(hojaEditada);
  if (encabezados.length === 0) return;

  const nombreColumna = encabezados[columna - 1] || `Columna ${columna}`;

  // 2. Identificar el Tablero (por columna 'Nombre')
  let idxNombre;
  try {
    idxNombre = getHeaderIndex(encabezados, CAMPOS.NOMBRE);
  } catch (err) {
    Logger.log("Advertencia: Columna 'Nombre' no encontrada para onEdit.");
    return;
  }
  const nombreTablero = hojaEditada.getRange(fila, idxNombre + 1).getValue();

  // 3. Obtener el estado actual
  let estadoActual = 'N/A';
  try {
    const idxEstado = getHeaderIndex(encabezados, CAMPOS.ESTADO);
    estadoActual = hojaEditada.getRange(fila, idxEstado + 1).getValue() || 'N/A';
  } catch (err) {
    Logger.log("Advertencia: Columna 'Estado' no encontrada para onEdit.");
  }

  if (nombreColumna.toUpperCase() === CAMPOS.ESTADO.toUpperCase()) {
    estadoActual = nuevoValor;
  }

  // 4. Detalle del Registro
  const detalle = `Modificación manual: ${nombreColumna}. Ant: "${valorAnterior || 'Vacío'}". Nuevo: "${nuevoValor}".`;

  // 5. Llamada a registrarHistorial
  registrarHistorial('Edición Manual', nombreTablero, estadoActual, detalle, e.user.getEmail(), nombreColumna);
}

// ================== PROCESOS: NUEVO TABLERO ==================

/** Crea una nueva fila en Tableros y Maps, y registra en Historial. */
function crearNuevoTablero(data) {
  try {
    const { headers, data: allData, indices } = _getDatosTableros(); // Optimizado
    const hoja = obtenerHoja(HOJA.TABLEROS);
    const hojaMaps = obtenerHoja(HOJA.MAPS);

    // Generar ID único
    let max = 0;
    allData.forEach(row => {
      const name = (row[indices[CAMPOS.NOMBRE]] || '').toString().trim();
      if (name.startsWith(PREFIJO_ID) && /^NSDS-\d+$/.test(name)) {
        const n = parseInt(name.split("-")[1], 10);
        if (n > max) max = n;
      }
    });
    const nuevoID = `${PREFIJO_ID}${String(max + 1).padStart(3, '0')}`;

    // Coordenadas iniciales PENDIENTES
    const lat = ESTADO.PENDIENTE;
    const lng = ESTADO.PENDIENTE;

    // Paso 1: Crear la nueva fila en HOJA.TABLEROS
    const nueva = new Array(headers.length).fill('');
    nueva[indices[CAMPOS.NOMBRE]] = nuevoID;
    nueva[indices[CAMPOS.DIRECCION]] = data.direccion;
    nueva[indices[CAMPOS.MARCADO]] = data.marcado;
    nueva[indices[CAMPOS.ESTADO]] = ESTADO.PEDIDO;
    nueva[indices[CAMPOS.ENERGIA]] = ESTADO.PENDIENTE;
    nueva[indices[CAMPOS.CONECTIVIDAD]] = ESTADO.PENDIENTE;
    nueva[indices[CAMPOS.OBSERVACIONES]] = data.observaciones;
    nueva[indices[CAMPOS.PROVEEDOR]] = ESTADO.PENDIENTE;
    nueva[indices[CAMPOS.TIPO]] = ESTADO.PENDIENTE;
    nueva[indices[CAMPOS.SWITCH]] = ESTADO.PENDIENTE;
    nueva[indices[CAMPOS.FECHA_SWITCH]] = ESTADO.PENDIENTE;

    hoja.appendRow(nueva);

    // Paso 2: Registrar en HOJA.MAPS 
    const requiredMapsHeaders = [CAMPOS.NOMBRE, CAMPOS.DIRECCION, CAMPOS.LATITUD, CAMPOS.LONGITUD];
    let headersMaps = getHeaders(hojaMaps);
    let colIndexesMaps = {};

    if (headersMaps.length === 0 || headersMaps.join('').trim() === '') {
      hojaMaps.clear();
      hojaMaps.appendRow(requiredMapsHeaders);
      headersMaps = requiredMapsHeaders;
    }

    try {
      colIndexesMaps[CAMPOS.NOMBRE] = getHeaderIndex(headersMaps, CAMPOS.NOMBRE);
      colIndexesMaps[CAMPOS.DIRECCION] = getHeaderIndex(headersMaps, CAMPOS.DIRECCION);
      colIndexesMaps[CAMPOS.LATITUD] = getHeaderIndex(headersMaps, [CAMPOS.LATITUD, 'Lat']);
      colIndexesMaps[CAMPOS.LONGITUD] = getHeaderIndex(headersMaps, [CAMPOS.LONGITUD, 'Lng']);
    } catch (e) {
      Logger.log("Advertencia: Faltan encabezados estándar en Maps. Usando orden por defecto.");
      colIndexesMaps = {
        [CAMPOS.NOMBRE]: 0, [CAMPOS.DIRECCION]: 1,
        [CAMPOS.LATITUD]: 2, [CAMPOS.LONGITUD]: 3
      };
    }

    const mapRowLength = hojaMaps.getLastRow() > 0 ? hojaMaps.getLastColumn() : requiredMapsHeaders.length;
    const nuevaFilaMaps = new Array(mapRowLength).fill('');

    nuevaFilaMaps[colIndexesMaps[CAMPOS.NOMBRE]] = nuevoID;
    nuevaFilaMaps[colIndexesMaps[CAMPOS.DIRECCION]] = data.direccion;
    nuevaFilaMaps[colIndexesMaps[CAMPOS.LATITUD]] = lat;
    nuevaFilaMaps[colIndexesMaps[CAMPOS.LONGITUD]] = lng;

    hojaMaps.appendRow(nuevaFilaMaps);

    // Paso 3: Registrar en Historial
    registrarHistorial(
      ESTADO.CREACION,
      nuevoID,
      ESTADO.PEDIDO,
      `Dirección: ${data.direccion} | Coordenadas: ${lat},${lng} | Observaciones: ${data.observaciones || ''}`,
      usuarioActivo(),
      CAMPOS.ESTADO
    );

    return { success: true, id: nuevoID };

  } catch (e) {
    Logger.log("Error al crear nuevo tablero: " + e);
    return { success: false, message: e.message };
  }
}


// ================== PROCESOS: INSTALACIÓN ==================

/** Registra la instalación física de un tablero, actualizando coordenadas. */
function registrarInstalacionTablero(datos) {
  try {
    const usuario = usuarioActivo();
    if (!usuarioTienePermiso('INSTALACION', usuario)) {
      return { success: false, message: `Usuario ${usuario} sin permiso de Instalación` };
    }

    const hoja = obtenerHoja(HOJA.TABLEROS);
    const headers = getHeaders(hoja);
    const idxEstado = getHeaderIndex(headers, CAMPOS.ESTADO);
    const idxObs = getHeaderIndex(headers, [CAMPOS.OBSERVACIONES, 'Observación']);

    // 1. VALIDACIÓN
    if (!datos.latitud || !datos.longitud) {
      return { success: false, message: 'Coordenadas (latitud y longitud) son requeridas para la instalación.' };
    }
    const { latitud, longitud } = datos;

    const filaIndex = findRowByColValue(hoja, CAMPOS.NOMBRE, datos.id);
    if (filaIndex === -1) return { success: false, message: 'Tablero no encontrado' };

    const estadoActual = (hoja.getRange(filaIndex, idxEstado + 1).getValue() || '').toString().trim().toUpperCase();
    if (estadoActual !== ESTADO.PEDIDO) return { success: false, message: `Estado actual es ${estadoActual}, no es PEDIDO` };

    // 2. ACTUALIZAR ESTADO EN HOJA.TABLEROS
    hoja.getRange(filaIndex, idxEstado + 1).setValue(ESTADO.INSTALADO);

    // Anexar observaciones
    if (idxObs !== -1) {
      const obsExistentes = (hoja.getRange(filaIndex, idxObs + 1).getValue() || '').trim();
      hoja.getRange(filaIndex, idxObs + 1).setValue(`${obsExistentes}\n[INSTALACION] Coordenadas: ${latitud},${longitud}. ${datos.observaciones || ''}`);
    }

    // 3. ACTUALIZAR COORDENADAS EN HOJA.MAPS
    const hojaMaps = obtenerHoja(HOJA.MAPS);
    const headersMaps = getHeaders(hojaMaps);
    const filaMaps = findRowByColValue(hojaMaps, CAMPOS.NOMBRE, datos.id);

    if (filaMaps !== -1) {
      try {
        const idxLatMaps = getHeaderIndex(headersMaps, [CAMPOS.LATITUD, 'Lat']);
        const idxLngMaps = getHeaderIndex(headersMaps, [CAMPOS.LONGITUD, 'Lng']);
        if (idxLatMaps !== -1) hojaMaps.getRange(filaMaps, idxLatMaps + 1).setValue(latitud);
        if (idxLngMaps !== -1) hojaMaps.getRange(filaMaps, idxLngMaps + 1).setValue(longitud);
      } catch (e) {
        Logger.log("Error al actualizar coordenadas en MAPS: " + e.message);
      }
    }

    // 4. REGISTRAR EN HISTORIAL
    const detalleHistorial = `Instalación completada. Coordenadas: ${latitud},${longitud}. Obs: ${datos.observaciones || ''}`;
    registrarHistorial('Instalación', datos.id, ESTADO.INSTALADO, detalleHistorial, usuario, CAMPOS.ESTADO);

    return { success: true };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.message || e };
  }
}

// ================== PROCESOS: ENERGÍA ==================

function registrarEnergiaTablero(datos) {
  try {
    const usuario = usuarioActivo();
    if (!usuarioTienePermiso('ENERGIA', usuario)) {
      return { success: false, message: `Usuario ${usuario} sin permiso de Energía` };
    }

    const hoja = obtenerHoja(HOJA.TABLEROS);
    const headers = getHeaders(hoja);

    const idxEnergia = getHeaderIndex(headers, CAMPOS.ENERGIA);
    const idxObs = getHeaderIndex(headers, CAMPOS.OBSERVACIONES);

    const fila = findRowByColValue(hoja, CAMPOS.NOMBRE, datos.id);
    if (!fila || fila === -1) return { success: false, message: "Tablero no encontrado" };

    const energiaActual = (hoja.getRange(fila, idxEnergia + 1).getValue() || '').toString().trim().toUpperCase();
    if (energiaActual === ESTADO.CONECTADA) return { success: false, message: "La Energía ya está CONECTADA" };

    let proveedorFinal = datos.proveedor === "OTRA" ? datos.detalleProveedor : datos.proveedor;

    hoja.getRange(fila, idxEnergia + 1).setValue(ESTADO.CONECTADA);

    // Anexar observaciones
    if (datos.observaciones) {
      const obsExistentes = (hoja.getRange(fila, idxObs + 1).getValue() || '').trim();
      hoja.getRange(fila, idxObs + 1).setValue(`${obsExistentes}\n[ENERGIA] ${datos.observaciones || ''}`);
    }

    const detalleHistorial = `Proveedor:${proveedorFinal} | Obs:${datos.observaciones || ""}`;
    registrarHistorial('Energía', datos.id, ESTADO.CONECTADA, detalleHistorial, usuario, CAMPOS.ENERGIA);

    return { success: true };

  } catch (e) {
    Logger.log("Error registrarEnergiaTablero: " + e);
    return { success: false, message: e.message || e };
  }
}

// ================== PROCESOS: SWITCH (CONECTIVIDAD COMPLETA) ==================


/**
 * Vincula el nombre del Switch y actualiza los campos de conectividad.
 * Incluye la actualización del estado de CONECTIVIDAD a 'CONECTADA'.
 * @param {Object} data - Objeto con id, switch, proveedor, detalleProveedor, tipo y observaciones.
 * @returns {Object} Resultado de la operación.
 */
function vincularSwitchTablero(data) {
  try {
    const hoja = obtenerHoja(HOJA.TABLEROS);
    const headers = getHeaders(hoja);

    const idxSwitch = getHeaderIndex(headers, CAMPOS.SWITCH);
    const idxProveedor = getHeaderIndex(headers, CAMPOS.PROVEEDOR);
    const idxTipo = getHeaderIndex(headers, CAMPOS.TIPO);
    const idxFechaSwitch = getHeaderIndex(headers, CAMPOS.FECHA_SWITCH);
    const idxObs = getHeaderIndex(headers, CAMPOS.OBSERVACIONES);

    // ÍNDICE CRÍTICO: Para establecer el estado de CONECTIVIDAD
    const idxConectividad = getHeaderIndex(headers, CAMPOS.CONECTIVIDAD);

    const filaIndex = findRowByColValue(hoja, CAMPOS.NOMBRE, data.id);
    if (filaIndex === -1) throw new Error(`Tablero ID ${data.id} no encontrado.`);

    const fechaActual = new Date();
    const proveedorFinal = data.proveedor === 'OTRA' ? data.detalleProveedor : data.proveedor;

    // 1. Escribir los datos del Switch y Conectividad
    if (idxSwitch !== -1) hoja.getRange(filaIndex, idxSwitch + 1).setValue(data.switch);
    if (idxTipo !== -1) hoja.getRange(filaIndex, idxTipo + 1).setValue(data.tipo);
    if (idxProveedor !== -1) hoja.getRange(filaIndex, idxProveedor + 1).setValue(proveedorFinal);
    if (idxFechaSwitch !== -1) hoja.getRange(filaIndex, idxFechaSwitch + 1).setValue(fechaActual);

    // 2. Establecer estado de CONECTIVIDAD a CONECTADA (YA ESTÁ CORRECTO)
    if (idxConectividad !== -1) {
      hoja.getRange(filaIndex, idxConectividad + 1).setValue(ESTADO.CONECTADA);
    }

    // 3. Añadir observación
    if (idxObs !== -1) {
      let obsExistente = (hoja.getRange(filaIndex, idxObs + 1).getValue() || '').toString().trim();
      let nuevaObs = `[Switch/${data.tipo}] ${data.switch} con ${proveedorFinal}. Obs: ${data.observaciones || ''} (${fechaActual.toLocaleDateString()})`;
      let obsFinal = obsExistente + (obsExistente ? '\n' : '') + nuevaObs;
      hoja.getRange(filaIndex, idxObs + 1).setValue(obsFinal);
    }

    // 4. Registrar en Historial
    const detalleHistorial = `Switch: ${data.switch}, Tipo: ${data.tipo}, Proveedor: ${proveedorFinal}. Obs: ${data.observaciones || ''}`;
    registrarHistorial('Switch', data.id, ESTADO.CONECTADA, detalleHistorial, usuarioActivo(), CAMPOS.SWITCH);


    return { success: true, id: data.id };

  } catch (e) {
    Logger.log("Error en vincularSwitchTablero: " + e);
    return { success: false, message: e.message };
  }
}


// ================== PROCESOS: CONECTIVIDAD (Sin Switch) ==================

/** Registra la conexión de conectividad (proveedor, tipo) en un tablero. */
function registrarConectividadTablero(datos) {
  try {
    const usuario = usuarioActivo();
    if (!usuarioTienePermiso('CONECTIVIDAD', usuario)) {
      return { success: false, message: `Usuario ${usuario} sin permiso de Conectividad` };
    }

    const hoja = obtenerHoja(HOJA.TABLEROS);
    const headers = getHeaders(hoja);

    const idxConectividad = getHeaderIndex(headers, CAMPOS.CONECTIVIDAD);
    const idxProveedor = getHeaderIndex(headers, CAMPOS.PROVEEDOR);
    const idxTipo = getHeaderIndex(headers, CAMPOS.TIPO);
    const idxObs = getHeaderIndex(headers, CAMPOS.OBSERVACIONES);

    const fila = findRowByColValue(hoja, CAMPOS.NOMBRE, datos.id);
    if (!fila || fila === -1) return { success: false, message: "Tablero no encontrado" };

    const conectActual = (hoja.getRange(fila, idxConectividad + 1).getValue() || '').toString().trim().toUpperCase();
    if (conectActual === ESTADO.CONECTADA) return { success: false, message: "La Conectividad ya está CONECTADA" };

    let proveedorFinal = datos.proveedor === "OTRA" ? datos.detalleProveedor : datos.proveedor;

    // Actualizar campos
    if (idxProveedor !== -1) hoja.getRange(fila, idxProveedor + 1).setValue(proveedorFinal);
    if (idxTipo !== -1) hoja.getRange(fila, idxTipo + 1).setValue(datos.tipo);
    if (idxConectividad !== -1) hoja.getRange(fila, idxConectividad + 1).setValue(ESTADO.CONECTADA);

    // Anexar observaciones
    if (datos.observaciones && idxObs !== -1) {
      const obsExistentes = (hoja.getRange(fila, idxObs + 1).getValue() || '').trim();
      hoja.getRange(fila, idxObs + 1).setValue(`${obsExistentes}\n[CONECTIVIDAD] ${datos.observaciones || ''}`);
    }

    const detalleHistorial = `Proveedor: ${proveedorFinal} | Tipo: ${datos.tipo} | Obs: ${datos.observaciones || ""}`;
    registrarHistorial('Conectividad', datos.id, ESTADO.CONECTADA, detalleHistorial, usuario, CAMPOS.CONECTIVIDAD);

    return { success: true };

  } catch (e) {
    Logger.log("Error registrarConectividadTablero: " + e);
    return { success: false, message: e.message || e };
  }
}

// ================== FUNCIONES DE LISTADO (OPTIMIZADAS) ==================
/**
 * Función de utilidad para filtrar y mapear datos de tableros.
 * @param {Function} filtroFn - Función que recibe (row, indices) y retorna boolean.
 * @returns {Array<Object>} Lista de tableros filtrados.
 */
/**
 * RE-DEFINICIÓN DE _filtrarTableros para que sea universal
 * Agregamos los campos de conectividad para que no falten en ninguna vista.
 */
function _filtrarTableros(filtroFn) {
  try {
    const { data, indices } = _getDatosTableros();
    const out = [];

    data.forEach(row => {
      if (filtroFn(row, indices)) {
        out.push({
          id: row[indices[CAMPOS.NOMBRE]] || '',
          direccion: row[indices[CAMPOS.DIRECCION]] || '',
          estado: row[indices[CAMPOS.ESTADO]] || '',
          energia: row[indices[CAMPOS.ENERGIA]] || '',
          conectividad: row[indices[CAMPOS.CONECTIVIDAD]] || '',
          // Agregamos estos para que el autocompletado de Switch tenga la info
          switch: row[indices[CAMPOS.SWITCH]] || '',
          proveedorConectividad: row[indices[CAMPOS.PROVEEDOR]] || '',
          tipoConectividad: row[indices[CAMPOS.TIPO]] || ''
        });
      }
    });
    return out;
  } catch (e) {
    Logger.log('_filtrarTableros error: ' + e);
    return [];
  }
}

function listarTablerosPedido() {
  return _filtrarTableros((row, indices) => {
    return (row[indices[CAMPOS.ESTADO]] || '').toString().toUpperCase() === ESTADO.PEDIDO;
  });
}

function listarTablerosParaEnergia() {
  return _filtrarTableros((row, indices) => {
    const estado = (row[indices[CAMPOS.ESTADO]] || '').toString().toUpperCase();
    return estado === ESTADO.INSTALADO || estado === ESTADO.CONECTADA;
  });
}

/** Filtra: Estado = INSTALADO AND Conectividad != CONECTADA */
function listarTablerosParaConectividad() {
  return _filtrarTableros((row, indices) => {
    const estado = (row[indices[CAMPOS.ESTADO]] || '').toString().toUpperCase();
    return estado === ESTADO.INSTALADO || estado === ESTADO.CONECTADA;
  });
}

/** Retorna tableros listos para recibir un Switch.  */
/** Filtra: Estado = INSTALADO o CONECTADA para el formulario de Switch */
function listarTablerosParaSwitch() {
  return _filtrarTableros((row, idx) => {
    const estado = (row[idx[CAMPOS.ESTADO]] || '').toString().trim().toUpperCase();
    // Permitimos ambos estados para poder vincular o actualizar el switch
    return estado === ESTADO.INSTALADO || estado === ESTADO.CONECTADA;
  });
}

// ================== INFORMES Y FILTRADO ==================

/**
 * Función central para obtener, combinar y filtrar los datos de tableros.
 * @param {object} filters - Objeto con filtros (estado, energia, conectividad, direccionContains).
 * @returns {Array<object>} Lista de tableros filtrados mapeados a objetos.
 */
function obtenerTablerosFiltrados(filters) {
  // 1. OBTENER DATOS PRINCIPALES (Optimizado)
  const { headers, data: dataPrincipal, indices } = _getDatosTableros();
  let datosCombinados = dataPrincipal;

  // 2. OBTENER Y COMBINAR DATOS DE HOJA SECUNDARIA (Paradas Seguras)
  try {
    const hojaSecundaria = obtenerHoja(HOJA.PARADAS_SEGURAS);
    const valsSecundarios = hojaSecundaria.getDataRange().getValues();

    if (valsSecundarios && valsSecundarios.length > 1) {
      const headersSecundarios = valsSecundarios[0];

      for (let i = 1; i < valsSecundarios.length; i++) {
        const rowSecundaria = valsSecundarios[i];
        const rowPrincipal = new Array(headers.length).fill('');

        headersSecundarios.forEach((h, j) => {
          const targetIdx = headers.indexOf(h);
          if (targetIdx !== -1) {
            rowPrincipal[targetIdx] = rowSecundaria[j];
          }
        });
        // Evitar duplicados por nombre
        const nombreSecundario = rowSecundaria[getHeaderIndex(headersSecundarios, CAMPOS.NOMBRE)];
        const existe = datosCombinados.some(r => r[indices[CAMPOS.NOMBRE]] === nombreSecundario);
        if (!existe) {
          datosCombinados.push(rowPrincipal);
        }
      }
    }
  } catch (e) {
    Logger.log('Advertencia: No se pudo cargar la hoja ' + HOJA.PARADAS_SEGURAS + '. Error: ' + e.message);
  }

  // 3. APLICAR FILTROS
  const filasFiltradas = datosCombinados.filter(row => {
    // Filtro por Estado
    if (filters && filters.estado && filters.estado.toString().trim() !== '') {
      const estado = (row[indices[CAMPOS.ESTADO]] || '').toString().trim().toUpperCase();
      if (estado !== filters.estado.toString().trim().toUpperCase()) return false;
    }
    // Filtro por Energía
    if (filters && filters.energia && filters.energia.toString().trim() !== '') {
      const energia = (row[indices[CAMPOS.ENERGIA]] || '').toString().trim().toUpperCase();
      if (energia !== filters.energia.toString().trim().toUpperCase()) return false;
    }
    // Filtro por Conectividad
    if (filters && filters.conectividad && filters.conectividad.toString().trim() !== '') {
      const conectividad = (row[indices[CAMPOS.CONECTIVIDAD]] || '').toString().trim().toUpperCase();
      if (conectividad !== filters.conectividad.toString().trim().toUpperCase()) return false;
    }
    // Filtro de búsqueda por Dirección
    if (filters && filters.direccionContains && filters.direccionContains.toString().trim() !== '') {
      const q = filters.direccionContains.toString().trim().toUpperCase();
      const nombre = (row[indices[CAMPOS.NOMBRE]] || '').toString().toUpperCase();
      const direccion = (row[indices[CAMPOS.DIRECCION]] || '').toString().toUpperCase();

      // Busca por Nombre O por Dirección
      if (direccion.indexOf(q) === -1 && nombre.indexOf(q) === -1) return false;
    }
    return true;
  });

  // 4. MAPEAR A OBJETOS DE SALIDA
  return filasFiltradas.map(r => ({
    id: r[indices[CAMPOS.NOMBRE]] || '',
    direccion: r[indices[CAMPOS.DIRECCION]] || '',
    estado: r[indices[CAMPOS.ESTADO]] || '',
    energia: r[indices[CAMPOS.ENERGIA]] || '',
    conectividad: r[indices[CAMPOS.CONECTIVIDAD]] || '',
    proveedor: (indices[CAMPOS.PROVEEDOR] !== -1 ? r[indices[CAMPOS.PROVEEDOR]] : '') || '',
    tipo: (indices[CAMPOS.TIPO] !== -1 ? r[indices[CAMPOS.TIPO]] : '') || '',
    marcado: r[indices[CAMPOS.MARCADO]] || ''
  }));
}

/** Devuelve los tableros filtrados para la vista previa en el HTML. */
function listarTablerosFiltrados(filters) {
  try {
    const tableros = obtenerTablerosFiltrados(filters);
    return tableros;
  } catch (e) {
    Logger.log('listarTablerosFiltrados error: ' + e);
    throw e;
  }
}

/** Genera el informe PDF a partir de los datos filtrados. */
function generarInformePDF(filters) {
  try {
    const filas = obtenerTablerosFiltrados(filters);
    const usuario = usuarioActivo();
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    if (!filas || filas.length === 0) {
      const errorHtml = '<html><body><h2>Informe sin resultados</h2><p>No se encontraron tableros con los filtros aplicados.</p></body></html>';
      return Utilities.base64Encode(Utilities.newBlob(errorHtml, 'text/html', 'error.html').getBytes());
    }

    // --- Generación de HTML para PDF ---
    let html = '<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Arial,Helvetica,sans-serif;font-size:12px}h2,p{margin:6px 0}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left}th{background:#f5f5f5}</style></head><body>';

    html += `<h2>Informe de Tableros (${filas.length} encontrados)</h2>`;
    html += `<p>Generado por: <b>${usuario}</b><br>Fecha: ${fecha}</p>`;

    html += '<table><tr><th>ID</th><th>Dirección</th><th>Estado</th><th>Energía</th><th>Conectividad</th><th>Proveedor</th><th>Tipo</th><th>Marcado</th></tr>';

    filas.forEach(r => {
      html += '<tr>';
      html += `<td>${r.id}</td>`;
      html += `<td>${r.direccion}</td>`;
      html += `<td>${r.estado}</td>`;
      html += `<td>${r.energia}</td>`;
      html += `<td>${r.conectividad}</td>`;
      html += `<td>${r.proveedor}</td>`;
      html += `<td>${r.tipo}</td>`;
      html += `<td>${r.marcado}</td>`;
      html += '</tr>';
    });

    html += `</table></body></html>`;

    // --- Crear PDF ---
    const blob = Utilities.newBlob(html, 'text/html', 'informe.html');
    const pdf = blob.getAs('application/pdf').setName(`Informe_Tablero_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.pdf`);

    registrarHistorial('Descarga Informe', 'TODAS', 'INFORME', `Informe generado`, usuario, 'Informe');

    // Devuelve el PDF codificado en Base64 para que el cliente (HTML) lo descargue
    return Utilities.base64Encode(pdf.getBytes());
  } catch (e) {
    Logger.log('generarInformePDF error: ' + e);
    throw e;
  }
}