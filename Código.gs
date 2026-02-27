// ==========================================
// CONFIGURACIÓN GLOBAL
// ==========================================

const CONFIG = {
  SPREADSHEET_IDS: {
    USUARIOS: '1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM',
    ASISTENCIA: '1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk'
  },

  HOJAS: {
    USUARIOS: 'BD_SLIMAPP'
  },

  COLUMNAS: {
    USUARIOS: {
      RUT: 0,                           // A
      RUT_VALIDADO: 1,                  // B
      FECHA_INGRESO: 2,                 // C
      NOMBRE: 3,                        // D
      CARGO: 4,                         // E
      CORREO: 5,                        // F
      SITE: 6,                          // G
      REGION: 7,                        // H
      SEXO: 8,                          // I
      ESTADO: 9,                        // J
      DETALLE_DESVINCULACION: 10,       // K
      ID_CREDENCIAL: 11,                // L
      CORREO_REGISTRADO: 12,            // M
      CONTACTO: 13,                     // N
      ROL: 14,                          // O
      LINK_REGISTRO: 15,                // P
      QR_REGISTRO: 16,                  // Q
      BANCO: 17,                        // R
      TIPO_CUENTA: 18,                  // S
      NUMERO_CUENTA: 19,                // T
      ESTADO_NEG_COLECT_2026: 20        // U
    },
    ASISTENCIA: {
      FECHA_HORA: 0,                    // A
      RUT: 1,                           // B
      NOMBRE: 2,                        // C
      ASAMBLEA: 3,                      // D
      TIPO_ASISTENCIA: 4,               // E
      GESTION: 5                        // F
    }
  },

  WEB_APP: {
    URL: 'https://script.google.com/a/~/macros/s/AKfycbzrmy_GgdzMpOLfycvxxUPHU6iyuL9Jv6As_4kxG7mG8oQ4RbV-ALUZw0oeSJnqbvvc/exec'
  }
};

// ==========================================
// FUNCIONES AUXILIARES
// ==========================================

/**
 * Obtiene una hoja específica de un spreadsheet usando la clave de CONFIG.
 */
function getSheet(spreadsheetKey, sheetName) {
  const spreadsheetId = CONFIG.SPREADSHEET_IDS[spreadsheetKey];
  if (!spreadsheetId) throw new Error(`Spreadsheet key "${spreadsheetKey}" no encontrado en CONFIG`);

  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sheet) throw new Error(`Hoja "${sheetName}" no encontrada en el spreadsheet`);

  return sheet;
}

/**
 * Limpia un RUT chileno (elimina puntos, guiones y espacios).
 */
function cleanRut(rut) {
  if (!rut) return '';
  return rut.toString().replace(/[.\-\s]/g, '').toUpperCase();
}

/**
 * Genera el link de registro y la fórmula QR para un RUT dado.
 * @returns {{ link: string, formulaQR: string }}
 */
function buildRegistroData(rutLimpio) {
  const link = `${CONFIG.WEB_APP.URL}?action=register&rut=${rutLimpio}`;
  const formulaQR = `=IMAGE("https://quickchart.io/qr?size=300&text=${encodeURIComponent(link)}")`;
  return { link, formulaQR };
}

// ==========================================
// MENÚ PERSONALIZADO
// ==========================================

/**
 * Crea el menú personalizado al abrir el archivo.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 SLIM - Herramientas')
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('📱 QR Asistencia')
        .addItem('🔄 Generar QR para TODOS los usuarios', 'ejecutarGenerarQRTodos')
        .addSeparator()
        .addItem('📋 Instrucciones de configuración', 'mostrarInstruccionesQR')
    )
    .addToUi();
}

/**
 * Solicita confirmación y ejecuta la generación masiva de QR.
 */
function ejecutarGenerarQRTodos() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Generar QR de Registro',
    '¿Estás seguro de que deseas generar los códigos QR para TODOS los usuarios?\n\n' +
    'Esta acción sobrescribirá los datos actuales en las columnas P (Link Registro) y Q (QR Registro).',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Generando QR... Por favor espera.', '⏳ Procesando', -1);

  const resultado = generarQRRegistroUsuarios();

  ss.toast('Proceso finalizado', '✅ Listo', 3);
  ui.alert(resultado.success ? '✅ Completado' : '❌ Error', resultado.message, ui.ButtonSet.OK);
}

/**
 * Muestra las instrucciones para configurar el sistema QR.
 */
function mostrarInstruccionesQR() {
  SpreadsheetApp.getUi().alert(
    '📚 Instrucciones',
    '📋 INSTRUCCIONES PARA CONFIGURAR QR DE ASISTENCIA\n\n' +
    '1️⃣ Despliega la Web App:\n' +
    '   • En Apps Script: Implementar → Nueva implementación\n' +
    '   • Tipo: Aplicación web\n' +
    '   • Ejecutar como: Yo\n' +
    '   • Acceso: Cualquier persona\n\n' +
    '2️⃣ Copia la URL generada (ejemplo: https://script.google.com/...)\n\n' +
    '3️⃣ Pégala en el código:\n' +
    '   • Busca: CONFIG.WEB_APP.URL\n' +
    '   • Reemplaza con tu URL real\n\n' +
    '4️⃣ Guarda el proyecto (Ctrl+S)\n\n' +
    '5️⃣ Vuelve a este menú y ejecuta "Generar QR para TODOS"\n\n' +
    '✅ Los códigos QR aparecerán en las columnas P y Q',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==========================================
// GENERADOR DE QR DE REGISTRO
// ==========================================

/**
 * Genera links y códigos QR de registro para todos los usuarios de la hoja.
 * Escribe los resultados en las columnas P (Link Registro) y Q (QR Registro).
 */
function generarQRRegistroUsuarios() {
  try {
    const sheet = getSheet('USUARIOS', CONFIG.HOJAS.USUARIOS);
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.USUARIOS;

    const updates = [];
    let contadorGenerados = 0;

    for (let i = 1; i < data.length; i++) {
      const rut = data[i][COL.RUT];

      if (!rut || rut === '') {
        updates.push(['', '']);
        continue;
      }

      const { link, formulaQR } = buildRegistroData(cleanRut(rut));
      updates.push([link, formulaQR]);
      contadorGenerados++;
    }

    if (updates.length > 0) {
      sheet.getRange(2, COL.LINK_REGISTRO + 1, updates.length, 2).setValues(updates);
    }

    return {
      success: true,
      message: `✅ Se generaron ${contadorGenerados} links y códigos QR correctamente.`
    };

  } catch (e) {
    return { success: false, message: '❌ Error: ' + e.toString() };
  }
}

/**
 * Regenera el link y QR de registro para un usuario específico (por RUT).
 */
function regenerarQRUsuario(rutInput) {
  try {
    const sheet = getSheet('USUARIOS', CONFIG.HOJAS.USUARIOS);
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.USUARIOS;
    const rutLimpio = cleanRut(rutInput);

    for (let i = 1; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) !== rutLimpio) continue;

      const { link, formulaQR } = buildRegistroData(rutLimpio);
      sheet.getRange(i + 1, COL.LINK_REGISTRO + 1).setValue(link);
      sheet.getRange(i + 1, COL.QR_REGISTRO + 1).setValue(formulaQR);

      return { success: true, message: `✅ QR regenerado para ${data[i][COL.NOMBRE]}` };
    }

    return { success: false, message: '❌ Usuario no encontrado' };

  } catch (e) {
    return { success: false, message: '❌ Error: ' + e.toString() };
  }
}

// ==========================================
// VALIDACIÓN DE RUT CHILENO
// ==========================================

/**
 * Valida el dígito verificador de un RUT chileno (algoritmo módulo 11).
 * @param {string} rutCompleto - RUT con o sin formato (ej: "12.345.678-9" o "123456789")
 * @returns {boolean}
 */
function validarDigitoVerificadorRut(rutCompleto) {
  try {
    const rutLimpio = cleanRut(String(rutCompleto));
    if (!rutLimpio || rutLimpio.length < 2) return false;

    const dv = rutLimpio.slice(-1);
    const cuerpo = rutLimpio.slice(0, -1);

    if (!/^\d+$/.test(cuerpo)) return false;

    let suma = 0;
    let factor = 2;

    for (let i = cuerpo.length - 1; i >= 0; i--) {
      suma += parseInt(cuerpo.charAt(i)) * factor;
      factor = factor === 7 ? 2 : factor + 1;
    }

    const resto = suma % 11;
    const dvEsperado = resto === 1 ? 'K' : resto === 0 ? '0' : String(11 - resto);

    return dv === dvEsperado;

  } catch (e) {
    Logger.log('❌ Error en validarDigitoVerificadorRut: ' + e.toString());
    return false;
  }
}

/**
 * Aplica estilos visuales a una celda según si el RUT es válido o no.
 */
function aplicarEstiloRut(celda, esValido) {
  if (esValido) {
    celda.setValue('RUT VÁLIDO');
    celda.setFontColor('#166534');
    celda.setBackground('#dcfce7');
  } else {
    celda.setValue('RUT NO VÁLIDO');
    celda.setFontColor('#991b1b');
    celda.setBackground('#fee2e2');
  }
  celda.setFontWeight('bold');
}

/**
 * Trigger automático: valida el RUT al editar la columna A de BD_SLIMAPP
 * y escribe el resultado en la columna B.
 *
 * IMPORTANTE: Instalar manualmente como trigger de tipo "onEdit".
 */
function onEditValidarRut(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();

    if (sheet.getName() !== CONFIG.HOJAS.USUARIOS) return;
    if (range.getColumn() !== 1 || range.getRow() === 1) return;

    const valorRut = String(range.getValue()).trim();
    const celdaResultado = sheet.getRange(range.getRow(), 2);

    if (!valorRut || valorRut === '' || valorRut === 'undefined') {
      celdaResultado.setValue('');
      celdaResultado.setFontColor(null);
      celdaResultado.setBackground(null);
      celdaResultado.setFontWeight('normal');
      Logger.log('🧹 Fila ' + range.getRow() + ': RUT vacío → columna B limpiada.');
      return;
    }

    const esValido = validarDigitoVerificadorRut(valorRut);
    aplicarEstiloRut(celdaResultado, esValido);
    Logger.log(`${esValido ? '✅' : '❌'} Fila ${range.getRow()}: RUT "${valorRut}" → ${esValido ? 'VÁLIDO' : 'NO VÁLIDO'}`);

  } catch (err) {
    Logger.log('❌ Error en onEditValidarRut: ' + err.toString());
  }
}

/**
 * Valida en lote todos los RUTs de la hoja que aún no tienen valor en columna B.
 * Ejecutar manualmente una sola vez para poblar registros históricos.
 */
function validarRutsExistentesEnLote() {
  try {
    const sheet = getSheet('USUARIOS', CONFIG.HOJAS.USUARIOS);
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      Logger.log('⚠️ No hay datos en la hoja.');
      return;
    }

    const rangoA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const rangoB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

    const valoresB      = [];
    const coloresFuente = [];
    const coloresFondo  = [];
    const pesos         = [];

    let validados = 0, invalidos = 0, omitidos = 0, sinRut = 0;

    for (let i = 0; i < rangoA.length; i++) {
      const rutRaw       = String(rangoA[i][0]).trim();
      const valorBActual = String(rangoB[i][0]).trim();

      // Si ya tiene valor en columna B, conservar sin cambios
      if (valorBActual !== '' && valorBActual !== '0' && valorBActual.toLowerCase() !== 'false') {
        valoresB.push([valorBActual]);
        coloresFuente.push([null]);
        coloresFondo.push([null]);
        pesos.push([null]);
        omitidos++;
        continue;
      }

      // Sin RUT: dejar vacío
      if (!rutRaw || rutRaw === '0' || rutRaw.toLowerCase() === 'false') {
        valoresB.push(['']);
        coloresFuente.push([null]);
        coloresFondo.push([null]);
        pesos.push([null]);
        sinRut++;
        continue;
      }

      const esValido = validarDigitoVerificadorRut(rutRaw);

      valoresB.push([esValido ? 'RUT VÁLIDO' : 'RUT NO VÁLIDO']);
      coloresFuente.push([esValido ? '#166534' : '#991b1b']);
      coloresFondo.push([esValido ? '#dcfce7' : '#fee2e2']);
      pesos.push(['bold']);
      esValido ? validados++ : invalidos++;
    }

    const rangoEscritura = sheet.getRange(2, 2, lastRow - 1, 1);
    rangoEscritura.setValues(valoresB);
    rangoEscritura.setFontColors(coloresFuente);
    rangoEscritura.setBackgrounds(coloresFondo);
    rangoEscritura.setFontWeights(pesos);

    Logger.log('══════════════════════════════════════════');
    Logger.log('📊 RESUMEN — validarRutsExistentesEnLote');
    Logger.log('   ✅ RUTs válidos    : ' + validados);
    Logger.log('   ❌ RUTs no válidos : ' + invalidos);
    Logger.log('   ⏭️  Ya tenían valor : ' + omitidos);
    Logger.log('   ⚠️  Sin RUT         : ' + sinRut);
    Logger.log('══════════════════════════════════════════');

  } catch (e) {
    Logger.log('❌ Error en validarRutsExistentesEnLote: ' + e.toString());
    throw e;
  }
}
