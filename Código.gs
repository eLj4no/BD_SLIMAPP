// ==========================================
// CONFIGURACIÓN GLOBAL
// ==========================================

const CONFIG = {
  SPREADSHEET_IDS: {
    USUARIOS: '1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM',
    ASISTENCIA: '1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk'
  },
  
  COLUMNAS: {
    USUARIOS: {
      RUT: 0,                           // A
      RUT_VALIDADO: 1,                  // B
      FECHA_INGRESO: 2,                 // C
      NOMBRE: 3,                        // D
      CARGO: 4,                         // E
      CORREO: 5,                        // F
      SITE: 6,                         // G
      REGION: 7,                        // H
      SEXO: 8,                          // I
      ESTADO: 9,                       // J
      DETALLE_DESVINCULACION: 10,       // K
      ID_CREDENCIAL: 11,                // L
      CORREO_REGISTRADO: 12,            // M
      CONTACTO: 13,                     // N
      ROL: 14,                          // O
      LINK_REGISTRO: 15,                // P ⭐
      QR_REGISTRO: 16,                  // Q ⭐
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
 * Obtiene una hoja específica de un spreadsheet
 */
function getSheet(spreadsheetKey, sheetName) {
  const spreadsheetId = CONFIG.SPREADSHEET_IDS[spreadsheetKey];
  if (!spreadsheetId) {
    throw new Error(`Spreadsheet key "${spreadsheetKey}" no encontrado en CONFIG`);
  }
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Hoja "${sheetName}" no encontrada en el spreadsheet`);
  }
  return sheet;
}

/**
 * Limpia un RUT chileno (elimina puntos, guiones y espacios)
 */
function cleanRut(rut) {
  if (!rut) return "";
  return rut.toString().replace(/[.\-\s]/g, '').toUpperCase();
}

// ==========================================
// MENÚ PERSONALIZADO
// ==========================================

/**
 * Crea menús personalizados en Google Sheets al abrir el archivo
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🔧 SLIM - Herramientas')
    .addSubMenu(ui.createMenu('📱 QR Asistencia')
      .addItem('🔄 Generar QR para TODOS los usuarios', 'ejecutarGenerarQRTodos')
      .addSeparator()
      .addItem('📋 Instrucciones de configuración', 'mostrarInstruccionesQR')
    )
    .addToUi();
}

/**
 * Ejecuta la generación masiva de QR con confirmación
 */
function ejecutarGenerarQRTodos() {
  const ui = SpreadsheetApp.getUi();
  
  // Confirmación antes de ejecutar
  const response = ui.alert(
    'Generar QR de Registro',
    '¿Estás seguro de que deseas generar los códigos QR para TODOS los usuarios?\n\n' +
    'Esta acción sobrescribirá los datos actuales en las columnas Q (Link Registro) y R (QR Registro).',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // Mostrar mensaje de procesamiento
    const toast = SpreadsheetApp.getActiveSpreadsheet();
    toast.toast('Generando QR... Por favor espera.', '⏳ Procesando', -1);
    
    // Ejecutar la función
    const resultado = generarQRRegistroUsuarios();
    
    // Mostrar resultado
    if (resultado.success) {
      ui.alert('✅ Completado', resultado.message, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Error', resultado.message, ui.ButtonSet.OK);
    }
    
    toast.toast('Proceso finalizado', '✅ Listo', 3);
  }
}

/**
 * Muestra las instrucciones para configurar el sistema QR
 */
function mostrarInstruccionesQR() {
  const ui = SpreadsheetApp.getUi();
  
  const instrucciones = 
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
    '✅ Los códigos QR aparecerán en las columnas Q y R';
  
  ui.alert('📚 Instrucciones', instrucciones, ui.ButtonSet.OK);
}

// ==========================================
// GENERADOR AUTOMÁTICO DE QR DE REGISTRO
// ==========================================

/**
 * Genera automáticamente los Links y QR de registro para todos los usuarios
 */
function generarQRRegistroUsuarios() {
  try {
    if (CONFIG.WEB_APP.URL.includes('REEMPLAZAR')) {
      throw new Error("⚠️ Primero configura la URL de tu Web App en CONFIG.WEB_APP.URL");
    }
    
    const sheet = getSheet('USUARIOS', 'BD_SLIMAPP');
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.USUARIOS;
    
    const COL_LINK_REGISTRO = 15;  // Columna P
    const COL_QR_REGISTRO = 16;    // Columna Q
    
    const urlBase = CONFIG.WEB_APP.URL;
    const updates = [];
    let contadorGenerados = 0;
    
    for (let i = 1; i < data.length; i++) {
      const rut = data[i][COL.RUT];
      
      if (!rut || rut === "") {
        updates.push(["", ""]);
        continue;
      }
      
      const rutLimpio = cleanRut(rut);
      
      // ✅ Link de registro
      const linkRegistro = `${urlBase}?action=register&rut=${rutLimpio}`;
      
      // ✅ Fórmula QR con encodeURIComponent
      const formulaQR = `=IMAGE("https://quickchart.io/qr?size=300&text=${encodeURIComponent(linkRegistro)}")`;
      
      updates.push([linkRegistro, formulaQR]);
      contadorGenerados++;
    }
    
    if (updates.length > 0) {
      const rangeToUpdate = sheet.getRange(2, COL_LINK_REGISTRO + 1, updates.length, 2);
      rangeToUpdate.setValues(updates);
    }
    
    return {
      success: true,
      message: `✅ Se generaron ${contadorGenerados} links y códigos QR correctamente.`,
      total: contadorGenerados
    };
    
  } catch (e) {
    return {
      success: false,
      message: "❌ Error: " + e.toString()
    };
  }
}

/**
 * Genera QR para un usuario específico (por RUT)
 * Útil para regenerar QR de un solo usuario sin afectar a los demás
 */
function regenerarQRUsuario(rutInput) {
  try {
    if (CONFIG.WEB_APP.URL.includes('REEMPLAZAR')) {
      throw new Error("Primero configura la URL de Web App en CONFIG.WEB_APP.URL");
    }
    
    const sheet = getSheet('USUARIOS', 'BD_SLIMAPP');
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.USUARIOS;
    
    // ✅ COLUMNAS CORRECTAS: P y Q
    const COL_LINK_REGISTRO = 15;  // Columna P
    const COL_QR_REGISTRO = 16;    // Columna Q
    
    const rutLimpio = cleanRut(rutInput);
    
    // Buscar el usuario
    for (let i = 1; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
        const urlBase = CONFIG.WEB_APP.URL;
        const linkRegistro = `${urlBase}?action=register&rut=${rutLimpio}`;
        const formulaQR = `=IMAGE("https://quickchart.io/qr?size=300&text=${encodeURIComponent(linkRegistro)}")`;
        
        // Actualizar solo esa fila
        sheet.getRange(i + 1, COL_LINK_REGISTRO + 1).setValue(linkRegistro);
        sheet.getRange(i + 1, COL_QR_REGISTRO + 1).setValue(formulaQR);
        
        return {
          success: true,
          message: `✅ QR regenerado para ${data[i][COL.NOMBRE]}`
        };
      }
    }
    
    return {
      success: false,
      message: "❌ Usuario no encontrado"
    };
    
  } catch (e) {
    return {
      success: false,
      message: "❌ Error: " + e.toString()
    };
  }
}

// ==========================================
// VALIDACIÓN AUTOMÁTICA DE RUT - HOJA BD_SLIMAPP
// Detecta cambios en columna A (RUT) y escribe el
// resultado en columna B (RUT VALIDADO)
// ==========================================

/**
 * Valida el dígito verificador de un RUT chileno.
 * Utiliza el algoritmo estándar de módulo 11.
 * @param {string} rutCompleto - RUT con o sin formato (ej: "12.345.678-9" o "123456789")
 * @returns {boolean} true si el RUT es matemáticamente válido
 */
function validarDigitoVerificadorRut(rutCompleto) {
  try {
    // Limpiar usando la función existente del proyecto
    const rutLimpio = cleanRut(String(rutCompleto));
    
    if (!rutLimpio || rutLimpio.length < 2) return false;
    
    // Separar cuerpo y dígito verificador
    const dv    = rutLimpio.slice(-1).toUpperCase();
    const cuerpo = rutLimpio.slice(0, -1);
    
    // El cuerpo debe ser numérico
    if (!/^\d+$/.test(cuerpo)) return false;
    
    // Calcular dígito verificador esperado con módulo 11
    let suma    = 0;
    let factor  = 2;
    
    for (let i = cuerpo.length - 1; i >= 0; i--) {
      suma   += parseInt(cuerpo.charAt(i)) * factor;
      factor  = factor === 7 ? 2 : factor + 1;
    }
    
    const resto    = suma % 11;
    const dvEsperado = resto === 1 ? 'K' : resto === 0 ? '0' : String(11 - resto);
    
    return dv === dvEsperado;
    
  } catch (e) {
    Logger.log('❌ Error en validarDigitoVerificadorRut: ' + e.toString());
    return false;
  }
}

/**
 * Trigger automático que se ejecuta al editar cualquier celda
 * de la hoja BD_SLIMAPP en el Spreadsheet de USUARIOS.
 * 
 * - Solo actúa cuando se edita la COLUMNA A (RUT)
 * - Escribe "RUT VÁLIDO" o "RUT NO VÁLIDO" en COLUMNA B (RUT VALIDADO)
 * - Ignora la fila de encabezado (fila 1)
 * - Si la celda A queda vacía, limpia el valor de B
 * 
 * IMPORTANTE: Este trigger debe instalarse manualmente una sola vez
 * (ver instrucciones de implementación).
 */
function onEditValidarRut(e) {
  try {
    const range  = e.range;
    const sheet  = range.getSheet();
    const col    = range.getColumn();
    const fila   = range.getRow();
    
    // Solo actuar en la hoja BD_SLIMAPP
    if (sheet.getName() !== CONFIG.HOJAS.USUARIOS) return;
    
    // Solo actuar en la columna A (columna 1) y no en el encabezado
    if (col !== 1 || fila === 1) return;
    
    const valorRut = String(range.getValue()).trim();
    const celdaResultado = sheet.getRange(fila, 2); // Columna B
    
    // Si la celda quedó vacía, limpiar el resultado
    if (!valorRut || valorRut === '' || valorRut === 'undefined') {
      celdaResultado.setValue('');
      Logger.log('🧹 Fila ' + fila + ': RUT vacío → columna B limpiada.');
      return;
    }
    
    // Validar el RUT
    const esValido = validarDigitoVerificadorRut(valorRut);
    
    if (esValido) {
      celdaResultado.setValue('RUT VÁLIDO');
      celdaResultado.setFontColor('#166534');       // Verde oscuro
      celdaResultado.setBackground('#dcfce7');      // Verde claro
      celdaResultado.setFontWeight('bold');
      Logger.log('✅ Fila ' + fila + ': RUT "' + valorRut + '" → VÁLIDO');
    } else {
      celdaResultado.setValue('RUT NO VÁLIDO');
      celdaResultado.setFontColor('#991b1b');       // Rojo oscuro
      celdaResultado.setBackground('#fee2e2');      // Rojo claro
      celdaResultado.setFontWeight('bold');
      Logger.log('❌ Fila ' + fila + ': RUT "' + valorRut + '" → NO VÁLIDO');
    }
    
  } catch (err) {
    Logger.log('❌ Error en onEditValidarRut: ' + err.toString());
  }
}

/**
 * Valida en lote todos los RUTs existentes en la hoja BD_SLIMAPP
 * que aún no tienen valor en la columna B (RUT VALIDADO).
 * 
 * Ejecutar manualmente SOLO UNA VEZ desde el editor de Apps Script
 * para poblar los registros históricos.
 */
function validarRutsExistentesEnLote() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BD_SLIMAPP');
    
    if (!sheet) {
      Logger.log('❌ No se encontró la hoja: BD_SLIMAPP');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('⚠️ No hay datos en la hoja.');
      return;
    }
    
    // Leer columnas A y B en bloque (más eficiente)
    const rangoA = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Col A: RUT
    const rangoB = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Col B: RUT VALIDADO
    
    let validados   = 0;
    let invalidos   = 0;
    let omitidos    = 0;
    let sinRut      = 0;
    
    // Preparar arrays de valores y estilos para escritura en bloque
    const valoresB      = [];
    const coloresFuente = [];
    const coloresFondo  = [];
    const pesos         = [];
    
    for (let i = 0; i < rangoA.length; i++) {
      const rutRaw      = String(rangoA[i][0]).trim();
      const valorBActual = String(rangoB[i][0]).trim();
      
      // Si ya tiene valor en columna B, omitir
      if (valorBActual !== '' && valorBActual !== '0' && valorBActual.toLowerCase() !== 'false') {
        valoresB.push([valorBActual]);
        coloresFuente.push([null]);
        coloresFondo.push([null]);
        pesos.push([null]);
        omitidos++;
        continue;
      }
      
      // Si no hay RUT, dejar vacío
      if (!rutRaw || rutRaw === '' || rutRaw === '0' || rutRaw.toLowerCase() === 'false') {
        valoresB.push(['']);
        coloresFuente.push([null]);
        coloresFondo.push([null]);
        pesos.push([null]);
        sinRut++;
        continue;
      }
      
      // Validar RUT
      const esValido = validarDigitoVerificadorRut(rutRaw);
      
      if (esValido) {
        valoresB.push(['VÁLIDO']);
        coloresFuente.push(['#166534']);
        coloresFondo.push(['#dcfce7']);
        pesos.push(['bold']);
        validados++;
      } else {
        valoresB.push(['NO VÁLIDO']);
        coloresFuente.push(['#991b1b']);
        coloresFondo.push(['#fee2e2']);
        pesos.push(['bold']);
        invalidos++;
      }
    }
    
    // Escribir todos los resultados en bloque
    const rangoEscritura = sheet.getRange(2, 2, lastRow - 1, 1);
    rangoEscritura.setValues(valoresB);
    rangoEscritura.setFontColors(coloresFuente);
    rangoEscritura.setBackgrounds(coloresFondo);
    rangoEscritura.setFontWeights(pesos);
    
    Logger.log('══════════════════════════════════════════');
    Logger.log('📊 RESUMEN — validarRutsExistentesEnLote');
    Logger.log('   ✅ RUTs válidos      : ' + validados);
    Logger.log('   ❌ RUTs no válidos   : ' + invalidos);
    Logger.log('   ⏭️  Ya tenían valor   : ' + omitidos);
    Logger.log('   ⚠️  Sin RUT           : ' + sinRut);
    Logger.log('══════════════════════════════════════════');
    
  } catch (e) {
    Logger.log('❌ Error en validarRutsExistentesEnLote: ' + e.toString());
    throw e;
  }
}
