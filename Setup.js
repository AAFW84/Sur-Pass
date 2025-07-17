// =====================================================
// ARCHIVO: Setup.js - VERSIÓN CORREGIDA Y OPTIMIZADA
// FUNCIÓN PRINCIPAL ROBUSTA - Maneja todos los casos posibles
// Versión: 3.0 - Compatible con funciones.js
// =====================================================

/**
 * FUNCIÓN PRINCIPAL ROBUSTA - Maneja todos los casos posibles
 * Asegura que todas las hojas requeridas para SurPass existan o se creen.
 * ✅ OPTIMIZADO: Elimina duplicidades y mejora la validación.
 */
function createSurPassSheets() {
    console.log('🚀 Iniciando SurPass Setup - Versión 3.0 Corregida');

    let spreadsheet = null;
    let creadoNuevo = false;

    try {
        // PASO 1: Obtener spreadsheet con múltiples estrategias
        console.log('📋 Paso 1: Obteniendo spreadsheet...');
        const resultado = obtenerSpreadsheetRobusto();
        
        // ✅ VALIDACIÓN CRÍTICA: Verificar que el resultado sea válido
        if (!resultado || typeof resultado !== 'object') {
            throw new Error('obtenerSpreadsheetRobusto() devolvió un resultado inválido');
        }
        
        spreadsheet = resultado.spreadsheet;
        creadoNuevo = resultado.creadoNuevo;

        // ✅ VALIDACIÓN CRÍTICA: Verificar que spreadsheet sea válido
        if (!spreadsheet) {
            throw new Error('No se pudo obtener o crear una hoja de cálculo válida.');
        }

        // ✅ VALIDACIÓN ADICIONAL: Verificar que spreadsheet tenga métodos esperados
        if (typeof spreadsheet.getSheetByName !== 'function' || typeof spreadsheet.insertSheet !== 'function') {
            throw new Error('El objeto spreadsheet no tiene los métodos necesarios.');
        }

        console.log(`✅ Spreadsheet obtenido: ${spreadsheet.getName()} (ID: ${spreadsheet.getId()})`);
        if (creadoNuevo) {
            SpreadsheetApp.getUi().alert(
                '¡Hoja de Cálculo Creada!',
                'Se ha creado una nueva hoja de cálculo para SurPass. Por favor, no cambie el nombre.',
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }

        // PASO 2: Verificar y crear hojas requeridas
        console.log('📝 Paso 2: Verificando y creando hojas requeridas...');
        const hojasRequeridas = ['Configuracion', 'Historial', 'BaseDeDatos', 'Evacuacion'];
        const hojasCreadas = [];
        const hojasExistentes = [];

        hojasRequeridas.forEach(nombreHoja => {
            let hoja = spreadsheet.getSheetByName(nombreHoja);
            if (!hoja) {
                hoja = spreadsheet.insertSheet(nombreHoja);
                hojasCreadas.push(nombreHoja);
                console.log(`➕ Hoja '${nombreHoja}' creada.`);
                // Opcional: Configurar encabezados por defecto para nuevas hojas
                if (nombreHoja === 'Configuracion') {
                    hoja.getRange('A1').setValue('Clave');
                    hoja.getRange('B1').setValue('Valor');
                } else if (nombreHoja === 'Historial') {
                    hoja.getRange('A1').setValue('Timestamp');
                    hoja.getRange('B1').setValue('TipoEvento');
                    hoja.getRange('C1').setValue('Detalles');
                    hoja.getRange('D1').setValue('Usuario');
                } else if (nombreHoja === 'BaseDeDatos') {
                    hoja.getRange('A1').setValue('ID');
                    hoja.getRange('B1').setValue('Nombre');
                    hoja.getRange('C1').setValue('Apellido');
                    hoja.getRange('D1').setValue('Departamento');
                    hoja.getRange('E1').setValue('Estado'); // Ej: Presente, Ausente, Evacuado
                } else if (nombreHoja === 'Evacuacion') {
                    hoja.getRange('A1').setValue('EstadoEvacuacion'); // Activa, Inactiva
                    hoja.getRange('B1').setValue('TimestampInicio');
                    hoja.getRange('C1').setValue('TimestampFin');
                    hoja.getRange('D1').setValue('TipoEvacuacion'); // Real, Simulacro
                    hoja.getRange('E1').setValue('NotasSimulacro');
                    hoja.getRange('F1').setValue('Responsable');
                }
            } else {
                hojasExistentes.push(nombreHoja);
                console.log(`✔️ Hoja '${nombreHoja}' ya existe.`);
            }
        });

        // PASO 3: Mensaje de éxito final
        let mensajeFinal = '✅ SurPass Setup completado con éxito.\n\n';
        if (hojasCreadas.length > 0) {
            mensajeFinal += `Hojas creadas: ${hojasCreadas.join(', ')}.\n`;
        }
        if (hojasExistentes.length > 0) {
            mensajeFinal += `Hojas ya existentes: ${hojasExistentes.join(', ')}.\n`;
        }
        mensajeFinal += '\nEl sistema está listo para ser utilizado.';

        SpreadsheetApp.getUi().alert(
            '¡Setup Completo!',
            mensajeFinal,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        console.log('🎉 SurPass Setup finalizado con éxito.');

    } catch (error) {
        console.error('❌ Error crítico en createSurPassSheets:', error);
        SpreadsheetApp.getUi().alert(
            'Error de Setup',
            'Ha ocurrido un error crítico durante la configuración de SurPass: ' + error.message + '\n\nPor favor, contacte al administrador.',
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

/**
 * Función auxiliar robusta para obtener o crear una hoja de cálculo.
 * Intenta obtener la hoja de cálculo activa o la crea si no existe.
 * @returns {Object} Un objeto con la hoja de cálculo y un booleano indicando si fue creada.
 */
function obtenerSpreadsheetRobusto() {
    let spreadsheet = null;
    let creadoNuevo = false;

    // Intento 1: Obtener la hoja de cálculo activa
    try {
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        if (spreadsheet) {
            console.log('✅ Hoja de cálculo activa encontrada.');
            return { spreadsheet: spreadsheet, creadoNuevo: false };
        }
    } catch (e) {
        console.warn('⚠️ No se pudo obtener la hoja de cálculo activa:', e.message);
    }

    // Intento 2: Intentar abrir por ID si hay alguna propiedad guardada (no implementado en este fragmento, pero sería el siguiente paso)
    // if (propertiesService.getUserProperties().getProperty('spreadsheetId')) { ... }

    // Intento 3: Crear una nueva hoja de cálculo si no se encontró ninguna
    try {
        console.log('⚙️ Creando nueva hoja de cálculo...');
        spreadsheet = SpreadsheetApp.create('SurPass_Datos');
        creadoNuevo = true;
        console.log('✅ Nueva hoja de cálculo creada con éxito.');
        return { spreadsheet: spreadsheet, creadoNuevo: true };
    } catch (e) {
        console.error('❌ Error al crear nueva hoja de cálculo:', e.message);
        throw new Error('No se pudo obtener ni crear una hoja de cálculo. ' + e.message);
    }
}

/**
 * Verifica el estado completo del setup de SurPass.
 * @returns {Object} Objeto con el estado del setup.
 */
function verificarSetupCompleto() {
    console.log('🔍 Verificando setup completo...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojasRequeridas = ['Configuracion', 'Historial', 'BaseDeDatos', 'Evacuacion'];
    const hojasEncontradas = [];
    const hojasFaltantes = [];
    let spreadsheetId = null;

    if (!ss) {
        return {
            success: false,
            resumen: 'No se encontró una hoja de cálculo activa.',
            hojasEncontradas: [],
            hojasFaltantes: hojasRequeridas,
            spreadsheetId: null
        };
    }

    spreadsheetId = ss.getId();

    hojasRequeridas.forEach(nombreHoja => {
        if (ss.getSheetByName(nombreHoja)) {
            hojasEncontradas.push(nombreHoja);
        } else {
            hojasFaltantes.push(nombreHoja);
        }
    });

    const success = hojasFaltantes.length === 0;
    const resumen = success ?
        'Todas las hojas requeridas están presentes.' :
        `Faltan las siguientes hojas: ${hojasFaltantes.join(', ')}.`;

    console.log('✅ Verificación de setup completada:', { success, resumen, hojasEncontradas, hojasFaltantes });
    return {
        success: success,
        resumen: resumen,
        hojasEncontradas: hojasEncontradas,
        hojasFaltantes: hojasFaltantes,
        spreadsheetId: spreadsheetId
    };
}

/**
 * Ejecuta un diagnóstico de la configuración de la UI para el setup.
 * Es una función de alto nivel para el diagnóstico de la UI.
 * Se ha consolidado la lógica de diagnóstico con `verificarSetupCompleto`.
 */
function diagnosticarSetupUI() {
    try {
        console.log('🩺 Ejecutando diagnóstico de UI para Setup...');
        const estado = verificarSetupCompleto(); // Reutilizamos la lógica principal de verificación

        let mensaje = `Estado del Setup de SurPass:\n\n`;
        mensaje += `Resumen: ${estado.resumen}\n\n`;
        mensaje += `Hojas encontradas: ${estado.hojasEncontradas.join(', ') || 'Ninguna'}\n`;
        mensaje += `Hojas faltantes: ${estado.hojasFaltantes.join(', ') || 'Ninguna'}\n`;
        if (estado.spreadsheetId) {
            mensaje += `ID de Hoja de Cálculo: ${estado.spreadsheetId}\n`;
        }

        if (estado.success) {
            mensaje += '\nEl setup parece estar correcto. ¡Listo para usar!';
        } else {
            mensaje += '\nSe detectaron problemas. Por favor, ejecute `createSurPassSheets()` para corregir.';
        }

        SpreadsheetApp.getUi().alert(
            'Diagnóstico de Setup',
            mensaje,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        console.log('Diagnóstico de UI para Setup finalizado.');

    } catch (error) {
        console.error('❌ Error ejecutando diagnóstico de Setup UI:', error);
        SpreadsheetApp.getUi().alert(
            'Error en Diagnóstico',
            'Error ejecutando diagnóstico: ' + error.message,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

// =====================================================
// LOGGING SIMPLE PARA SETUP (Unificado con funciones.js)
// =====================================================

/**
 * ✅ FUNCIÓN DE LOGGING SIMPLE - Para cuando funciones.js no esté disponible
 * o para logs iniciales de setup.
 * Se recomienda usar logError de funciones.js una vez que esté cargado.
 */
function logSimple(mensaje, nivel = 'INFO') {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] [SETUP] [${nivel}] ${mensaje}`;
    console.log(logMessage);
    
    // También log a Google Apps Script Logger
    Logger.log(logMessage);
}

// =====================================================
// EXPORT PARA COMPATIBILIDAD
// =====================================================

/**
 * ✅ FUNCIÓN DE COMPATIBILIDAD - Para exportar configuración setup
 */
function exportarConfiguracionSetup() {
    try {
        const estado = verificarSetupCompleto();
        const config = {
            setupCompleto: estado.success,
            hojas: estado.hojasEncontradas,
            faltantes: estado.hojasFaltantes,
            spreadsheetId: estado.spreadsheetId,
            version: '3.0',
            timestamp: new Date().toISOString()
        };
        
        console.log('📤 Configuración setup exportada:', config);
        return config;
        
    } catch (error) {
        console.error('❌ Error exportando configuración setup:', error);
        return { error: error.message };
    }
}
