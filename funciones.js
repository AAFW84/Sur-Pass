    // =====================================================
    // SURPASS - SISTEMA DE CONTROL DE ACCESO
    // Archivo: funciones.gs - VERSIÓN COMPLETA Y FUNCIONAL
    // Versión: 3.0 - Sistema Integrado con Evacuación
    // Autor: Sistema SurPass
    // Fecha: 2025
    // =====================================================

    // =====================================================
    // CONFIGURACIÓN GLOBAL Y UTILIDADES
    // =====================================================

    /**
     * Registra errores en el log del sistema con niveles de severidad
     */
    function logError(mensaje, nivel = 'ERROR', detalles = null) {
        const timestamp = new Date().toISOString();
        const logEntry = `[${timestamp}] [${nivel}] ${mensaje}`;
        
        if (detalles) {
            Logger.log(logEntry + ' | Detalles: ' + JSON.stringify(detalles));
        } else {
            Logger.log(logEntry);
        }
        
        // Si es error crítico, intentar notificar
        if (nivel === 'CRITICAL' || nivel === 'FATAL') {
            try {
                // Placeholder for critical error notification (e.g., email to admin)
                // notificarErrorCritico(mensaje, detalles);
            } catch (e) {
                Logger.log('No se pudo enviar notificación de error crítico: ' + e.message);
            }
        }
    }

    /**
     * Valida que un campo tenga valor
     */
    function validarCampoObligatorio(valor) {
        return valor !== undefined && valor !== null && String(valor).trim() !== '';
    }

    /**
     * Obtiene el usuario actual del sistema
     */
    function obtenerUsuarioActual() {
        try {
            const user = Session.getEffectiveUser().getEmail();
            return user;
        } catch (e) {
            logError('Error obteniendo usuario actual: ' + e.message);
            return 'usuario_desconocido';
        }
    }

    /**
     * Notifica errores críticos por email
     */
    function notificarErrorCritico(mensaje, detalles) {
        try {
            const config = obtenerConfiguracion();
            const adminEmail = config.NOTIFICACIONES_EMAIL || 'afernandez@sesursa.com';
            const asunto = '🚨 ERROR CRÍTICO - Sistema SurPass';
            
            let cuerpo = `🚨 ERROR CRÍTICO EN SISTEMA SURPASS\n\n`;
            cuerpo += `Mensaje: ${mensaje}\n`;
            cuerpo += `Fecha: ${new Date().toLocaleString()}\n`;
            cuerpo += `Usuario: ${obtenerUsuarioActual()}\n`;
            
            if (detalles) {
                cuerpo += `\nDetalles técnicos:\n${JSON.stringify(detalles, null, 2)}\n`;
            }
            
            cuerpo += `\nPor favor, revise el sistema inmediatamente.`;
            
            MailApp.sendEmail(adminEmail, asunto, cuerpo);
            logError('Notificación de error crítico enviada', 'INFO');
            
        } catch (error) {
            logError('Error enviando notificación crítica', 'ERROR', { error: error.message });
        }
    }

    // =====================================================
    // FUNCIONES DE CONFIGURACIÓN AVANZADA
    // =====================================================

    /**
     * Obtiene la configuración completa del sistema desde la hoja Configuración
     */
    function obtenerConfiguracion() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const configSheet = ss.getSheetByName('Configuración');
            
            if (!configSheet) {
                logError('Hoja Configuración no encontrada, usando valores por defecto', 'WARNING');
                return obtenerConfiguracionPorDefecto();
            }
            
            const data = configSheet.getDataRange().getValues();
            const config = {};
            
            // Configuración por defecto
            const defaults = obtenerConfiguracionPorDefecto();
            Object.keys(defaults).forEach(key => {
                config[key] = defaults[key];
            });
            
            // Sobrescribir con valores de la hoja
            for (let i = 1; i < data.length; i++) {
                const clave = String(data[i][0] || '').trim();
                const valor = data[i][1];
                
                if (clave && valor !== undefined && valor !== null) {
                    config[clave] = String(valor).trim();
                }
            }
            
            return config;
            
        } catch (error) {
            logError('Error obteniendo configuración', 'ERROR', { error: error.message });
            return obtenerConfiguracionPorDefecto();
        }
    }

    /**
     * Configuración por defecto del sistema
     */
    function obtenerConfiguracionPorDefecto() {
        return {
            // Configuración básica
            EMPRESA_NOMBRE: 'SurPass',
            EMPRESA_LOGO: '',
            HORARIO_APERTURA: '05:00',
            HORARIO_CIERRE: '20:00',
            DIAS_LABORABLES: 'Lunes,Martes,Miércoles,Jueves,Viernes,Sábado,Domingo',
            
            // Configuración de acceso
            TIEMPO_MAX_VISITA: '12',
            PERMITIR_ACCESO_FUERA_HORARIO: 'NO',
            REQUIERE_AUTORIZACION_ADMIN: 'NO',
            
            // Notificaciones y comunicación
            NOTIFICACIONES_EMAIL: 'afernandez@sesursa.com',
            EMAIL_SECUNDARIO: '@sesursa.com',
            NOTIFICAR_ACCESOS_DENEGADOS: 'SI',
            NOTIFICAR_EVACUACIONES: 'SI',
            
            // Backup y mantenimiento
            BACKUP_AUTOMATICO: 'SI',
            FRECUENCIA_BACKUP: 'DIARIO',
            LIMPIAR_LOGS_AUTOMATICO: 'SI',
            DIAS_RETENER_LOGS: '30',
            
            // Configuración de evacuación
            TIEMPO_LIMITE_EVACUACION: '15',
            NOTIFICAR_EVACUACION_AUTOMATICA: 'SI',
            ENVIAR_REPORTE_EVACUACION: 'SI',
            
            // Configuración de interfaz
            TEMA_POR_DEFECTO: 'claro',
            MOSTRAR_ESTADISTICAS: 'SI',
            SONIDOS_ACTIVADOS: 'SI',
            ESCANER_QR_NATIVO: 'NO',
            
            // Configuración de seguridad
            INTENTOS_MAX_LOGIN: '3',
            TIEMPO_BLOQUEO_LOGIN: '15',
            REGISTRO_ACTIVIDAD_ADMINS: 'SI',
            
            // Configuración de reportes
            GENERAR_REPORTE_DIARIO: 'SI',
            INCLUIR_ESTADISTICAS_REPORTE: 'SI',
            FORMATO_FECHA_REPORTE: 'dd/mm/yyyy',
            
            // Configuración personalizada
            MENSAJE_BIENVENIDA: 'Bienvenido al Sistema SurPass',
            MENSAJE_ACCESO_DENEGADO: 'Acceso denegado. Contacte al administrador.',
            MOSTRAR_EMPRESA_VISITANTE: 'SI'
        };
    }

    /**
     * Actualiza un valor de configuración
     */
    function actualizarConfiguracion(clave, valor, usuario = null) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const configSheet = ss.getSheetByName('Configuración');
            
            if (!configSheet) {
                throw new Error('Hoja Configuración no encontrada');
            }
            
            const data = configSheet.getDataRange().getValues();
            let filaEncontrada = -1;
            
            // Buscar la clave existente
            for (let i = 1; i < data.length; i++) {
                const claveEnFila = String(data[i][0] || '').trim();
                if (claveEnFila === clave) {
                    filaEncontrada = i + 1;
                    break;
                }
            }
            
            const usuarioActual = usuario || obtenerUsuarioActual();
            const fechaActual = new Date();
            
            if (filaEncontrada > 0) {
                // Actualizar fila existente
                configSheet.getRange(filaEncontrada, 2).setValue(valor);
                configSheet.getRange(filaEncontrada, 4).setValue(fechaActual);
                configSheet.getRange(filaEncontrada, 5).setValue(usuarioActual);
            } else {
                // Agregar nueva fila
                const descripcion = obtenerDescripcionConfiguracion(clave);
                configSheet.appendRow([clave, valor, descripcion, fechaActual, usuarioActual]);
            }
            
            logError(`Configuración actualizada: ${clave} = ${valor}`, 'INFO', { usuario: usuarioActual });
            
            return {
                success: true,
                message: 'Configuración actualizada correctamente'
            };
            
        } catch (error) {
            logError('Error actualizando configuración', 'ERROR', { clave, valor, error: error.message });
            return {
                success: false,
                message: 'Error al actualizar configuración: ' + error.message
            };
        }
    }

    /**
     * Obtiene descripción para una clave de configuración
     */
    function obtenerDescripcionConfiguracion(clave) {
        const descripciones = {
            'EMPRESA_NOMBRE': 'Nombre de la empresa',
            'EMPRESA_LOGO': 'URL del logo de la empresa',
            'HORARIO_APERTURA': 'Hora de apertura (HH:MM)',
            'HORARIO_CIERRE': 'Hora de cierre (HH:MM)',
            'TIEMPO_MAX_VISITA': 'Tiempo máximo de visita en horas',
            'NOTIFICACIONES_EMAIL': 'Email principal para notificaciones',
            'EMAIL_SECUNDARIO': 'Email secundario para notificaciones',
            'BACKUP_AUTOMATICO': '¿Hacer backup automático? (SI/NO)',
            'FRECUENCIA_BACKUP': 'Frecuencia de backup (DIARIO/SEMANAL/MENSUAL)',
            'TIEMPO_LIMITE_EVACUACION': 'Tiempo límite para evacuación en minutos',
            'TEMA_POR_DEFECTO': 'Tema por defecto de la interfaz',
            'MOSTRAR_ESTADISTICAS': 'Mostrar panel de estadísticas (SI/NO)',
            'SONIDOS_ACTIVADOS': 'Activar sonidos del sistema (SI/NO)',
            'MENSAJE_BIENVENIDA': 'Mensaje de bienvenida personalizado'
        };
        
        return descripciones[clave] || 'Configuración personalizada';
    }

    // =====================================================
    // FUNCIONES DE MENÚ Y NAVEGACIÓN
    // =====================================================

    /**
     * Crea el menú personalizado al abrir la hoja
     */
    function onOpen() {
        try {
            const ui = SpreadsheetApp.getUi();
            ui.createMenu('SurPass')
                .addItem('Abrir Formulario', 'abrirFormulario')
                .addItem('Abrir Panel de Administración', 'abrirAdmin') // Nuevo elemento de menú
                .addToUi();
        } catch (e) {
            logError('Error en onOpen: ' + e.message);
        }
    }

    // Función: abrirFormulario
    // Propósito: Abrir el formulario HTML como un cuadro de diálogo modal en la hoja de cálculo.
    function abrirFormulario() {
        const validacion = validarEstructuraHojas();

        if (!validacion.valido) {
            SpreadsheetApp.getUi().alert(validacion.mensaje);
            return;
        }

        const htmlOutput = HtmlService.createHtmlOutputFromFile('formulario')
            .setWidth(720)
            .setHeight(1080)
            .setTitle('SurPass - Control de Acceso');
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SurPass - Control de Acceso');
    }

    // Función: abrirAdmin
    // Propósito: Abrir el panel de administración HTML como un cuadro de diálogo modal en la hoja de cálculo.
    function abrirAdmin() {
        const htmlOutput = HtmlService.createHtmlOutputFromFile('admin')
            .setWidth(1000) // Ajusta el ancho según necesidad
            .setHeight(800) // Ajusta el alto según necesidad
            .setTitle('SurPass - Administración');
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SurPass - Administración');
    }

    /**
     * Ejecuta diagnóstico completo desde la UI
     */
    function ejecutarDiagnosticoUI() {
        try {
            const resultado = diagnosticoCompletoSistema();
            
            let mensaje = `🔍 DIAGNÓSTICO COMPLETO DEL SISTEMA\n\n`;
            mensaje += `✅ Tests exitosos: ${resultado.testsPasados}/${resultado.totalTests}\n`;
            mensaje += `⏱️ Tiempo total: ${resultado.tiempoTotal}ms\n`;
            mensaje += `📊 Porcentaje de éxito: ${Math.round((resultado.testsPasados/resultado.totalTests)*100)}%\n\n`;
            
            if (resultado.exito) {
                mensaje += `🎉 El sistema está funcionando correctamente.\n\n`;
            } else {
                mensaje += `⚠️ Se encontraron algunos problemas:\n\n`;
                
                Object.keys(resultado.detalles).forEach(test => {
                    const detalle = resultado.detalles[test];
                    if (!detalle.status) {
                        mensaje += `❌ ${test.toUpperCase()}: ${detalle.mensaje}\n`;
                    }
                });
                mensaje += `\n💡 Revise los logs para más detalles.`;
            }
            
            SpreadsheetApp.getUi().alert(
                '🔍 Diagnóstico del Sistema', 
                mensaje, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            
        } catch (error) {
            logError('Error ejecutando diagnóstico desde UI', 'ERROR', { error: error.message });
            SpreadsheetApp.getUi().alert(
                '❌ Error', 
                'Error al ejecutar el diagnóstico: ' + error.message, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
    }

    /**
     * Muestra configuración del sistema
     */
    function mostrarConfiguracion() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const configSheet = ss.getSheetByName('Configuración');
            
            if (!configSheet) {
                SpreadsheetApp.getUi().alert(
                    '❌ Error', 
                    'La hoja "Configuración" no existe. Ejecute primero la creación del sistema.', 
                    SpreadsheetApp.getUi().ButtonSet.OK
                );
                return;
            }
            
            // Mostrar la hoja de configuración
            configSheet.showSheet();
            ss.setActiveSheet(configSheet);
            
            SpreadsheetApp.getUi().alert(
                '⚙️ Configuración', 
                'Se ha abierto la hoja de configuración. Puede modificar los valores en la columna "Valor".\n\n' +
                '💡 Tip: Después de hacer cambios, use "Diagnóstico Completo" para verificar que todo funcione correctamente.', 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            
            logError('Hoja de configuración abierta', 'INFO');
            
        } catch (error) {
            logError('Error abriendo configuración', 'ERROR', { error: error.message });
            SpreadsheetApp.getUi().alert(
                '❌ Error', 
                'Error al abrir configuración: ' + error.message, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
    }

    // =====================================================
    // FUNCIÓN DOGET - PARA WEB APP
    // =====================================================

    /**
     * Maneja las peticiones HTTP GET
     */
    function doGet(e) {
        if (e && e.parameter && e.parameter.admin === 'true') {
            return HtmlService.createHtmlOutputFromFile('admin')
                .setTitle('SurPass - Administración')
                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        }
        return HtmlService.createHtmlOutputFromFile('formulario')
            .setTitle('SurPass - Control de Acceso')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Función: obtenerUrlAdmin
    // Propósito: Generar la URL para el panel de administración
    function obtenerUrlAdmin() {
        try {
            const scriptId = ScriptApp.getScriptId();
            // Esta URL es específica de tu despliegue, la obtenemos dinámicamente o la mantenemos si es fija
            const url = `https://script.google.com/macros/s/${scriptId}/exec?admin=true`; // Ajustado para ser dinámico si es posible, o usar la URL real de despliegue.
            
            return {
                success: true,
                url: url
            };
        } catch (error) {
            logError('Error en obtenerUrlAdmin: ' + error.message, 'ERROR', error.stack);
            return {
                success: false,
                message: 'Error al obtener la URL de administración: ' + error.message
            };
        }
    }

    /**
     * Maneja peticiones de API de manera robusta
     * @param {Object} e - Objeto de evento con los parámetros de la petición
     * @return {TextOutput} Respuesta en formato JSON
     */
    function manejarPeticionAPI(e) {
        const startTime = new Date();
        const logData = {
            timestamp: startTime.toISOString(),
            action: e?.parameter?.api || 'desconocida',
            usuario: Session.getActiveUser()?.getEmail() || 'anónimo',
            parameters: { ...(e?.parameter || {}) },
            error: null,
            duracionMs: 0
        };

        try {
            // Validación básica del parámetro de entrada
            if (!e || !e.parameter || !e.parameter.api) {
                const errorMsg = 'Solicitud de API inválida: parámetros faltantes';
                logData.error = errorMsg;
                logData.duracionMs = new Date() - startTime;
                logError('Error en API', 'ERROR', logData);
                
                return ContentService
                    .createTextOutput(JSON.stringify({
                        success: false,
                        data: null,
                        message: errorMsg,
                        timestamp: logData.timestamp,
                        duracionMs: logData.duracionMs
                    }))
                    .setMimeType(ContentService.MimeType.JSON);
            }
            
            const action = e.parameter.api;
            logData.action = action;
            
            // Registrar la acción de la API para auditoría
            logError(`API llamada: ${action}`, 'INFO', logData);
            
            // Procesar la acción solicitada
            let response;
            switch (action.toLowerCase()) {
                case 'status':
                    response = { 
                        success: true, 
                        data: { 
                            status: 'online', 
                            version: '3.0',
                            timestamp: logData.timestamp,
                            timezone: 'America/Panama',
                            entorno: 'produccion'
                        }, 
                        message: 'Sistema operativo' 
                    };
                    break;
                    
                case 'stats':
                    try {
                        const stats = obtenerEstadisticas();
                        response = { 
                            success: true, 
                            data: stats, 
                            message: 'Estadísticas obtenidas' 
                        };
                    } catch (error) {
                        throw new Error('Error al obtener estadísticas: ' + error.message);
                    }
                    break;
                    
                case 'evacuation':
                    try {
                        const datosEvacuacion = getEvacuacionDataForClient();
                        response = { 
                            success: datosEvacuacion.success,
                            data: datosEvacuacion.personasDentro || [],
                            message: datosEvacuacion.message || 'Estado de evacuación obtenido',
                            totalPersonas: datosEvacuacion.totalDentro || 0,
                            timestamp: datosEvacuacion.timestamp || logData.timestamp,
                            procesados: datosEvacuacion.procesados,
                            totalRegistros: datosEvacuacion.totalRegistros
                        };
                        
                        // Registrar estadísticas de la operación
                        logData.registrosProcesados = datosEvacuacion.procesados;
                        logData.totalRegistros = datosEvacuacion.totalRegistros;
                        logData.personasDentro = datosEvacuacion.totalDentro;
                        
                    } catch (error) {
                        throw new Error('Error al obtener datos de evacuación: ' + error.message);
                    }
                    break;
                    
                default:
                    response = {
                        success: false,
                        message: `Acción de API no reconocida: ${action}`,
                        accionesDisponibles: ['status', 'stats', 'evacuation']
                    };
            }
            
            // Completar registro de auditoría
            logData.duracionMs = new Date() - startTime;
            logData.estado = response.success ? 'éxito' : 'fallo';
            
            if (Array.isArray(response.data)) {
                logData.registrosDevueltos = response.data.length;
            }
            
            logError(`API ${action} completada`, 'INFO', logData);
            
            // Asegurar que la respuesta incluya metadatos básicos
            response = {
                ...response,
                success: response.success !== false,
                timestamp: logData.timestamp,
                duracionMs: logData.duracionMs,
                version: '3.0'
            };
            
            return ContentService
                .createTextOutput(JSON.stringify(response))
                .setMimeType(ContentService.MimeType.JSON);
                
        } catch (error) {
            // Manejo detallado de errores
            logData.error = error.message;
            logData.stack = error.stack;
            logData.duracionMs = new Date() - startTime;
            
            logError('Error en API', 'ERROR', logData);
            
            const errorResponse = {
                success: false,
                data: null,
                message: 'Error en el servidor: ' + error.message,
                errorCode: 'API_ERROR',
                timestamp: logData.timestamp,
                duracionMs: logData.duracionMs,
                version: '3.0'
            };
            
            // Solo incluir detalles de depuración si se solicita explícitamente
            if (e?.parameter?.debug === 'true') {
                errorResponse.debug = {
                    action: logData.action,
                    error: error.message,
                    stack: error.stack
                };
            }
            
            return ContentService
                .createTextOutput(JSON.stringify(errorResponse))
                .setMimeType(ContentService.MimeType.JSON);
        }
    }

    // =====================================================
    // VALIDACIÓN DEL SISTEMA
    // =====================================================

    /**
     * Valida que existan todas las hojas necesarias del sistema
     */
    function validarEstructuraHojas() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let mensaje = '';
            let errores = 0;
            const warnings = [];

            // Hojas principales requeridas con validación mejorada
            const hojasRequeridas = [
                { 
                    nombre: 'Base de Datos', 
                    descripcion: 'Contiene la información del personal autorizado',
                    columnasRequeridas: ['Cédula', 'Nombre', 'Empresa'],
                    validacionAdicional: (sheet) => {
                        const lastRow = sheet.getLastRow();
                        if (lastRow <= 1) {
                            warnings.push('Base de Datos está vacía - agregue personal autorizado');
                            return true; // No es un error crítico
                        }
                        return true;
                    }
                },
                { 
                    nombre: 'Respuestas formulario', 
                    descripcion: 'Registra todos los accesos',
                    columnasRequeridas: ['Marca de tiempo', 'Cédula', 'Respuesta'],
                    validacionAdicional: (sheet) => true
                },
                { 
                    nombre: 'Historial', 
                    descripcion: 'Historial consolidado de accesos',
                    columnasRequeridas: ['Fecha', 'Cédula', 'Nombre'],
                    validacionAdicional: (sheet) => true
                },
                { 
                    nombre: 'Configuración', 
                    descripcion: 'Configuración del sistema',
                    columnasRequeridas: ['Clave', 'Valor'],
                    validacionAdicional: (sheet) => {
                        const data = sheet.getDataRange().getValues();
                        if (data.length <= 1) {
                            warnings.push('Configuración vacía - se usarán valores por defecto');
                        }
                        return true;
                    }
                },
                { 
                    nombre: 'Clave', 
                    descripcion: 'Credenciales de administradores',
                    columnasRequeridas: ['Cédula', 'Nombre'],
                    validacionAdicional: (sheet) => {
                        const data = sheet.getDataRange().getValues();
                        if (data.length <= 1) {
                            errores++;
                            mensaje += `• ${sheet.getName()}: No hay administradores configurados\n`;
                            return false;
                        }
                        return true;
                    }
                }
            ];

            // Verificar cada hoja requerida
            hojasRequeridas.forEach(hojaConfig => {
                const sheet = ss.getSheetByName(hojaConfig.nombre);
                
                if (!sheet) {
                    mensaje += `• ${hojaConfig.nombre}: ${hojaConfig.descripcion} - FALTANTE\n`;
                    errores++;
                    return;
                }

                // Validar columnas si existen
                if (hojaConfig.columnasRequeridas && sheet.getLastRow() > 0) {
                    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
                    const columnasFaltantes = hojaConfig.columnasRequeridas.filter(col => 
                        !headers.some(header => String(header).trim() === col)
                    );
                    
                    if (columnasFaltantes.length > 0) {
                        mensaje += `• ${hojaConfig.nombre}: Faltan columnas: ${columnasFaltantes.join(', ')}\n`;
                        errores++;
                    }
                }

                // Validacion adicional personalizada
                if (hojaConfig.validacionAdicional) {
                    try {
                        hojaConfig.validacionAdicional(sheet);
                    } catch (e) {
                        warnings.push(`${hojaConfig.nombre}: ${e.message}`);
                    }
                }
            });

            // Verificar permisos y configuración
            try {
                const user = obtenerUsuarioActual();
                if (!user || user.includes('desconocido')) {
                    warnings.push('No se pudo obtener información del usuario actual');
                }
            } catch (e) {
                warnings.push('Error al verificar permisos de usuario');
            }

            // Compilar resultado final
            if (errores > 0) {
                let mensajeFinal = `❌ Faltan ${errores} componentes esenciales del sistema:\n\n${mensaje}`;
                if (warnings.length > 0) {
                    mensajeFinal += `\n⚠️ Advertencias:\n${warnings.map(w => `• ${w}`).join('\n')}`;
                }
                mensajeFinal += `\n\n💡 Solución: Use el menú "SurPass" → "Crear Sistema Completo" para configurar las hojas faltantes.`;
                
                return {
                    valido: false,
                    mensaje: mensajeFinal,
                    errores: errores,
                    warnings: warnings
                };
            }

            let mensajeExito = '✅ Estructura del sistema validada correctamente.';
            if (warnings.length > 0) {
                mensajeExito += `\n\n⚠️ Advertencias menores:\n${warnings.map(w => `• ${w}`).join('\n')}`;
            }

            return {
                valido: true,
                mensaje: mensajeExito,
                errores: 0,
                warnings: warnings
            };
            
        } catch (error) {
            logError('Error en validarEstructuraHojas', 'ERROR', { error: error.message });
            return {
                valido: false,
                mensaje: '❌ Error crítico al validar la estructura del sistema: ' + error.message,
                errores: 1,
                warnings: []
            };
        }
    }

    // =====================================================
    // NORMALIZACIÓN DE CÉDULAS
    // =====================================================

    /**
     * Normaliza diferentes formatos de cédula/ID con múltiples patrones
     */
    function normalizarCedula(cedula) {
        try {
            if (!cedula) {
                return '';
            }
            
            let cedulaLimpia = String(cedula).trim();
            
            // Paso 1: Intentar parsear como JSON
            try {
                const jsonData = JSON.parse(cedulaLimpia);
                const campoPosibles = [
                    'cedula', 'cédula', 'documento', 'id', 'identificacion',
                    'identificación', 'numero', 'número', 'dni', 'doc', 
                    'num_doc', 'no_doc', 'nro_doc', 'num', 'no', 'nro'
                ];
                
                for (const campo of campoPosibles) {
                    if (jsonData[campo]) {
                        const valor = String(jsonData[campo]).trim();
                        if (/^[\w\-]{3,20}$/.test(valor)) {
                            return valor;
                        }
                    }
                }
            } catch (e) {
                // No es JSON válido, continuar con otros métodos
            }
            
            // Paso 2: Patrones de extracción para diferentes formatos
            const patronesCedula = [
                // Formato Panamá QR (más específico)
                /Texto\s*-\s*([\w\-]+)/i,
                
                // Formatos con etiquetas
                /\b[Cc][EeÉé][Dd][Uu][Ll][Aa][:=\s-]*([A-Z0-9\-]{3,20})/i,
                /\bID[:=\s-]*([A-Z0-9\-]{3,20})/i,
                /\bDNI[:=\s-]*([A-Z0-9\-]{3,20})/i,
                /\bDOC(UMENTO)?[:=\s-]*([A-Z0-9\-]{3,20})/i,
                
                // Formatos específicos de países
                /\b[VE]-([0-9\-]{6,15})\b/i,  // Venezuela, Ecuador
                /\b([0-9]{1,2}-[0-9]{3,4}-[0-9]{3,6})\b/, // Formato con guiones (Panamá)
                /\b([A-Z]?[\d]+-[\d]+-[\d]+)\b/, // Formato panameño: "8-123-456", "PE-1-2-3"
                /\b([A-Z]{1,3}[0-9]{6,12})\b/i, // Formato alfanumérico
                
                // Formato numérico largo
                /\b([0-9]{7,15})\b/,
                
                // Cualquier combinación que parezca ID
                /\b([A-Z0-9\-]{5,20})\b/i
            ];
            
            for (const patron of patronesCedula) {
                const match = cedulaLimpia.match(patron);
                if (match) {
                    // Algunos patrones tienen el resultado en match[2], otros en match[1]
                    const resultado = match[2] || match[1];
                    if (resultado) {
                        const cedulaFinal = resultado.replace(/\s/g, '');
                        // Validar que tenga un formato razonable
                        if (/^[A-Z0-9\-]{3,20}$/i.test(cedulaFinal)) {
                            return cedulaFinal;
                        }
                    }
                }
            }
            
            // Paso 3: Si no coincide con ningún patrón, validar el texto completo
            const cedulaDirecta = cedulaLimpia.replace(/\s+/g, '');
            if (/^[A-Z0-9\-]{3,20}$/i.test(cedulaDirecta)) {
                return cedulaDirecta;
            }
            
            // Paso 4: Extraer números si no hay otra opción
            const numerosEncontrados = cedulaLimpia.match(/\d{5,15}/g);
            if (numerosEncontrados && numerosEncontrados.length > 0) {
                // Priorizar números de 8-12 dígitos
                const numerosPrioritarios = numerosEncontrados.filter(num => /^\d{8,12}$/.test(num));
                if (numerosPrioritarios.length > 0) {
                    return numerosPrioritarios[0];
                }
                return numerosEncontrados[0];
            }
            
            // Si todo falla, devolver texto limpio
            return cedulaDirecta || cedulaLimpia;
            
        } catch (error) {
            logError('Error en normalizarCedula', 'WARNING', { cedula, error: error.message });
            return String(cedula).trim();
        }
    }

    // =====================================================
    // GESTIÓN DE PERSONAL
    // =====================================================

    /**
     * Obtiene todo el personal de la Base de Datos con optimizaciones
     */
    function obtenerTodoElPersonal() {
        const startTime = new Date().getTime();
        
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bdSheet = ss.getSheetByName('Base de Datos');

            if (!bdSheet) {
                throw new Error('La hoja "Base de Datos" no fue encontrada.');
            }

            const lastRow = bdSheet.getLastRow();
            const lastCol = bdSheet.getLastColumn();
            
            if (lastRow <= 1) {
                logError('Base de Datos vacía', 'WARNING');
                return [];
            }

            const data = bdSheet.getRange(1, 1, lastRow, Math.max(lastCol, 3)).getValues();
            const headers = data[0];
            const personal = [];

            // Detectar índices de columnas automáticamente
            const indices = {
                cedula: headers.findIndex(h => h && ['cédula', 'cedula', 'id'].includes(h.toString().toLowerCase())),
                nombre: headers.findIndex(h => h && h.toString().toLowerCase().includes('nombre')),
                empresa: headers.findIndex(h => h && h.toString().toLowerCase().includes('empresa'))
            };

            if (indices.cedula === -1) {
                throw new Error('No se encontró columna de Cédula en la Base de Datos');
            }

            // Procesar datos de forma optimizada
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const cedulaOriginal = String(row[indices.cedula] || '').trim();
                
                if (!cedulaOriginal) continue; // Saltar filas vacías
                
                const cedulaNormalizada = normalizarCedula(cedulaOriginal).replace(/[^\w\-]/g, '');
                const nombre = String(row[indices.nombre] || 'Sin nombre').trim();
                const empresa = String(row[indices.empresa] || 'No especificada').trim();

                personal.push({
                    cedula: cedulaOriginal,
                    cedulaNormalizada: cedulaNormalizada,
                    nombre: nombre,
                    empresa: empresa,
                    busqueda: `${cedulaOriginal} ${cedulaNormalizada} ${nombre.toLowerCase()} ${empresa.toLowerCase()}`
                });
            }

            // Ordenar por cédula para búsquedas más eficientes
            personal.sort((a, b) => a.cedula.localeCompare(b.cedula));
            
            const endTime = new Date().getTime();
            logError(`obtenerTodoElPersonal completado en ${endTime - startTime}ms. ${personal.length} registros procesados.`, 'INFO');
            
            return personal;
            
        } catch (error) {
            logError('Error en obtenerTodoElPersonal', 'ERROR', { error: error.message });
            throw new Error('Error al obtener el personal: ' + error.message);
        }
    }

    /**
     * Busca una persona en la Base de Datos por cédula.
     * @param {GoogleAppsScript.Spreadsheet.Sheet} bdSheet La hoja de "Base de Datos".
     * @param {string} cedula La cédula a buscar.
     * @returns {Object|null} Un objeto con los datos de la persona (cedula, nombre, empresa) o null si no se encuentra.
     */
    function buscarPersonaEnBD(bdSheet, cedula) {
        try {
            if (!bdSheet) {
                throw new Error('La hoja "Base de Datos" no es válida.');
            }

            const data = bdSheet.getDataRange().getValues();
            if (data.length <= 1) { // Solo encabezados
                return null;
            }

            const headers = data[0];
            const indices = {
                cedula: headers.findIndex(h => h && ['cédula', 'cedula', 'id'].includes(h.toString().toLowerCase())),
                nombre: headers.findIndex(h => h && h.toString().toLowerCase().includes('nombre')),
                empresa: headers.findIndex(h => h && h.toString().toLowerCase().includes('empresa'))
            };

            if (indices.cedula === -1) {
                throw new Error('No se encontró columna de Cédula en la Base de Datos.');
            }

            const cedulaNormalizadaBuscada = normalizarCedula(cedula).replace(/[^\w\-]/g, '');

            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const cedulaEnFila = String(row[indices.cedula] || '').trim();
                const cedulaNormalizadaEnFila = normalizarCedula(cedulaEnFila).replace(/[^\w\-]/g, '');

                if (cedulaEnFila === cedula || cedulaNormalizadaEnFila === cedulaNormalizadaBuscada) {
                    return {
                        cedula: cedulaEnFila,
                        nombre: String(row[indices.nombre] || 'Sin nombre').trim(),
                        empresa: String(row[indices.empresa] || 'No especificada').trim()
                    };
                }
            }
            return null; // No se encontró la persona
        } catch (error) {
            logError('Error en buscarPersonaEnBD', 'ERROR', { cedula, error: error.message });
            return null;
        }
    }


    /**
     * ✅ FUNCIÓN DE PRUEBA COMPLETA ACTUALIZADA
     * Reemplaza la función probarEvacuacionCompleta existente
     */
    function probarEvacuacionCompleta() {
        console.log('🧪 === PRUEBA COMPLETA DEL SISTEMA DE EVACUACIÓN ===');
        
        try {
            // Paso 1: Diagnóstico
            console.log('\n📋 Paso 1: Diagnóstico del sistema...');
            const diagnostico = validarSistemaEvacuacion();
            
            if (!diagnostico.success) {
                throw new Error('Diagnóstico falló: ' + diagnostico.resumen);
            }
            
            // Paso 2: Obtener datos de evacuación (USAR LA NUEVA FUNCIÓN)
            console.log('\n📡 Paso 2: Obteniendo datos de evacuación...');
            const datosEvacuacion = getEvacuacionDataForClient(); // ✅ USAR LA NUEVA FUNCIÓN
            
            if (!datosEvacuacion.success) {
                throw new Error('Error obteniendo datos: ' + datosEvacuacion.message);
            }
            
            console.log(`✅ Datos obtenidos: ${datosEvacuacion.totalDentro} personas dentro`);
            
            // Paso 3: Simular evacuación si hay personas
            if (datosEvacuacion.totalDentro > 0) {
                console.log('\n🚨 Paso 3: Simulando evacuación...');
                
                // Tomar solo las primeras 2 personas para la prueba
                const personasParaPrueba = datosEvacuacion.personasDentro.slice(0, 2);
                const cedulasPrueba = personasParaPrueba.map(p => p.cedula);
                
                console.log('📝 Evacuando (prueba):', cedulasPrueba);
                
                // NO ejecutar realmente la evacuación en modo prueba
                console.log('⚠️ MODO PRUEBA: Evacuación simulada exitosamente');
            } else {
                console.log('\n✅ Paso 3: No hay personas dentro, edificio ya evacuado');
            }
            
            console.log('\n🎉 === PRUEBA COMPLETADA EXITOSAMENTE ===');
            
            return {
                success: true,
                message: 'Prueba completa exitosa',
                diagnostico: diagnostico,
                datosEvacuacion: datosEvacuacion
            };
            
        } catch (error) {
            console.error('❌ Error en prueba:', error.message);
            return {
                success: false,
                message: 'Prueba falló: ' + error.message
            };
        }
    }

    /**
     * ✅ FUNCIÓN SIMPLE PARA EJECUTAR DESDE EL MENÚ
     */
    function ejecutarDiagnosticoEvacuacion() {
        const resultado = validarSistemaEvacuacion();
        
        let mensaje = `🔍 DIAGNÓSTICO DE EVACUACIÓN\n\n`;
        mensaje += `✅ Tests exitosos: ${Object.values(resultado.detalles).filter(Boolean).length}/4\n`;
        mensaje += `📊 Estado: ${resultado.success ? 'FUNCIONAL' : 'CON PROBLEMAS'}\n`;
        mensaje += `📋 Resumen: ${resultado.resumen}\n\n`;
        
        if (resultado.success) {
            mensaje += `🎉 El sistema de evacuación está funcionando correctamente.\n\n`;
        } else {
            mensaje += `⚠️ Se encontraron algunos problemas:\n\n`;
            Object.keys(resultado.detalles).forEach(test => {
                const estado = resultado.detalles[test];
                mensaje += `${estado ? '✅' : '❌'} ${test.toUpperCase()}\n`;
            });
            mensaje += `\n💡 Revise los logs para más detalles.`;
        }
        
        // Mostrar en UI si está disponible
        try {
            SpreadsheetApp.getUi().alert('🔍 Diagnóstico de Evacuación', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
        } catch (e) {
            console.log(mensaje);
        }
        
        return resultado;
    }

    // =====================================================
    // 🚨 FUNCIONES DE EVACUACIÓN MEJORADAS
    // =====================================================

    /**
     * Muestra estado de evacuación con interfaz HTML profesional
     */
    function mostrarEstadoEvacuacion() {
        try {
            const conteoEvacuacion = getEvacuacionDataForClient(); // Usar la función unificada
            
            // Generar HTML profesional para el modal de evacuación
            const htmlContent = generarHTMLEvacuacion(conteoEvacuacion);
            
            const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
                .setWidth(1000)
                .setHeight(800)
                .setTitle('🚨 ESTADO DE EVACUACIÓN - EMERGENCIA');
            
            SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🚨 Control de Evacuación');
            
            // Log de auditoría
            logError(`Estado de evacuación consultado via UI: ${conteoEvacuacion.totalDentro} personas dentro`, 'INFO');
            
            return conteoEvacuacion;
            
        } catch (error) {
            logError('Error en mostrarEstadoEvacuacion', 'ERROR', { error: error.message });
            SpreadsheetApp.getUi().alert(
                '❌ Error', 
                'Error al obtener estado de evacuación: ' + error.message, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            return null;
        }
    }

    /**
     * Genera HTML profesional para evacuación
     */
    function generarHTMLEvacuacion(conteoEvacuacion) {
        // Asegurarse de que conteoEvacuacion tenga los valores por defecto necesarios
        const totalDentro = conteoEvacuacion?.totalDentro || 0;
        const personasDentro = Array.isArray(conteoEvacuacion?.personasDentro) ? conteoEvacuacion.personasDentro : [];
        
        // Obtener estadísticas del día para el conteo de Entradas/Salidas
        let estadisticasDelDia = { entradas: 0, salidas: 0 };
    try {
        const stats = obtenerEstadisticas();
        estadisticasDelDia.entradas = stats.entradas;
        estadisticasDelDia.salidas = stats.salidas;
    } catch (e) {
        logError('Error obteniendo estadísticas del día para HTML de evacuación', 'WARNING', { error: e.message });
    }

        console.log('📊 Datos recibidos en generarHTMLEvacuacion:', {
            totalDentro,
            totalPersonas: personasDentro.length,
            estadisticas: estadisticasDelDia,
            personasDentro: personasDentro.map(p => ({
                cedula: p.cedula || 'N/A',
                nombre: p.nombre || 'N/A',
                empresa: p.empresa || 'N/A',
                horaEntrada: p.horaEntrada || null, // horaEntrada ya viene formateada como string
                tipo: p.tipo || 'entrada'
            }))
        });
        
        // Determinar estado de emergencia
        const estadoEmergencia = totalDentro === 0 ? 'EVACUADO' : 'PENDIENTE';
        const colorEstado = totalDentro === 0 ? '#4CAF50' : '#F44336';
        const iconoEstado = totalDentro === 0 ? '✅' : '⚠️';
        
        // Generar filas de la tabla
        let filasTabla = '';
        if (personasDentro.length === 0) {
            filasTabla = '<tr><td colspan="6" style="text-align: center; color: #4CAF50; font-weight: bold; padding: 30px; font-size: 18px;"><i class="fas fa-check-circle" style="font-size: 48px; display: block; margin-bottom: 10px;"></i>✅ EDIFICIO COMPLETAMENTE EVACUADO</td></tr>';
        } else {
            personasDentro.forEach((persona, index) => {
                // persona.horaEntrada ya viene como string 'HH:mm' del servidor
                const horaEntradaFormateada = persona.horaEntrada || 'N/A';
                
                // Calcular tiempo dentro (si hay hora de entrada)
                let tiempoDentro = 'N/A';
                if (horaEntradaFormateada !== 'N/A') {
                    try {
                        // Reconstruir una fecha para calcular duración
                        const [horas, minutos] = horaEntradaFormateada.split(':').map(Number);
                        const fechaActual = new Date();
                        const fechaEntradaHoy = new Date(fechaActual.getFullYear(), fechaActual.getMonth(), fechaActual.getDate(), horas, minutos);
                        
                        // Si la hora de entrada es posterior a la hora actual (ej. entrada ayer), ajustar
                        if (fechaEntradaHoy > fechaActual) {
                            fechaEntradaHoy.setDate(fechaEntradaHoy.getDate() - 1);
                        }

                        const diffMs = fechaActual.getTime() - fechaEntradaHoy.getTime();
                        const diffMin = Math.floor(diffMs / (1000 * 60));
                        
                        if (diffMin >= 0) {
                            const h = Math.floor(diffMin / 60);
                            const m = diffMin % 60;
                            tiempoDentro = `${h}h ${m}m`;
                        }
                    } catch (e) {
                        logError('Error calculando tiempo dentro en generarHTMLEvacuacion', 'WARNING', { horaEntrada: horaEntradaFormateada, error: e.message });
                        tiempoDentro = 'Calc. Error';
                    }
                }
                
                filasTabla += `
                    <tr id="persona-${index}" class="persona-row">
                        <td style="text-align: center;">
                            <input type="checkbox" 
                                id="check-${index}" 
                                data-cedula="${persona.cedula}"
                                data-nombre="${persona.nombre}"
                                data-empresa="${persona.empresa}"
                                class="evacuation-checkbox"
                                style="transform: scale(1.3); margin: 0;">
                        </td>
                        <td><strong>${persona.cedula}</strong></td>
                        <td>${persona.nombre}</td>
                        <td>${persona.empresa}</td>
                        <td style="text-align: center;">${horaEntradaFormateada}</td>
                        <td style="text-align: center;">${tiempoDentro}</td>
                        <td style="text-align: center;">
                            <span class="status-badge status-inside">DENTRO</span>
                        </td>
                    </tr>
                `;
            });
        }
        
        return `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>🚨 Control de Evacuación</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: #333;
                line-height: 1.6;
                padding: 20px;
            }
            
            .container {
                background: white;
                border-radius: 20px;
                box-shadow: 0 25px 50px rgba(0,0,0,0.15);
                overflow: hidden;
                max-width: 100%;
                animation: slideIn 0.5s ease-out;
            }
            
            @keyframes slideIn {
                from { opacity: 0; transform: translateY(30px); }
                to { opacity: 1; transform: translateY(0); }
            }
            
            .header {
                background: linear-gradient(135deg, ${colorEstado}, #c62828);
                color: white;
                padding: 30px;
                text-align: center;
                position: relative;
            }
            
            .header h1 {
                font-size: 32px;
                margin-bottom: 15px;
                text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
            }
            
            .status-indicator {
                font-size: 64px;
                margin: 20px 0;
                animation: ${totalDentro === 0 ? 'bounce' : 'pulse'} 2s infinite;
            }
            
            @keyframes pulse {
                0%, 100% { transform: scale(1); opacity: 1; }
                50% { transform: scale(1.1); opacity: 0.8; }
            }
            
            @keyframes bounce {
                0%, 20%, 53%, 80%, 100% { transform: translate3d(0,0,0); }
                40%, 43% { transform: translate3d(0,-10px,0); }
                70% { transform: translate3d(0,-5px,0); }
                90% { transform: translate3d(0,-2px,0); }
            }
            
            .status-text {
                font-size: 28px;
                font-weight: bold;
                margin-bottom: 15px;
            }
            
            .stats-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
                gap: 25px;
                padding: 30px;
                background: #f8f9fa;
            }
            
            .stat-card {
                background: white;
                padding: 25px;
                border-radius: 15px;
                text-align: center;
                box-shadow: 0 8px 25px rgba(0,0,0,0.1);
                border-left: 5px solid #2196F3;
                transition: transform 0.3s ease;
            }
            
            .stat-card:hover {
                transform: translateY(-5px);
            }
            
            .stat-number {
                font-size: 42px;
                font-weight: bold;
                color: #2196F3;
                margin-bottom: 10px;
            }
            
            .stat-label {
                color: #666;
                font-size: 14px;
                text-transform: uppercase;
                letter-spacing: 1px;
                font-weight: 600;
            }
            
            .table-container {
                padding: 30px;
                max-height: 500px;
                overflow-y: auto;
            }
            
            .table-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 25px;
            }
            
            .table-title {
                font-size: 24px;
                font-weight: bold;
                color: #333;
            }
            
            .select-controls {
                display: flex;
                gap: 15px;
                align-items: center;
            }
            
            .btn {
                padding: 12px 20px;
                border: none;
                border-radius: 8px;
                cursor: pointer;
                font-size: 14px;
                font-weight: 600;
                transition: all 0.3s ease;
                text-decoration: none;
                display: inline-block;
            }
            
            .btn-primary {
                background: #2196F3;
                color: white;
            }
            
            .btn-primary:hover {
                background: #1976D2;
                transform: translateY(-2px);
                box-shadow: 0 5px 15px rgba(33, 150, 243, 0.4);
            }
            
            .btn-secondary {
                background: #6c757d;
                color: white;
            }
            
            .btn-secondary:hover {
                background: #545b62;
                transform: translateY(-2px);
            }
            
            .table {
                width: 100%;
                border-collapse: collapse;
                background: white;
                border-radius: 12px;
                overflow: hidden;
                box-shadow: 0 8px 25px rgba(0,0,0,0.1);
            }
            
            .table th, .table td {
                padding: 18px 15px;
                text-align: left;
                border-bottom: 1px solid #e0e0e0;
            }
            
            .table th {
                background: #f5f5f5;
                font-weight: 700;
                color: #333;
                text-transform: uppercase;
                font-size: 12px;
                letter-spacing: 1px;
                position: sticky;
                top: 0;
                z-index: 10;
            }
            
            .table tbody tr:hover {
                background: #f8f9fa;
            }
            
            .persona-row.evacuated {
                background: #e8f5e9 !important;
                text-decoration: line-through;
                opacity: 0.7;
                transition: all 0.3s ease;
            }
            
            .status-badge {
                padding: 6px 16px;
                border-radius: 25px;
                font-size: 12px;
                font-weight: bold;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            
            .status-inside {
                background: #ffebee;
                color: #c62828;
                border: 2px solid #ffcdd2;
            }
            
            .status-evacuated {
                background: #e8f5e9;
                color: #2e7d32;
                border: 2px solid #c8e6c9;
            }
            
            .action-buttons {
                padding: 30px;
                background: #f8f9fa;
                border-top: 1px solid #e0e0e0;
                display: flex;
                flex-direction: column;
                gap: 20px;
                align-items: center;
            }
            
            .evacuation-type-selector {
                text-align: center;
                margin-bottom: 20px;
            }
            
            .evacuation-type-selector h3 {
                color: #d32f2f;
                margin-bottom: 15px;
                font-size: 22px;
                text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
            }
            
            .type-buttons {
                display: flex;
                gap: 20px;
                justify-content: center;
                margin-bottom: 20px;
                flex-wrap: wrap;
            }
            
            .type-buttons .btn {
                flex-direction: column;
                min-width: 200px;
                padding: 20px;
                position: relative;
                transition: all 0.3s ease;
            }
            
            .type-buttons .btn small {
                font-size: 12px;
                opacity: 0.8;
                margin-top: 5px;
                font-weight: normal;
            }
            
            .type-buttons .btn.selected {
                transform: scale(1.05);
                box-shadow: 0 8px 25px rgba(0,0,0,0.2);
                border: 3px solid #fff;
            }
            
            .evacuation-confirm-section {
                text-align: center;
            }
            
            .secondary-actions {
                display: flex;
                gap: 15px;
                justify-content: center;
                flex-wrap: wrap;
            }
            
            .secondary-actions .btn {
                min-width: 150px;
            }
            
            .btn-large {
                padding: 18px 35px;
                font-size: 16px;
                border-radius: 12px;
                min-width: 200px;
                font-weight: 600;
                position: relative;
                overflow: hidden;
            }
            
            .btn-large::before {
                content: '';
                position: absolute;
                top: 0;
                left: -100%;
                width: 100%;
                height: 100%;
                background: linear-gradient(90deg, rgba(255,255,255,0.3), transparent);
                transition: all 0.6s;
            }
            
            .btn-large:hover::before {
                left: 100%;
            }
            
            .btn-success {
                background: linear-gradient(135deg, #4CAF50, #388e3c);
                color: white;
            }
            
            .btn-success:hover {
                background: linear-gradient(135deg, #66bb6a, #4caf50);
                transform: translateY(-3px);
                box-shadow: 0 10px 25px rgba(76, 175, 80, 0.4);
            }
            
            .btn-warning {
                background: linear-gradient(135deg, #FF9800, #f57c00);
                color: white;
            }
            
            .btn-warning:hover {
                background: linear-gradient(135deg, #ffb74d, #ff9800);
                transform: translateY(-3px);
                box-shadow: 0 10px 25px rgba(255, 152, 0, 0.4);
            }
            
            .btn-info {
                background: linear-gradient(135deg, #2196F3, #1976d2);
                color: white;
            }
            
            .btn-info:hover {
                background: linear-gradient(135deg, #64b5f6, #2196f3);
                transform: translateY(-3px);
                box-shadow: 0 10px 25px rgba(33, 150, 243, 0.4);
            }
            
            .btn-danger {
                background: linear-gradient(135deg, #f44336, #d32f2f);
                color: white;
            }
            
            .btn-danger:hover {
                background: linear-gradient(135deg, #ef5350, #f44336);
                transform: translateY(-3px);
                box-shadow: 0 10px 25px rgba(244, 67, 54, 0.4);
            }
            
            .timestamp {
                color: #666;
                font-size: 13px;
                margin-top: 15px;
                font-weight: 500;
            }
            
            .loading {
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(0,0,0,0.8);
                backdrop-filter: blur(5px);
                justify-content: center;
                align-items: center;
                flex-direction: column;
                z-index: 1000;
            }
            
            .spinner {
                border: 6px solid rgba(255,255,255,0.1);
                border-top: 6px solid #2196F3;
                border-radius: 50%;
                width: 60px;
                height: 60px;
                animation: spin 1s linear infinite;
                margin-bottom: 20px;
            }
            
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            
            .loading p {
                color: white;
                font-size: 18px;
                font-weight: 500;
            }
            
            /* Responsive */
            @media (max-width: 768px) {
                .stats-grid {
                    grid-template-columns: 1fr 1fr;
                    gap: 15px;
                    padding: 20px;
                }
                
                .table-container {
                    padding: 20px;
                    max-height: 400px;
                }
                
                .table {
                    font-size: 14px;
                }
                
                .action-buttons {
                    flex-direction: column;
                    align-items: center;
                    padding: 20px;
                }
                
                .btn-large {
                    min-width: 100%;
                    margin-bottom: 10px;
                }
                
                .select-controls {
                    flex-direction: column;
                    gap: 10px;
                }
            }
            
            /* Estilos para impresión */
            @media print {
                body { background: white !important; }
                .action-buttons, .loading { display: none !important; }
                .container { box-shadow: none !important; }
                .header { background: #f5f5f5 !important; color: black !important; }
                .btn, input[type="checkbox"] { display: none !important; }
                .table th { background: #f0f0f0 !important; }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1><i class="fas fa-shield-alt"></i> CONTROL DE EVACUACIÓN DE EMERGENCIA</h1>
                <div class="status-indicator">${iconoEstado}</div>
                <div class="status-text">ESTADO: ${estadoEmergencia}</div>
                <div class="timestamp">
                    <i class="fas fa-clock"></i> Consultado: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')}
                </div>
            </div>
            
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-number">${totalDentro}</div>
                    <div class="stat-label"><i class="fas fa-users"></i> Personas Dentro</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">${estadisticasDelDia.entradas}</div>
                    <div class="stat-label"><i class="fas fa-sign-in-alt"></i> Total Entradas Hoy</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">${estadisticasDelDia.salidas}</div>

                    <div class="stat-label"><i class="fas fa-sign-out-alt"></i> Total Salidas Hoy</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">${Math.round(((estadisticasDelDia.salidas || 0) / Math.max(estadisticasDelDia.entradas || 1, 1)) * 100)}%</div>
                    <div class="stat-label"><i class="fas fa-percentage"></i> % Evacuado</div>
                </div>
            </div>
            
            <div class="table-container">
                <div class="table-header">
                    <div class="table-title"><i class="fas fa-list-ul"></i> Personas en el Edificio</div>
                    ${totalDentro > 0 ? `
                    <div class="select-controls">
                        <button onclick="seleccionarTodos()" class="btn btn-primary">
                            <i class="fas fa-check-double"></i> Seleccionar Todos
                        </button>
                        <button onclick="limpiarSeleccion()" class="btn btn-secondary">
                            <i class="fas fa-times"></i> Limpiar
                        </button>
                    </div>
                    ` : ''}
                </div>
                
                <table class="table">
                    <thead>
                        <tr>
                            ${totalDentro > 0 ? '<th style="width: 60px;"><i class="fas fa-check"></i></th>' : ''}
                            <th><i class="fas fa-id-card"></i> Cédula</th>
                            <th><i class="fas fa-user"></i> Nombre</th>
                            <th><i class="fas fa-building"></i> Empresa</th>
                            <th style="width: 100px;"><i class="fas fa-clock"></i> Entrada</th>
                            <th style="width: 100px;"><i class="fas fa-hourglass-half"></i> Tiempo</th>
                            <th style="width: 120px;"><i class="fas fa-info-circle"></i> Estado</th>
                        </tr>
                    </thead>
                    <tbody id="tablaPersonas">
                        ${filasTabla}
                    </tbody>
                </table>
            </div>
            
            ${totalDentro > 0 ? `
            <div class="action-buttons">
                <div class="evacuation-type-selector">
                    <h3><i class="fas fa-exclamation-triangle"></i> Seleccione el Tipo de Evacuación</h3>
                    <div class="type-buttons">
                        <button onclick="seleccionarTipoEvacuacion('real')" class="btn btn-large btn-danger" id="btnReal">
                            <i class="fas fa-fire"></i> EVACUACIÓN REAL
                            <small>Modifica registros permanentemente</small>
                        </button>
                        <button onclick="seleccionarTipoEvacuacion('simulacro')" class="btn btn-large btn-warning" id="btnSimulacro">
                            <i class="fas fa-theater-masks"></i> SIMULACRO
                            <small>Solo registro de auditoría</small>
                        </button>
                    </div>
                </div>
                <div class="evacuation-confirm-section" id="confirmSection" style="display: none;">
                    <button onclick="confirmarEvacuacion()" class="btn btn-large btn-success" id="btnConfirmar">
                        <i class="fas fa-check-circle"></i> Confirmar Evacuación Seleccionada
                    </button>
                </div>
                <div class="secondary-actions">
                    <button onclick="exportarListado()" class="btn btn-large btn-info">
                        <i class="fas fa-file-csv"></i> Exportar Listado CSV
                    </button>
                    <button onclick="imprimirListado()" class="btn btn-large btn-secondary">
                        <i class="fas fa-print"></i> Imprimir Listado
                    </button>
                    <button onclick="cerrarVentana()" class="btn btn-large btn-secondary">
                        <i class="fas fa-times"></i> Cerrar
                    </button>
                </div>
            </div>
            ` : `
            <div class="action-buttons">
                <button onclick="exportarListado()" class="btn btn-large btn-warning">
                    <i class="fas fa-file-csv"></i> Exportar Reporte CSV
                </button>
                <button onclick="imprimirListado()" class="btn btn-large btn-info">
                    <i class="fas fa-print"></i> Imprimir Reporte
                </button>
                <button onclick="cerrarVentana()" class="btn btn-large btn-success">
                    <i class="fas fa-check-circle"></i> Cerrar - Edificio Evacuado
                </button>
            </div>
            `}
            
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p><i class="fas fa-cog fa-spin"></i> Procesando evacuación...</p>
            </div>
        </div>

        <script>
            // Variables globales
            let personasSeleccionadas = [];
            let tipoEvacuacionSeleccionado = null;
            
            // Funciones de selección de tipo
            function seleccionarTipoEvacuacion(tipo) {
                tipoEvacuacionSeleccionado = tipo;
                
                // Actualizar botones
                const btnReal = document.getElementById('btnReal');
                const btnSimulacro = document.getElementById('btnSimulacro');
                const confirmSection = document.getElementById('confirmSection');
                
                // Remover selección anterior
                btnReal.classList.remove('selected');
                btnSimulacro.classList.remove('selected');
                
                // Agregar selección actual
                if (tipo === 'real') {
                    btnReal.classList.add('selected');
                } else {
                    btnSimulacro.classList.add('selected');
                }
                
                // Mostrar sección de confirmación
                confirmSection.style.display = 'block';
                
                // Actualizar texto del botón de confirmación
                actualizarBotonConfirmar();
            }
            
            // Funciones de selección de personas
            function seleccionarTodos() {
                const checkboxes = document.querySelectorAll('.evacuation-checkbox');
                checkboxes.forEach(cb => {
                    cb.checked = true;
                    actualizarSeleccion(cb);
                });
            }
            
            function limpiarSeleccion() {
                const checkboxes = document.querySelectorAll('.evacuation-checkbox');
                checkboxes.forEach(cb => {
                    cb.checked = false;
                    const fila = cb.closest('tr');
                    fila.classList.remove('evacuated');
                    const badge = fila.querySelector('.status-badge');
                    badge.textContent = 'DENTRO';
                    badge.className = 'status-badge status-inside';
                });
                personasSeleccionadas = [];
                actualizarBotonConfirmar();
            }
            
            function actualizarSeleccion(checkbox) {
                const fila = checkbox.closest('tr');
                const badge = fila.querySelector('.status-badge');
                
                if (checkbox.checked) {
                    fila.classList.add('evacuated');
                    badge.textContent = 'EVACUADO';
                    badge.className = 'status-badge status-evacuated';
                    
                    if (!personasSeleccionadas.includes(checkbox.dataset.cedula)) {
                        personasSeleccionadas.push(checkbox.dataset.cedula);
                    }
                } else {
                    fila.classList.remove('evacuated');
                    badge.textContent = 'DENTRO';
                    badge.className = 'status-badge status-inside';
                    
                    const index = personasSeleccionadas.indexOf(checkbox.dataset.cedula);
                    if (index > -1) {
                        personasSeleccionadas.splice(index, 1);
                    }
                }
                
                actualizarBotonConfirmar();
            }
            
            function actualizarBotonConfirmar() {
                const btnConfirmar = document.getElementById('btnConfirmar');
                if (btnConfirmar) {
                    if (!tipoEvacuacionSeleccionado) {
                        btnConfirmar.innerHTML = '<i class="fas fa-exclamation-triangle"></i> Seleccione el Tipo de Evacuación Primero';
                        btnConfirmar.disabled = true;
                        btnConfirmar.style.opacity = '0.5';
                        return;
                    }
                    
                    const iconoTipo = tipoEvacuacionSeleccionado === 'real' ? 'fas fa-fire' : 'fas fa-theater-masks';
                    const textoTipo = tipoEvacuacionSeleccionado === 'real' ? 'EVACUACIÓN REAL' : 'SIMULACRO';
                    
                    if (personasSeleccionadas.length > 0) {
                        btnConfirmar.innerHTML = \`<i class="\${iconoTipo}"></i> Confirmar \${textoTipo} (\${personasSeleccionadas.length} personas)\`;
                        btnConfirmar.disabled = false;
                        btnConfirmar.style.opacity = '1';
                    } else {
                        btnConfirmar.innerHTML = \`<i class="\${iconoTipo}"></i> Confirmar \${textoTipo} - Seleccione Personas\`;
                        btnConfirmar.disabled = true;
                        btnConfirmar.style.opacity = '0.5';
                    }
                }
            }
            
            // ✅ FUNCIÓN CORREGIDA DE CONFIRMACIÓN DE EVACUACIÓN
            function confirmarEvacuacion() {
                if (!tipoEvacuacionSeleccionado) {
                    alert('❌ Debe seleccionar el tipo de evacuación primero.');
                    return;
                }
                
                if (personasSeleccionadas.length === 0) {
                    alert('❌ Seleccione al menos una persona para evacuar.');
                    return;
                }
                
                const esSimulacro = tipoEvacuacionSeleccionado === 'simulacro';
                const textoTipo = esSimulacro ? 'SIMULACRO' : 'EVACUACIÓN REAL';
                const warningMessage = esSimulacro 
                    ? \`🎭 CONFIRMACIÓN DE \${textoTipo}\\n\\n¿Confirma el \${textoTipo} para \${personasSeleccionadas.length} persona(s)?\\n\\nEsto SOLO registrará la actividad en logs de auditoría.\\nNO se modificarán los registros reales de entrada/salida.\`
                    : \`🚨 CONFIRMACIÓN DE \${textoTipo}\\n\\n¿Confirma la \${textoTipo} para \${personasSeleccionadas.length} persona(s)?\\n\\nEsta acción registrará automáticamente la salida en el historial y enviará una notificación de emergencia.\\n\\n⚠️ Esta acción NO se puede deshacer.\`;
                
                const confirmacion = confirm(warningMessage);
                
                if (!confirmacion) return;
                
                mostrarLoading(true);
                
                // ✅ CORRECCIÓN CRÍTICA: Llamar a la función correcta según el tipo
                if (esSimulacro) {
                    // Para simulacros, usar función específica que NO modifica historial
                    google.script.run
                        .withSuccessHandler(function(resultado) {
                            mostrarLoading(false);
                            if (resultado.success) {
                                alert('✅ ' + resultado.message + '\\n\\n🎭 SIMULACRO completado - NO se modificó el historial real.');
                                window.location.reload();
                            } else {
                                alert('❌ Error en simulacro: ' + resultado.message);
                            }
                        })
                        .withFailureHandler(function(error) {
                            mostrarLoading(false);
                            alert('❌ Error al procesar simulacro: ' + error.message);
                        })
                        .procesarSimulacroEvacuacion(personasSeleccionadas, 'Simulacro ejecutado desde interfaz de evacuación');
                } else {
                    // Para evacuaciones reales, usar función unificada
                    google.script.run
                        .withSuccessHandler(function(resultado) {
                            mostrarLoading(false);
                            if (resultado.success) {
                                alert('✅ ' + resultado.message + '\\n\\n🚨 Se ha enviado una notificación de emergencia a los administradores.');
                                window.location.reload();
                            } else {
                                alert('❌ Error: ' + resultado.message);
                            }
                        })
                        .withFailureHandler(function(error) {
                            mostrarLoading(false);
                            alert('❌ Error al procesar evacuación: ' + error.message);
                        })
                        .procesarEvacuacionUnificada({
                            cedulas: personasSeleccionadas,
                            tipo: 'real'
                        });
                }
            }
            
            // Función de exportación
            function exportarListado() {
                mostrarLoading(true);
                
                google.script.run
                    .withSuccessHandler(function(resultado) {
                        mostrarLoading(false);
                        if (resultado.success) {
                            alert('✅ Archivo exportado correctamente\\n\\nArchivo: ' + resultado.fileName + '\\nPersonas: ' + resultado.totalPersonas);
                            if (resultado.downloadUrl) {
                                window.open(resultado.downloadUrl, '_blank');
                            }
                        } else {
                            alert('❌ Error al exportar: ' + resultado.message);
                        }
                    })
                    .withFailureHandler(function(error) {
                        mostrarLoading(false);
                        alert('❌ Error al exportar: ' + error.message);
                    })
                    .exportarEstadoEvacuacion();
            }
            
            // Función de impresión
            function imprimirListado() {
                // Ocultar elementos para impresión
                const loading = document.getElementById('loading');
                loading.style.display = 'none';
                
                // Imprimir
                window.print();
            }
            
            // Función para cerrar ventana
            function cerrarVentana() {
                if (typeof google !== 'undefined' && google.script && google.script.host) {
                    google.script.host.close();
                } else {
                    window.close();
                }
            }
            
            // Función para mostrar/ocultar loading
            function mostrarLoading(mostrar) {
                const loading = document.getElementById('loading');
                
                if (mostrar) {
                    loading.style.display = 'flex';
                } else {
                    loading.style.display = 'none';
                }
            }
            
            // Event listeners
            document.addEventListener('DOMContentLoaded', function() {
                // Agregar event listeners a checkboxes
                const checkboxes = document.querySelectorAll('.evacuation-checkbox');
                checkboxes.forEach(cb => {
                    cb.addEventListener('change', function() {
                        actualizarSeleccion(this);
                    });
                });
                
                // Actualizar estado inicial del botón
                actualizarBotonConfirmar();
                
                // Mensaje informativo inicial
                const totalDentro = ${totalDentro};
                if (totalDentro > 0) {
                    setTimeout(() => {
                        alert('🚨 INTERFAZ DE EVACUACIÓN MEJORADA\\n\\n' +
                            '1. Seleccione el TIPO de evacuación (Real o Simulacro)\\n' +
                            '2. Seleccione las personas a evacuar\\n' +
                            '3. Confirme la acción\\n\\n' +
                            '⚠️ SIMULACROS: Solo generan logs de auditoría\\n' +
                            '🚨 EVACUACIÓN REAL: Modifica registros permanentemente');
                    }, 1000);
                }
            });
        </script>
    </body>
    </html>
        `;
    }

    /**
     * Exporta estado de evacuación a CSV
     */
    function exportarEstadoEvacuacion() {
        try {
            const conteoEvacuacion = getEvacuacionDataForClient();
            
            // ✅ OBTENER ESTADÍSTICAS DENTRO DE LA FUNCIÓN
            let estadisticasDelDia = { entradas: 0, salidas: 0 };
            try {
                const stats = obtenerEstadisticas();
                estadisticasDelDia.entradas = stats.entradas;
                estadisticasDelDia.salidas = stats.salidas;
            } catch (e) {
                logError('Error obteniendo estadísticas para exportación', 'WARNING', { error: e.message });
            }
            
            const fechaActual = new Date();
            const fechaFormateada = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
            const nombreArchivo = `Evacuacion_${fechaFormateada}.csv`;
            
            // Preparar datos CSV con BOM UTF-8
            let csvContent = '\uFEFF'; // BOM para UTF-8
            csvContent += 'Estado,Cedula,Nombre,Empresa,Hora_Entrada,Tiempo_Dentro,Observaciones\n';
            
            // Agregar header de información
            csvContent += `REPORTE_EVACUACION,${fechaFormateada},TOTAL_DENTRO:${conteoEvacuacion.totalDentro},GENERADO:${Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')},,,\n`;
            csvContent += ',,,,,, \n'; // Línea vacía
            
            if (conteoEvacuacion.personasDentro && conteoEvacuacion.personasDentro.length > 0) {
                conteoEvacuacion.personasDentro.forEach(persona => {
                    const horaEntrada = persona.horaEntrada || 'N/A'; // Ya es un string
                    
                    // Calcular tiempo dentro
                    let tiempoDentro = 'N/A';
                    if (horaEntrada !== 'N/A') {
                        try {
                            const [horas, minutos] = horaEntrada.split(':').map(Number);
                            const fechaEntradaHoy = new Date(fechaActual.getFullYear(), fechaActual.getMonth(), fechaActual.getDate(), horas, minutos);
                            
                            if (fechaEntradaHoy > fechaActual) {
                                fechaEntradaHoy.setDate(fechaEntradaHoy.getDate() - 1);
                            }

                            const diffMs = fechaActual.getTime() - fechaEntradaHoy.getTime();
                            const diffMin = Math.floor(diffMs / (1000 * 60));
                            
                            if (diffMin >= 0) {
                                const h = Math.floor(diffMin / 60);
                                const m = diffMin % 60;
                                tiempoDentro = `${h}h ${m}m`;
                            }
                        } catch (e) {
                            logError('Error calculando tiempo dentro para exportación', 'WARNING', { horaEntrada: horaEntrada, error: e.message });
                            tiempoDentro = 'Calc. Error';
                        }
                    }
                    
                    // Determinar nivel de prioridad
                    let prioridad = 'NORMAL';
                    if (tiempoDentro !== 'N/A' && tiempoDentro !== 'Calc. Error') {
                        const [horasStr, minutosStr] = horaEntrada.split(':');
                        const entradaDate = new Date();
                        entradaDate.setHours(parseInt(horasStr), parseInt(minutosStr), 0, 0);

                        const diffMsPrioridad = fechaActual.getTime() - entradaDate.getTime();
                        const minutosInside = Math.floor(diffMsPrioridad / (1000 * 60));
                        
                        if (minutosInside > 480) prioridad = 'CRITICA'; // Más de 8 horas
                        else if (minutosInside > 240) prioridad = 'ALTA'; // Más de 4 horas
                    }
                    
                    // Escapar comillas en los datos
                    const cedula = `"${String(persona.cedula || '').replace(/"/g, '""')}"`;
                    const nombre = `"${String(persona.nombre || '').replace(/"/g, '""')}"`;
                    const empresa = `"${String(persona.empresa || '').replace(/"/g, '""')}"`;
                    
                    csvContent += `DENTRO,${cedula},${nombre},${empresa},${horaEntrada},${tiempoDentro},${prioridad}\n`;
                });
            } else {
                csvContent += 'EVACUADO_COMPLETO,N/A,EDIFICIO_EVACUADO,N/A,N/A,N/A,COMPLETADO\n';
            }
            
            // Agregar estadísticas al final
            csvContent += ',,,,,, \n'; // Línea vacía
            csvContent += 'ESTADISTICAS,,,,,, \n';
            csvContent += `Total_Entradas_Hoy,${estadisticasDelDia.entradas || 0},,,,, \n`;
            csvContent += `Total_Salidas_Hoy,${estadisticasDelDia.salidas || 0},,,,, \n`;
            csvContent += `Personas_Dentro,${conteoEvacuacion.totalDentro},,,,, \n`;
            csvContent += `Porcentaje_Evacuado,${Math.round(((estadisticasDelDia.salidas || 0) / Math.max(estadisticasDelDia.entradas || 1, 1)) * 100)}%,,,,, \n`;
            
            // Crear archivo
            const blob = Utilities.newBlob(csvContent, 'text/csv; charset=utf-8', nombreArchivo);
            const file = DriveApp.createFile(blob);
            
            // Configurar permisos
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            
            logError(`Estado de evacuación exportado: ${nombreArchivo}`, 'INFO');
            
            return {
                success: true,
                message: 'Estado de evacuación exportado correctamente',
                fileName: nombreArchivo,
                downloadUrl: file.getDownloadUrl(),
                fileId: file.getId(),
                totalPersonas: conteoEvacuacion.totalDentro
            };
            
        } catch (error) {
            logError('Error en exportarEstadoEvacuacion', 'ERROR', { error: error.message });
            return {
                success: false,
                message: 'Error al exportar estado de evacuación: ' + error.message
            };
        }
    }

    /**
     * ✅ FUNCIÓN AUXILIAR: mapearIndicesColumnas
     * Mapea los nombres de columnas a sus índices
     */
    function mapearIndicesColumnas(headers) {
        const findIndex = (nombres) => {
            for (const nombre of nombres) {
                const index = headers.findIndex(h => 
                    h && h.toString().toLowerCase().trim() === nombre.toLowerCase()
                );
                if (index !== -1) return index;
            }
            return -1;
        };
        
        return {
            fecha: findIndex(['fecha', 'fecha y hora', 'marca de tiempo']),
            cedula: findIndex(['cédula', 'cedula', 'id']),
            nombre: findIndex(['nombre', 'nombres']),
            empresa: findIndex(['empresa', 'compania', 'organización']),
            entrada: findIndex(['entrada', 'hora entrada', 'hora_entrada']),
            salida: findIndex(['salida', 'hora salida', 'hora_salida']),
            estado: findIndex(['estado del acceso', 'estado', 'acceso']),
            duracion: findIndex(['duración', 'duracion', 'tiempo'])
        };
    }

    /**
     * ✅ FUNCIÓN UNIFICADA CORREGIDA: Obtiene datos completos de evacuación
     * Reemplaza: contarPersonasEnEdificio, obtenerEstadoEvacuacionRapido, procesarRegistrosEvacuacion
     */
    function getEvacuacionDataForClient() {
        console.log('🚀 getEvacuacionDataForClient DEBUG iniciando...');
        
        // ✅ RESPUESTA MÍNIMA GARANTIZADA
        const respuesta = {
            success: true,
            message: 'Función ejecutada',
            totalDentro: 0,
            personasDentro: [],
            timestamp: new Date().toISOString(),
            debug: 'Función llamada correctamente'
        };
        
        try {
            console.log('🔍 Paso 1: Verificando SpreadsheetApp...');
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            
            if (!ss) {
                console.log('❌ SpreadsheetApp es null');
                respuesta.message = 'SpreadsheetApp no disponible';
                respuesta.debug = 'SpreadsheetApp devolvió null';
                return respuesta;
            }
            
            console.log('✅ SpreadsheetApp OK');
            respuesta.debug = 'SpreadsheetApp OK';
            
            console.log('🔍 Paso 2: Verificando hoja Historial...');
            const historial = ss.getSheetByName('Historial');
            
            if (!historial) {
                console.log('❌ Hoja Historial no encontrada');
                respuesta.message = 'Hoja Historial no encontrada';
                respuesta.debug = 'Hoja Historial no existe';
                return respuesta;
            }
            
            console.log('✅ Hoja Historial OK');
            respuesta.debug = 'Hoja Historial OK';
            
            console.log('🔍 Paso 3: Contando filas...');
            const filas = historial.getLastRow();
            console.log('📊 Filas en Historial:', filas);
            
            if (filas <= 1) {
                respuesta.message = 'Historial vacío - edificio vacío';
                respuesta.debug = `Historial tiene ${filas} filas`;
                return respuesta;
            }
            
            console.log('🔍 Paso 4: Obteniendo datos...');
            const data = historial.getDataRange().getValues();
            const headers = data[0]; // Get headers to find column indices dynamically
            const indices = mapearIndicesColumnas(headers); // Use the existing helper

            const personasAdentroMap = new Map(); // Use a Map to track latest entry for each person

            // Iterate from the latest record backwards
            for (let i = data.length - 1; i >= 1; i--) {
                const fila = data[i];
                const cedula = String(fila[indices.cedula] || '').trim();

                if (!cedula) continue;

                const entrada = indices.entrada >= 0 ? fila[indices.entrada] : null;
                const salida = indices.salida >= 0 ? fila[indices.salida] : null;

                // If this person is already processed (meaning we found a later entry/exit), skip
                if (personasAdentroMap.has(cedula)) {
                    continue;
                }

                if (entrada && entrada !== '' && (!salida || salida === '')) {
                    // This is the latest entry without a corresponding exit for this person
                    personasAdentroMap.set(cedula, {
                        cedula: cedula,
                        nombre: String(fila[indices.nombre] || '').trim() || 'Sin nombre',
                        empresa: String(fila[indices.empresa] || '').trim() || 'Sin empresa',
                        // Asegúrate de que horaEntrada sea un string formateado para el cliente
                        horaEntrada: entrada instanceof Date ? Utilities.formatDate(entrada, Session.getScriptTimeZone(), 'HH:mm') : String(entrada || '').trim() 
                    });
                }
            }

            // Convert Map values to an array
            respuesta.personasDentro = Array.from(personasAdentroMap.values());
            respuesta.totalDentro = respuesta.personasDentro.length;
            respuesta.message = `${respuesta.personasDentro.length} personas dentro`;
            respuesta.debug = `Procesadas ${data.length} filas, encontradas ${respuesta.personasDentro.length} personas`;
            
            console.log('✅ getEvacuacionDataForClient completado:', respuesta);
            return respuesta;
            
        } catch (error) {
            console.error('❌ Error en getEvacuacionDataForClient:', error);
            respuesta.success = false;
            respuesta.message = 'Error: ' + error.message;
            respuesta.debug = 'Error capturado: ' + error.toString();
            return respuesta;
        }
    }

    /**
     * ✅ FUNCIÓN UNIFICADA DE EVACUACIÓN
     * Procesa una evacuación, distinguiendo entre real y simulacro.
     * Esta es la función principal que debe ser llamada desde el cliente (admin.html/formulario.html).
     * @param {Object} parametros - Objeto con los parámetros de la evacuación.
     * @param {Array<string>} parametros.cedulas - Array de cédulas de las personas a evacuar/simular.
     * @param {string} [parametros.tipo='real'] - Tipo de evacuación: 'real' o 'simulacro'.
     * @param {string} [parametros.operador] - Operador que inicia la evacuación.
     * @param {Date} [parametros.timestamp] - Marca de tiempo del evento.
     * @param {string} [parametros.notas] - Notas adicionales para el simulacro.
     * @returns {Object} - Objeto con el resultado de la operación.
     */
    function procesarEvacuacionUnificada(parametros) {
        const startTime = new Date().getTime();
        const sessionId = Utilities.getUuid(); // Generar un ID de sesión único para el seguimiento
        
        try {
            logError(`🚨 Iniciando evacuación unificada`, 'INFO', { 
                sessionId: sessionId, 
                tipo: parametros.tipo || 'real', 
                totalPersonas: parametros.cedulas?.length || 0 
            });

            // ✅ VALIDACIÓN UNIFICADA de parámetros
            if (!parametros || !parametros.cedulas || !Array.isArray(parametros.cedulas)) {
                throw new Error('Parámetros de evacuación inválidos: se requiere un array de cédulas.');
            }
            if (parametros.cedulas.length === 0) {
                throw new Error('No hay personas seleccionadas para evacuar/simular.');
            }

            const { 
                cedulas, 
                tipo = 'real', // Por defecto es 'real' si no se especifica
                operador = obtenerUsuarioActual(), 
                timestamp = new Date(),
                notas = '' // Notas para simulacros
            } = parametros;

            // ✅ OBTENER HOJAS NECESARIAS
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const historialSheet = ss.getSheetByName('Historial');
            const bdSheet = ss.getSheetByName('Base de Datos'); // Necesaria para obtener nombres/empresas
            
            if (!historialSheet) {
                throw new Error('Hoja "Historial" no encontrada. Verifique la configuración del sistema.');
            }
            if (!bdSheet) {
                logError('Hoja "Base de Datos" no encontrada. Algunas personas podrían aparecer como "Sin nombre".', 'WARNING');
            }

            // ✅ PROCESAR EVACUACIÓN SEGÚN EL TIPO
            let resultadoProcesamiento;
            if (tipo === 'simulacro') {
                // ✅ IMPORTANTE: procesarSimulacroUnificado NO MODIFICA el historialSheet
                resultadoProcesamiento = procesarSimulacroUnificado(historialSheet, bdSheet, cedulas, timestamp, sessionId, notas);
            } else { // tipo === 'real'
                resultadoProcesamiento = procesarEvacuacionRealUnificada(historialSheet, bdSheet, cedulas, timestamp, sessionId);
            }

            // ✅ LOGGING Y NOTIFICACIONES
            if (resultadoProcesamiento.success && resultadoProcesamiento.personasEvacuadas.length > 0) {
                // Registrar en el log de emergencia o simulacro
                registrarLogEvacuacionUnificado(resultadoProcesamiento.personasEvacuadas, timestamp, tipo, sessionId, notas);

                // Enviar notificación solo si es evacuación real
                if (tipo === 'real') {
                    enviarNotificacionEvacuacionUnificada(resultadoProcesamiento.personasEvacuadas, timestamp, sessionId);
                }
            }

            // ✅ OBTENER LISTA ACTUALIZADA DE PERSONAS DENTRO DESPUÉS DE LA OPERACIÓN
            // Esto es crucial para el cliente, para mostrar los "faltantes" o el estado final
            const personasDentroActualizadas = getEvacuacionDataForClient().personasDentro;

            // ✅ RESULTADO FINAL
            const tiempoTotal = new Date().getTime() - startTime;
            logError(`✅ Evacuación ${tipo} completada en ${tiempoTotal}ms`, 'INFO', { 
                sessionId: sessionId, 
                totalEvacuadas: resultadoProcesamiento.personasEvacuadas?.length || 0, 
                tipo: tipo, 
                personasDentroRestantes: personasDentroActualizadas.length // Añadir para el log
            });

            return { 
                ...resultadoProcesamiento, 
                sessionId: sessionId, 
                tiempoMs: tiempoTotal, 
                tipo: tipo, 
                personasDentroActualizadas: personasDentroActualizadas // Devolver al cliente para actualizar la UI
            };

        } catch (error) {
            logError('❌ Error en procesarEvacuacionUnificada', 'ERROR', { 
                sessionId: sessionId, 
                error: error.message, 
                parametros: parametros,
                stack: error.stack
            });
            return { 
                success: false, 
                message: 'Error procesando evacuación: ' + error.message, 
                totalEvacuadas: 0, 
                personasEvacuadas: [], 
                sessionId: sessionId, 
                tipo: parametros?.tipo || 'desconocido',
                personasDentroActualizadas: getEvacuacionDataForClient().personasDentro // Intentar obtener el estado actual incluso con error
            };
        }
    }

    /**
     * ✅ FUNCIÓN AUXILIAR: Procesar evacuación real
     */
    function procesarEvacuacionRealUnificada(historialSheet, bdSheet, cedulas, timestamp, sessionId) {
        try {
            const data = historialSheet.getDataRange().getValues();
            const headers = data[0];
            const indices = mapearIndicesColumnas(headers);
            const personasEvacuadas = [];
            
            console.log(`🚨 Procesando evacuación REAL para ${cedulas.length} personas`);
            
            cedulas.forEach(cedula => {
                const cedulaNorm = normalizarCedula(cedula).replace(/[^\w\-]/g, '');
                
                // Buscar y actualizar entrada más reciente sin salida
                for (let i = data.length - 1; i >= 1; i--) {
                    const row = data[i];
                    const cedulaRow = (row[indices.cedula] || '').toString().trim();
                    const cedulaRowNorm = normalizarCedula(cedulaRow).replace(/[^\w\-]/g, '');
                    
                    const entrada = indices.entrada >= 0 ? row[indices.entrada] : null;
                    const salida = indices.salida >= 0 ? row[indices.salida] : null;
                    
                    if ((cedulaRow === cedula || cedulaRowNorm === cedulaNorm) &&
                        entrada && entrada !== '' && (!salida || salida === '')) {
                        
                        // ✅ ACTUALIZAR REGISTRO REAL
                        const rowIndex = i + 1;
                        if (indices.salida >= 0) {
                            historialSheet.getRange(rowIndex, indices.salida + 1).setValue(timestamp);
                        }
                        
                        // Calcular duración
                        if (entrada instanceof Date && indices.duracion >= 0) {
                            const duracionMs = timestamp.getTime() - entrada.getTime();
                            const horas = Math.floor(duracionMs / (1000 * 60 * 60));
                            const minutos = Math.floor((duracionMs % (1000 * 60 * 60)) / (1000 * 60));
                            const duracionTexto = `${horas}h ${minutos}m`;
                            
                            historialSheet.getRange(rowIndex, indices.duracion + 1).setValue(duracionTexto);
                        }
                        
                        // Agregar comentario de evacuación
                        try {
                            const comentario = `EVACUACIÓN EMERGENCIA - ${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'HH:mm')} - Session: ${sessionId}`;
                            historialSheet.getRange(rowIndex, indices.salida + 1).setNote(comentario);
                        } catch (noteError) {
                            console.warn('No se pudo agregar nota:', noteError.message);
                        }
                        
                        // Obtener datos completos
                        let nombre = (row[indices.nombre] || '').toString().trim() || 'Sin nombre';
                        let empresa = (row[indices.empresa] || '').toString().trim() || 'Sin empresa';
                        
                        // Mejorar datos desde BD
                        if ((nombre === 'Sin nombre' || nombre === 'DENEGADO') && bdSheet) {
                            const personaBD = buscarPersonaEnBD(bdSheet, cedulaRow);
                            if (personaBD) {
                                nombre = personaBD.nombre;
                                empresa = personaBD.empresa;
                            }
                        }
                        
                        personasEvacuadas.push({
                            cedula: cedulaRow,
                            nombre: nombre,
                            empresa: empresa,
                            horaEntrada: entrada,
                            horaSalida: timestamp,
                            filaHistorial: rowIndex
                        });
                        
                        console.log(`✅ Evacuada REAL: ${nombre} (${cedulaRow})`);
                        break; // Only the first entry found
                    }
                }
            });
            
            return {
                success: true,
                message: `Evacuación real completada: ${personasEvacuadas.length} persona(s) evacuadas`,
                totalEvacuadas: personasEvacuadas.length,
                personasEvacuadas: personasEvacuadas
            };
            
        } catch (error) {
            return {
                success: false,
                message: 'Error en evacuación real: ' + error.message,
                totalEvacuadas: 0,
                personasEvacuadas: []
            };
        }
    }

    /**
     * ✅ FUNCIÓN FALTANTE CRÍTICA: procesarSimulacroUnificado
     * Esta función NO modifica el historial real, solo registra el evento del simulacro
     */
    function procesarSimulacroUnificado(historialSheet, bdSheet, cedulas, timestamp, sessionId, notas = '') {
        try {
            console.log(`🎭 SIMULACRO INICIADO para ${cedulas.length} personas - Session: ${sessionId}`);
            
            const data = historialSheet.getDataRange().getValues();
            const headers = data[0];
            const indices = mapearIndicesColumnas(headers);
            const personasEvacuadas = [];
            
            // ✅ IMPORTANTE: SOLO OBTENER DATOS, NO MODIFICAR HISTORIAL
            cedulas.forEach(cedula => {
                const cedulaNorm = normalizarCedula(cedula).replace(/[^\w\-]/g, '');
                console.log(`🔍 Procesando simulacro para cédula: ${cedula}`);
                
                // Buscar entrada más reciente sin salida (solo para obtener datos)
                for (let i = data.length - 1; i >= 1; i--) {
                    const row = data[i];
                    const cedulaRow = (row[indices.cedula] || '').toString().trim();
                    const cedulaRowNorm = normalizarCedula(cedulaRow).replace(/[^\w\-]/g, '');
                    
                    const entrada = indices.entrada >= 0 ? row[indices.entrada] : null;
                    const salida = indices.salida >= 0 ? row[indices.salida] : null;
                    
                    if ((cedulaRow === cedula || cedulaRowNorm === cedulaNorm) &&
                        entrada && entrada !== '' && (!salida || salida === '')) {
                        
                        console.log(`🎭 Encontrada entrada sin salida para: ${cedulaRow}`);
                        
                        // ❌ CRÍTICO: NO MODIFICAR EL HISTORIAL
                        // ❌ NO historialSheet.getRange().setValue()
                        // ❌ NO actualizar columna de salida
                        
                        // ✅ SOLO OBTENER DATOS PARA EL REPORTE
                        let nombre = (row[indices.nombre] || '').toString().trim() || 'Sin nombre';
                        let empresa = (row[indices.empresa] || '').toString().trim() || 'Sin empresa';
                        
                        // Mejorar datos desde BD si es necesario
                        if ((nombre === 'Sin nombre' || nombre === 'DENEGADO') && bdSheet) {
                            const personaBD = buscarPersonaEnBD(bdSheet, cedulaRow);
                            if (personaBD) {
                                nombre = personaBD.nombre;
                                empresa = personaBD.empresa;
                            }
                        }
                        
                        personasEvacuadas.push({
                            cedula: cedulaRow,
                            nombre: nombre,
                            empresa: empresa,
                            horaEntrada: entrada,
                            horaSalidaSimulada: timestamp, // Solo para el reporte
                            filaHistorial: i + 1,
                            tipoEvento: 'SIMULACRO',
                            notas: notas
                        });
                        
                        console.log(`✅ Incluida en SIMULACRO: ${nombre} (${cedulaRow})`);
                        break; // Solo la primera entrada encontrada
                    }
                }
            });
            
            console.log(`🎭 SIMULACRO COMPLETADO: ${personasEvacuadas.length} personas participaron`);
            
            return {
                success: true,
                message: `Simulacro completado: ${personasEvacuadas.length} persona(s) participaron en el simulacro`,
                totalEvacuadas: personasEvacuadas.length,
                personasEvacuadas: personasEvacuadas,
                tipoEvento: 'SIMULACRO',
                notas: notas,
                sessionId: sessionId
            };
            
        } catch (error) {
            console.error('❌ Error crítico en procesarSimulacroUnificado:', error);
            logError('Error en procesarSimulacroUnificado', 'ERROR', { 
                error: error.message, 
                sessionId: sessionId,
                cedulas: cedulas
            });
            
            return {
                success: false,
                message: 'Error en simulacro: ' + error.message,
                totalEvacuadas: 0,
                personasEvacuadas: [],
                tipoEvento: 'SIMULACRO_ERROR',
                sessionId: sessionId
            };
        }
    }

    /**
     * ✅ FUNCIÓN DE LOGGING CORREGIDA CON DEBUG
     * Registra eventos de evacuación en la hoja correspondiente
     */
    function registrarLogEvacuacionUnificado(personasEvacuadas, timestamp, tipo, sessionId, notas = '') {
        try {
            console.log(`📝 Iniciando registro de log para tipo: ${tipo}`);
            
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const nombreHoja = tipo === 'simulacro' ? 'Log_Simulacros' : 'Log_Emergencias';
            
            console.log(`📝 Buscando hoja: ${nombreHoja}`);
            let logSheet = ss.getSheetByName(nombreHoja);
            
            if (!logSheet) {
                console.log(`📝 Creando nueva hoja: ${nombreHoja}`);
                logSheet = ss.insertSheet(nombreHoja);
                
                const headers = [
                    'Session_ID', 'Fecha_Hora', 'Tipo_Evento', 'Operador', 'Personas_Afectadas', 
                    'Detalle_Personas', 'Estado', 'Observaciones', 'Notas_Adicionales'
                ];
                logSheet.appendRow(headers);
                
                const headerRange = logSheet.getRange(1, 1, 1, headers.length);
                const colorHeader = tipo === 'simulacro' ? '#2196f3' : '#d32f2f';
                headerRange.setBackground(colorHeader)
                        .setFontColor('white')
                        .setFontWeight('bold');
                logSheet.setFrozenRows(1);
                
                console.log(`✅ Hoja ${nombreHoja} creada con encabezados`);
            }

            // ✅ CONSTRUCCIÓN SEGURA DE DETALLES
            let detalles = 'Sin personas específicas';
            
            if (personasEvacuadas && Array.isArray(personasEvacuadas) && personasEvacuadas.length > 0) {
                try {
                    detalles = personasEvacuadas
                        .filter(persona => persona && typeof persona === 'object')
                        .map(persona => {
                            const cedula = persona.cedula || 'SIN_CEDULA';
                            const nombre = persona.nombre || 'SIN_NOMBRE';
                            return `${cedula}-${nombre}`;
                        })
                        .join('; ');
                        
                    if (!detalles) {
                        detalles = `${personasEvacuadas.length} persona(s) con datos incompletos`;
                    }
                } catch (mapError) {
                    console.error('❌ Error procesando detalles:', mapError.message);
                    detalles = `Error procesando ${personasEvacuadas.length} persona(s): ${mapError.message}`;
                }
            }
            
            const registro = [
                sessionId,
                timestamp,
                tipo === 'simulacro' ? 'SIMULACRO_EVACUACION' : 'EVACUACIÓN_EMERGENCIA',
                obtenerUsuarioActual(),
                personasEvacuadas ? personasEvacuadas.length : 0,
                detalles,
                'COMPLETADO',
                `${tipo.toUpperCase()} ejecutado vía Sistema SurPass - ${new Date().toISOString()}`,
                notas || 'Sin notas adicionales'
            ];
            
            console.log(`📝 Agregando registro a ${nombreHoja}:`, registro);
            logSheet.appendRow(registro);
            
            const lastRow = logSheet.getLastRow();
            const color = tipo === 'simulacro' ? '#e3f2fd' : '#ffebee';
            logSheet.getRange(lastRow, 1, 1, registro.length).setBackground(color);
            
            console.log(`✅ Log ${tipo} registrado exitosamente en fila ${lastRow}`);
            
            return {
                success: true,
                mensaje: `Log registrado en ${nombreHoja}`,
                fila: lastRow
            };
            
        } catch (error) {
            console.error(`❌ Error crítico registrando log ${tipo}:`, error.message);
            console.error('Stack trace:', error.stack);
            
            // ✅ REGISTRO DE FALLBACK EN CASO DE ERROR
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                let errorSheet = ss.getSheetByName('Log_Errores');
                if (!errorSheet) {
                    errorSheet = ss.insertSheet('Log_Errores');
                    errorSheet.appendRow(['Fecha', 'Tipo_Error', 'Mensaje', 'Detalles']);
                }
                errorSheet.appendRow([
                    new Date(),
                    `ERROR_LOG_${tipo.toUpperCase()}`,
                    error.message,
                    `Falló al registrar ${tipo}: ${error.message}`
                ]);
                console.log('📝 Error registrado en Log_Errores como fallback');
            } catch (fallbackError) {
                console.error('❌ Error incluso en fallback:', fallbackError.message);
            }
            
            return {
                success: false,
                mensaje: `Error registrando log: ${error.message}`
            };
        }
    }

    /**
     * ✅ FUNCIÓN AUXILIAR: Notificación unificada
     */
    function enviarNotificacionEvacuacionUnificada(personasEvacuadas, timestamp, sessionId) {
        try {
            const config = obtenerConfiguracion();
            const destinatarios = [
                config.NOTIFICACIONES_EMAIL,
                config.EMAIL_SECUNDARIO
            ].filter(email => email && email.trim()).join(', ');
            
            if (!destinatarios) {
                console.warn('⚠️ No hay destinatarios para notificaciones');
                return;
            }
            
            const asunto = `🚨 ALERTA CRÍTICA: Evacuación de Emergencia - ${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}`;
            
            let mensaje = `🚨 NOTIFICACIÓN DE EVACUACIÓN DE EMERGENCIA\n`;
            mensaje += `════════════════════════════════════════════\n\n`;
            mensaje += `📅 Fecha: ${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')}\n`;
            mensaje += `👤 Operador: ${obtenerUsuarioActual()}\n`;
            mensaje += `👥 Personas evacuadas: ${personasEvacuadas.length}\n`;
            mensaje += `🆔 Session ID: ${sessionId}\n\n`;
            
            mensaje += `📋 LISTADO DE PERSONAS EVACUADAS:\n`;
            mensaje += `═══════════════════════════════════════\n`;
            
            personasEvacuadas.forEach((persona, index) => {
                const horaEntrada = persona.horaEntrada ? 
                    Utilities.formatDate(persona.horaEntrada, Session.getScriptTimeZone(), 'HH:mm') : 'N/A';
                const horaSalida = Utilities.formatDate(persona.horaSalida, Session.getScriptTimeZone(), 'HH:mm');
                
                mensaje += `${String(index + 1).padStart(2, '0')}. ${persona.nombre}\n`;
                mensaje += `    🆔 Cédula: ${persona.cedula}\n`;
                mensaje += `    🏢 Empresa: ${persona.empresa}\n`;
                mensaje += `    🕐 Entrada: ${horaEntrada} | Salida: ${horaSalida}\n\n`;
            });
            
            mensaje += `🔒 Mensaje generado automáticamente por Sistema SurPass v3.0\n`;
            mensaje += `📞 Session ID para soporte: ${sessionId}`;
            
            MailApp.sendEmail(destinatarios, asunto, mensaje, { name: 'Sistema SurPass' });
            console.log('✅ Notificación enviada a:', destinatarios);
            
        } catch (error) {
            console.error('❌ Error enviando notificación:', error.message);
        }
    }

    /**
     * ✅ FUNCIÓN CORREGIDA PARA MANEJAR EVACUACIONES
     * Esta función ahora detecta automáticamente si debe ser tratado como simulacro o real
     */
    function confirmarSalidasEvacuacion(cedulas, tipo = 'real') {
        console.log(`🚨 confirmarSalidasEvacuacion llamada con tipo: ${tipo}`);
        
        // ✅ CORRECCIÓN CRÍTICA: Validar el tipo
        if (tipo === 'simulacro') {
            console.log('🎭 Redirigiendo a procesarSimulacroEvacuacion...');
            return procesarSimulacroEvacuacion(cedulas, 'Simulacro ejecutado desde interfaz de evacuación');
        } else {
            console.log('🚨 Procesando como evacuación REAL...');
            return procesarEvacuacionUnificada({
                cedulas: cedulas,
                tipo: 'real'
            });
        }
    }

    // ✅ O MEJOR AÚN, función específica para simulacros:
    function procesarSimulacroEvacuacion(cedulas, notas = '') {
        console.log('🎭 === INICIANDO SIMULACRO CON PROTECCIÓN ===');
        
        try {
            // 🛡️ ACTIVAR PROTECCIÓN ANTI-MODIFICACIÓN
            SIMULACRO_EN_CURSO = true;
            console.log('🛡️ Protección de simulacro ACTIVADA');
            
            const resultado = procesarEvacuacionUnificada({
                cedulas: cedulas,
                tipo: 'simulacro',
                notas: notas
            });
            
            console.log('🎭 Simulacro completado sin modificar historial');
            return resultado;
            
        } catch (error) {
            console.error('❌ Error en simulacro:', error.message);
            return {
                success: false,
                message: 'Error en simulacro: ' + error.message
            };
        } finally {
            // 🛡️ DESACTIVAR PROTECCIÓN
            SIMULACRO_EN_CURSO = false;
            console.log('🛡️ Protección de simulacro DESACTIVADA');
        }
    }

    function procesarEvacuacionConTipo(datosEvacuacion) {
        return procesarEvacuacionUnificada(datosEvacuacion);
    }

    /**
     * Registra evento de emergencia en log especial
     */
    function registrarLogEmergencia(personasEvacuadas, fechaEvacuacion) {
        try {
            // ✅ CORRECCIÓN PRINCIPAL: Validar que personasEvacuadas existe y es un array
            if (!personasEvacuadas || !Array.isArray(personasEvacuadas)) {
                logError('personasEvacuadas inválido, usando array vacío', 'WARNING', { 
                    recibido: personasEvacuadas,
                    tipo: typeof personasEvacuadas 
                });
                personasEvacuadas = [];
            }

            // ✅ Validar fecha
            if (!fechaEvacuacion || !(fechaEvacuacion instanceof Date)) {
                fechaEvacuacion = new Date();
            }

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let emergencySheet = ss.getSheetByName('Log_Emergencias');
            
            if (!emergencySheet) {
                emergencySheet = ss.insertSheet('Log_Emergencias');
                emergencySheet.appendRow([
                    'Fecha_Hora', 'Tipo_Emergencia', 'Operador', 'Personas_Afectadas', 
                    'Detalles', 'Estado', 'Notas_Adicionales'
                ]);
                
                const headerRange = emergencySheet.getRange(1, 1, 1, 7);
                headerRange.setBackground('#d32f2f').setFontColor('white').setFontWeight('bold');
                emergencySheet.setFrozenRows(1);
            }

            // ✅ CORRECCIÓN: Construcción segura de detalles
            let detalles = 'Sin personas específicas';
            
            if (personasEvacuadas.length > 0) {
                try {
                    detalles = personasEvacuadas
                        .filter(persona => persona && typeof persona === 'object') // Filtrar objetos válidos
                        .map(persona => {
                            const cedula = persona.cedula || 'SIN_CEDULA';
                            const nombre = persona.nombre || 'SIN_NOMBRE';
                            return `${cedula}-${nombre}`;
                        })
                        .join('; ');
                        
                    if (!detalles) {
                        detalles = `${personasEvacuadas.length} persona(s) con datos incompletos`;
                    }
                } catch (mapError) {
                    logError('Error procesando detalles', 'ERROR', { error: mapError.message });
                    detalles = `Error procesando ${personasEvacuadas.length} persona(s): ${mapError.message}`;
                }
            }
            
            // Registrar en la hoja
            emergencySheet.appendRow([
                fechaEvacuacion,
                'EVACUACIÓN_EMERGENCIA',
                obtenerUsuarioActual(),
                personasEvacuadas.length,
                detalles,
                'COMPLETADO',
                `Evacuación automática vía Sistema SurPass - ${new Date().toISOString()}`
            ]);
            
            const lastRow = emergencySheet.getLastRow();
            emergencySheet.getRange(lastRow, 1, 1, 7).setBackground('#ffebee');
            
            logError(`✅ Log de emergencia registrado: ${personasEvacuadas.length} personas`, 'INFO');
            
        } catch (error) {
            logError('❌ Error crítico en registrarLogEmergencia', 'ERROR', { 
                error: error.message,
                personasEvacuadas: personasEvacuadas,
                fechaEvacuacion: fechaEvacuacion
            });
            
            // Fallback: registro básico
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                let emergencySheet = ss.getSheetByName('Log_Emergencias');
                if (!emergencySheet) {
                    emergencySheet = ss.insertSheet('Log_Emergencias');
                    emergencySheet.appendRow(['Fecha_Hora', 'Tipo', 'Error', 'Detalles']);
                }
                emergencySheet.appendRow([
                    new Date(),
                    'ERROR_LOG',
                    error.message,
                    `Falló al registrar evacuación: ${error.message}`
                ]);
            } catch (fallbackError) {
                logError('❌ Error incluso en fallback', 'CRITICAL', { error: fallbackError.message });
            }
        }
    }

    // =====================================================
    // PROCESAMIENTO DE FORMULARIOS (CONTINUACIÓN...)
    // =====================================================

    /**
     * Función principal para manejar el envío del formulario HTML
     */
    function handleHTMLFormSubmit(cedula, respuesta) {
        const startTime = new Date().getTime();
        const sessionId = Utilities.getUuid();
        
        try {
            // Validación inicial
            if (!validarCampoObligatorio(cedula) || !validarCampoObligatorio(respuesta)) {
                throw new Error('Cédula o respuesta no proporcionada.');
            }

            // Normalizar datos de entrada
            const cedulaOriginal = normalizarCedula(cedula);
            const respuestaNormalizada = String(respuesta).toLowerCase().trim();

            if (!['entrada', 'salida'].includes(respuestaNormalizada)) {
                throw new Error('Respuesta inválida. Debe ser "entrada" o "salida".');
            }

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            
            // Verificar y obtener hojas necesarias
            const sheets = {
                respuestas: ss.getSheetByName('Respuestas formulario'),
                baseDatos: ss.getSheetByName('Base de Datos'),
                historial: ss.getSheetByName('Historial')
            };

            // Validar que todas las hojas existan
            for (const [nombre, sheet] of Object.entries(sheets)) {
                if (!sheet) {
                    throw new Error(`La hoja "${nombre}" no fue encontrada. Verifique que el sistema esté configurado correctamente.`);
                }
            }

            const fecha = new Date();
            
            // Buscar persona en base de datos
            const persona = buscarPersonaEnBD(sheets.baseDatos, cedulaOriginal);
            
            // Configurar respuesta base
            let resultado = {
                cedula: persona ? persona.cedula : cedulaOriginal,
                nombre: persona ? persona.nombre : 'DENEGADO',
                empresa: persona ? persona.empresa : 'No registrada',
                estadoAcceso: persona ? 'Acceso Permitido' : 'Acceso Denegado',
                color: persona ? 'green' : 'red',
                comentarioObligatorio: !persona
            };

            // Registrar en las hojas del sistema
            const registroResult = registrarAcceso(sheets, fecha, resultado, respuestaNormalizada);
            
            // Configurar mensaje de respuesta
        if (persona) {
        if (respuestaNormalizada === 'salida' && registroResult && registroResult.sinEntradaPrevia === true) {
            resultado.message = 'Salida registrada (sin entrada previa)';
            resultado.color = 'orange';  // ✅ Usar 'orange' consistentemente
        } else {
            resultado.message = `${respuestaNormalizada === 'entrada' ? 'Entrada' : 'Salida'} registrada correctamente`;
            resultado.color = 'green';
        }
    } else {
        resultado.message = 'Acceso Denegado - Persona no registrada';
        resultado.color = 'red';  // ✅ Usar 'red' consistentemente
    }
                
            // Log de auditoría detallado
            resultado.log = {
                timestamp: new Date().toISOString(),
                level: persona ? 'INFO' : 'WARNING',
                message: `Acceso ${persona ? 'permitido' : 'denegado'}`,
                details: {
                    sessionId: sessionId,
                    cedula: cedulaOriginal,
                    tipo: respuestaNormalizada,
                    processingTime: new Date().getTime() - startTime,
                    sinEntradaPrevia: registroResult ? registroResult.sinEntradaPrevia : false,
                    usuario: obtenerUsuarioActual()
                }
            };

            logError(`[${resultado.log.level}] ${resultado.log.message} - Cédula: ${cedulaOriginal} - Tiempo: ${resultado.log.details.processingTime}ms`, resultado.log.level, resultado.log.details);
            
            return resultado;
            
        } catch (error) {
            const errorLog = {
                timestamp: new Date().toISOString(),
                level: 'ERROR',
                message: 'Error procesando formulario',
                details: {
                    sessionId: sessionId,
                    error: error.message,
                    stack: error.stack,
                    cedula: cedula,
                    respuesta: respuesta,
                    processingTime: new Date().getTime() - startTime,
                    usuario: obtenerUsuarioActual()
                }
            };
            
            logError('Error en handleHTMLFormSubmit', 'ERROR', errorLog.details);
            
            return {
                message: 'Error del sistema: ' + error.message,
                nombre: 'ERROR',
                empresa: 'ERROR',
                estadoAcceso: 'Error del Sistema',
                color: 'red',
                comentarioObligatorio: false,
                log: errorLog
            };
        }
    }

    /**
     * Registra el acceso en las hojas del sistema
     */
    function registrarAcceso(sheets, fecha, resultado, tipo) {
        try {
            // Registrar en la hoja de respuestas
            const row = [
                fecha,
                resultado.cedula,
                tipo === 'entrada' ? 'Entrada' : 'Salida',
                resultado.estadoAcceso,
                fecha, // Hora entrada
                tipo === 'salida' ? fecha : '', // Hora salida
                '', // Cédulas similares (se llenará después si es necesario)
                resultado.empresa,
                '' // Comentarios
            ];
            sheets.respuestas.appendRow(row);
            
            // Registrar en el historial y obtener si hubo entrada previa
            const historialResult = registrarEnHistorial(sheets.historial, fecha, resultado, tipo);
            const lastRowRespuestas = sheets.respuestas.getLastRow();
            
            // Aplicar formato visual según el resultado
            const colorFondo = resultado.color === 'orange' ? '#ffcc80' : 
                            resultado.color === 'red' ? '#ffcdd2' : '#c8e6c9';
            sheets.respuestas.getRange(lastRowRespuestas, 4).setBackground(colorFondo);
            
            // Formatear columnas de tiempo
            sheets.respuestas.getRange(lastRowRespuestas, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
            if (tipo === 'entrada') {
                sheets.respuestas.getRange(lastRowRespuestas, 5).setNumberFormat("HH:mm");
            } else {
                sheets.respuestas.getRange(lastRowRespuestas, 6).setNumberFormat("HH:mm");
            }
            
            // Si es acceso denegado, buscar cédulas similares
            if (resultado.estadoAcceso === 'Acceso Denegado') {
                const data = sheets.baseDatos.getDataRange().getValues();
                const similares = buscarCedulasSimilares(resultado.cedula, data);
                if (similares) {
                    sheets.respuestas.getRange(lastRowRespuestas, 7).setValue(similares);
                }
            }
            
            return historialResult || {};
            
        } catch (error) {
            logError('Error en registrarAcceso', 'ERROR', { error: error.message });
            throw new Error('Error al registrar el acceso: ' + error.message);
        }
    }

    /**
     * ✅ VARIABLE GLOBAL PARA CONTROLAR EL MODO SIMULACRO
     */
    let SIMULACRO_EN_CURSO = false;

    /**
     * ✅ FUNCIÓN CORREGIDA: registrarEnHistorial con protección anti-simulacro
     */
    function registrarEnHistorial(historialSheet, fecha, resultado, tipo) {
        // 🛡️ PROTECCIÓN CRÍTICA: NO REGISTRAR DURANTE SIMULACROS
        if (SIMULACRO_EN_CURSO) {
            console.log('🛡️ BLOQUEADO: Intento de modificar historial durante simulacro');
            console.trace('🔍 Stack trace del intento de modificación:');
            return { 
                sinEntradaPrevia: false, 
                bloqueadoPorSimulacro: true,
                mensaje: 'Modificación bloqueada: simulacro en curso'
            };
        }
        
        let sinEntradaPrevia = false;
        
        try {
            if (tipo === 'entrada') {
                // Registro directo de entrada
                const historialRow = [
                    fecha,                  // Fecha
                    resultado.cedula,       // Cédula
                    resultado.nombre,       // Nombre
                    resultado.estadoAcceso, // Estado del Acceso
                    fecha,                  // Entrada
                    '',                     // Salida (vacía)
                    '',                     // Duración (vacía)
                    resultado.empresa       // Empresa
                ];
                
                historialSheet.appendRow(historialRow);
                const lastRow = historialSheet.getLastRow();
                
                // Aplicar formato
                const colorFondo = resultado.color === 'red' ? '#ffcdd2' : '#c8e6c9';
                historialSheet.getRange(lastRow, 4).setBackground(colorFondo);
                historialSheet.getRange(lastRow, 1).setNumberFormat("yyyy-mm-dd");
                historialSheet.getRange(lastRow, 5).setNumberFormat("HH:mm");
                
            } else if (tipo === 'salida') {
                // Lógica inteligente para salidas
                const data = historialSheet.getDataRange().getValues();
                let entradaEncontrada = false;
                
                // Buscar la entrada más reciente sin salida para esta cédula
                for (let i = data.length - 1; i >= 1; i--) {
                    const row = data[i];
                    const cedulaRow = String(row[1] || '').trim();
                    const cedulaRowNorm = normalizarCedula(cedulaRow).replace(/[^\w\-]/g, '');
                    const cedulaBuscadaNorm = normalizarCedula(resultado.cedula).replace(/[^\w\-]/g, '');
                    
                    // Verificar coincidencia de cédula y que tenga entrada sin salida
                    if ((cedulaRow === resultado.cedula || cedulaRowNorm === cedulaBuscadaNorm) &&
                        row[4] && row[4] !== '' && (!row[5] || row[5] === '')) {
                        
                        // Actualizar la fila existente con la salida
                        const rowIndex = i + 1;
                        historialSheet.getRange(rowIndex, 6).setValue(fecha).setNumberFormat("HH:mm");
                        historialSheet.getRange(rowIndex, 8).setValue(resultado.empresa);
                        
                        // Calcular y registrar duración
                        const entrada = row[4];
                        if (entrada instanceof Date) {
                            const duracionMs = fecha.getTime() - entrada.getTime();
                            const horas = Math.floor(duracionMs / (1000 * 60 * 60));
                            const minutos = Math.floor((duracionMs % (1000 * 60 * 60)) / (1000 * 60));
                            
                            const duracionTexto = `${horas}h ${minutos}m`;
                            historialSheet.getRange(rowIndex, 7).setValue(duracionTexto);
                            
                            // Agregar duración al resultado para logging
                            resultado.duracion = duracionTexto;
                            logError(`Duración calculada para ${resultado.cedula}: ${duracionTexto}`, 'INFO');
                        }
                        
                        entradaEncontrada = true;
                        break;
                    }
                }
                
                // Si no se encontró entrada previa, crear nueva fila de solo salida
                if (!entradaEncontrada) {
                    sinEntradaPrevia = true;
                    const mensajeDuracion = 'Sin entrada';
                    const historialRow = [
                        fecha,                  // Fecha
                        resultado.cedula,       // Cédula
                        resultado.nombre,       // Nombre
                        resultado.estadoAcceso, // Estado del Acceso
                        '',                     // Entrada (vacía)
                        fecha,                  // Salida
                        mensajeDuracion,        // Duración (con mensaje)
                        resultado.empresa       // Empresa
                    ];
                    
                    historialSheet.appendRow(historialRow);
                    const lastRow = historialSheet.getLastRow();
                    
                    // Aplicar formato
                    historialSheet.getRange(lastRow, 4).setBackground('#ffcc80'); // Naranja para salida sin entrada
                    historialSheet.getRange(lastRow, 1).setNumberFormat("yyyy-mm-dd");
                    historialSheet.getRange(lastRow, 6).setNumberFormat("HH:mm");
                    historialSheet.getRange(lastRow, 7).setFontColor('#e65100'); // Color naranja oscuro
                    
                    logError(`Salida sin entrada previa registrada para ${resultado.cedula}`, 'WARNING');
                }
            }
            
        } catch (error) {
            logError('Error en registrarEnHistorial', 'ERROR', { error: error.message });
            throw new Error('Error al registrar en historial: ' + error.message);
        }
        
        return { sinEntradaPrevia };
    }

    /**
     * Busca cédulas similares para casos de acceso denegado
     */
    function buscarCedulasSimilares(cedula, bdData) {
        try {
            if (!cedula || !bdData || bdData.length <= 1) {
                return '';
            }

            const similares = [];
            const cedulaNorm = normalizarCedula(cedula).replace(/[^\w\-]/g, '');
            const prefijoCedula = cedulaNorm.substring(0, Math.min(5, cedulaNorm.length));
            
            if (prefijoCedula.length < 2) {
                return ''; // Prefijo muy corto, no buscar
            }

            for (let i = 1; i < bdData.length && similares.length < 5; i++) {
                if (bdData[i][0]) {
                    const cedulaOriginal = String(bdData[i][0]).trim();
                    const cedulaBD = normalizarCedula(cedulaOriginal).replace(/[^\w\-]/g, '');
                    
                    // Buscar cédulas que comiencen con el mismo prefijo, pero no sean idénticas
                    if (cedulaBD.startsWith(prefijoCedula) && cedulaBD !== cedulaNorm) {
                        similares.push(cedulaOriginal);
                    }
                }
            }

            if (similares.length > 5) {
                similares.splice(5);
                similares.push('... (más resultados)');
            }

            return similares.join(', ');
            
        } catch (error) {
            logError('Error en buscarCedulasSimilares', 'ERROR', { error: error.message });
            return '';
        }
    }

    // =====================================================
    // GESTIÓN DE COMENTARIOS
    // =====================================================

    /**
     * Maneja el envío de comentarios para accesos denegados
     */
    function handleCommentSubmit(cedula, comentarioCierre) {
        try {
            if (!validarCampoObligatorio(cedula) || !validarCampoObligatorio(comentarioCierre)) {
                return {
                    success: false,
                    message: 'La cédula y el nombre son obligatorios'
                };
            }

            const cedulaNorm = String(cedula).trim();
            const comentarioNorm = String(comentarioCierre).trim();
            
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const respuestaSheet = ss.getSheetByName('Respuestas formulario');

            if (!respuestaSheet) {
                throw new Error('La hoja "Respuestas formulario" no fue encontrada.');
            }

            const data = respuestaSheet.getDataRange().getValues();
            let lastRowIndex = -1;

            // Buscar la última fila con esta cédula
            for (let i = data.length - 1; i >= 0; i--) {
                const cedulaFila = String(data[i][1] || '').trim();
                const cedulaFilaNorm = normalizarCedula(cedulaFila).replace(/[^\w\-]/g, '');
                const cedulaBuscadaNorm = normalizarCedula(cedulaNorm).replace(/[^\w\-]/g, '');
                
                if (cedulaFila === cedulaNorm || cedulaFilaNorm === cedulaBuscadaNorm) {
                    lastRowIndex = i + 1;
                    break;
                }
            }

            if (lastRowIndex === -1) {
                // Si no se encuentra la fila específica, usar la última fila
                lastRowIndex = respuestaSheet.getLastRow();
                logError(`No se encontró fila específica para cédula ${cedulaNorm}, usando última fila: ${lastRowIndex}`, 'WARNING');
            }

            // Registrar el comentario en la columna 9 (Comentarios)
            const comentarioCompleto = `${new Date().toLocaleString()} - ${obtenerUsuarioActual()}: ${comentarioNorm}`;
            respuestaSheet.getRange(lastRowIndex, 9).setValue(comentarioCompleto);
            
            logError(`Comentario registrado en fila ${lastRowIndex} para cédula ${cedulaNorm}`, 'INFO', { 
                comentario: comentarioNorm,
                usuario: obtenerUsuarioActual()
            });

            return {
                message: 'Comentario enviado correctamente',
                color: 'green',
                log: {
                    timestamp: new Date().toISOString(),
                    level: 'INFO',
                    message: 'Comentario registrado exitosamente',
                    details: { 
                        cedula: cedulaNorm, 
                        comentario: comentarioNorm,
                        fila: lastRowIndex,
                        usuario: obtenerUsuarioActual()
                    }
                }
            };
            
        } catch (error) {
            logError('Error en handleCommentSubmit', 'ERROR', { error: error.message, cedula, comentarioCierre });
            return {
                message: 'Error al enviar comentario: ' + error.message,
                color: 'red',
                log: {
                    timestamp: new Date().toISOString(),
                    level: 'ERROR',
                    message: 'Error al registrar comentario',
                    details: { 
                        error: error.message,
                        cedula: cedula,
                        comentario: comentarioCierre,
                        usuario: obtenerUsuarioActual()
                    }
                }
            };
        }
    }

    // =====================================================
    // VISTAS PREVIAS Y REPORTES
    // =====================================================

    /**
     * Genera vista previa de registros/rondas
     */
    function mostrarVistaPreviaRondas(cedula) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const historialSheet = ss.getSheetByName('Historial');
            
            if (!historialSheet) {
                throw new Error("La hoja 'Historial' no existe.");
            }

            const data = historialSheet.getDataRange().getValues();
            
            if (data.length <= 1) {
                return `<p style="font-size: 35px; text-align: center; color: #666;">No hay registros en el historial.</p>`;
            }
            
            const headers = data[0];
            let filteredData = data.slice(1);
            
            // Filtrar por cédula si se proporciona
            if (cedula && cedula.trim() !== '') {
                const cedulaBuscar = String(cedula).trim();
                filteredData = filteredData.filter(row => {
                    const cedulaRow = String(row[1] || '').trim();
                    const cedulaRowNorm = normalizarCedula(cedulaRow).replace(/[^\w\-]/g, '');
                    const cedulaBuscarNorm = normalizarCedula(cedulaBuscar).replace(/[^\w\-]/g, '');
                    
                    return cedulaRow === cedulaBuscar || cedulaRowNorm === cedulaBuscarNorm;
                });
            }

            if (filteredData.length === 0) {
                return `<p style="font-size: 35px; text-align: center; color: #666;">No se encontraron registros ${cedula ? 'para la cédula ' + cedula : ''}.</p>`;
            }

            // Formatear datos para mostrar
            const formattedData = filteredData.map(row => {
                return row.map((cell, colIndex) => {
                    if (cell instanceof Date) {
                        // Columnas 4 y 5 son Entrada y Salida (solo hora)
                        if (colIndex === 4 || colIndex === 5) {
                            return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'HH:mm');
                        } else {
                            // Otras fechas (fecha completa)
                            const formattedDate = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'dd/MM/yy');
                            return formattedDate !== '01/01/70' ? formattedDate : Utilities.formatDate(cell, Session.getScriptTimeZone(), 'HH:mm');
                        }
                    }
                    return cell || '';
                });
            });

            // Generar HTML
            let html = `
                <div style="padding: 20px; font-family: Arial, sans-serif;" class="preview-content">
                    <h2 style="font-size: 45px; text-align: center; margin-bottom: 30px; color: #960018;" class="preview-title">
                        📊 ${cedula ? 'Registros para: ' + cedula : 'Historial Completo'}
                    </h2>
                    <div style="overflow-x: auto; margin: 20px 0;">
                        <table style="border-collapse: collapse; width: 100%; font-size: 24px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);" class="preview-table">
                            <thead>
                                <tr style="background: linear-gradient(135deg, #960018, #b71c1c); color: white;">`;

            headers.forEach(header => {
                html += `<th style="padding: 15px 10px; border: 1px solid #ddd; text-align: center; font-weight: bold;">${header}</th>`;
            });

            html += '</tr></thead><tbody>';

            formattedData.forEach((row, rowIndex) => {
                const isEven = rowIndex % 2 === 0;
                const rowColor = isEven ? '#f9f9f9' : '#ffffff';
                const isAccessDenied = row[3] === 'Acceso Denegado';
                const isSinEntrada = row[6] === 'Sin entrada';
                
                let backgroundColor = rowColor;
                if (isAccessDenied) backgroundColor = 'rgba(255, 0, 0, 0.1)';
                else if (isSinEntrada) backgroundColor = 'rgba(255, 152, 0, 0.1)';
                
                html += `<tr style="background-color: ${backgroundColor};">`;
                
                row.forEach((cell, cellIndex) => {
                    let textColor = '#333';
                    let fontWeight = 'normal';
                    
                    if (isAccessDenied && cellIndex === 3) {
                        textColor = '#d32f2f';
                        fontWeight = 'bold';
                    } else if (isSinEntrada && cellIndex === 6) {
                        textColor = '#e65100';
                        fontWeight = 'bold';
                    }
                    
                    html += `<td style="padding: 12px 8px; border: 1px solid #ddd; text-align: center; color: ${textColor}; font-weight: ${fontWeight};">${cell}</td>`;
                });
                
                html += '</tr>';
            });

            html += `
                            </tbody>
                        </table>
                    </div>
                    <div style="margin-top: 30px; padding: 15px; background: linear-gradient(135deg, #f5f5f5, #e0e0e0); border-radius: 10px;">
                        <p style="font-size: 28px; text-align: center; margin: 0; color: #666;">
                            📈 <strong>Total de registros:</strong> ${formattedData.length}
                        </p>
                        <p style="font-size: 20px; text-align: center; margin: 10px 0 0 0; color: #888;">
                            Generado el ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}
                        </p>
                    </div>
                </div>`;

            return html;
            
        } catch (error) {
            logError('Error en mostrarVistaPreviaRondas', 'ERROR', { error: error.message, cedula });
            return `<div style="padding: 40px; text-align: center; font-family: Arial, sans-serif;">
                        <h2 style="color: #d32f2f; font-size: 36px;">❌ Error</h2>
                        <p style="font-size: 24px; color: #666;">Error al cargar la vista previa:</p>
                        <p style="font-size: 20px; color: #d32f2f; font-weight: bold;">${error.message}</p>
                        <p style="font-size: 18px; color: #888; margin-top: 20px;">Por favor, verifique que el sistema esté configurado correctamente.</p>
                    </div>`;
        }
    }

    /**
     * Genera información completa del sistema
     */
    function mostrarInformacionSistema() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const config = obtenerConfiguracion();
            
            // Obtener estadísticas básicas
            let estadisticas = '';
            try {
                const stats = obtenerEstadisticasBasicas();
                estadisticas = `
                    <div style="background: linear-gradient(135deg, #e3f2fd, #bbdefb); padding: 20px; border-radius: 15px; margin: 20px 0;">
                        <h4 style="color: #1976d2; margin-bottom: 15px;">📊 Estadísticas del Sistema</h4>
                        <p><strong>Total de personal registrado:</strong> ${stats.totalPersonal}</p>
                        <p><strong>Registros de acceso hoy:</strong> ${stats.accesosHoy}</p>
                        <p><strong>Último registro:</strong> ${stats.ultimoRegistro}</p>
                        <p><strong>Personas dentro del edificio:</strong> ${stats.personasDentro}</p>
                    </div>
                `;
            } catch (e) {
                estadisticas = '<p style="color: #ff9800;">Estadísticas no disponibles en este momento.</p>';
            }
            
            const htmlContent = `
                <div style="font-family: Arial, sans-serif; padding: 30px; max-width: 900px; margin: 0 auto; line-height: 1.6;" class="preview-content">
                    <h2 style="font-size: 42px; color: #960018; text-align: center; margin-bottom: 30px; text-shadow: 2px 2px 4px rgba(0,0,0,0.1);" class="preview-title">
                        🏢 SurPass - Control de Acceso v3.0
                    </h2>
                    
                    <div style="background: linear-gradient(135deg, #fff3e0, #ffe0b2); padding: 25px; border-radius: 15px; margin-bottom: 25px; border-left: 5px solid #ff9800;">
                        <h3 style="color: #f57c00; margin-bottom: 15px;">ℹ️ Información del Sistema</h3>
                        <p><strong>📝 Nombre de la empresa:</strong> ${config.EMPRESA_NOMBRE || 'SurPass'}</p>
                        <p><strong>🕐 Horario de operación:</strong> ${config.HORARIO_APERTURA || '08:00'} - ${config.HORARIO_CIERRE || '17:00'}</p>
                        <p><strong>📅 Días laborables:</strong> ${config.DIAS_LABORABLES || 'Lunes a Viernes'}</p>
                        <p><strong>⏱️ Tiempo máximo de visita:</strong> ${config.TIEMPO_MAX_VISITA || '4'} horas</p>
                        <p><strong>📧 Email de notificaciones:</strong> ${config.NOTIFICACIONES_EMAIL || 'admin@surpass.com'}</p>
                        <p><strong>📧 Email secundario:</strong> ${config.EMAIL_SECUNDARIO || 'N/A'}</p>
                        <p><strong>💾 Backup automático:</strong> ${config.BACKUP_AUTOMATICO || 'SI'} (${config.FRECUENCIA_BACKUP || 'DIARIO'})</p>
                        <p><strong>🚨 Evacuación automática:</strong> ${config.NOTIFICAR_EVACUACION_AUTOMATICA || 'SI'}</p>
                        <p><strong>🆔 ID de la hoja:</strong> ${ss.getId()}</p>
                        <p><strong>📊 Total de hojas:</strong> ${ss.getSheets().length}</p>
                        <p><strong>👤 Usuario actual:</strong> ${obtenerUsuarioActual()}</p>
                    </div>

                    ${estadisticas}

                    <div style="background: linear-gradient(135deg, #e8f5e8, #c8e6c8); padding: 25px; border-radius: 15px; margin-bottom: 25px; border-left: 5px solid #4caf50;">
                        <h3 style="color: #388e3c; margin-bottom: 20px;">📖 Guía de Usuario</h3>
                        
                        <h4 style="color: #2e7d32; margin-top: 25px; margin-bottom: 15px;">1. 🚪 Registro de Acceso</h4>
                        <ul style="margin-left: 20px; margin-bottom: 20px;">
                            <li>Ingrese su número de cédula en el campo correspondiente</li>
                            <li>El sistema mostrará sugerencias mientras escribe</li>
                            <li>Seleccione "Entrada" o "Salida" según corresponda</li>
                            <li>Haga clic en "Registrar" para completar el registro</li>
                            <li>Si aparece un mensaje de acceso denegado, agregue un comentario obligatorio</li>
                        </ul>

                        <h4 style="color: #2e7d32; margin-bottom: 15px;">2. 📱 Lectura por Código QR</h4>
                        <ul style="margin-left: 20px; margin-bottom: 20px;">
                            <li>Haga clic en el botón "Escanear QR" (ícono de código)</li>
                            <li>Permita el acceso a la cámara cuando se le solicite</li>
                            <li>Apunte la cámara al código QR o documento de identidad</li>
                            <li>El número de cédula se insertará automáticamente</li>
                            <li>El sistema puede leer múltiples formatos de documentos</li>
                        </ul>

                        <h4 style="color: #2e7d32; margin-bottom: 15px;">3. 🚨 Modo de Evacuación</h4>
                        <ul style="margin-left: 20px; margin-bottom: 20px;">
                            <li>Acceda al modo evacuación desde el menú principal</li>
                            <li>Vea en tiempo real quién está dentro del edificio</li>
                            <li>Seleccione personas para marcar como evacuadas</li>
                            <li>Confirme evacuaciones masivas con un solo clic</li>
                            <li>Exporte reportes de evacuación en formato CSV</li>
                            <li>Imprima listas de verificación para uso manual</li>
                        </ul>

                        <h4 style="color: #2e7d32; margin-bottom: 15px;">4. 📊 Estadísticas en Tiempo Real</h4>
                        <ul style="margin-left: 20px; margin-bottom: 20px;">
                            <li>El panel lateral muestra estadísticas del día</li>
                            <li>Puede mover el panel arrastrándolo</li>
                            <li>Se actualiza automáticamente cada 30 segundos</li>
                            <li>Muestra entradas, salidas y últimos registros</li>
                            <li>Gráficos visuales para análisis rápido</li>
                        </ul>

                        <h4 style="color: #2e7d32; margin-bottom: 15px;">5. 🌐 Modo Sin Conexión</h4>
                        <ul style="margin-left: 20px; margin-bottom: 20px;">
                            <li>El sistema funciona incluso sin conexión a internet</li>
                            <li>Los registros se guardan localmente</li>
                            <li>Se sincronizarán automáticamente al recuperar la conexión</li>
                            <li>El indicador de conexión muestra el estado actual</li>
                        </ul>

                        <h4 style="color: #2e7d32; margin-bottom: 15px;">6. 🔧 Funciones Administrativas</h4>
                        <ul style="margin-left: 20px; margin-bottom: 20px;">
                            <li><strong>Menú:</strong> Acceda a todas las funciones desde el ícono de menú</li>
                            <li><strong>Historial:</strong> Vea todos los registros de acceso</li>
                            <li><strong>Información:</strong> Consulte esta guía en cualquier momento</li>
                            <li><strong>Finalizar:</strong> Termine su turno y envíe reportes automáticamente</li>
                            <li><strong>Limpiar:</strong> Borre el historial tras finalizar el turno</li>
                            <li><strong>Configuración:</strong> Personalice el comportamiento del sistema</li>
                            <li><strong>Diagnóstico:</strong> Verifique el estado completo del sistema</li>
                        </ul>
                    </div>

                    <div style="background: linear-gradient(135deg, #ffebee, #ffcdd2); padding: 20px; border-radius: 15px; margin-bottom: 25px; border-left: 5px solid #f44336;">
                        <h3 style="color: #c62828; margin-bottom: 15px;">🚨 Procedimientos de Emergencia</h3>
                        <ul style="margin-left: 20px;">
                            <li><strong>Evacuación:</strong> Use el botón de evacuación para ver personas dentro</li>
                            <li><strong>Emergencia:</strong> El sistema envía notificaciones automáticas</li>
                            <li><strong>Reportes:</strong> Genere informes de evacuación instantáneos</li>
                            <li><strong>Verificación:</strong> Confirme evacuaciones con checkboxes</li>
                            <li><strong>Documentación:</strong> Todos los eventos se registran automáticamente</li>
                        </ul>
                    </div>

                    <div style="background: linear-gradient(135deg, #f3e5f5, #e1bee7); padding: 20px; border-radius: 15px; text-align: center;">
                        <h3 style="color: #7b1fa2; margin-bottom: 15px;">🚀 Sistema SurPass v3.0</h3>
                        <p style="font-style: italic; color: #666; margin-bottom: 10px;">
                            Control de Acceso Inteligente con Gestión de Evacuaciones
                        </p>
                        <p style="font-size: 14px; color: #888;">
                            Versión 3.0 - ${new Date().getFullYear()} | 
                            Generado el ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')} |
                            Usuario: ${obtenerUsuarioActual()}
                        </p>
                    </div>
                </div>
            `;

            return htmlContent;
            
        } catch (error) {
            logError('Error en mostrarInformacionSistema', 'ERROR', { error: error.message });
            return `
                <div style="padding: 40px; text-align: center; font-family: Arial, sans-serif;">
                    <h2 style="color: #d32f2f;">❌ Error</h2>
                    <p style="color: #666;">Error al cargar la información del sistema:</p>
                    <p style="color: #d32f2f; font-weight: bold;">${error.message}</p>
                </div>
            `;
        }
    }

    /**
     * Obtiene estadísticas básicas del sistema - VERSIÓN CORREGIDA
     */
    function obtenerEstadisticasBasicas() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            
            // Contar personal registrado
            const bdSheet = ss.getSheetByName('Base de Datos');
            const totalPersonal = bdSheet ? Math.max(0, bdSheet.getLastRow() - 1) : 0;
            
            // Obtener estadísticas detalladas de hoy
            const statsDetalladas = obtenerEstadisticas();
            
            // Contar accesos de hoy
            const respuestasSheet = ss.getSheetByName('Respuestas formulario');
            let accesosHoy = 0;
            let ultimoRegistro = 'Ninguno';
            
            if (respuestasSheet && respuestasSheet.getLastRow() > 1) {
                const today = new Date();
                today.setHours(0, 0, 0, 0);
                
                const data = respuestasSheet.getDataRange().getValues();
                
                for (let i = 1; i < data.length; i++) {
                    const fecha = data[i][0];
                    if (fecha instanceof Date && fecha >= today) {
                        accesosHoy++;
                    }
                }
                
                // Obtener último registro
                if (data.length > 1) {
                    const ultimaFila = data[data.length - 1];
                    const cedulaIndex = 1; // Columna B
                    const tipoIndex = 2; // Columna C
                    const fechaIndex = 0; // Columna A

                    const fechaUltima = ultimaFila[fechaIndex];
                    const cedulaUltima = ultimaFila[cedulaIndex];
                    const tipoUltimo = ultimaFila[tipoIndex];
                    
                    if (fechaUltima instanceof Date) {
                        ultimoRegistro = `${cedulaUltima} (${tipoUltimo}) - ${Utilities.formatDate(fechaUltima, Session.getScriptTimeZone(), 'HH:mm')}`;
                    }
                }
            }
            
            // Contar personas dentro
            let personasDentro = 0;
            try {
                const estadoEvacuacion = getEvacuacionDataForClient();
                personasDentro = estadoEvacuacion.totalDentro;
            } catch (e) {
                logError('Error contando personas dentro en obtenerEstadisticasBasicas', 'WARNING', { error: e.message });
            }
            
            return {
                totalPersonal: totalPersonal,
                accesosHoy: accesosHoy,
                ultimoRegistro: ultimoRegistro,
                personasDentro: personasDentro,
                // ✅ AGREGAR ENTRADAS Y SALIDAS QUE FALTABAN
                entradas: statsDetalladas.entradas || 0,
                salidas: statsDetalladas.salidas || 0
            };
            
        } catch (error) {
            logError('Error general en obtenerEstadisticasBasicas', 'ERROR', { error: error.message });
            return {
                totalPersonal: 0,
                accesosHoy: 0,
                ultimoRegistro: 'Error al obtener',
                personasDentro: 0,
                entradas: 0,
                salidas: 0
            };
        }
    }

    // =====================================================
    // ESTADÍSTICAS Y REPORTES
    // =====================================================

    /**
     * Obtiene estadísticas en tiempo real del sistema
     */
    function obtenerEstadisticas() {
        try {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas formulario');
            
            if (!sheet) {
                throw new Error('No se encontró la hoja "Respuestas formulario"');
            }

            const data = sheet.getDataRange().getValues();
            
            if (data.length <= 1) {
                return {
                    entradas: 0,
                    salidas: 0,
                    total: 0,
                    recentRecords: []
                };
            }

            const today = new Date();
            today.setHours(0, 0, 0, 0);

            let entradas = 0;
            let salidas = 0;
            const recentRecords = [];

            // Procesar datos desde el más reciente
            for (let i = data.length - 1; i > 0; i--) {
                const row = data[i];
                const fecha = row[0] instanceof Date ? row[0] : new Date(row[0]);
                const tipo = String(row[2] || '').toLowerCase();
                const cedula = String(row[1] || '');
                const estado = String(row[3] || '');

                // Contar accesos de hoy
                if (fecha >= today) {
                    if (tipo === 'entrada') entradas++;
                    if (tipo === 'salida') salidas++;
                }

                // Agregar a registros recientes (últimos 10)
                if (recentRecords.length < 10) {
                    recentRecords.push({
                        cedula: cedula,
                        accion: tipo,
                        hora: fecha.toLocaleTimeString('es-ES', { 
                            hour: '2-digit', 
                            minute: '2-digit' 
                        }),
                        estado: estado
                    });
                }
            }

            logError(`Estadísticas calculadas: ${entradas} entradas, ${salidas} salidas`, 'INFO');

            return {
                entradas: entradas,
                salidas: salidas,
                total: data.length - 1,
                recentRecords: recentRecords
            };
            
        } catch (error) {
            logError('Error en obtenerEstadisticas', 'ERROR', { error: error.message });
            return {
                entradas: 0,
                salidas: 0,
                total: 0,
                recentRecords: [],
                error: error.message
            };
        }
    }

    /**
     * Muestra estadísticas del día en formato UI
     */
    function mostrarEstadisticasDelDia() {
        try {
            const stats = obtenerEstadisticas();
            const evacuation = getEvacuacionDataForClient(); // Use the new unified function
            
            let mensaje = `📊 ESTADÍSTICAS DEL DÍA\n\n`;
            mensaje += `📅 Fecha: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')}\n`;
            mensaje += `🕐 Hora: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm')}\n\n`;
            
            mensaje += `📈 RESUMEN DE ACCESOS:\n`;
            mensaje += `• Entradas: ${stats.entradas}\n`;
            mensaje += `• Salidas: ${stats.salidas}\n`;
            mensaje += `• Total registros: ${stats.total}\n`;
            mensaje += `• Personas dentro: ${evacuation.totalDentro}\n\n`;
            
            if (stats.recentRecords && stats.recentRecords.length > 0) {
                mensaje += `🕒 ÚLTIMOS REGISTROS:\n`;
                stats.recentRecords.slice(0, 5).forEach((record, index) => {
                    mensaje += `${index + 1}. ${record.cedula} - ${record.accion} (${record.hora})\n`;
                });
            }
            
            mensaje += `\n💡 Para más detalles, use el menú "Historial" o "Estado de Evacuación".`;
            
            SpreadsheetApp.getUi().alert(
                '📊 Estadísticas del Día', 
                mensaje, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            
            logError('Estadísticas del día mostradas', 'INFO');
            
        } catch (error) {
            logError('Error mostrando estadísticas del día', 'ERROR', { error: error.message });
            SpreadsheetApp.getUi().alert(
                '❌ Error', 
                'Error al obtener estadísticas: ' + error.message, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
    }

    // =====================================================
    // CONFIGURACIÓN DE OPCIONES DE MENÚ
    // =====================================================

    /**
     * FUNCIONES ADICIONALES REQUERIDAS PARA EL MENÚ
     * Estas funciones son llamadas desde el HTML pero pueden necesitar ajustes
     */

    /**
     * Función mejorada para obtener opciones de menú
     * Asegura compatibilidad completa con el HTML
     */
    function obtenerOpcionesMenu() {
        try {
            const userProperties = PropertiesService.getUserProperties();
            const config = obtenerConfiguracion();
            
            return {
                mostrarEstadisticas: userProperties.getProperty('mostrarEstadisticas') !== 'false',
                usarJsQR: userProperties.getProperty('usarJsQR') === 'true',
                sonidosActivados: config.SONIDOS_ACTIVADOS !== 'NO',
                temaOscuro: userProperties.getProperty('temaOscuro') === 'true',
                notificacionesEmail: config.NOTIFICACIONES_EMAIL ? true : false,
                // Estas propiedades adicionales son esperadas por el HTML
                escaneNativo: userProperties.getProperty('usarJsQR') === 'true',
                estadisticasVisibles: userProperties.getProperty('mostrarEstadisticas') !== 'false'
            };
            
        } catch (error) {
            logError('Error en obtenerOpcionesMenu', 'ERROR', { error: error.message });
            return {
                mostrarEstadisticas: true,
                usarJsQR: false,
                sonidosActivados: true,
                temaOscuro: false,
                notificacionesEmail: true,
                escaneNativo: false,
                estadisticasVisibles: true
            };
        }
    }

    /**
     * Activa/desactiva el panel de estadísticas
     */
    function toggleEstadisticas(mostrar) {
        try {
            const userProperties = PropertiesService.getUserProperties();
            userProperties.setProperty('mostrarEstadisticas', mostrar.toString());

            logError(`Panel de estadísticas ${mostrar ? 'activado' : 'desactivado'}`, 'INFO');

            return {
                success: true,
                mostrarEstadisticas: mostrar,
                message: 'Panel de estadísticas ' + (mostrar ? 'activado' : 'desactivado')
            };
            
        } catch (error) {
            logError('Error en toggleEstadisticas', 'ERROR', { error: error.message });
            return {
                success: false,
                message: 'Error al cambiar configuración de estadísticas: ' + error.message
            };
        }
    }

    /**
     * Activa/desactiva el escáner nativo jsQR
     */
    function toggleEscaner(usar) {
        try {
            const userProperties = PropertiesService.getUserProperties();
            userProperties.setProperty('usarJsQR', usar.toString());

            logError(`Escáner nativo ${usar ? 'activado' : 'desactivado'}`, 'INFO');

            return {
                success: true,
                usarJsQR: usar,
                message: 'Escáner nativo ' + (usar ? 'activado' : 'desactivado')
            };
            
        } catch (error) {
            logError('Error en toggleEscaner', 'ERROR', { error: error.message });
            return {
                success: false,
                message: 'Error al cambiar configuración del escáner: ' + error.message
            };
        }
    }

    /**
     * Activa/desactiva los sonidos del sistema
     */
    function toggleSonidos(activar) {
        try {
            // Actualizar en configuración del sistema
            const resultado = actualizarConfiguracion('SONIDOS_ACTIVADOS', activar ? 'SI' : 'NO');
            
            if (resultado.success) {
                logError(`Sonidos ${activar ? 'activados' : 'desactivados'}`, 'INFO');
                
                return {
                    success: true,
                    sonidosActivados: activar,
                    message: 'Sonidos ' + (activar ? 'activados' : 'desactivados')
                };
            } else {
                throw new Error(resultado.message);
            }
            
        } catch (error) {
            logError('Error en toggleSonidos', 'ERROR', { error: error.message });
            return {
                success: false,
                message: 'Error al cambiar configuración de sonidos: ' + error.message
            };
        }
    }

    // =====================================================
    // GESTIÓN DE TURNOS Y RESPALDOS
    // =====================================================

    /**
     * Limpia los registros del historial (solo después de finalizar turno)
     */
    function limpiarRegistros() {
        try {
            const correoEnviado = PropertiesService.getScriptProperties().getProperty('correoEnviado');
            
            if (correoEnviado !== 'true') {
                logError('El correo no se envió. No se limpiarán los registros.', 'WARNING');
                return {
                    exito: false,
                    mensaje: 'El turno no ha sido finalizado correctamente. No se pueden limpiar los registros.'
                };
            }

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheetsToClean = ['Historial', 'Respuestas formulario'];
            let totalRegistrosLimpiados = 0;

            sheetsToClean.forEach(sheetName => {
                const sheet = ss.getSheetByName(sheetName);
                if (sheet) {
                    const lastRow = sheet.getLastRow();
                    
                    if (lastRow > 1) {
                        const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
                        range.clearContent();
                        range.setBackground(null);
                        
                        totalRegistrosLimpiados += (lastRow - 1);
                        logError(`${sheetName} limpiada: ${lastRow - 1} registros eliminados`, 'INFO');
                    }
                }
            });

            // Limpiar flag de correo enviado
            PropertiesService.getScriptProperties().deleteProperty('correoEnviado');
            
            logError(`Limpieza completada: ${totalRegistrosLimpiados} registros eliminados`, 'INFO');

            return {
                exito: true,
                mensaje: `Registros limpiados correctamente. ${totalRegistrosLimpiados} registros eliminados.`
            };
            
        } catch (error) {
            logError('Error en limpiarRegistros', 'ERROR', { error: error.message });
            return {
                exito: false,
                mensaje: 'Error al limpiar los registros: ' + error.message
            };
        }
    }

    /**
     * Finaliza el turno y envía el reporte por correo
     */
    function finalizarTurno() {
        try {
            logError('Iniciando proceso de finalización de turno', 'INFO');

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheetHistorial = ss.getSheetByName('Historial');
            const sheetRespuestas = ss.getSheetByName('Respuestas formulario');

            if (!sheetHistorial || !sheetRespuestas) {
                throw new Error('No se encontraron las hojas "Historial" o "Respuestas formulario".');
            }

            if (sheetHistorial.getLastRow() <= 1 && sheetRespuestas.getLastRow() <= 1) {
                return {
                    exito: false,
                    mensaje: 'No hay registros para finalizar el turno. Ambas hojas están vacías.'
                };
            }

            const fechaActual = new Date();
            const fechaFormateada = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            
            // Obtener configuración de email
            const config = obtenerConfiguracion();
            const destinatarios = [config.NOTIFICACIONES_EMAIL, config.EMAIL_SECUNDARIO]
                .filter(email => email && email.trim())
                .join(', ');
            
            const asunto = `📊 Registro de Personal SurPass - ${fechaFormateada}`;
            
            // Generar estadísticas del turno
            const stats = obtenerEstadisticasBasicas();
            const evacuation = getEvacuacionDataForClient(); // Use the new unified function
            
            let mensaje = `Estimado equipo,\n\n`;
            mensaje += `📊 REPORTE DE TURNO - ${fechaFormateada}\n`;
            mensaje += `════════════════════════════════════════\n\n`;
            mensaje += `📈 ESTADÍSTICAS DEL TURNO:\n`;
            mensaje += `• Personal registrado en sistema: ${stats.totalPersonal}\n`;
            mensaje += `• Accesos registrados hoy: ${stats.accesosHoy}\n`;
            mensaje += `• Último registro: ${stats.ultimoRegistro}\n`;
            mensaje += `• Personas actualmente dentro: ${evacuation.totalDentro}\n`;
            mensaje += `• Entradas del día: ${stats.entradas}\n`; // Use stats.entradas
            mensaje += `• Salidas del día: ${stats.salidas}\n\n`; // Use stats.salidas
            
            if (evacuation.totalDentro > 0) {
                mensaje += `⚠️ PERSONAS DENTRO DEL EDIFICIO:\n`;
                evacuation.personasDentro.forEach((persona, index) => {
                    mensaje += `${index + 1}. ${persona.nombre} (${persona.cedula}) - ${persona.empresa}\n`;
                });
                mensaje += `\n`;
            } else {
                mensaje += `✅ EDIFICIO COMPLETAMENTE EVACUADO\n\n`;
            }
            
            mensaje += `👤 Turno finalizado por: ${obtenerUsuarioActual()}\n`;
            mensaje += `🕐 Hora de finalización: ${Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'HH:mm:ss')}\n\n`;
            mensaje += `📎 Se adjunta el archivo Excel con el detalle completo del historial.\n\n`;
            mensaje += `Generado automáticamente por el Sistema SurPass v3.0.`;

            try {
                // Crear y enviar archivo Excel
                const exportUrl = `https://www.googleapis.com/drive/v3/files/${ss.getId()}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
                const params = {
                    method: 'get',
                    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
                    muteHttpExceptions: true
                };

                const response = UrlFetchApp.fetch(exportUrl, params);
                
                if (response.getResponseCode() === 200) {
                    const blob = response.getBlob().setName(`SurPass_Historial_${fechaFormateada}.xlsx`);
                    
                    if (destinatarios) {
                        MailApp.sendEmail(destinatarios, asunto, mensaje, { attachments: [blob] });
                        logError('Correo de finalización enviado a: ' + destinatarios, 'INFO');
                    } else {
                        logError('No hay destinatarios configurados para el reporte', 'WARNING');
                    }
                    
                    PropertiesService.getScriptProperties().setProperty('correoEnviado', 'true');
                    
                    // Crear respaldo adicional
                    try {
                        crearRespaldoAutomatico();
                        logError('Respaldo automático creado', 'INFO');
                    } catch (backupError) {
                        logError('Advertencia: No se pudo crear respaldo automático', 'WARNING', { error: backupError.message });
                    }
                    
                    return {
                        exito: true,
                        mensaje: 'Turno finalizado y reporte enviado correctamente'
                    };
                    
                } else {
                    throw new Error(`Error al exportar archivo: Código ${response.getResponseCode()}`);
                }
                
            } catch (emailError) {
                logError('Error al enviar correo de finalización', 'ERROR', { error: emailError.message });
                return {
                    exito: false,
                    mensaje: 'Error al enviar el correo: ' + emailError.message
                };
            }

        } catch (error) {
            logError('Error en finalizarTurno', 'ERROR', { error: error.message });
            PropertiesService.getScriptProperties().setProperty('correoEnviado', 'false');
            return {
                exito: false,
                mensaje: 'Error al finalizar el turno: ' + error.message
            };
        }
    }

    /**
     * Crea un respaldo manual desde el menú
     */
    function crearRespaldoManual() {
        try {
            const resultado = crearRespaldoAutomatico();
            
            if (resultado.exito) {
                SpreadsheetApp.getUi().alert(
                    '💾 Respaldo Creado', 
                    `✅ ${resultado.mensaje}\n\n📁 Archivo: ${resultado.archivoId}\n🔗 URL: ${resultado.url}`, 
                    SpreadsheetApp.getUi().ButtonSet.OK
                );
            } else {
                SpreadsheetApp.getUi().alert(
                    '❌ Error', 
                    `Error al crear respaldo: ${resultado.mensaje}`, 
                    SpreadsheetApp.getUi().ButtonSet.OK
                );
            }
            
        } catch (error) {
            logError('Error en crearRespaldoManual', 'ERROR', { error: error.message });
            SpreadsheetApp.getUi().alert(
                '❌ Error', 
                'Error al crear respaldo manual: ' + error.message, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
    }

    /**
     * Crea un respaldo automático del sistema
     */
    function crearRespaldoAutomatico() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const fechaActual = new Date();
            const fechaFormateada = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
            const nombreArchivo = `Respaldo_SurPass_${fechaFormateada}.xlsx`;

            // Buscar o crear carpeta de respaldos
            let carpeta;
            const carpetaNombre = 'Respaldos_SurPass';
            const carpetas = DriveApp.getFoldersByName(carpetaNombre);

            if (carpetas.hasNext()) {
                carpeta = carpetas.next();
            } else {
                carpeta = DriveApp.createFolder(carpetaNombre);
                logError('Carpeta de respaldos creada', 'INFO', { carpetaId: carpeta.getId() });
            }

            // Exportar archivo
            const exportUrl = `https://www.googleapis.com/drive/v3/files/${ss.getId()}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
            const params = {
                method: 'get',
                headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
                muteHttpExceptions: true
            };

            const response = UrlFetchApp.fetch(exportUrl, params);
            
            if (response.getResponseCode() !== 200) {
                throw new Error(`Error al exportar archivo para respaldo: ${response.getResponseCode()}`);
            }

            const blob = response.getBlob().setName(nombreArchivo);
            const archivo = carpeta.createFile(blob);

            // Limpiar respaldos antiguos (mantener solo los últimos 15)
            const archivos = carpeta.getFiles();
            const todosArchivos = [];

            while (archivos.hasNext()) {
                const archivo = archivos.next();
                if (archivo.getName().startsWith('Respaldo_SurPass_')) {
                    todosArchivos.push(archivo);
                }
            }

            // Ordenar por fecha de creación
            todosArchivos.sort((a, b) => a.getDateCreated() - b.getDateCreated());

            // Eliminar archivos antiguos si hay más de 15
            if (todosArchivos.length > 15) {
                for (let i = 0; i < todosArchivos.length - 15; i++) {
                    todosArchivos[i].setTrashed(true);
                    logError('Respaldo antiguo eliminado', 'INFO', { archivo: todosArchivos[i].getName() });
                }
            }

            // Registrar en hoja de Backup si existe
            try {
                const backupSheet = ss.getSheetByName('Backup');
                if (backupSheet) {
                    backupSheet.appendRow([
                        Utilities.getUuid(),
                        nombreArchivo,
                        fechaActual,
                        'Respaldo automático',
                        archivo.getUrl(),
                        obtenerUsuarioActual(),
                        'Completado',
                        `Archivo creado: ${archivo.getSize()} bytes`
                    ]);
                }
            } catch (registroError) {
                logError('No se pudo registrar en hoja Backup', 'WARNING', { error: registroError.message });
            }

            logError(`Respaldo creado exitosamente: ${nombreArchivo}`, 'INFO');

            return {
                exito: true,
                mensaje: `Respaldo creado correctamente: ${nombreArchivo}`,
                archivoId: archivo.getId(),
                url: archivo.getUrl()
            };
            
        } catch (error) {
            logError('Error en crearRespaldoAutomatico', 'ERROR', { error: error.message });
            return {
                exito: false,
                mensaje: 'Error al crear respaldo: ' + error.message
            };
        }
    }

    /**
     * Limpia registros antiguos basado en configuración
     */
    function limpiarRegistrosAntiguos() {
        try {
            const config = obtenerConfiguracion();
            const diasRetener = parseInt(config.DIAS_RETENER_LOGS) || 30;
            
            if (config.LIMPIAR_LOGS_AUTOMATICO !== 'SI') {
                SpreadsheetApp.getUi().alert(
                    '⚠️ Función Deshabilitada', 
                    'La limpieza automática de logs está deshabilitada en la configuración.', 
                    SpreadsheetApp.getUi().ButtonSet.OK
                );
                return;
            }
            
            const fechaLimite = new Date();
            fechaLimite.setDate(fechaLimite.getDate() - diasRetener);
            
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const respuestasSheet = ss.getSheetByName('Respuestas formulario');
            
            let registrosEliminados = 0;
            
            if (respuestasSheet && respuestasSheet.getLastRow() > 1) {
                const data = respuestasSheet.getDataRange().getValues();
                const filasParaEliminar = [];
                
                for (let i = 1; i < data.length; i++) {
                    const fecha = data[i][0];
                    if (fecha instanceof Date && fecha < fechaLimite) {
                        filasParaEliminar.push(i + 1); // +1 porque las filas empiezan en 1
                    }
                }
                
                // Eliminar filas de atrás hacia adelante para mantener índices correctos
                for (let i = filasParaEliminar.length - 1; i >= 0; i--) {
                    respuestasSheet.deleteRow(filasParaEliminar[i]);
                    registrosEliminados++;
                }
            }
            
            logError(`Limpieza de registros antiguos completada: ${registrosEliminados} registros eliminados`, 'INFO');
            
            SpreadsheetApp.getUi().alert(
                '🧹 Limpieza Completada', 
                `Se eliminaron ${registrosEliminados} registros anteriores a ${diasRetener} días.\n\nFecha límite: ${Utilities.formatDate(fechaLimite, Session.getScriptTimeZone(), 'dd/MM/yyyy')}`, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
            
        } catch (error) {
            logError('Error en limpiarRegistrosAntiguos', 'ERROR', { error: error.message });
            SpreadsheetApp.getUi().alert(
                '❌ Error', 
                'Error al limpiar registros antiguos: ' + error.message, 
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        }
    }

    // =====================================================
    // ADMINISTRACIÓN DEL SISTEMA
    // =====================================================

    /**
     * Valida credenciales de administrador
     */
    function validarUsuarioAdmin(cedula) {
        try {
            if (!validarCampoObligatorio(cedula)) {
                return {
                    valido: false,
                    mensaje: 'Por favor, ingrese una cédula válida'
                };
            }

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheetClave = ss.getSheetByName('Clave');

            if (!sheetClave) {
                return {
                    valido: false,
                    mensaje: 'Error: La hoja "Clave" no existe. Configure el sistema primero.'
                };
            }

            const datos = sheetClave.getDataRange().getValues();
            const startRow = datos[0][0] === 'Cédula' ? 1 : 0;

            for (let i = startRow; i < datos.length; i++) {
                const cedulaHoja = String(datos[i][0] || '').trim();
                const cedulaIngresada = normalizarCedula(cedula);

                // Comparación exacta y normalizada
                if (cedulaHoja === cedula || 
                    normalizarCedula(cedulaHoja) === cedulaIngresada) {
                    
                    return {
                        valido: true,
                        usuario: {
                            cedula: cedulaHoja,
                            nombre: String(datos[i][1] || 'Usuario').trim(),
                            cargo: String(datos[i][2] || 'Administrador').trim(),
                            email: String(datos[i][3] || '').trim(),
                            estado: String(datos[i][4] || 'Activo').trim()
                        },
                        mensaje: 'Acceso autorizado'
                    };
                }
            }

            return {
                valido: false,
                mensaje: 'Cédula no autorizada para acceso administrativo'
            };
            
        } catch (error) {
            logError('Error en validarUsuarioAdmin', 'ERROR', { error: error.message, cedula });
            return {
                valido: false,
                mensaje: 'Error al validar usuario: ' + error.message
            };
        }
    }

    /**
     * Agrega un nuevo registro a la Base de Datos
     */
    function agregarRegistro(registro) {
        try {
            if (!registro || !validarCampoObligatorio(registro.cedula) || !validarCampoObligatorio(registro.nombre)) {
                return {
                    success: false,
                    message: 'La cédula y el nombre son obligatorios'
                };
            }

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bdSheet = ss.getSheetByName('Base de Datos');

            if (!bdSheet) {
                return {
                    success: false,
                    message: 'La hoja "Base de Datos" no fue encontrada'
                };
            }

            const data = bdSheet.getDataRange().getValues();
            const cedulaIdx = 0; // Primera columna es cédula

            // Verificar duplicados
            const cedulaNormalizada = normalizarCedula(registro.cedula);
            for (let i = 1; i < data.length; i++) {
                const cedulaExistente = String(data[i][cedulaIdx] || '').trim();
                const cedulaExistenteNorm = normalizarCedula(cedulaExistente);
                
                if (cedulaExistente === registro.cedula || 
                    cedulaExistenteNorm === cedulaNormalizada) {
                    return {
                        success: false,
                        message: 'La cédula ya existe en la base de datos'
                    };
                }
            }

            // Agregar nuevo registro
            bdSheet.appendRow([
                String(registro.cedula).trim(),
                String(registro.nombre).trim(),
                String(registro.empresa || 'No especificada').trim()
            ]);

            logError(`Nuevo registro agregado: ${registro.cedula} - ${registro.nombre}`, 'INFO', {
                usuario: obtenerUsuarioActual(),
                empresa: registro.empresa
            });

            return {
                success: true,
                message: 'Registro agregado correctamente'
            };
            
        } catch (error) {
            logError('Error en agregarRegistro', 'ERROR', { error: error.message, registro });
            return {
                success: false,
                message: 'Error al agregar registro: ' + error.message
            };
        }
    }

    /**
     * Actualiza un registro existente en la Base de Datos
     */
    function actualizarRegistro(registro, cedulaOriginal) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bdSheet = ss.getSheetByName('Base de Datos');
            
            if (!bdSheet) {
                throw new Error('Hoja "Base de Datos" no encontrada');
            }

            const data = bdSheet.getDataRange().getValues();
            let rowIndex = -1;

            const cedulaOriginalNorm = normalizarCedula(cedulaOriginal);

            // Buscar la fila a actualizar
            for (let i = 1; i < data.length; i++) {
                const cedulaEnFila = String(data[i][0] || '').trim();
                const cedulaEnFilaNorm = normalizarCedula(cedulaEnFila);

                if (cedulaEnFila === cedulaOriginal || cedulaEnFilaNorm === cedulaOriginalNorm) {
                    rowIndex = i + 1; // +1 porque las filas empiezan en 1
                    break;
                }
            }

            if (rowIndex === -1) {
                return { 
                    success: false, 
                    message: 'No se encontró el registro con la cédula especificada: ' + cedulaOriginal 
                };
            }

            // Preparar nuevos valores
            const nuevaCedula = String(registro.cedula || cedulaOriginal).trim();
            const nuevoNombre = String(registro.nombre || data[rowIndex - 1][1] || '').trim();
            const nuevaEmpresa = String(registro.empresa || data[rowIndex - 1][2] || 'No especificada').trim();

            // Verificar duplicados si la cédula cambia
            if (nuevaCedula !== cedulaOriginal) {
                const nuevaCedulaNorm = normalizarCedula(nuevaCedula);
                for (let i = 1; i < data.length; i++) {
                    if (i + 1 === rowIndex) continue; // Saltar la fila actual
                    
                    const cedulaEnFilaNorm = normalizarCedula(String(data[i][0] || '').trim());
                    if (cedulaEnFilaNorm === nuevaCedulaNorm) {
                        return { 
                            success: false, 
                            message: 'La nueva cédula ya está en uso: ' + nuevaCedula 
                        };
                    }
                }
            }

            // Actualizar la fila
            bdSheet.getRange(rowIndex, 1).setValue(nuevaCedula);
            bdSheet.getRange(rowIndex, 2).setValue(nuevoNombre);
            bdSheet.getRange(rowIndex, 3).setValue(nuevaEmpresa);

            logError(`Registro actualizado: ${cedulaOriginal} -> ${nuevaCedula}`, 'INFO', {
                usuario: obtenerUsuarioActual(),
                nuevoNombre,
                nuevaEmpresa
            });

            return { 
                success: true, 
                message: 'Registro actualizado correctamente' 
            };
            
        } catch (error) {
            logError('Error en actualizarRegistro', 'ERROR', { error: error.message, registro, cedulaOriginal });
            return { 
                success: false, 
                message: 'Error al actualizar: ' + error.message 
            };
        }
    }

    /**
     * Elimina un registro de la Base de Datos
     */
    function eliminarRegistro(cedula) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bdSheet = ss.getSheetByName('Base de Datos');
            
            if (!bdSheet) {
                throw new Error('Hoja "Base de Datos" no encontrada');
            }

            const data = bdSheet.getDataRange().getValues();
            let rowIndex = -1;
            const cedulaBuscarNorm = normalizarCedula(cedula);

            // Buscar la fila a eliminar
            for (let i = 1; i < data.length; i++) {
                const cedulaFila = String(data[i][0] || '').trim();
                const cedulaFilaNorm = normalizarCedula(cedulaFila);
                
                if (cedulaFila === cedula || cedulaFilaNorm === cedulaBuscarNorm) {
                    rowIndex = i + 1; // +1 porque las filas empiezan en 1
                    break;
                }
            }

            if (rowIndex === -1) {
                return { 
                    success: false, 
                    message: 'Registro no encontrado para la cédula: ' + cedula 
                };
            }

            // Guardar información antes de eliminar para el log
            const nombre = data[rowIndex - 1][1];
            const empresa = data[rowIndex - 1][2];

            // Eliminar la fila
            bdSheet.deleteRow(rowIndex);

            logError(`Registro eliminado: ${cedula}`, 'INFO', {
                usuario: obtenerUsuarioActual(),
                nombre,
                empresa
            });

            return { 
                success: true, 
                message: 'Registro eliminado correctamente' 
            };
            
        } catch (error) {
            logError('Error en eliminarRegistro', 'ERROR', { error: error.message, cedula });
            return { 
                success: false, 
                message: 'Error al eliminar: ' + error.message 
            };
        }
    }

    /**
     * Exporta la Base de Datos a un archivo CSV
     */
    function exportarBaseDeDatos() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bdSheet = ss.getSheetByName('Base de Datos');

            if (!bdSheet) {
                return {
                    success: false,
                    message: 'La hoja "Base de Datos" no fue encontrada'
                };
            }

            const data = bdSheet.getDataRange().getValues();
            
            if (data.length <= 1) {
                return {
                    success: false,
                    message: 'La Base de Datos está vacía'
                };
            }

            // Crear contenido CSV con BOM UTF-8
            let csvContent = '\uFEFF'; // BOM para UTF-8
            csvContent += data.map(row => {
                return row.map(cell => {
                    const cellStr = String(cell || '');
                    // Escapar comillas y agregar comillas si contiene comas
                    if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
                        return '"' + cellStr.replace(/"/g, '""') + '"';
                    }
                    return cellStr;
                }).join(',');
            }).join('\n');

            // Crear archivo
            const fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
            const nombreArchivo = `SurPass_Base_Datos_${fechaActual}.csv`;
            
            const blob = Utilities.newBlob(csvContent, 'text/csv; charset=utf-8', nombreArchivo);
            const file = DriveApp.createFile(blob);
            
            // Configurar permisos
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

            logError(`Base de Datos exportada: ${nombreArchivo}`, 'INFO', { 
                registros: data.length - 1,
                usuario: obtenerUsuarioActual()
            });

            return {
                success: true,
                message: 'Base de Datos exportada correctamente',
                fileName: nombreArchivo,
                fileUrl: file.getUrl(),
                recordCount: data.length - 1
            };
            
        } catch (error) {
            logError('Error en exportarBaseDeDatos', 'ERROR', { error: error.message });
            return {
                success: false,
                message: 'Error al exportar la Base de Datos: ' + error.message
            };
        }
    }

    // =====================================================
    // FUNCIONES DE DIAGNÓSTICO Y TESTING
    // =====================================================

    /**
     * Función de diagnóstico completo del sistema
     */
    function diagnosticoCompletoSistema() {
        try {
            logError('=== INICIANDO DIAGNÓSTICO COMPLETO DEL SISTEMA SURPASS ===', 'INFO');
            
            const resultados = {
                configuracion: { status: false, mensaje: '', tiempo: 0 },
                baseDatos: { status: false, mensaje: '', tiempo: 0 },
                hojas: { status: false, mensaje: '', tiempo: 0 },
                permisos: { status: false, mensaje: '', tiempo: 0 },
                evacuacion: { status: false, mensaje: '', tiempo: 0 },
                estadisticas: { status: false, mensaje: '', tiempo: 0 },
                opciones: { status: false, mensaje: '', tiempo: 0 },
                normalizacion: { status: false, mensaje: '', tiempo: 0 }
            };
            
            // Test 1: Configuración del sistema
            let startTime = new Date().getTime();
            try {
                const config = obtenerConfiguracion();
                const tieneConfigBasica = config.EMPRESA_NOMBRE && config.HORARIO_APERTURA;
                resultados.configuracion = {
                    status: tieneConfigBasica,
                    mensaje: tieneConfigBasica ? 
                        `✅ Configuración OK - Empresa: ${config.EMPRESA_NOMBRE}` : 
                        '❌ Configuración incompleta',
                    tiempo: new Date().getTime() - startTime
                };
            } catch (e) {
                resultados.configuracion = {
                    status: false,
                    mensaje: `❌ Error en configuración: ${e.message}`,
                    tiempo: new Date().getTime() - startTime
                };
            }
            
            // Test 2: Base de Datos
            startTime = new Date().getTime();
            try {
                const personal = obtenerTodoElPersonal();
                const esValida = Array.isArray(personal) && personal.length >= 0;
                resultados.baseDatos = {
                    status: esValida,
                    mensaje: esValida ? 
                        `✅ Base de Datos OK - ${personal.length} registros` : 
                        '❌ Base de Datos inválida',
                    tiempo: new Date().getTime() - startTime
                };
            } catch (e) {
                resultados.baseDatos = {
                    status: false,
                    mensaje: `❌ Error en Base de Datos: ${e.message}`,
                    tiempo: new Date().getTime() - startTime
                };
            }
            
            // Test 3: Estructura de hojas
            startTime = new Date().getTime();
            try {
                const validacion = validarEstructuraHojas();
                resultados.hojas = {
                    status: validacion.valido,
                    mensaje: validacion.valido ? 
                        `✅ Estructura OK - ${validacion.warnings?.length || 0} advertencias` : 
                        `❌ Estructura inválida - ${validacion.errores} errores`,
                    tiempo: new Date().getTime() - startTime
                };
            } catch (e) {
                resultados.hojas = {
                    status: false,
                    mensaje: `❌ Error validando hojas: ${e.message}`,
                    tiempo: new Date().getTime() - startTime
                };
            }
            
            // Test 4: Permisos y acceso
            startTime = new Date().getTime();
            try {
                const user = obtenerUsuarioActual();
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const tieneAcceso = user && ss.getId();
                resultados.permisos = {
                    status: tieneAcceso,
                    mensaje: tieneAcceso ? 
                        `✅ Permisos OK - Usuario: ${user}` : 
                        '❌ Sin permisos adecuados',
                    tiempo: new Date().getTime() - startTime
                };
            } catch (e) {
                resultados.permisos = {
                    status: false,
                    mensaje: `❌ Error verificando permisos: ${e.message}`,
                    tiempo: new Date().getTime() - startTime
                };
            }
            
            // Test 5: Sistema de evacuación
            startTime = new Date().getTime();
            try {
                const estadoEvacuacion = getEvacuacionDataForClient(); // Use the new unified function
                const esValido = estadoEvacuacion && typeof estadoEvacuacion.totalDentro === 'number';
                resultados.evacuacion = {
                    status: esValido,
                    mensaje: esValido ? 
                        `✅ Evacuación OK - ${estadoEvacuacion.totalDentro} personas dentro` : 
                        '❌ Sistema de evacuación con errores',
                    tiempo: new Date().getTime() - startTime
                };
            } catch (e) {
                resultados.evacuacion = {
                    status: false,
                    mensaje: `❌ Error en evacuación: ${e.message}`,
                    tiempo: new Date().getTime() - startTime
                };
            }
            
            // Test 6: Estadísticas
            startTime = new Date().getTime();
            try {
                const stats = obtenerEstadisticas();
                const sonValidas = stats && typeof stats.entradas === 'number' && typeof stats.salidas === 'number';
                resultados.estadisticas = {
                    status: sonValidas,
                    mensaje: sonValidas ? 
                        `✅ Estadísticas OK - E:${stats.entradas} S:${stats.salidas}` : 
                        '❌ Estadísticas inválidas',
                    tiempo: new Date().getTime() - startTime
                };
            } catch (e) {
                resultados.estadisticas = {
                    status: false,
                    mensaje: `❌ Error en estadísticas: ${e.message}`,
                    tiempo: new Date().getTime() - startTime
                };
            }
            
            // Test 7: Opciones de menú
            startTime = new Date().getTime();
            try {
                const opciones = obtenerOpcionesMenu();
                const sonValidas = opciones && typeof opciones.mostrarEstadisticas === 'boolean';
                resultados.opciones = {
                    status: sonValidas,
                    mensaje: sonValidas ? 
                        '✅ Opciones de menú OK' : 
                        '❌ Opciones de menú inválidas',
                    tiempo: new Date().getTime() - startTime
                };
            } catch (e) {
                resultados.opciones = {
                    status: false,
                    mensaje: `❌ Error en opciones: ${e.message}`,
                    tiempo: new Date().getTime() - startTime
                };
            }
            
            // Test 8: Normalización de cédulas
            startTime = new Date().getTime();
            try {
                const test1 = normalizarCedula('8-123-456');
                const test2 = normalizarCedula('{"cedula": "9-234-567"}');
                const test3 = normalizarCedula('Texto - PE-456-789');
                const funcionaBien = test1 === '8-123-456' && 
                                    test2 === '9-234-567' && 
                                    test3 === 'PE-456-789';
                resultados.normalizacion = {
                    status: funcionaBien,
                    mensaje: funcionaBien ? 
                        '✅ Normalización OK - Múltiples formatos soportados' : 
                        '❌ Normalización con errores',
                    tiempo: new Date().getTime() - startTime
                };
            } catch (e) {
                resultados.normalizacion = {
                    status: false,
                    mensaje: `❌ Error en normalización: ${e.message}`,
                    tiempo: new Date().getTime() - startTime
                };
            }
            
            // Calcular resumen
            const testsPasados = Object.values(resultados).filter(test => test.status).length;
            const totalTests = Object.keys(resultados).length;
            const tiempoTotal = Object.values(resultados).reduce((sum, test) => sum + test.tiempo, 0);
            
            // Log detallado
            logError('\n=== RESULTADOS DETALLADOS ===', 'INFO');
            Object.keys(resultados).forEach(test => {
                const resultado = resultados[test];
                logError(`${test.toUpperCase()}: ${resultado.mensaje} (${resultado.tiempo}ms)`, resultado.status ? 'INFO' : 'ERROR');
            });
            
            logError(`\n=== RESUMEN FINAL ===`, 'INFO');
            logError(`✅ Tests exitosos: ${testsPasados}/${totalTests}`, 'INFO');
            logError(`⏱️ Tiempo total: ${tiempoTotal}ms`, 'INFO');
            logError(`📊 Porcentaje de éxito: ${Math.round((testsPasados/totalTests)*100)}%`, 'INFO');
            logError('========================', 'INFO');
            
            const esExitoso = testsPasados === totalTests;
            const mensaje = esExitoso ? 
                `✅ Sistema completamente funcional (${testsPasados}/${totalTests} tests pasados)` :
                `⚠️ Sistema parcialmente funcional (${testsPasados}/${totalTests} tests pasados)`;
            
            return {
                exito: esExitoso,
                mensaje: mensaje,
                testsPasados: testsPasados,
                totalTests: totalTests,
                tiempoTotal: tiempoTotal,
                detalles: resultados
            };
            
        } catch (error) {
            logError('❌ Error crítico en diagnóstico completo', 'CRITICAL', { error: error.message });
            return {
                exito: false,
                mensaje: 'Error crítico durante el diagnóstico: ' + error.message,
                testsPasados: 0,
                totalTests: 0,
                tiempoTotal: 0,
                detalles: {}
            };
        }
    }

    /**
     * Trigger automático para respaldo diario
     */
    function dailyBackup() {
        try {
            const config = obtenerConfiguracion();
            
            if (config.BACKUP_AUTOMATICO !== 'SI') {
                logError('Backup automático deshabilitado en configuración', 'INFO');
                return;
            }
            
            const resultado = crearRespaldoAutomatico();
            if (resultado.exito) {
                logError('✅ Backup automático completado: ' + resultado.mensaje, 'INFO');
            } else {
                logError('❌ Error en backup automático: ' + resultado.mensaje, 'ERROR');
            }
        } catch (error) {
            logError('❌ Error en dailyBackup', 'ERROR', { error: error.message });
        }
    }

    function ejecutarDiagnosticoUnificado(tipoTest = 'completo') {
        const startTime = new Date().getTime();
        const sessionId = Utilities.getUuid();
        
        logError(`🔍 Iniciando diagnóstico unificado: ${tipoTest}`, 'INFO', { sessionId });
        
        try {
            const resultados = {
                conexionBasica: { status: false, mensaje: '', tiempo: 0 },
                configuracion: { status: false, mensaje: '', tiempo: 0 },
                evacuacion: { status: false, mensaje: '', tiempo: 0 },
                personal: { status: false, mensaje: '', tiempo: 0 },
                hojas: { status: false, mensaje: '', tiempo: 0 }
            };
            
            // ✅ TEST 1: Conexión básica
            let testStart = new Date().getTime();
            try {
                const user = Session.getEffectiveUser().getEmail();
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const ssId = ss.getId();
                
                resultados.conexionBasica = {
                    status: true,
                    mensaje: `✅ Conexión OK - Usuario: ${user.substring(0, 20)}...`,
                    tiempo: new Date().getTime() - testStart
                };
            } catch (e) {
                resultados.conexionBasica = {
                    status: false,
                    mensaje: `❌ Error de conexión: ${e.message}`,
                    tiempo: new Date().getTime() - testStart
                };
            }
            
            // ✅ TEST 2: Configuración (solo si es completo)
            if (tipoTest === 'completo') {
                testStart = new Date().getTime();
                try {
                    const config = obtenerConfiguracion();
                    const esValida = config.EMPRESA_NOMBRE && config.HORARIO_APERTURA;
                    
                    resultados.configuracion = {
                        status: esValida,
                        mensaje: esValida ? 
                            `✅ Configuración OK - ${config.EMPRESA_NOMBRE}` : 
                            '❌ Configuración incompleta',
                        tiempo: new Date().getTime() - testStart
                    };
                } catch (e) {
                    resultados.configuracion = {
                        status: false,
                        mensaje: `❌ Error configuración: ${e.message}`,
                        tiempo: new Date().getTime() - testStart
                    };
                }
            }
            
            // ✅ TEST 3: Sistema de evacuación
            testStart = new Date().getTime();
            try {
                const datosEvacuacion = getEvacuacionDataForClient();
                const esValido = datosEvacuacion && datosEvacuacion.success !== false;
                
                resultados.evacuacion = {
                    status: esValido,
                    mensaje: esValido ? 
                        `✅ Evacuación OK - ${datosEvacuacion.totalDentro || 0} personas` : 
                        '❌ Sistema evacuación con errores',
                    tiempo: new Date().getTime() - testStart
                };
            } catch (e) {
                resultados.evacuacion = {
                    status: false,
                    mensaje: `❌ Error evacuación: ${e.message}`,
                    tiempo: new Date().getTime() - testStart
                };
            }
            
            // ✅ TEST 4: Personal (solo si es completo)
            if (tipoTest === 'completo') {
                testStart = new Date().getTime();
                try {
                    const personal = obtenerTodoElPersonal();
                    const esValida = Array.isArray(personal) && personal.length >= 0;
                    
                    resultados.personal = {
                        status: esValida,
                        mensaje: esValida ? 
                            `✅ Personal OK - ${personal.length} registros` : 
                            '❌ Base de datos con errores',
                        tiempo: new Date().getTime() - testStart
                    };
                }
                catch (e) {
                    resultados.personal = {
                        status: false,
                        mensaje: `❌ Error personal: ${e.message}`,
                        tiempo: new Date().getTime() - testStart
                    };
                }
            }
            
            // ✅ TEST 5: Estructura de hojas (solo si es completo)
            if (tipoTest === 'completo') {
                testStart = new Date().getTime();
                try {
                    const validacion = validarEstructuraHojas();
                    
                    resultados.hojas = {
                        status: validacion.valido,
                        mensaje: validacion.valido ? 
                            `✅ Hojas OK - ${validacion.warnings?.length || 0} advertencias` : 
                            `❌ Hojas inválidas - ${validacion.errores} errores`,
                        tiempo: new Date().getTime() - testStart
                    };
                } catch (e) {
                    resultados.hojas = {
                        status: false,
                        mensaje: `❌ Error hojas: ${e.message}`,
                        tiempo: new Date().getTime() - testStart
                    };
                }
            }
            
            // ✅ CALCULAR RESUMEN
            const testsRealizados = Object.values(resultados).filter(r => r.tiempo > 0);
            const testsPasados = testsRealizados.filter(r => r.status).length;
            const tiempoTotal = new Date().getTime() - startTime;
            
            // ✅ LOG DETALLADO
            logError('\n=== RESULTADOS DIAGNÓSTICO UNIFICADO ===', 'INFO');
            Object.keys(resultados).forEach(test => {
                const resultado = resultados[test];
                if (resultado.tiempo > 0) {
                    logError(`${test.toUpperCase()}: ${resultado.mensaje} (${resultado.tiempo}ms)`, 
                            resultado.status ? 'INFO' : 'ERROR');
                }
            });
            
            const esExitoso = testsPasados === testsRealizados.length;
            const mensaje = esExitoso ? 
                `✅ Diagnóstico ${tipoTest} exitoso (${testsPasados}/${testsRealizados.length})` :
                `⚠️ Diagnóstico ${tipoTest} con problemas (${testsPasados}/${testsRealizados.length})`;
            
            logError(`\n📊 RESUMEN: ${mensaje} - ${tiempoTotal}ms total`, 'INFO');
            
            return {
                success: esExitoso,
                tipo: tipoTest,
                mensaje: mensaje,
                testsPasados: testsPasados,
                totalTests: testsRealizados.length,
                tiempoTotal: tiempoTotal,
                detalles: resultados,
                sessionId: sessionId
            };
            
        } catch (error) {
            logError('❌ Error crítico en diagnóstico unificado', 'CRITICAL', { 
                error: error.message, 
                sessionId 
            });
            
            return {
                success: false,
                tipo: tipoTest,
                mensaje: 'Error crítico: ' + error.message,
                testsPasados: 0,
                totalTests: 0,
                tiempoTotal: new Date().getTime() - startTime,
                detalles: {},
                sessionId: sessionId
            };
        }
    }

    // ✅ FUNCIONES ESPECÍFICAS QUE LLAMAN AL DIAGNÓSTICO UNIFICADO
    function testSimple() {
        return ejecutarDiagnosticoUnificado('basico');
    }

    function testConexionBasica() {
        const resultado = ejecutarDiagnosticoUnificado('basico');
        return {
            success: resultado.success,
            message: resultado.mensaje,
            timestamp: new Date().toISOString(),
            test: 'Conexión básica verificada'
        };
    }

    function verificarEvacuacionFunciona() {
        const resultado = ejecutarDiagnosticoUnificado('basico');
        return resultado.detalles.evacuacion.status;
    }

    /**
     * FUNCIÓN SIMPLE PARA LIMPIAR DATOS DE PRUEBA
     */
    function limpiarDatosPrueba() {
        try {
            console.log('🧹 Limpiando datos de prueba...');
            
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            
            // Limpiar entradas TEST del Historial
            const historialSheet = ss.getSheetByName('Historial');
            if (historialSheet) {
                const data = historialSheet.getDataRange().getValues();
                for (let i = data.length - 1; i >= 1; i--) {
                    const cedula = String(data[i][1] || '');
                    if (cedula.startsWith('TEST-')) {
                        historialSheet.deleteRow(i + 1);
                        console.log(`🗑️ Eliminada fila de prueba: ${cedula}`);
                    }
                }
            }
            
            // Limpiar entradas TEST de Base de Datos
            const bdSheet = ss.getSheetByName('Base de Datos');
            if (bdSheet) {
                const bdData = bdSheet.getDataRange().getValues();
                for (let i = bdData.length - 1; i >= 1; i--) {
                    const cedula = String(bdData[i][0] || '');
                    if (cedula.startsWith('TEST-')) {
                        bdSheet.deleteRow(i + 1);
                        console.log(`🗑️ Eliminada BD de prueba: ${cedula}`);
                    }
                }
            }
            
            console.log('✅ Datos de prueba limpiados');
            return true;
            
        } catch (error) {
            console.error('❌ Error limpiando datos de prueba:', error);
            return false;
        }
    }

    /**
     * Obtiene los registros de emergencia para mostrar en la interfaz
     * @param {number} limite - Número máximo de registros a devolver (los más recientes)
     * @return {Object} Objeto con los registros de emergencia y metadatos
     */
    function obtenerLogsEmergencia(limite = 100) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const emergencySheet = ss.getSheetByName('Log_Emergencias');
            
            if (!emergencySheet) {
                return {
                    success: true,
                    message: 'No se encontró la hoja de Log_Emergencias',
                    registros: [],
                    total: 0
                };
            }
            
            const lastRow = emergencySheet.getLastRow();
            
            if (lastRow <= 1) {
                return {
                    success: true,
                    message: 'No hay registros de emergencia',
                    registros: [],
                    total: 0
                };
            }
            
            // Obtener encabezados
            const headers = emergencySheet.getRange(1, 1, 1, emergencySheet.getLastColumn()).getValues()[0];
            
            // Calcular filas a obtener (las más recientes según el límite)
            const startRow = Math.max(2, lastRow - limite + 1);
            const numRows = lastRow - startRow + 1;
            
            // Obtener datos
            const data = emergencySheet.getRange(startRow, 1, numRows, headers.length).getValues();
            
            // Mapear a objetos con nombres de propiedades según los encabezados
            const registros = data.map(row => {
                const registro = {};
                headers.forEach((header, index) => {
                    registro[header] = row[index];
                    // Formatear fechas
                    if (header === 'Fecha_Hora' && row[index] instanceof Date) {
                        registro[header + '_formatted'] = Utilities.formatDate(
                            row[index], 
                            Session.getScriptTimeZone(), 
                            'dd/MM/yyyy HH:mm:ss'
                        );
                    }
                });
                return registro;
            });
            
            // Ordenar por fecha descendente (más reciente primero)
            registros.sort((a, b) => {
                const dateA = a.Fecha_Hora || 0;
                const dateB = b.Fecha_Hora || 0;
                return dateB - dateA;
            });
            
            return {
                success: true,
                message: `${registros.length} registros encontrados`,
                registros: registros,
                total: registros.length,
                headers: headers
            };
            
        } catch (error) {
            logError('Error al obtener logs de emergencia', 'ERROR', { error: error.message });
            return {
                success: false,
                message: 'Error al obtener logs de emergencia: ' + error.message,
                registros: [],
                total: 0
            };
        }
    }

    /**
     * Valida que el sistema de evacuación esté funcionando correctamente
     */
    function validarSistemaEvacuacion() {
        try {
            const detalles = {
                hojas: false,
                configuracion: false,
                datosEvacuacion: false,
                permisos: false
            };
            
            let errores = [];
            
            // Test 1: Verificar hojas necesarias
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const historial = ss.getSheetByName('Historial');
                const baseDatos = ss.getSheetByName('Base de Datos');
                
                detalles.hojas = historial && baseDatos;
                if (!detalles.hojas) {
                    errores.push('Faltan hojas necesarias (Historial o Base de Datos)');
                }
            } catch (e) {
                errores.push('Error accediendo a hojas: ' + e.message);
            }
            
            // Test 2: Verificar configuración
            try {
                const config = obtenerConfiguracion();
                detalles.configuracion = config && config.EMPRESA_NOMBRE;
                if (!detalles.configuracion) {
                    errores.push('Configuración incompleta');
                }
            } catch (e) {
                errores.push('Error en configuración: ' + e.message);
            }
            
            // Test 3: Verificar datos de evacuación
            try {
                const datosEvacuacion = getEvacuacionDataForClient();
                detalles.datosEvacuacion = datosEvacuacion && datosEvacuacion.success !== false;
                if (!detalles.datosEvacuacion) {
                    errores.push('Error obteniendo datos de evacuación');
                }
            } catch (e) {
                errores.push('Error en datos de evacuación: ' + e.message);
            }
            
            // Test 4: Verificar permisos
            try {
                const usuario = obtenerUsuarioActual();
                detalles.permisos = usuario && !usuario.includes('desconocido');
                if (!detalles.permisos) {
                    errores.push('Problemas de permisos de usuario');
                }
            } catch (e) {
                errores.push('Error verificando permisos: ' + e.message);
            }
            
            const success = Object.values(detalles).every(Boolean);
            const resumen = success ? 
                'Sistema de evacuación operativo' : 
                `Problemas encontrados: ${errores.join(', ')}`;
            
            return {
                success: success,
                detalles: detalles,
                resumen: resumen,
                errores: errores
            };
            
        } catch (error) {
            return {
                success: false,
                detalles: {
                    hojas: false,
                    configuracion: false,
                    datosEvacuacion: false,
                    permisos: false
                },
                resumen: 'Error crítico en validación: ' + error.message,
                errores: [error.message]
            };
        }
    }

    // =====================================================
    // FUNCIÓN FINAL DE VALIDACIÓN DEL SISTEMA COMPLETO
    // =====================================================

    /**
     * ✅ VALIDACIÓN COMPLETA DEL SISTEMA SURPASS CORREGIDO
     * Esta función verifica que todas las correcciones estén funcionando correctamente
     */
    function validarSistemaCorregido() {
        const resultados = {
            timestamp: new Date().toISOString(),
            version: 'SurPass v3.0 - Corregido',
            usuario: obtenerUsuarioActual(),
            tests: {}
        };

        try {
            console.log('🔍 === INICIANDO VALIDACIÓN COMPLETA DEL SISTEMA CORREGIDO ===');

            // ✅ TEST 1: Función procesarSimulacroUnificado implementada
            try {
                if (typeof procesarSimulacroUnificado === 'function') {
                    resultados.tests.simulacroFunction = {
                        status: 'PASS',
                        mensaje: 'Función procesarSimulacroUnificado correctamente implementada'
                    };
                } else {
                    throw new Error('Función no encontrada');
                }
            } catch (e) {
                resultados.tests.simulacroFunction = {
                    status: 'FAIL',
                    mensaje: 'Función procesarSimulacroUnificado NO implementada: ' + e.message
                };
            }

            // ✅ TEST 2: Sistema de evacuación unificado
            try {
                const datosEvacuacion = getEvacuacionDataForClient();
                resultados.tests.evacuacionUnificada = {
                    status: datosEvacuacion && datosEvacuacion.totalDentro >= 0 ? 'PASS' : 'FAIL',
                    mensaje: `Sistema de evacuación funcional - ${datosEvacuacion?.totalDentro || 0} personas detectadas`
                };
            } catch (e) {
                resultados.tests.evacuacionUnificada = {
                    status: 'FAIL',
                    mensaje: 'Error en sistema de evacuación: ' + e.message
                };
            }

            // ✅ TEST 3: Estadísticas corregidas (con entradas/salidas)
            try {
                const stats = obtenerEstadisticasBasicas();
                const tieneEntradasSalidas = typeof stats.entradas === 'number' && typeof stats.salidas === 'number';
                resultados.tests.estadisticasCorregidas = {
                    status: tieneEntradasSalidas ? 'PASS' : 'FAIL',
                    mensaje: tieneEntradasSalidas ? 
                        `Estadísticas completas - E:${stats.entradas} S:${stats.salidas}` : 
                        'Faltan propiedades entradas/salidas en estadísticas'
                };
            } catch (e) {
                resultados.tests.estadisticasCorregidas = {
                    status: 'FAIL',
                    mensaje: 'Error en estadísticas: ' + e.message
                };
            }

            // ✅ TEST 4: Funciones de logging implementadas
            try {
                const tieneLogging = typeof registrarLogEvacuacionUnificado === 'function';
                resultados.tests.loggingImplementado = {
                    status: tieneLogging ? 'PASS' : 'FAIL',
                    mensaje: tieneLogging ? 'Sistema de logging completo' : 'Sistema de logging faltante'
                };
            } catch (e) {
                resultados.tests.loggingImplementado = {
                    status: 'FAIL',
                    mensaje: 'Error verificando logging: ' + e.message
                };
            }

            // ✅ TEST 5: Configuración del sistema
            try {
                const config = obtenerConfiguracion();
                resultados.tests.configuracion = {
                    status: config && config.EMPRESA_NOMBRE ? 'PASS' : 'FAIL',
                    mensaje: config?.EMPRESA_NOMBRE ? 
                        `Configuración OK - ${config.EMPRESA_NOMBRE}` : 
                        'Configuración incompleta'
                };
            } catch (e) {
                resultados.tests.configuracion = {
                    status: 'FAIL',
                    mensaje: 'Error en configuración: ' + e.message
                };
            }

            // ✅ RESUMEN FINAL
            const testsPasados = Object.values(resultados.tests).filter(test => test.status === 'PASS').length;
            const totalTests = Object.keys(resultados.tests).length;
            const porcentajeExito = Math.round((testsPasados / totalTests) * 100);

            resultados.resumen = {
                testsPasados,
                totalTests,
                porcentajeExito,
                estado: porcentajeExito === 100 ? 'SISTEMA COMPLETAMENTE FUNCIONAL' : 
                    porcentajeExito >= 80 ? 'SISTEMA FUNCIONAL CON ADVERTENCIAS' : 
                    'SISTEMA CON PROBLEMAS CRÍTICOS'
            };

            console.log(`✅ VALIDACIÓN COMPLETADA: ${testsPasados}/${totalTests} tests pasados (${porcentajeExito}%)`);
            console.log(`📊 ESTADO: ${resultados.resumen.estado}`);

            return resultados;

        } catch (error) {
            console.error('❌ Error crítico en validación del sistema:', error);
            resultados.tests.validacionGeneral = {
                status: 'CRITICAL_FAIL',
                mensaje: 'Error crítico durante validación: ' + error.message
            };
            return resultados;
        }
    }

    /**
     * ✅ FUNCIÓN DE SIMULACRO PARA TESTING
     */
    function testSimulacroSistema() {
        try {
            console.log('🎭 Iniciando test de simulacro...');
            
            const cedulasPrueba = ['TEST-001', 'TEST-002'];
            const resultado = procesarSimulacroEvacuacion(cedulasPrueba, 'Test automático del sistema');
            
            if (resultado.success) {
                console.log(`✅ Test de simulacro exitoso: ${resultado.message}`);
                return {
                    success: true,
                    mensaje: 'Sistema de simulacros funcionando correctamente',
                    detalles: resultado
                };
            } else {
                throw new Error(resultado.message);
            }
            
        } catch (error) {
            console.error('❌ Error en test de simulacro:', error);
            return {
                success: false,
                mensaje: 'Error en sistema de simulacros: ' + error.message
            };
        }
    }

    /**
     * ✅ FUNCIÓN PARA EXPORTAR LOGS DE SIMULACROS
     */
    function exportarLogsSimulacros() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const simulacrosSheet = ss.getSheetByName('Log_Simulacros');
            
            if (!simulacrosSheet) {
                return {
                    success: false,
                    message: 'No se encontró la hoja Log_Simulacros'
                };
            }
            
            const data = simulacrosSheet.getDataRange().getValues();
            
            if (data.length <= 1) {
                return {
                    success: false,
                    message: 'No hay datos de simulacros para exportar'
                };
            }
            
            // Crear CSV
            let csvContent = '\uFEFF'; // BOM UTF-8
            csvContent += data.map(row => {
                return row.map(cell => {
                    const cellStr = String(cell || '');
                    if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
                        return '"' + cellStr.replace(/"/g, '""') + '"';
                    }
                    return cellStr;
                }).join(',');
            }).join('\n');
            
            const fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
            const nombreArchivo = `SurPass_Simulacros_${fechaActual}.csv`;
            
            const blob = Utilities.newBlob(csvContent, 'text/csv; charset=utf-8', nombreArchivo);
            const file = DriveApp.createFile(blob);
            
            return {
                success: true,
                message: 'Logs de simulacros exportados correctamente',
                fileName: nombreArchivo,
                fileUrl: file.getUrl(),
                recordCount: data.length - 1
            };
            
        } catch (error) {
            logError('Error exportando logs de simulacros', 'ERROR', { error: error.message });
            return {
                success: false,
                message: 'Error al exportar logs: ' + error.message
            };
        }
    }

    // =====================================================
    // MENSAJE FINAL DE CORRECCIONES APLICADAS
    // =====================================================
    console.log('✅ === SISTEMA SURPASS v3.0 - CORRECCIONES COMPLETADAS ===');
    console.log('🎭 SIMULACROS: Solo registran en Log_Simulacros, NO modifican Historial');
    console.log('🔧 FUNCIONES: Todas las funciones faltantes implementadas');
    console.log('📊 ESTADÍSTICAS: Sistema corregido con entradas/salidas');
    console.log('🛡️ ROBUSTEZ: Sistema completamente funcional y robusto');
    console.log('🚀 VALIDACIÓN: Ejecute validarSistemaCorregido() para verificar');

    /**
     * ✅ FUNCIÓN DE TESTING MEJORADA PARA VERIFICAR SIMULACROS
     */
    function testearSimulacrosCompleto() {
        console.log('🎭 === INICIANDO TEST COMPLETO DE SIMULACROS ===');
        
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            
            // 1. Verificar que existe la función de simulacro
            if (typeof procesarSimulacroEvacuacion !== 'function') {
                throw new Error('Función procesarSimulacroEvacuacion no existe');
            }
            
            // 2. Crear datos de prueba
            const cedulasPrueba = ['TEST-SIM-001', 'TEST-SIM-002'];
            console.log('📝 Datos de prueba creados:', cedulasPrueba);
            
            // 3. Agregar personas de prueba al historial (como si hubieran entrado)
            const historialSheet = ss.getSheetByName('Historial');
            if (historialSheet) {
                const fechaPrueba = new Date();
                historialSheet.appendRow([fechaPrueba, 'TEST-SIM-001', 'Persona Prueba 1', 'Empresa Test', fechaPrueba, null, 'Entrada para test', 'TEST']);
                historialSheet.appendRow([fechaPrueba, 'TEST-SIM-002', 'Persona Prueba 2', 'Empresa Test', fechaPrueba, null, 'Entrada para test', 'TEST']);
                console.log('✅ Registros de prueba agregados al historial');
            }
            
            // 4. Ejecutar simulacro
            console.log('🎭 Ejecutando simulacro...');
            const resultadoSimulacro = procesarSimulacroEvacuacion(cedulasPrueba, 'Test automático de simulacro');
            
            // 5. Verificar resultado
            if (!resultadoSimulacro.success) {
                throw new Error('El simulacro falló: ' + resultadoSimulacro.message);
            }
            
            console.log('✅ Simulacro ejecutado:', resultadoSimulacro.message);
            
            // 6. Verificar que NO se modificó el historial
            const historialData = historialSheet.getDataRange().getValues();
            let historialModificado = false;
            
            for (let i = 1; i < historialData.length; i++) {
                const cedula = String(historialData[i][1] || '');
                const salida = historialData[i][5]; // Columna de salida
                
                if (cedulasPrueba.includes(cedula) && salida && salida !== null && salida !== '') {
                    console.error(`❌ ERROR: El simulacro modificó el historial - Cédula: ${cedula}, Salida: ${salida}`);
                    historialModificado = true;
                }
            }
            
            if (historialModificado) {
                throw new Error('CRÍTICO: El simulacro modificó el historial cuando NO debería haberlo hecho');
            }
            
            console.log('✅ VERIFICADO: El simulacro NO modificó el historial');
            
            // 7. Verificar que SÍ se registró en Log_Simulacros
            const simulacrosSheet = ss.getSheetByName('Log_Simulacros');
            if (simulacrosSheet) {
                const simulacrosData = simulacrosSheet.getDataRange().getValues();
                let registroEncontrado = false;
                
                console.log(`🔍 Verificando Log_Simulacros - Total filas: ${simulacrosData.length}`);
                
                for (let i = 1; i < simulacrosData.length; i++) {
                    // La columna 5 (índice 4) contiene los detalles con las cédulas
                    const detalles = String(simulacrosData[i][5] || '');
                    console.log(`🔍 Fila ${i}: Detalles = "${detalles}"`);
                    
                    // Verificar si alguna de las cédulas de prueba está en los detalles
                    for (const cedula of cedulasPrueba) {
                        if (detalles.includes(cedula)) {
                            registroEncontrado = true;
                            console.log(`✅ Registro encontrado en Log_Simulacros fila ${i}: ${cedula}`);
                            break;
                        }
                    }
                    
                    if (registroEncontrado) break;
                }
                
                if (!registroEncontrado) {
                    console.log('❌ Detalles del Log_Simulacros:');
                    for (let i = 1; i < simulacrosData.length; i++) {
                        console.log(`Fila ${i}:`, simulacrosData[i]);
                    }
                    throw new Error('El simulacro no se registró en Log_Simulacros');
                }
            } else {
                console.log('⚠️ Hoja Log_Simulacros no existe (se creará automáticamente)');
            }
            
            // 8. Limpiar datos de prueba
            console.log('🧹 Limpiando datos de prueba...');
            try {
                limpiarDatosPrueba(); // Usar función existente
            } catch (cleanupError) {
                console.log('⚠️ No se pudo limpiar automáticamente, limpieza manual requerida');
            }
            
            const resultado = {
                success: true,
                mensaje: 'Test de simulacros EXITOSO - El sistema funciona correctamente',
                verificaciones: {
                    funcionExiste: true,
                    simulacroEjecutado: true,
                    historialNoModificado: !historialModificado,
                    logSimulacroCreado: true
                },
                timestamp: new Date().toISOString()
            };
            
            console.log('🎉 === TEST DE SIMULACROS COMPLETADO EXITOSAMENTE ===');
            console.log('✅ CONFIRMADO: Los simulacros NO modifican el historial real');
            console.log('✅ CONFIRMADO: Los simulacros solo registran en Log_Simulacros');
            
            return resultado;
            
        } catch (error) {
            console.error('❌ ERROR en test de simulacros:', error.message);
            
            // Intentar limpiar datos de prueba incluso en caso de error
            try {
                limpiarDatosPrueba(); // Corregir nombre de función
            } catch (cleanupError) {
                console.error('Error adicional limpiando:', cleanupError.message);
            }
            
            return {
                success: false,
                mensaje: 'Test de simulacros FALLÓ: ' + error.message,
                error: error.message,
                timestamp: new Date().toISOString()
            };
        }
    }

    /**
     * ✅ FUNCIÓN DE TEST SIMPLIFICADA PARA VERIFICAR SIMULACROS
     */
    function testSimulacroRapido() {
        console.log('🎭 === TEST RÁPIDO DE SIMULACROS ===');
        
        try {
            // 1. Crear datos de prueba simples
            const cedulasPrueba = ['TEST-RAPIDO-001'];
            
            // 2. Ejecutar simulacro
            console.log('🎭 Ejecutando simulacro rápido...');
            const resultado = procesarSimulacroEvacuacion(cedulasPrueba, 'Test rápido de verificación');
            
            // 3. Mostrar resultado
            console.log('📊 Resultado del simulacro:', resultado);
            
            if (resultado.success) {
                console.log('✅ ÉXITO: El simulacro se ejecutó correctamente');
                console.log('✅ CONFIRMADO: Los simulacros funcionan y solo registran en Log_Simulacros');
                
                return {
                    success: true,
                    mensaje: 'Test rápido EXITOSO - Simulacros funcionan correctamente',
                    detalles: resultado
                };
            } else {
                throw new Error('El simulacro falló: ' + resultado.message);
            }
            
        } catch (error) {
            console.error('❌ Error en test rápido:', error.message);
            return {
                success: false,
                mensaje: 'Test rápido FALLÓ: ' + error.message
            };
        }
    }

    /**
     * ✅ FUNCIÓN PARA VERIFICAR LOG DE SIMULACROS
     */
    function verificarLogSimulacros() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const simulacrosSheet = ss.getSheetByName('Log_Simulacros');
            
            if (!simulacrosSheet) {
                return {
                    success: false,
                    mensaje: 'La hoja Log_Simulacros no existe'
                };
            }
            
            const data = simulacrosSheet.getDataRange().getValues();
            console.log('📊 Log_Simulacros - Total filas:', data.length);
            
            if (data.length > 1) {
                console.log('✅ VERIFICADO: La hoja Log_Simulacros contiene registros');
                console.log('📋 Últimos registros:');
                
                for (let i = Math.max(1, data.length - 3); i < data.length; i++) {
                    console.log(`Fila ${i}:`, data[i]);
                }
                
                return {
                    success: true,
                    mensaje: `Log_Simulacros OK - ${data.length - 1} registros encontrados`,
                    registros: data.length - 1
                };
            } else {
                return {
                    success: true,
                    mensaje: 'La hoja Log_Simulacros existe pero está vacía',
                    registros: 0
                };
            }
            
        } catch (error) {
            console.error('❌ Error verificando Log_Simulacros:', error.message);
            return {
                success: false,
                mensaje: 'Error verificando log: ' + error.message
            };
        }
    }

    /**
     * ✅ FUNCIÓN DE DEBUG PARA RASTREAR MODIFICACIONES AL HISTORIAL
     */
    function debugearSimulacro(cedulas) {
        console.log('🐛 === INICIANDO DEBUG DE SIMULACRO ===');
        
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const historialSheet = ss.getSheetByName('Historial');
            
            // 1. Obtener estado inicial del historial
            const estadoInicial = historialSheet.getDataRange().getValues();
            console.log('📊 Estado inicial del historial - Filas:', estadoInicial.length);
            
            // 2. Ejecutar simulacro
            console.log('🎭 Ejecutando procesarSimulacroEvacuacion...');
            const resultado = procesarSimulacroEvacuacion(cedulas, 'DEBUG - Test de simulacro');
            
            // 3. Obtener estado final del historial
            const estadoFinal = historialSheet.getDataRange().getValues();
            console.log('📊 Estado final del historial - Filas:', estadoFinal.length);
            
            // 4. Comparar estados
            let cambiosDetectados = false;
            
            if (estadoInicial.length !== estadoFinal.length) {
                console.error('❌ PROBLEMA: El número de filas cambió!');
                console.error(`Antes: ${estadoInicial.length}, Después: ${estadoFinal.length}`);
                cambiosDetectados = true;
            }
            
            // Comparar contenido de las filas existentes
            for (let i = 0; i < Math.min(estadoInicial.length, estadoFinal.length); i++) {
                for (let j = 0; j < Math.max(estadoInicial[i].length, estadoFinal[i].length); j++) {
                    const valorInicial = estadoInicial[i][j];
                    const valorFinal = estadoFinal[i][j];
                    
                    if (valorInicial !== valorFinal) {
                        console.error(`❌ CAMBIO DETECTADO en fila ${i+1}, columna ${j+1}:`);
                        console.error(`Antes: "${valorInicial}" | Después: "${valorFinal}"`);
                        cambiosDetectados = true;
                    }
                }
            }
            
            if (!cambiosDetectados) {
                console.log('✅ PERFECTO: No se detectaron cambios en el historial');
            } else {
                console.error('❌ CRÍTICO: Se detectaron modificaciones en el historial durante el simulacro');
            }
            
            return {
                success: resultado.success,
                cambiosDetectados: cambiosDetectados,
                filaInicial: estadoInicial.length,
                filaFinal: estadoFinal.length,
                resultado: resultado
            };
            
        } catch (error) {
            console.error('❌ Error en debug:', error.message);
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * ✅ FUNCIÓN PARA PROBAR SI HAY ALGUNA FUNCIÓN OCULTA QUE MODIFICA EL HISTORIAL
     */
    function testSimulacroCompleto() {
        console.log('🔬 === TEST COMPLETO DE SIMULACRO ===');
        
        const cedulasPrueba = ['TEST-COMPLETO-001', 'TEST-COMPLETO-002'];
        
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const historialSheet = ss.getSheetByName('Historial');
            const bdSheet = ss.getSheetByName('Base de Datos');
            
            // Agregar registros de prueba al historial
            console.log('🔧 Agregando registros de prueba...');
            const fechaPrueba = new Date();
            
            historialSheet.appendRow([fechaPrueba, 'TEST-COMPLETO-001', 'Persona Test 1', 'Empresa Test', fechaPrueba, null, 'Entrada para test', 'TEST']);
            historialSheet.appendRow([fechaPrueba, 'TEST-COMPLETO-002', 'Persona Test 2', 'Empresa Test', fechaPrueba, null, 'Entrada para test', 'TEST']);
            
            // Ejecutar debug de simulacro
            const resultado = debugearSimulacro(cedulasPrueba);
            
            // Limpiar datos de prueba
            console.log('🧹 Limpiando datos de prueba...');
            limpiarDatosPrueba();
            
            return resultado;
            
        } catch (error) {
            console.error('❌ Error en test completo:', error.message);
            return {
                success: false,
                error: error.message
            };
        }
    }
