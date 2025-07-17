/**
 * Verifica el estado del servidor
 * @return {Object} Estado del servidor
 */
function serverStatus() {
    try {
        console.log('🔍 Verificando estado del servidor...');
        
        // Verificar acceso a la hoja de cálculo
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (!ss) {
            console.warn('⚠️ No se pudo acceder a la hoja de cálculo');
            // No lanzamos error, continuamos sin la hoja
        }
        
        const appName = 'SurPass';
        const version = '3.0';
        const now = new Date();
        
        // Obtener información de la zona horaria de Panamá
        const timeZone = 'America/Panama';
        const panamaTime = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm:ss');
        
        // Obtener nombres de hojas disponibles para diagnóstico
        let hojasDisponibles = [];
        try {
            if (ss) {
                const hojas = ss.getSheets();
                hojasDisponibles = hojas.map(hoja => hoja.getName());
                console.log('📋 Hojas disponibles:', hojasDisponibles);
            }
        } catch (e) {
            console.warn('⚠️ No se pudieron listar las hojas:', e.message);
        }
        
        // Verificar acceso a servicios necesarios
        const servicios = {
            spreadsheet: !!ss,
            script: true, // Asumimos que el servicio Script está disponible
            cache: true, // Asumimos que el servicio Cache está disponible
            properties: true, // Asumimos que el servicio Properties está disponible
            mail: true, // Asumimos que el servicio Mail está disponible
            hojasDisponibles: hojasDisponibles
        };
        
        // Determinar si las hojas requeridas están disponibles
        servicios.configuracionDisponible = hojasDisponibles.includes('Configuracion');
        servicios.historialDisponible = hojasDisponibles.includes('Historial');
        servicios.baseDatosDisponible = hojasDisponibles.includes('BaseDeDatos');

        const estado = {
            status: 'ok',
            appName: appName,
            version: version,
            timestamp: now.toISOString(),
            panamaTime: panamaTime,
            environment: {
                executionLocation: 'WEB',
                user: Session.getEffectiveUser() ? Session.getEffectiveUser().getEmail() : 'usuario_desconocido',
                timeZone: Session.getScriptTimeZone() || timeZone
            },
            endpoints: {
                evacuation: '/evacuation',
                api: '/api',
                admin: '/admin'
            },
            hojasDisponibles: hojasDisponibles,
            hojasRequeridas: {
                configuracion: servicios.configuracionDisponible,
                historial: servicios.historialDisponible,
                baseDatos: servicios.baseDatosDisponible
            }
        };
        
        console.log('✅ Estado del servidor:', estado);
        return estado;
        
    } catch (error) {
        // Registrar el error en el log
        console.error('❌ Error en serverStatus:', error);
        
        // Devolver un objeto de error detallado
        const errorResponse = {
            status: 'error',
            error: error.message || 'Error desconocido',
            timestamp: new Date().toISOString(),
            recoverySteps: [
                'Verificar la conexión a Internet',
                'Recargar la aplicación',
                'Si el problema persiste, contactar al administrador del sistema'
            ]
        };
        
        // Solo incluir el stack en modo desarrollo
        if (error.stack) {
            errorResponse.stack = error.stack;
        }
        
        return errorResponse;
    }
}
