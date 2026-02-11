/**
 * Script para verificar que las animaciones estén funcionando correctamente
 * Este script se puede incluir en cualquier página para diagnosticar problemas
 */

(function() {
    'use strict';
    
    // Verificar si el DOM está listo
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initAnimationChecker);
    } else {
        initAnimationChecker();
    }
    
    function initAnimationChecker() {
        // Crear botón de diagnóstico
        createDiagnosticButton();
        
        // Verificar animaciones automáticamente
        setTimeout(checkAnimations, 2000);
    }
    
    function createDiagnosticButton() {
        // Solo crear el botón si estamos en una página de test
        if (window.location.href.includes('test') || window.location.href.includes('verify')) {
            const button = document.createElement('button');
            button.id = 'animationChecker';
            button.textContent = '🔍 Verificar Animaciones';
            button.style.cssText = `
                position: fixed;
                bottom: 20px;
                right: 20px;
                background: #2A56A1;
                color: white;
                border: none;
                padding: 12px 20px;
                border-radius: 30px;
                cursor: pointer;
                font-weight: bold;
                box-shadow: 0 4px 12px rgba(0,0,0,0.2);
                z-index: 10000;
            `;
            
            button.addEventListener('click', checkAnimations);
            document.body.appendChild(button);
        }
    }
    
    function checkAnimations() {
        console.log('🧪 Iniciando verificación de animaciones...');
        
        const results = {
            excelIcon: false,
            progressBar: false,
            rowFill: false,
            indicators: false
        };
        
        // Verificar animación del icono Excel
        const excelIcon = document.querySelector('.excel-icon');
        if (excelIcon) {
            const computedStyle = window.getComputedStyle(excelIcon);
            const animationName = computedStyle.animationName;
            results.excelIcon = animationName && animationName !== 'none';
            console.log('📊 Icono Excel animación:', results.excelIcon ? '✅ ACTIVA' : '❌ INACTIVA');
        }
        
        // Verificar animación de barra de progreso
        const progressFill = document.querySelector('.progress-fill');
        if (progressFill) {
            const computedStyle = window.getComputedStyle(progressFill);
            const animationName = computedStyle.animationName;
            results.progressBar = animationName && animationName !== 'none';
            console.log('📊 Barra de progreso animación:', results.progressBar ? '✅ ACTIVA' : '❌ INACTIVA');
        }
        
        // Verificar animación de filas
        const excelRows = document.querySelectorAll('.excel-row::after');
        if (excelRows.length > 0) {
            // Verificar al menos una fila
            const firstRow = document.querySelector('.excel-row');
            if (firstRow) {
                const computedStyle = window.getComputedStyle(firstRow, '::after');
                const animationName = computedStyle.animationName;
                results.rowFill = animationName && animationName !== 'none';
                console.log('📊 Filas animación:', results.rowFill ? '✅ ACTIVA' : '❌ INACTIVA');
            }
        }
        
        // Verificar indicadores activos
        const activeIndicators = document.querySelectorAll('.indicator.active');
        results.indicators = activeIndicators.length > 0;
        console.log('📊 Indicadores activos:', results.indicators ? '✅ ACTIVOS' : '❌ INACTIVOS');
        
        // Mostrar resultados
        showResults(results);
        
        return results;
    }
    
    function showResults(results) {
        const allWorking = Object.values(results).every(result => result === true);
        
        const message = allWorking 
            ? '🎉 ¡Todas las animaciones están funcionando correctamente!' 
            : '⚠️ Algunas animaciones pueden no estar funcionando';
            
        const bgColor = allWorking ? '#d4edda' : '#fff3cd';
        const borderColor = allWorking ? '#c3e6cb' : '#ffeaa7';
        const textColor = allWorking ? '#155724' : '#856404';
        
        // Crear notificación
        const notification = document.createElement('div');
        notification.innerHTML = `
            <div style="
                position: fixed;
                top: 20px;
                left: 50%;
                transform: translateX(-50%);
                background: ${bgColor};
                border: 1px solid ${borderColor};
                color: ${textColor};
                padding: 15px 25px;
                border-radius: 8px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                z-index: 10001;
                font-weight: 500;
                max-width: 90%;
                text-align: center;
            ">
                ${message}
                <div style="font-size: 0.9em; margin-top: 8px;">
                    Excel: ${results.excelIcon ? '✅' : '❌'} | 
                    Progreso: ${results.progressBar ? '✅' : '❌'} | 
                    Filas: ${results.rowFill ? '✅' : '❌'} | 
                    Indicadores: ${results.indicators ? '✅' : '❌'}
                </div>
            </div>
        `;
        
        document.body.appendChild(notification);
        
        // Eliminar notificación después de 5 segundos
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 5000);
        
        console.log('📋 Resultados completos:', results);
    }
    
    // Exponer función globalmente para pruebas manuales
    window.checkAnimations = checkAnimations;
    
})();