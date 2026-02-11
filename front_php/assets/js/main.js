(function() {
  "use strict";
  
  // API base provided by index.php via window.API_BASE
  const API_BASE = window.API_BASE || '';
  
  // TRM cache global (se persiste en localStorage)
  let LAST_TRM = { usd: 0, eur: 0 };
  function actualizarCamposModeloDeudaTRM() {
    const usdEl = document.getElementById('usd_md');
    const eurEl = document.getElementById('eur_md');
    if (usdEl) {
      usdEl.value = nfPlain.format(Number(LAST_TRM.usd || 0));
    }
    if (eurEl) {
      eurEl.value = nfPlain.format(Number(LAST_TRM.eur || 0));
    }
  }

  // Eliminamos el uso de localStorage para evitar cache
  console.log('🚀 Inicializando aplicación...');
  console.log('🌐 API_BASE:', API_BASE);
  
  // Variables globales para controlar requests
  let trmController = new AbortController();
  let formController = null;
  let isUpdatingTRM = false;
  
  // Performance monitoring
  const performanceTracker = {
    startTimes: {},
    metrics: [],
  
    start(operation) {
      this.startTimes[operation] = performance.now();
    },
  
    end(operation) {
      if (this.startTimes[operation]) {
        const duration = performance.now() - this.startTimes[operation];
        this.metrics.push({
          operation,
          duration: Math.round(duration),
          timestamp: new Date().toISOString(),
        });
        delete this.startTimes[operation];
        console.log(`⏱️ ${operation}: ${duration.toFixed(0)}ms`);
        return duration;
      }
      return 0;
    },
  
    getStats() {
      return {
        totalOperations: this.metrics.length,
        averageTime:
          this.metrics.length > 0
            ? Math.round(
                this.metrics.reduce((sum, m) => sum + m.duration, 0) /
                  this.metrics.length
              )
            : 0,
        recentMetrics: this.metrics.slice(-10),
      };
    },
  };
  
  // ===== Helpers de formato y parseo numérico =====
  const nfES2 = new Intl.NumberFormat('es-CO', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  const nfESflex = new Intl.NumberFormat('es-CO', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
    useGrouping: true,
  });
  const nfPlain = new Intl.NumberFormat('es-CO', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
    useGrouping: false,
  });
  
  function parseNumeroFlexible(input) {
    if (input === null || input === undefined) return 0;
    let s = String(input).trim();
    if (!s) return 0;
    s = s.replace(/\s+/g, '');
    const hasComma = s.includes(',');
    const hasDot = s.includes('.');
    let decimalSep = '.';
    let thousandSep = '';
  
    if (hasComma && hasDot) {
      decimalSep = s.lastIndexOf(',') > s.lastIndexOf('.') ? ',' : '.';
      thousandSep = decimalSep === ',' ? '.' : ',';
    } else if (hasComma) {
      decimalSep = ',';
    } else if (hasDot) {
      decimalSep = '.';
    }
  
    if (thousandSep) {
      const reThousand = new RegExp('\\' + thousandSep, 'g');
      s = s.replace(reThousand, '');
    }
    if (decimalSep === ',') s = s.replace(/,/g, '.');
    const n = Number(s);
    return isFinite(n) ? n : 0;
  }
  
  // ================= Loader mejorado - versión simple =================
  function mostrarLoader(mensaje = 'Procesando...', tipo = 'general') {
    const loader = document.getElementById('processing-loader');
    const loaderMessage = document.getElementById('loader-message');
    const loaderSubmessage = document.getElementById('loader-submessage');
    const progressStatus = document.getElementById('progress-status');
    
    // Mensajes personalizados según el tipo de proceso
    const mensajes = {
      'cartera': {
        titulo: 'Procesando Cartera',
        subtitulo: 'Generando Excel con vencimientos...',
        estados: ['Leyendo CSV...', 'Calculando...', 'Creando Excel...', 'Finalizando...']
      },
      'anticipos': {
        titulo: 'Procesando Anticipos',
        subtitulo: 'Convirtiendo valores...',
        estados: ['Leyendo CSV...', 'Procesando...', 'Generando Excel...', 'Finalizando...']
      },
      'modelo_deuda': {
        titulo: 'Generando Modelo de Deuda',
        subtitulo: 'Consolidando datos...',
        estados: ['Cargando archivos...', 'Aplicando TRM...', 'Consolidando...', 'Finalizando...']
      },
      'focus': {
        titulo: 'Actualizando FOCUS',
        subtitulo: 'Procesando archivos...',
        estados: ['Leyendo archivos...', 'Actualizando...', 'Aplicando TRM...', 'Finalizando...']
      },
      'general': {
        titulo: mensaje,
        subtitulo: 'Esto puede tomar unos segundos...',
        estados: ['Preparando...', 'Procesando...', 'Generando...', 'Finalizando...']
      }
    };
    
    const config = mensajes[tipo] || mensajes['general'];
    
    if (loaderMessage) loaderMessage.textContent = config.titulo;
    if (loaderSubmessage) loaderSubmessage.textContent = config.subtitulo;
    if (progressStatus) progressStatus.textContent = config.estados[0];
    
    // Mostrar loader
    if (loader) {
      loader.style.display = 'flex';
      document.body.style.overflow = 'hidden';
    }
    
    // Animar indicadores de progreso después de un pequeño delay
    setTimeout(() => {
      animarIndicadores(config.estados);
    }, 300);
  }
  
  function animarIndicadores(estados) {
    const indicators = document.querySelectorAll('.indicator');
    const progressStatus = document.getElementById('progress-status');
    let currentStep = 0;
    
    // Limpiar estados anteriores
    indicators.forEach(ind => ind.classList.remove('active'));
    
    // Animar cada paso
    const interval = setInterval(() => {
      if (currentStep < indicators.length) {
        // Activar indicador actual
        if (indicators[currentStep]) {
          indicators[currentStep].classList.add('active');
        }
        
        // Actualizar texto de estado
        if (progressStatus && estados[currentStep]) {
          progressStatus.textContent = estados[currentStep];
        }
        
        currentStep++;
      } else {
        clearInterval(interval);
      }
    }, 1000); // Cambiar cada segundo
    
    // Guardar el interval para poder limpiarlo después
    const loader = document.getElementById('processing-loader');
    if (loader) {
      loader.animationInterval = interval;
    }
  }
  
  function ocultarLoader() {
    const loader = document.getElementById('processing-loader');
    if (!loader) return;
    
    // Limpiar interval de animación si existe
    if (loader.animationInterval) {
      clearInterval(loader.animationInterval);
      loader.animationInterval = null;
    }
    
    // Limpiar estados de indicadores
    const indicators = document.querySelectorAll('.indicator');
    indicators.forEach(ind => ind.classList.remove('active'));
    
    // Ocultar loader con transición
    loader.style.opacity = '0';
    loader.style.visibility = 'hidden';
    
    setTimeout(() => {
      loader.style.display = 'none';
      document.body.style.overflow = '';
    }, 300);
  }

  // ================= TRM =================
  async function actualizarTRM() {
    const trmLabel = document.getElementById('trm-label');
    if (!trmLabel) {
        console.log('⚠️  Elemento TRM no encontrado');
        return;
    }
  
    if (isUpdatingTRM) {
        console.log('⏳ Ya hay una actualización de TRM en curso');
        return;
    }
  
    isUpdatingTRM = true;
    console.log('🔄 Iniciando actualización de TRM...');
    performanceTracker.start('actualizarTRM');
  
    try {
        if (trmController && trmController.signal && trmController.signal.aborted) {
            trmController = new AbortController();
        }
  
        // Usar ruta directa a trm.php con ruta relativa
        // Evitar cache agregando timestamp
        const timestamp = new Date().getTime();
        const base = window.API_BASE || '';
        const separator = base && !base.endsWith('/') ? '/' : '';
        const trmUrl = `${base}${separator}trm.php?t=${timestamp}`;
        console.log('🔍 Solicitando TRM desde:', trmUrl);

        const res = await fetch(trmUrl, {
            method: 'GET',
            signal: trmController.signal,
            headers: {
                'Accept': 'application/json',
                'Cache-Control': 'no-cache',
                'Pragma': 'no-cache'
            },
            credentials: 'same-origin'
        });
  
        if (!res.ok) {
            throw new Error(`Error ${res.status}: ${res.statusText}`);
        }
  
        const data = await res.json();
        console.log('✅ TRM actualizada desde backend:', data);
  
        const usdInput = document.querySelector('input[name="usd"]');
        const eurInput = document.querySelector('input[name="eur"]');
  
        if (usdInput && data.usd) {
            usdInput.value = nfPlain.format(Number(data.usd));
        }
        if (eurInput && data.eur) {
            eurInput.value = nfPlain.format(Number(data.eur));
        }

        LAST_TRM = {
            usd: Number(data.usd || 0),
            eur: Number(data.eur || 0)
        };
        actualizarCamposModeloDeudaTRM();
  
        if (trmLabel) {
            const usdTxt = nfPlain.format(Number(data.usd || 0));
            const eurTxt = nfPlain.format(Number(data.eur || 0));
            trmLabel.textContent = `TRM USD: ${usdTxt} | EUR: ${eurTxt}`;
        }

        showNotification('✅ TRM actualizada correctamente', 'success');
    } catch (error) {
        if (error.name === 'AbortError') {
            console.log('ℹ️  Solicitud TRM cancelada por el usuario');
            return;
        }
        console.error('❌ Error al actualizar TRM desde backend:', error);
  
        // No usamos localStorage, mostramos error directamente
        console.error('❌ No se pudo cargar TRM. Usando valores por defecto.');
        showNotification('❌ Error al cargar TRM. Usando valores por defecto.', 'error');
  
        const defaultData = { usd: 4000.0, eur: 4500.0 };
        LAST_TRM = defaultData;
        actualizarCamposModeloDeudaTRM();

        if (trmLabel) {
            const usdTxt = nfPlain.format(Number(defaultData.usd));
            const eurTxt = nfPlain.format(Number(defaultData.eur));
            trmLabel.textContent = `TRM USD: ${usdTxt} | EUR: ${eurTxt}`;
        }
    } finally {
        performanceTracker.end('actualizarTRM');
        isUpdatingTRM = false;
        console.log('✅ Finalizada actualización de TRM');
    }
  }
  
  function abrirTRM() {
    console.log('🔓 Abriendo modal TRM');
    const modal = document.getElementById('modalTRM');
    if (modal) {
      modal.classList.add('show');
      modal.style.display = 'flex';
    }
  
    actualizarTRM().catch((error) => {
      console.error('Error al actualizar TRM automáticamente:', error);
    });
  }
  
  function cerrarTRM() {
    console.log('🔒 Cerrando modal TRM');
    const modal = document.getElementById('modalTRM');
    if (modal) {
      modal.classList.remove('show');
      modal.style.display = 'none';
    }
  }
  
  // Hacer funciones globales para el HTML
  window.abrirTRM = abrirTRM;
  window.cerrarTRM = cerrarTRM;
  
  // ================= Formularios =================
  function mostrarFormulario() {
    const val = document.getElementById('proceso')?.value;
    const cont = document.getElementById('formularios');
  
    console.log('📝 Mostrando formulario para:', val);
    console.log('🔍 Contenedor encontrado:', !!cont);
  
    if (!cont) {
      console.error('❌ Elemento formularios no encontrado');
      return;
    }
  
    cont.innerHTML = '';
  
    if (!val) {
      cont.innerHTML = `
        <div class="form-container">
          <div class="process-card">
            <div class="process-icon">
              <i class="fas fa-cogs"></i>
            </div>
            <h6 style="font-size: 1.2rem; margin-bottom: 10px; color: var(--primary-color);">Sistema de Procesamiento de Cartera</h6>
            <p style="color: var(--text-light); margin-bottom: 20px; text-align: center;">
              Seleccione un proceso de la lista superior para comenzar.<br><br>
              <strong>Procesos disponibles:</strong><br>
              • <strong>Cartera:</strong> Procesar archivos de cartera<br>
              • <strong>Anticipos:</strong> Procesar archivos de anticipos<br>
              • <strong>Modelo Deuda:</strong> Generar modelo consolidado<br>
              • <strong>FOCUS:</strong> Procesar archivos FOCUS<br>
            </p>
          </div>
        </div>`;
      return;
    }
  
    if (val === 'cartera') {
      cont.innerHTML = `
        <div class="form-container">
          <div class="process-card">
            <div class="process-icon">
              <i class="fas fa-file-invoice"></i>
            </div>
            <h6 style="font-size: 1.2rem; margin-bottom: 10px; color: var(--primary-color);">Procesamiento de Cartera</h6>
            <p style="color: var(--text-light); margin-bottom: 20px;">Procesa archivos de cartera con saldos, facturas y vencimientos.</p>
          </div>
          <form id="formCartera">
            <div class="form-row">
              <label><i class="fas fa-upload"></i> Archivo de Cartera:</label>
              <div class="file-input-wrapper">
                <input type="file" name="cartera" id="cartera" required>
                <label for="cartera" class="file-input-label">
                  <i class="fas fa-cloud-upload-alt"></i>
                  <span>Seleccionar archivo de cartera</span>
                </label>
              </div>
            </div>
            <div class="form-row">
              <label><i class="fas fa-calendar"></i> Fecha de Cierre:</label>
              <input type="date" name="fecha_cierre" required>
            </div>
            <button type="submit">
              <i class="fas fa-play"></i>
              Procesar Cartera
            </button>
          </form>
        </div>`;
    } else if (val === 'anticipos') {
      cont.innerHTML = `
        <div class="form-container">
          <div class="process-card">
            <div class="process-icon">
              <i class="fas fa-coins"></i>
            </div>
            <h6 style="font-size: 1.2rem; margin-bottom: 10px; color: var(--primary-color);">Procesamiento de Anticipos</h6>
            <p style="color: var(--text-light); margin-bottom: 20px;">Procesa archivos de anticipos de clientes en formato CSV o Excel.</p>
          </div>
          <form id="formAnticipos">
            <div class="form-row">
              <label><i class="fas fa-upload"></i> Archivo de Anticipos:</label>
              <div class="file-input-wrapper">
                <input type="file" name="anticipos" id="anticipos" accept=".csv,.xlsx" required>
                <label for="anticipos" class="file-input-label">
                  <i class="fas fa-cloud-upload-alt"></i>
                  <span>Seleccionar archivo de anticipos</span>
                </label>
              </div>
            </div>
            <button type="submit">
              <i class="fas fa-play"></i>
              Procesar Anticipos
            </button>
          </form>
        </div>`;
    } else if (val === 'modelo_deuda') {
      cont.innerHTML = `
        <div class="form-container">
          <div class="process-card">
            <div class="process-icon">
              <i class="fas fa-chart-area"></i>
            </div>
            <h6 style="font-size: 1.2rem; margin-bottom: 10px; color: var(--primary-color);">Modelo de Deuda</h6>
            <p style="color: var(--text-light); margin-bottom: 20px;">Genera un modelo consolidado con cartera, anticipos y tasas de cambio.</p>
          </div>
          <form id="formModeloDeuda">
            <div class="form-row">
              <label><i class="fas fa-upload"></i> Archivo de Cartera:</label>
              <div class="file-input-wrapper">
                <input type="file" name="cartera" id="cartera_md" required>
                <label for="cartera_md" class="file-input-label">
                  <i class="fas fa-cloud-upload-alt"></i>
                  <span>Seleccionar archivo de cartera</span>
                </label>
              </div>
            </div>
            <div class="form-row">
              <label><i class="fas fa-upload"></i> Archivo de Anticipos:</label>
              <div class="file-input-wrapper">
                <input type="file" name="anticipos" id="anticipos_md" required>
                <label for="anticipos_md" class="file-input-label">
                  <span>Seleccionar archivo de anticipos</span>
                </label>
              </div>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
              <div class="form-row">
                <label><i class="fas fa-dollar-sign"></i> USD (TRM actual):</label>
                <input type="text" name="usd" id="usd_md" placeholder="Se toma de TRM" disabled>
              </div>
              <div class="form-row">
                <label><i class="fas fa-euro-sign"></i> EUR (TRM actual):</label>
                <input type="text" name="eur" id="eur_md" placeholder="Se toma de TRM" disabled>
              </div>
            </div>
            <button type="submit">
              <i class="fas fa-play"></i>
              Procesar Modelo Deuda
            </button>
          </form>
        </div>`;
      actualizarCamposModeloDeudaTRM();
    } else if (val === 'focus') {
      cont.innerHTML = `
        <div class="form-container">
          <div class="process-card">
            <div class="process-icon">
              <i class="fas fa-crosshairs"></i>
            </div>
            <h6 style="font-size: 1.2rem; margin-bottom: 10px; color: var(--primary-color);">Procesamiento FOCUS</h6>
            <p style="color: var(--text-light); margin-bottom: 20px; text-align: center;">
              Procese los siguientes archivos para actualizar el archivo FOCUS:<br>
              <span style="font-size: 0.9em; color: #ffcc00;"><i class="fas fa-exclamation-circle"></i> Todos los archivos son obligatorios</span>
            </p>
          </div>
          <form id="formFocus">
            <div style="display: grid; grid-template-columns: minmax(300px, 600px); gap: 15px; margin: 0 auto; max-width: 600px;">
              <div class="form-row">
                <label><i class="fas fa-balance-scale"></i> Archivo BALANCE:</label>
                <div class="file-input-wrapper">
                  <input type="file" name="balance" id="balance" accept=".xlsx,.xls" required>
                  <label for="balance" class="file-input-label">
                    <i class="fas fa-cloud-upload-alt"></i>
                    <span>Seleccionar archivo BALANCE</span>
                  </label>
                </div>
                <small class="file-info">Archivo que contiene el balance general</small>
              </div>
              
              <div class="form-row">
                <label><i class="fas fa-chart-pie"></i> Archivo SITUACIÓN:</label>
                <div class="file-input-wrapper">
                  <input type="file" name="situacion" id="situacion" accept=".xlsx,.xls" required>
                  <label for="situacion" class="file-input-label">
                    <i class="fas fa-cloud-upload-alt"></i>
                    <span>Seleccionar archivo SITUACIÓN</span>
                  </label>
                </div>
                <small class="file-info">Archivo con la situación financiera</small>
              </div>
              
              <div class="form-row">
                <label><i class="fas fa-layer-group"></i> Archivo ACUMULADO PRUEBA :</label>
                <div class="file-input-wrapper">
                  <input type="file" name="acumulado" id="acumulado" accept=".xlsx,.xls" required>
                  <label for="acumulado" class="file-input-label">
                    <i class="fas fa-cloud-upload-alt"></i>
                    <span>Seleccionar archivo ACUMULADO PRUEBA </span>
                  </label>
                </div>
                <small class="file-info">Archivo con los valores acumulados</small>
              </div>
              
              <div class="form-row">
                <label><i class="fas fa-crosshairs"></i> Archivo FOCUS MES ANTERIOR:</label>
                <div class="file-input-wrapper">
                  <input type="file" name="focus" id="focus" accept=".xlsx,.xls" required>
                  <label for="focus" class="file-input-label">
                    <i class="fas fa-cloud-upload-alt"></i>
                    <span>Seleccionar archivo FOCUS MES ANTERIOR</span>
                  </label>
                </div>
                <small class="file-info">Archivo FOCUS a actualizar</small>
              </div>
              
              <div class="form-row">
                <label><i class="fas fa-file-invoice"></i> Archivo MODELO DEUDA:</label>
                <div class="file-input-wrapper">
                  <input type="file" name="modelo_deuda" id="modelo_deuda" accept=".xlsx,.xls" required>
                  <label for="modelo_deuda" class="file-input-label">
                    <i class="fas fa-cloud-upload-alt"></i>
                    <span>Seleccionar archivo MODELO DEUDA</span>
                  </label>
                </div>
                <small class="file-info">Archivo con los vencimientos de cartera</small>
              </div>
              <div id="valores_vencimientos" style="display: none; margin-top: 20px; background: #f8f9fa; padding: 15px; border-radius: 5px; border-left: 4px solid var(--primary-color);">
                <h6 style="margin-top: 0; color: var(--primary-color);"><i class="fas fa-calculator"></i> Valores de Vencimientos</h6>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 10px;">
                  <div><strong>Vencido 30 días:</strong> <span id="vencido_30">-</span></div>
                  <div><strong>Vencido 60+ días:</strong> <span id="vencido_60">-</span></div>
                  <div><strong>Total Vencido:</strong> <span id="total_vencido">-</span></div>
                  <div><strong>Provisión:</strong> <span id="provision">-</span></div>
                </div>
              </div>
            </div>
            <div style="margin-top: 20px; text-align: center;">
              <button type="submit" style="padding: 12px 30px; font-size: 1.1em;">
                <i class="fas fa-play"></i>
                Procesar FOCUS
              </button>
            </div>
          </form>
        </div>`;
      console.log('✅ Formulario FOCUS generado');
    } else {
      console.warn('⚠️ Opción no reconocida:', val);
      cont.innerHTML = `
        <div class="form-container">
          <div class="process-card" style="border-left: 4px solid var(--warning-color);">
            <div class="process-icon">
              <i class="fas fa-exclamation-triangle"></i>
            </div>
            <h6 style="font-size: 1.2rem; margin-bottom: 10px; color: var(--warning-color);">Opción no válida</h6>
            <p style="color: var(--text-light); margin-bottom: 20px;">La opción seleccionada no es válida o no está implementada.</p>
            <p style="color: var(--text-light); font-size: 0.9em;">Valor recibido: <strong>${val}</strong></p>
          </div>
        </div>`;
    }
  
    configurarFormularios();
    console.log('🔧 Eventos de formularios configurados');
    console.log('✅ mostrarFormulario() completado exitosamente');
  }
  
  // ================= File input labels =================
  function updateFileLabel(input) {
    if (!input) return;
    const label = input.nextElementSibling;
    if (!label) return;
    const span = label.querySelector('span');
    if (!span) return;
    
    if (input.files.length > 0) {
      span.textContent = input.files[0].name;
      label.classList.add('has-file');
      const icon = label.querySelector('i');
      if (icon) icon.className = 'fas fa-check-circle';
    } else {
      span.textContent = 'Seleccionar archivo';
      label.classList.remove('has-file');
      const icon = label.querySelector('i');
      if (icon) icon.className = 'fas fa-cloud-upload-alt';
    }
  }
  
  // ================= Notificaciones =================
  function showNotification(message, type = 'info') {
    let notification = document.getElementById('notification');
    if (!notification) {
      notification = document.createElement('div');
      notification.id = 'notification';
      notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 16px 24px;
        border-radius: 12px;
        color: white;
        font-weight: 600;
        z-index: 10001;
        transform: translateX(400px);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 8px 32px rgba(0,0,0,0.15);
        max-width: 400px;
      `;
      document.body.appendChild(notification);
    }
  
    const colors = {
      success: 'linear-gradient(135deg, #00d084, #00a86b)',
      error: 'linear-gradient(135deg, #ff3b30, #dc2626)',
      info: 'linear-gradient(135deg, #0055a5, #003366)',
    };
  
    notification.style.background = colors[type] || colors.info;
    notification.textContent = message;
    notification.style.transform = 'translateX(0)';
  
    setTimeout(() => {
      notification.style.transform = 'translateX(400px)';
    }, 4000);
  }
  
  // ================= Utilidad: Descargar archivo =================
  async function descargarArchivo(response, defaultFilename) {
    const dispo = response.headers.get('Content-Disposition') || '';
  
    let filename = defaultFilename;
    const matchUtf8 = dispo.match(/filename\*=UTF-8''([^;]+)/i);
    const matchAscii = dispo.match(/filename="?([^";]+)"?/i);
  
    if (matchUtf8 && matchUtf8[1]) {
      filename = decodeURIComponent(matchUtf8[1].trim());
    } else if (matchAscii && matchAscii[1]) {
      filename = matchAscii[1].trim();
    }
    // Sanitizar nombre (evitar sufijos raros como 'xlsx_' o espacios)
    try {
      filename = filename.replace(/[\r\n\t\0]/g, '').trim();
      // Si termina en guión bajo o punto, recortarlo
      filename = filename.replace(/[_.]+$/g, '.xlsx');
      // Asegurar extensión .xlsx
      if (!/\.xlsx$/i.test(filename)) {
        filename = filename.replace(/\.+$/g, '') + '.xlsx';
      }
    } catch (_) {}
  
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
  
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
  
    window.URL.revokeObjectURL(url);
  }
  
  // ================= Configurar formularios =================
  function configurarFormularios() {
    const fileInputs = document.querySelectorAll('input[type="file"]');
    fileInputs.forEach((input) => {
      if (!input.dataset.listenerAttached) {
        input.addEventListener('change', function () {
          updateFileLabel(this);
        });
        input.dataset.listenerAttached = 'true';
      }
    });
  
    const formContainer = document.getElementById('formularios');
  
    if (formContainer && !formContainer.dataset.submitListenerAttached) {
      async function enviarArchivo(endpoint, formData, nombreDescarga, mensajeExito, submitBtn, tipoLoader = 'general') {
        try {
            if (submitBtn) submitBtn.disabled = true;
            mostrarLoader(mensajeExito, tipoLoader);
            
            // Verificar que el FormData tenga el archivo
            const hasFile = formData.has('archivo') || formData.has('provision') || formData.has('anticipos') || 
                           formData.has('balance') || formData.has('situacion') || formData.has('focus') || 
                           formData.has('acumulado') || formData.has('modelo');
            
            if (!hasFile) {
                throw new Error('No se detectó ningún archivo en el formulario. Por favor, seleccione un archivo.');
            }
            
            const res = await fetch(endpoint, { method: 'POST', body: formData });

            if (!res.ok) {
                // Clonar la respuesta para poder leerla múltiples veces
                const clonedRes = res.clone();
                let errorMsg;
                let errorDetails = null;
                try {
                    const errorData = await res.json();
                    errorMsg = errorData.error || 'Error en el servidor';
                    errorDetails = errorData;
                    console.error('❌ Detalles del error:', errorData);
                } catch {
                    try {
                        errorMsg = await clonedRes.text();
                    } catch {
                        errorMsg = `Error ${res.status} en el servidor`;
                    }
                }
                
                // Mensaje más descriptivo según el tipo de error
                if (errorDetails && errorDetails.upload_error_code) {
                    errorMsg = errorDetails.error || errorMsg;
                }
                
                throw new Error(errorMsg);
            }

            await descargarArchivo(res, nombreDescarga);
            showNotification(`✅ ${mensajeExito}`, 'success');
        } catch (err) {
            console.error('❌ Error al ejecutar Python:', err.message);
            console.error('❌ Stack trace:', err.stack);
            showNotification(`❌ Error: ${err.message}`, 'error');
        } finally {
            ocultarLoader();
            if (submitBtn) submitBtn.disabled = false;
        }
      }
  
      formContainer.addEventListener('submit', async function (e) {
        e.preventDefault();
        const form = e.target;
        const formId = form.id;
        const submitBtn = form.querySelector('button[type="submit"]');
  
        if (formId === 'formCartera') {
          const formData = new FormData(form);
          formData.append('tipo', 'cartera');
          const fileInput = form.querySelector('input[type="file"]');
          if (!fileInput || fileInput.files.length === 0) {
            showNotification('Por favor seleccione un archivo', 'error');
            return;
          }
          formData.delete('cartera');
          formData.append('archivo', fileInput.files[0]);
  
          await enviarArchivo('procesar.php', formData, 'cartera.xlsx', 'Archivo de cartera procesado', submitBtn, 'cartera');
  
        } else if (formId === 'formAnticipos') {
          const formData = new FormData();
          formData.append('tipo', 'anticipos');
          const fileInput = form.querySelector('input[name="anticipos"]');
          if (!fileInput || fileInput.files.length === 0) {
            showNotification('Por favor seleccione el archivo de anticipos', 'error');
            return;
          }
          formData.append('archivo', fileInput.files[0]);
  
          await enviarArchivo('procesar.php', formData, 'anticipos.xlsx', 'Archivo de anticipos procesado', submitBtn, 'anticipos');
  
        } else if (formId === 'formModeloDeuda') {
          const formData = new FormData();
          formData.append('tipo', 'modelo_deuda');
          const carteraFile = document.getElementById('cartera_md')?.files[0];
          const anticiposFile = document.getElementById('anticipos_md')?.files[0];
  
          if (!carteraFile || !anticiposFile) {
            showNotification('Por favor complete todos los campos correctamente', 'error');
            return;
          }
  
          formData.append('provision', carteraFile);
          formData.append('anticipos', anticiposFile);
  
          await enviarArchivo('procesar.php', formData, 'modelo_deuda.xlsx', 'Modelo de deuda procesado', submitBtn, 'modelo_deuda');
  
        } else if (formId === 'formFocus') {
          const formData = new FormData();
          formData.append('tipo', 'focus');
          const balanceFile = document.getElementById('balance')?.files[0];
          const situacionFile = document.getElementById('situacion')?.files[0];
          const focusFile = document.getElementById('focus')?.files[0];
          const acumuladoFile = document.getElementById('acumulado')?.files[0];
          const modeloDeudaFile = document.getElementById('modelo_deuda')?.files[0];
  
          const requiredFiles = [
            { file: balanceFile, name: 'BALANCE' },
            { file: situacionFile, name: 'SITUACIÓN' },
            { file: focusFile, name: 'FOCUS' },
            { file: acumuladoFile, name: 'ACUMULADO' },
            { file: modeloDeudaFile, name: 'MODELO DEUDA' }
          ];
  
          const missingFiles = requiredFiles.filter(item => !item.file).map(item => item.name);
          if (missingFiles.length > 0) {
            showNotification(`❌ Faltan archivos obligatorios: ${missingFiles.join(', ')}`, 'error');
            return;
          }
  
          formData.append('balance', balanceFile);
          formData.append('situacion', situacionFile);
          formData.append('focus', focusFile);
          formData.append('acumulado', acumuladoFile);
          formData.append('modelo', modeloDeudaFile);
  
          await enviarArchivo('procesar.php', formData, 'FOCUS_ACTUALIZADO.xlsx', 'Archivo FOCUS procesado', submitBtn, 'focus');
        }
      });
  
      formContainer.dataset.submitListenerAttached = 'true';
    }
  }
  
  // ================= Inicialización =================
  document.addEventListener('DOMContentLoaded', function() {
    // Configuración global de la aplicación
    window.API_BASE = window.API_BASE || ''; // Fallback por si no se define en index.php
    console.log(' Configuración de API_BASE:', window.API_BASE);
    
    // Ocultar page loader cuando la página esté completamente cargada
    window.addEventListener('load', function() {
      const pageLoader = document.getElementById('page-loader');
      if (pageLoader) {
        // Agregar clase para ocultar con transición
        pageLoader.classList.add('hidden');
        
        // Asegurar que se oculte completamente después de la transición
        setTimeout(function() {
          pageLoader.style.display = 'none';
        }, 500);
      }
    });
    
    // Fallback para asegurar que el loader se oculte
    const pageLoader = document.getElementById('page-loader');
    if (pageLoader && pageLoader.style.display !== 'none') {
      pageLoader.classList.add('hidden');
      setTimeout(function() {
        pageLoader.style.display = 'none';
      }, 500);
    }
    
    console.log('🚀 === INICIALIZACIÓN DE LA APLICACIÓN ===');
  
    const procesoSelect = document.getElementById('proceso');
    if (procesoSelect) {
      procesoSelect.addEventListener('change', mostrarFormulario);
    } else {
      console.error('❌ ERROR: No se encontró el elemento #proceso');
      return;
    }
  
    procesoSelect.value = '';
    mostrarFormulario();
  
    try {
      actualizarTRM();
      setInterval(actualizarTRM, 600000);
    } catch (err) {
      console.error('❌ Error al inicializar TRM:', err);
    }
  
    const btnTRM = document.getElementById('btnTRM');
    if (btnTRM) btnTRM.addEventListener('click', abrirTRM);
  
    const btnCancelarTRM = document.getElementById('btnCancelarTRM');
    if (btnCancelarTRM) btnCancelarTRM.addEventListener('click', cerrarTRM);
  
    const formTRM = document.getElementById('formTRM');
    if (formTRM) {
      formTRM.addEventListener('submit', async function (e) {
        e.preventDefault();

        const usdRaw = this.querySelector('input[name="usd"]')?.value;
        const eurRaw = this.querySelector('input[name="eur"]')?.value;
        const usd = parseNumeroFlexible(usdRaw);
        const eur = parseNumeroFlexible(eurRaw);

        if (isNaN(usd) || usd <= 0 || isNaN(eur) || eur <= 0) {
          showNotification('❌ Ingrese valores válidos de TRM', 'error');
          return;
        }

        const fd = new FormData();
        fd.append('usd', usd);
        fd.append('eur', eur);

        try {
            const base = window.API_BASE || '';
            const separator = base && !base.endsWith('/') ? '/' : '';
            const response = await fetch(`${base}${separator}trm.php`, { 
                method: 'POST', 
                body: JSON.stringify({
                    usd: usd,
                    eur: eur
                }),
                headers: {
                    'Content-Type': 'application/json'
                }
            });

            if (!response.ok) {
                throw new Error(`Error ${response.status}: ${response.statusText}`);
            }

            await actualizarTRM();
            cerrarTRM();
            this.reset();
            showNotification('✅ TRM actualizada correctamente', 'success');
        } catch (err) {
            console.error('Error actualizando TRM:', err);
            showNotification('❌ Error actualizando TRM', 'error');
        }
      });
    }
  
    console.log('✅ === APLICACIÓN INICIALIZADA ===');
  });
  
})();