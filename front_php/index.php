<?php
require_once 'config.php';
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Procesador de Cartera - Grupo Planeta</title>
    <link rel="icon" type="image/x-icon" href="assets/img/favicon.ico">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet" crossorigin="anonymous">
    <link href="assets/css/styles.css" rel="stylesheet">
</head>
<body>
    <!-- Page Loader -->
    <div id="page-loader">
        <div class="loader-content">
            <!-- Imagen del planeta cargando -->
            <div class="planet-loader">
                <img src="assets/img/planetacargando.gif" alt="Cargando..." class="loader-image">
            </div>
            <div class="loader-text">
                <h3>Cargando aplicación...</h3>
                <p>Por favor espere mientras se carga la interfaz</p>
            </div>
        </div>
    </div>
    
    <!-- Processing Loader -->
    <div id="processing-loader" style="display: none;">
        <div class="loader-content">
            <!-- Imagen del planeta cargando -->
            <div class="planet-loader">
                <img src="assets/img/planetacargando.gif" alt="Cargando..." class="loader-image">
            </div>
            
            <!-- Texto y barra de progreso -->
            <div class="loader-text">
                <h3 id="loader-message">Procesando archivos...</h3>
                <p id="loader-submessage">Esto puede tomar algunos segundos</p>
            </div>
            
            <div class="progress-container">
                <div class="progress-bar">
                    <div class="progress-fill"></div>
                </div>
                <div class="progress-text">
                    <span id="progress-status">Preparando...</span>
                </div>
            </div>
            
            <!-- Indicadores de proceso -->
            <div class="process-indicators">
                <div class="indicator" data-step="1">
                    <i class="fas fa-file-upload"></i>
                    <span>Cargando</span>
                </div>
                <div class="indicator" data-step="2">
                    <i class="fas fa-cog"></i>
                    <span>Procesando</span>
                </div>
                <div class="indicator" data-step="3">
                    <i class="fas fa-file-excel"></i>
                    <span>Generando</span>
                </div>
                <div class="indicator" data-step="4">
                    <i class="fas fa-check-circle"></i>
                    <span>Finalizando</span>
                </div>
            </div>
        </div>
    </div>

    <!-- Header -->
    <header>
        <div class="logo-container">
            <img src="assets/img/logo_blanco_azul-sin_fondo.png" alt="Logo Grupo Planeta" class="logo">
            <h2>Procesador de Cartera</h2>
        </div>
        <div class="trm-container">
            <div class="trm-info">
                <div class="trm-title">Tasas de Cambio</div>
                <div class="trm-label" id="trm-label">TRM: Cargando..</div>
            </div>
            <button id="btnTRM" class="trm-button">
                <i class="fas fa-sync-alt"></i>
                Actualizar
            </button>
        </div>
    </header>

    <!-- Main -->
    <main>
        <div class="content">
            <h5>
                <i class="fas fa-cogs"></i>
                Seleccione un proceso
            </h5>
            
            <select id="proceso">
                <option value="">-- Seleccione un proceso --</option>
                <option value="cartera">📊 Cartera</option>
                <option value="anticipos">💰 Anticipos</option>
                <option value="modelo_deuda">👻 Modelo Deuda</option>
                <option value="focus">🎯 Focus</option>
            </select>
            
            <div id="formularios"></div>
            <div id="resultado" style="margin-top: 20px;"></div>
        </div>
    </main>

    <!-- Footer -->
    <footer>
        © 2026 Grupo Planeta - Sistema de Procesamiento de Cartera
    </footer>

    <!-- Modal TRM - Rediseñado -->
    <div id="modalTRM" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-exchange-alt"></i> Actualizar Tasas de Cambio</h3>
                <button type="button" class="close-modal" onclick="cerrarTRM()">&times;</button>
            </div>
            <div class="modal-body">
                <form id="formTRM">
                    <div class="modal-form-group">
                        <label>Dólar Americano (USD)</label>
                        <div class="input-field">
                            <i class="fas fa-dollar-sign"></i>
                            <input type="text" inputmode="decimal" name="usd" placeholder="Ej: 3.921,58" required>
                        </div>
                    </div>
                    <div class="modal-form-group">
                        <label>Euro (EUR)</label>
                        <div class="input-field">
                            <i class="fas fa-euro-sign"></i>
                            <input type="text" inputmode="decimal" name="eur" placeholder="Ej: 4.000,00" required>
                        </div>
                    </div>
                    <div class="modal-buttons">
                        <button type="button" id="btnCancelarTRM" class="btn btn-secondary">
                            <i class="fas fa-times"></i> Cancelar
                        </button>
                        <button type="submit" id="btn-guardar-trm" class="btn-guardar">
                            <i class="fas fa-save"></i> Guardar Cambios
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <!-- JavaScript -->
    <script>
        // Definir API_BASE antes de cargar main.js
        window.API_BASE = '';
    </script>
    <script src="assets/js/main.js?v=<?php echo time() ?>"></script>
</body>
</html>