# Sistema de Procesamiento de Cartera v3.0.0

## Descripción General

El Sistema de Procesamiento de Cartera es una aplicación web desarrollada en PHP con procesamiento backend en Python que permite transformar y analizar archivos de cartera, anticipos y otros documentos financieros. El sistema está diseñado para el Grupo Planeta y proporciona funcionalidades avanzadas para el procesamiento de datos financieros con múltiples formatos de entrada y salida.

Este sistema automatiza el procesamiento de información financiera compleja, reduciendo el tiempo de análisis manual y minimizando errores en los cálculos. Proporciona una interfaz intuitiva para usuarios no técnicos mientras mantiene la potencia de procesamiento necesaria para grandes volúmenes de datos.

## Flujo Completo del Proceso

### 1. Preparación Inicial

1. **Instalación del sistema**: Ejecutar el script `instalar_sistema_completo.bat` para configurar todas las dependencias necesarias.
2. **Configuración del entorno**: Verificar que la ruta de Python esté correctamente configurada en `front_php/config_local.php`.
3. **Inicio del servidor web**: Asegurarse de que Apache (o el servidor web elegido) esté en funcionamiento.

### 2. Acceso al Sistema

1. **Acceder a la interfaz web**: Abrir un navegador y navegar a `http://localhost/modelo-deuda-python/cartera_v3.0.0/front_php/`
2. **Verificar la conectividad**: Confirmar que todos los componentes del sistema estén disponibles.

### 3. Carga de Archivos

1. **Seleccionar tipo de procesamiento**: Elegir entre Cartera, Anticipos, Modelo de Deuda o FOCUS.
2. **Preparar archivos de entrada**: Asegurarse de que los archivos cumplan con los formatos requeridos.
3. **Subir archivos**: Utilizar el formulario de carga para subir los archivos necesarios.

### 4. Procesamiento

1. **Validación de archivos**: El sistema verifica la integridad y formato de los archivos subidos.
2. **Ejecución del script correspondiente**: El sistema llama al script Python apropiado según el tipo de procesamiento seleccionado.
3. **Procesamiento de datos**: Los scripts Python realizan los cálculos y transformaciones necesarios.
4. **Registro de eventos**: Todos los pasos se registran en el sistema de logs centralizado.

### 5. Generación de Resultados

1. **Creación de archivos de salida**: Los resultados se generan en formato Excel con la estructura correspondiente.
2. **Validación de resultados**: El sistema verifica la integridad de los datos procesados.
3. **Almacenamiento en directorio de salidas**: Los archivos se guardan en `Python_principales/salidas/`.

### 6. Descarga de Resultados

1. **Notificación de finalización**: El sistema notifica al usuario que el procesamiento ha terminado.
2. **Disponibilidad de archivos**: Los archivos procesados están disponibles para descarga.
3. **Descarga de resultados**: El usuario puede descargar los archivos procesados desde la interfaz web.

### 7. Verificación y Validación

1. **Revisión de logs**: Verificar los registros de procesamiento en caso de errores o advertencias.
2. **Validación de resultados**: Confirmar que los resultados cumplen con las expectativas y requisitos.
3. **Análisis de resultados**: Utilizar los archivos generados para análisis financieros o informes.

### 8. Cierre del Proceso

1. **Limpieza de archivos temporales**: Los archivos temporales se eliminan según la política del sistema.
2. **Cierre de sesiones**: Las sesiones de usuario se cierran de forma segura.
3. **Registro de auditoría**: Se registran los eventos finales para fines de auditoría y seguimiento.

## Arquitectura del Sistema

### Componentes Principales

1. **Frontend PHP**: Interfaz web desarrollada en PHP que permite la interacción con el usuario
2. **Backend Python**: Scripts en Python que realizan el procesamiento intensivo de datos
3. **Sistema de Logs Centralizado**: Registro unificado de eventos del sistema
4. **Gestión de Archivos**: Manejo de uploads, procesamiento y descargas de archivos

### Estructura de Directorios

```
cartera_v3.0.0/
├── Python_principales/          # Scripts Python de procesamiento
│   ├── procesador_cartera.py    # Procesador principal de cartera
│   ├── procesador_anticipos.py  # Procesador de anticipos
│   ├── modelo_deuda.py          # Generador de modelo de deuda
│   ├── procesar_y_actualizar_focus.py  # Actualizador de archivos FOCUS
│   ├── log_bridge.py            # Conector de logs entre Python y PHP
│   ├── trm_config.py            # Configuración de tasas de cambio
│   └── salidas/                 # Directorio de archivos de salida
├── front_php/                   # Aplicación web PHP
│   ├── index.php                # Página principal
│   ├── procesar.php             # Controlador de procesamiento
│   ├── config_local.php         # Configuración local
│   ├── includes/                # Librerías auxiliares
│   │   └── LogHelper.php        # Gestor de logs en PHP
│   ├── logs/                    # Sistema de logs
│   │   ├── system_log.txt       # Archivo de logs centralizado
│   │   └── viewer.php           # Visor web de logs
│   ├── uploads/                 # Archivos subidos por usuarios
│   └── assets/                  # Recursos estáticos (CSS, JS, imágenes)
└── instalar_sistema_completo.bat # Script de instalación
```

### Flujo de Datos

1. Usuario accede a la interfaz web
2. Archivos son subidos al servidor
3. PHP valida y mueve los archivos a directorios de trabajo
4. Scripts Python procesan los datos
5. Resultados se guardan en directorio de salidas
6. Usuario descarga archivos procesados

## Funcionalidades Principales

### 1. Procesamiento de Cartera

Transforma archivos CSV de cartera en formatos Excel con análisis detallado de vencimientos:

- **Cálculo de días vencidos**: Determina cuántos días han pasado desde la fecha de vencimiento
- **Cálculo de días por vencer**: Calcula cuántos días faltan para la fecha de vencimiento
- **Clasificación por rangos**: Agrupa saldos en rangos de vencimiento (30, 60, 90, 180, 360, +360 días)
- **Análisis de saldos**: Calcula saldos vencidos, por vencer y totales
- **Generación de columnas históricas**: Crea columnas con saldos por meses históricos

#### Formato de Entrada Cartera

El archivo CSV de cartera debe contener las siguientes columnas obligatorias:
- **FECHA VTO**: Fecha de vencimiento en formato YYYY-MM-DD
- **SALDO**: Valor monetario del saldo
- **MONEDA**: Código de moneda (COP, USD, EUR)

#### Proceso de Cálculo

1. Lectura y validación del archivo CSV
2. Detección automática de codificación
3. Parseo de fechas y valores
4. Cálculo de días vencidos/por vencer
5. Clasificación en rangos de vencimiento
6. Generación de estructura Excel con formato contable

### 2. Procesamiento de Anticipos

Convierte archivos CSV de anticipos en formato Excel con estructura específica:

- **Limpieza de datos**: Elimina caracteres no imprimibles y normaliza formatos
- **Conversión de valores**: Transforma valores numéricos y aplica signo negativo
- **Formateo de fechas**: Convierte fechas a formato estándar
- **Renombrado de columnas**: Mapea columnas de diferentes formatos a estructura estándar

#### Formato de Entrada Anticipos

El archivo CSV de anticipos debe contener las siguientes columnas:
- **CODIGO**: Código identificador
- **VALOR**: Valor monetario (será convertido a negativo)
- **FECHA**: Fecha de registro

### 3. Modelo de Deuda

Genera un modelo consolidado que combina información de cartera y anticipos:

- **Integración de datos**: Combina múltiples fuentes de información financiera
- **Aplicación de TRM**: Utiliza tasas de cambio para conversión de monedas
- **Compensación de anticipos**: Aplica correctamente la compensación de anticipos (valores negativos) con cartera (valores positivos)
- **Cálculos consolidados**: Realiza cálculos financieros complejos sobre datos combinados

#### Parámetros de Entrada

- Archivo de cartera procesado
- Archivo de anticipos procesado
- Tasas de cambio TRM (USD/EUR)

#### Cálculos Realizados

1. Conversión de monedas usando TRM
2. Consolidación de saldos
3. Cálculo neto (cartera - anticipos)
4. Generación de informe detallado

### 4. Actualización FOCUS

Procesa y actualiza archivos FOCUS con la información más reciente. Esta funcionalidad es crítica para el sistema ya que maneja cálculos financieros complejos que deben ser precisos según las normativas contables del Grupo Planeta.

#### Archivos de Entrada Requeridos

1. **Archivo de Balance**: Contiene información de cuentas contables
2. **Archivo de Situación**: Datos de situación financiera
3. **Archivo de Modelo Deuda**: Información de vencimientos
4. **Archivo Acumulado**: Datos históricos
5. **Archivo FOCUS Base**: Plantilla a actualizar

#### Cálculos Críticos en FOCUS

##### Celda Q22 - Deuda bruta NO Grupo (Final)
En la celda Q22 del archivo FOCUS debe aplicarse la siguiente fórmula: + Facturación del mes - Columna no vencida. Este cálculo es crítico para el informe financiero y debe verificarse en cada actualización para garantizar precisión en los saldos reportados.

##### Otros Cálculos Importantes
- **Q15**: Total Balance dividido entre 1000
- **Q16**: Saldos Mes dividido entre 1000
- **Q17**: Vencido 60+ días dividido entre 1000
- **Q19**: Vencido 30 días dividido entre 1000
- **Q21**: Provisión dividida entre 1000
- **F16**: Facturación del mes - No vencida
- **F15**: Deuda bruta NO Grupo (Final - No vencida)

#### Reglas de Procesamiento

##### Del Archivo BALANCE
Se toman las siguientes cuentas y se suman la columna con nombre "Saldo AAF variación":
- Total cuenta objeto 43001
- 0080.43002.20
- 0080.43002.21
- 0080.43002.15
- 0080.43002.28
- 0080.43002.31
- 0080.43002.63
- Total cuenta objeto 43008
- Total cuenta objeto 43042

##### Del Archivo Situación
Se toma el valor de TOTAL 01010 Columna SALDOS MES

##### Archivo FOCUS - Reglas de Actualización
1. Tipos de cambio: Cambiar el mes de cierre y actualizar tasas de cambio
2. Deuda bruta NO Grupo (Inicial) = Deuda bruta NO Grupo (Final)
3. Dotaciones Acumuladas (Inicial) = '+/- Provisión acumulada (Final)
4. Cobro de mes - Vencida = Deuda bruta NO Grupo (Inicial) Vencidas - Total vencido de 60 días en adelante / 1000
5. Cobro mes mes - Total Deuda = COBROS SITUACION (SALDO MES) / -1000
6. Cobros del mes - No Vencida = H15-D15
7. +/- Vencidos en el mes – vencido = VENCIDO MES 30 días signo positivo
8. +/- Vencidos en el mes – No vencido = D17
9. '+/- Vencidos en el mes – Total deuda = D17 - F17
10. + Facturación del mes – vencida = 0
11. + Facturación del mes – no vencida = primero borrar el resultado y calcular =+Q22-H22
    '+ Facturación del mes - Columna no vencida - Deuda bruta NO Grupo (Final)-total deuda

##### Dotación del mes
Primero borrar la dotación y calcular:
Del archivo de datos de la provisión del mes Interco RESTO
- Dotaciones Acumuladas (Inicial) - Provisión del mes

##### ACUMULADO
Del archivo formato acumulado Focus prueba
Copiar las formulas de B54 a F54 al archivo focus

## Instalación y Configuración

### Requisitos del Sistema

- **Servidor Web**: Apache (recomendado WAMP en Windows)
- **PHP**: Versión 7.4 o superior
- **Python**: Versión 3.8 o superior
- **Librerías Python**: pandas, openpyxl, xlsxwriter, chardet
- **Permisos de escritura**: En directorios de uploads, salidas y logs

### Instalación

1. **Clonar o descargar el repositorio**
2. **Ejecutar el script de instalación**:
   ```
   instalar_sistema_completo.bat
   ```
3. **Configurar rutas en `front_php/config_local.php`**:
   - Verificar que `LOCAL_PYTHON_PATH` apunte a la instalación de Python correcta
   - Ajustar rutas si es necesario

### Configuración

El archivo `front_php/config_local.php` contiene las configuraciones principales:

```php
// Ruta a Python - CAMBIAR SEGÚN TU INSTALACIÓN
define('LOCAL_PYTHON_PATH', 'python'); // O ruta completa a python.exe

// Directorios base
define('LOCAL_BACKEND_DIR', LOCAL_BASE_DIR . '/Python_principales');
define('LOCAL_SALIDAS_DIR', LOCAL_BACKEND_DIR . '/salidas');
define('LOCAL_UPLOADS_DIR', __DIR__ . '/uploads');
```

### Variables de Entorno

- **PYTHONPATH**: Debe incluir la ruta a los scripts Python
- **LOG_LEVEL**: Nivel de detalle de logs (INFO, DEBUG, WARNING, ERROR)
- **MAX_FILE_SIZE**: Tamaño máximo de archivos permitidos (en bytes)

## Uso del Sistema

### Acceso a la Interfaz

1. **Iniciar servidor web** (Apache)
2. **Acceder a la aplicación** mediante navegador web:
   ```
   http://localhost/modelo-deuda-python/cartera_v3.0.0/front_php/
   ```

### Procesos Disponibles

#### Cartera
- **Entrada**: Archivo CSV con datos de cartera
- **Parámetros**: Fecha de cierre (obligatorio)
- **Salida**: Archivo Excel con análisis de vencimientos
- **Tiempo estimado**: 2-5 minutos para archivos de 100MB

#### Anticipos
- **Entrada**: Archivo CSV con datos de anticipos
- **Salida**: Archivo Excel con estructura estandarizada
- **Tiempo estimado**: 1-2 minutos para archivos de 50MB

#### Modelo de Deuda
- **Entrada**: Archivos de cartera y anticipos
- **Parámetros**: Tasas de cambio (TRM USD/EUR)
- **Salida**: Modelo consolidado en Excel
- **Tiempo estimado**: 3-7 minutos para archivos combinados

#### FOCUS
- **Entrada**: Múltiples archivos (balance, situación, acumulado, FOCUS anterior, modelo deuda)
- **Salida**: Archivo FOCUS actualizado
- **Tiempo estimado**: 5-10 minutos para procesamiento completo

## Sistema de Logs

### Logs Centralizados

El sistema utiliza un archivo de logs centralizado (`front_php/logs/system_log.txt`) que registra eventos de ambos componentes (PHP y Python). Todos los logs se guardan en formato JSON para facilitar el análisis automatizado.

```json
{
  "timestamp": "2025-11-24T14:35:10.804882Z",
  "source": "PY_PROCESADOR_CARTERA",
  "level": "INFO",
  "message": "[INFO] SALDO TOTAL VENCIDO: $7,432,242,163",
  "context": {
    "filename": "procesador_cartera.py",
    "lineno": 605,
    "func": "procesar_cartera"
  }
}
```

### Niveles de Log

- **DEBUG**: Información detallada para diagnóstico
- **INFO**: Eventos generales del sistema
- **WARNING**: Advertencias que no detienen el proceso
- **ERROR**: Errores que pueden afectar el resultado
- **CRITICAL**: Errores graves que detienen el proceso

### Visor de Logs Web

Acceso mediante:
```
http://localhost/modelo-deuda-python/cartera_v3.0.0/front_php/logs/viewer.php
```

**Características**:
- Filtrado por nivel (INFO, ERROR, WARNING, DEBUG)
- Filtrado por componente (PHP_PROCESAR, PY_PROCESADOR_CARTERA, etc.)
- Filtrado por proceso (Cartera, Anticipos)
- Búsqueda de texto
- Visualización compacta de registros
- Exportación de logs filtrados

## Desarrollo y Mantenimiento

### Scripts Python Principales

#### procesador_cartera.py

**Funcionalidades clave**:
- Detección automática de codificación de archivos
- Procesamiento robusto de fechas con múltiples formatos
- Cálculo de días vencidos y por vencer
- Generación de rangos de vencimiento
- Formato de salida en Excel con miles

**Uso desde línea de comandos**:
```bash
python procesador_cartera.py archivo_entrada.csv 2025-12-31
```

#### procesador_anticipos.py

**Funcionalidades clave**:
- Renombrado de columnas según mapeo definido
- Limpieza de caracteres no imprimibles
- Conversión de valores numéricos
- Formato de fechas
- Generación de archivo Excel

**Uso desde línea de comandos**:
```bash
python procesador_anticipos.py archivo_entrada.csv
```

#### procesar_y_actualizar_focus.py

**Funcionalidades clave**:
- Procesamiento de múltiples archivos de entrada
- Cálculos financieros complejos
- Validación de integridad de datos
- Actualización de plantillas FOCUS
- Limpieza de datos inconsistentes

**Uso desde línea de comandos**:
```bash
python procesar_y_actualizar_focus.py archivo_balance.xlsx archivo_situacion.xlsx archivo_focus.xlsx archivo_acumulado.xls archivo_modelo.xlsx
```

### API y Controladores PHP

#### procesar.php

Controlador principal que maneja las solicitudes de procesamiento:

1. **Validación de entrada**: Verifica tipo de proceso y archivos
2. **Movimiento de archivos**: Mueve archivos subidos al directorio de trabajo
3. **Ejecución de scripts Python**: Llama a los scripts correspondientes
4. **Manejo de resultados**: Procesa salida y envía archivo al usuario
5. **Registro de eventos**: Registra todos los pasos en el sistema de logs

#### LogHelper.php

Gestor de logs en PHP que escribe en formato JSON al archivo centralizado:

```php
LogHelper::log('PHP_PROCESAR', 'INFO', 'Inicio de procesamiento', [
    'request_method' => $_SERVER['REQUEST_METHOD'],
    'content_type' => $_SERVER['CONTENT_TYPE']
]);
```

## Solución de Problemas

### Errores Comunes

#### Python no disponible
- **Mensaje**: "Python no está disponible"
- **Solución**: Verificar `LOCAL_PYTHON_PATH` en `config_local.php`
- **Diagnóstico**: Ejecutar `python --version` en la terminal

#### Permisos de escritura
- **Mensaje**: "El directorio no es escribible"
- **Solución**: Verificar permisos en directorios `uploads` y `salidas`
- **Comando Windows**: `icacls "ruta\directorio" /grant Usuarios:F`

#### Archivos no encontrados
- **Mensaje**: "Script Python no encontrado"
- **Solución**: Verificar rutas en `LOCAL_PY_SCRIPTS`
- **Verificación**: Confirmar existencia de archivos en `Python_principales/`

#### Memoria insuficiente
- **Mensaje**: "MemoryError" en logs de Python
- **Solución**: Aumentar memoria asignada a PHP y Python
- **Configuración**: Editar `php.ini` y variables de entorno de Python

### Monitoreo de Logs

Utilizar el visor de logs web para identificar problemas:
```
http://localhost/modelo-deuda-python/cartera_v3.0.0/front_php/logs/viewer.php
```

Filtrar por nivel ERROR para ver problemas críticos.

### Diagnóstico de Rendimiento

1. **Archivos grandes**: Monitorear tiempo de procesamiento
2. **Conexiones lentas**: Verificar recursos del servidor
3. **Errores recurrentes**: Analizar patrones en logs
4. **Uso de CPU/Memoria**: Monitorear durante procesos intensivos

## Seguridad

### Validaciones Implementadas

1. **Validación de tipos de archivo**: Solo se aceptan formatos específicos
2. **Verificación de carga**: Se comprueban errores en la subida de archivos
3. **Limpieza de nombres**: Se sanitizan nombres de archivos para evitar inyecciones
4. **Control de acceso**: Solo se permiten solicitudes POST válidas

### Buenas Prácticas

1. **No exponer rutas del sistema** en mensajes de error
2. **Validar todas las entradas** del usuario
3. **Mantener actualizadas** las dependencias de Python
4. **Realizar copias de seguridad** de archivos importantes
5. **Monitorear logs regularmente** para detectar anomalías
6. **Limitar tamaño de archivos** para prevenir ataques DoS

### Consideraciones de Seguridad

- **Sanitización de entradas**: Todos los nombres de archivo y parámetros son sanitizados
- **Prevención de inyecciones**: Uso de funciones seguras para ejecución de comandos
- **Control de acceso**: Validación de sesiones y permisos
- **Cifrado de datos sensibles**: Protección de información financiera
- **Auditoría de logs**: Registro completo de todas las operaciones

## Personalización

### Agregar Nuevo Proceso

1. **Crear script Python** en `Python_principales/`
2. **Agregar entrada en `LOCAL_PY_SCRIPTS`** en `config_local.php`
3. **Actualizar selector de procesos** en `front_php/index.php`
4. **Crear formulario correspondiente** en `front_php/assets/js/main.js`
5. **Registrar nuevo tipo de log** en `LogHelper.php`

### Modificar Formatos de Salida

1. **Editar scripts Python** para cambiar estructura de columnas
2. **Modificar estilos CSS** en `front_php/assets/css/styles.css`
3. **Actualizar mapeos de columnas** en los scripts de procesamiento
4. **Ajustar validaciones** en el frontend PHP
5. **Actualizar documentación** en este README

### Extender Funcionalidades

1. **Nuevos formatos de entrada**: Agregar parsers en scripts Python
2. **Nuevos cálculos**: Implementar lógica en módulos correspondientes
3. **Nuevas salidas**: Crear generadores de reportes adicionales
4. **Integraciones externas**: Desarrollar conectores API
5. **Automatización**: Crear jobs programados para procesamiento batch

## Contribuidores

- Desarrollo interno del Grupo Planeta
- Mantenimiento y actualizaciones continuas

## Licencia

Sistema de uso interno del Grupo Planeta. No distribuible sin autorización.

## Versiones

### v3.0.0 (Actual)
- Mejoras en procesamiento de cartera
- Optimización de logs
- Interfaz web mejorada
- Sistema de monitoreo de procesos
- Validación de cálculos financieros críticos
- Documentación completa del sistema

### Historial de Versiones
- v2.0.0: Integración de procesamiento de anticipos
- v1.0.0: Versión inicial con procesamiento de cartera básico

## Contacto

Para soporte técnico y mantenimiento, contactar al equipo de desarrollo del Grupo Planeta.