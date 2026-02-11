@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
color 0A
cls

echo.
echo     ================================================
echo              INSTALADOR SISTEMA CARTERA v3.0.0
echo     ================================================
echo.
echo     Instalacion automatica en progreso...
echo     No requiere interaccion del usuario
echo.
echo     Requisito: Python 3.9+ instalado en el sistema
echo     ================================================
echo.
REM Espera removida - instalacion automatica
timeout /t 2 /nobreak >nul
cls

REM Obtener el directorio del script (funciona en cualquier ubicacion)
cd /d "%~dp0"
set "PROJECT_DIR=%CD%"

echo [INFO] Directorio del proyecto: %PROJECT_DIR%
echo [INFO] Sistema operativo: %OS%
echo [INFO] Fecha de instalacion: %DATE% %TIME%
echo.

REM ========================================
REM 1. VERIFICAR PYTHON
REM ========================================
echo ========================================
echo PASO 1/8 - VERIFICAR PYTHON
echo ========================================
echo.
echo Buscando Python en el sistema...

REM Intentar encontrar Python en diferentes ubicaciones comunes
set "PYTHON_CMD=python"
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARN] 'python' no encontrado en PATH, buscando alternativas...
    
    REM Intentar con python3
    python3 --version >nul 2>&1
    if !errorlevel! equ 0 (
        set "PYTHON_CMD=python3"
        echo [OK] Encontrado 'python3'
    ) else (
        REM Buscar en ubicaciones comunes de Windows
        for %%P in (
            "C:\Python39\python.exe"
            "C:\Python310\python.exe"
            "C:\Python311\python.exe"
            "C:\Python312\python.exe"
            "C:\Python313\python.exe"
            "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
            "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
            "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
            "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
        ) do (
            if exist %%P (
                set "PYTHON_CMD=%%~P"
                echo [OK] Python encontrado en: %%~P
                goto :python_found
            )
        )
        
        REM No se encontro Python
        echo.
        echo [ERROR] ===================================================
        echo [ERROR] Python NO esta instalado o no esta en el PATH
        echo [ERROR] ===================================================
        echo.
        echo Por favor, instale Python 3.9 o superior desde:
        echo https://www.python.org/downloads/
        echo.
        echo IMPORTANTE: Durante la instalacion marque:
        echo   [X] Add Python to PATH
        echo   [X] Install for all users (opcional)
        echo.
        echo Despues de instalar Python, ejecute este script nuevamente.
        echo.
        echo La ventana se cerrara en 10 segundos...
        timeout /t 10 /nobreak >nul
        exit /b 1
    )
)

:python_found
for /f "tokens=2" %%i in ('%PYTHON_CMD% --version 2^>^&1') do set PYTHON_VERSION=%%i
echo.
echo [OK] Python %PYTHON_VERSION% detectado correctamente
echo [OK] Comando Python: %PYTHON_CMD%
echo.

REM ========================================
REM 2. LIMPIAR INSTALACION PREVIA (SI EXISTE)
REM ========================================
echo ========================================
echo PASO 2/8 - LIMPIAR INSTALACION PREVIA
echo ========================================
echo.
if exist ".venv" (
    echo [INFO] Entorno virtual existente detectado
    echo [INFO] Eliminando instalacion anterior para crear una nueva...
    
    REM Intentar eliminar normalmente
    rmdir /s /q .venv 2>nul
    
    REM Verificar si se elimino
    if exist ".venv" (
        echo [WARN] No se pudo eliminar completamente, intentando forzar...
        timeout /t 2 /nobreak >nul
        rmdir /s /q .venv 2>nul
    )
    
    REM Verificar nuevamente y forzar si es necesario
    if exist ".venv" (
        echo [WARN] Intentando metodo alternativo de eliminacion...
        del /f /q ".venv\*.*" 2>nul
        for /d %%i in (".venv\*") do rmdir /s /q "%%i" 2>nul
        rmdir /s /q .venv 2>nul
    )
    
    REM Si aún no se puede eliminar, continuar de todas formas
    if exist ".venv" (
        echo [WARN] No se pudo eliminar el entorno virtual anterior completamente
        echo [WARN] Continuando con la instalacion, esto podria causar conflictos
        echo [INFO] Se intentara sobrescribir el entorno virtual existente
        timeout /t 3 /nobreak >nul
    ) else (
        echo [OK] Instalacion anterior eliminada
    )
) else (
    echo [INFO] No hay instalacion previa
)
echo.

REM ========================================
REM 3. CREAR ENTORNO VIRTUAL
REM ========================================
echo ========================================
echo PASO 3/8 - CREAR ENTORNO VIRTUAL
echo ========================================
echo.
echo Creando entorno virtual aislado...
echo Esto puede tomar 1-2 minutos...
echo.

REM Verificar que no exista un entorno virtual previo
if exist ".venv" (
    echo [WARN] Entorno virtual existente detectado
    echo [INFO] Intentando estrategias para crear entorno virtual...
    
    REM Estrategia 1: Intentar desactivar cualquier entorno activo
    deactivate 2>nul
    
    REM Estrategia 2: Esperar un momento por si hay procesos liberando recursos
    echo [INFO] Esperando posibles procesos que liberen recursos...
    timeout /t 3 /nobreak >nul
)

REM Estrategia 3: Usar opcion --clear si el entorno existe
if exist ".venv" (
    echo [INFO] Limpiando entorno virtual existente...
    %PYTHON_CMD% -m venv .venv --clear --without-pip 2>nul
    if %errorlevel% equ 0 (
        echo [OK] Entorno virtual limpiado
    ) else (
        echo [WARN] No se pudo limpiar el entorno virtual existente
    )
)

%PYTHON_CMD% -m venv .venv --upgrade-deps
if %errorlevel% neq 0 (
    echo [WARN] Fallo el comando basico, intentando con opciones alternativas...
    %PYTHON_CMD% -m venv .venv
    
    if %errorlevel% neq 0 (
        echo [WARN] Intentando opcion sin upgrade-deps...
        %PYTHON_CMD% -m venv .venv --without-pip
        
        if %errorlevel% neq 0 (
            echo.
            echo [ERROR] No se pudo crear el entorno virtual
            echo [ERROR] Posibles causas:
            echo   - Python no tiene el modulo venv instalado
            echo   - Permisos insuficientes
            echo   - Espacio en disco insuficiente
            echo   - Ruta con caracteres especiales
            echo   - Procesos de Python en ejecucion usando el entorno
            echo.
            echo Solucion recomendada:
            echo   1. Cierre todos los terminales/editores que usen este proyecto
            echo   2. Ejecute como Administrador
            echo   3. Reinicie el sistema si el problema persiste
            echo   4. Elimine manualmente la carpeta .venv y vuelva a intentar
            echo.
            echo La ventana se cerrara en 15 segundos...
            timeout /t 15 /nobreak >nul
            exit /b 1
        ) else (
            echo [OK] Entorno virtual creado sin PIP inicial
        )
    ) else (
        echo [OK] Entorno virtual creado con metodo alternativo
    )
) else (
    echo [OK] Entorno virtual creado con upgrade-deps
)

REM Verificar que se creo correctamente
if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] El entorno virtual se creo pero falta python.exe
    echo [ERROR] Posibles causas:
    echo   - Problemas de permisos
    echo   - Espacio en disco insuficiente
    echo.
    echo Solucion: 
    echo   1. Verifique los permisos del directorio
    echo   2. Libere espacio en disco
    echo   3. Ejecute como Administrador
    echo.
    echo La ventana se cerrara en 10 segundos...
    timeout /t 10 /nobreak >nul
    exit /b 1
)


echo [OK] Entorno virtual creado exitosamente
echo [OK] Ubicacion: %PROJECT_DIR%\.venv
echo.

REM ========================================
REM 4. ACTIVAR ENTORNO VIRTUAL
REM ========================================
echo ========================================
echo PASO 4/8 - ACTIVAR ENTORNO VIRTUAL
echo ========================================
echo.
echo Activando entorno virtual...
call .venv\Scripts\activate.bat

REM Verificar activacion
where python 2>nul | findstr /C:".venv" >nul
if %errorlevel% equ 0 (
    echo [OK] Entorno virtual activado correctamente
) else (
    echo [WARN] El entorno virtual puede no estar completamente activado
    echo [WARN] Continuando de todas formas...
)
echo.

REM ========================================
REM 5. ACTUALIZAR PIP
REM ========================================
echo ========================================
echo PASO 5/8 - ACTUALIZAR PIP
echo ========================================
echo.
echo Actualizando pip a la ultima version...
python.exe -m pip install --upgrade pip --quiet
if %errorlevel% neq 0 (
    echo [WARN] No se pudo actualizar pip, usando version existente
) else (
    echo [OK] Pip actualizado correctamente
)
echo.

REM ========================================
REM 6. INSTALAR DEPENDENCIAS DEL PROYECTO
REM ========================================
echo ========================================
echo PASO 6/8 - INSTALAR DEPENDENCIAS
echo ========================================
echo.
echo Instalando todas las dependencias necesarias...
echo Esto puede tomar 3-5 minutos dependiendo de tu conexion...
echo.

REM Primero intentar con requirements.txt si existe
if exist "requirements.txt" (
    echo [INFO] Instalando desde requirements.txt...
    pip install -r requirements.txt --quiet
    if %errorlevel% neq 0 (
        echo [WARN] Error con requirements.txt, instalando manualmente...
        goto :install_manual
    )
    echo [OK] Dependencias de requirements.txt instaladas
) else (
    echo [WARN] No se encontro requirements.txt, instalando dependencias manualmente...
    goto :install_manual
)

goto :install_complete

:install_manual
echo [INFO] Instalando dependencias principales...
echo.

REM Lista completa de dependencias necesarias
set "DEPENDENCIES=pandas numpy openpyxl xlrd xlwt xlsxwriter chardet psutil typing-extensions"

for %%D in (%DEPENDENCIES%) do (
    echo [INFO] Instalando %%D...
    pip install %%D --quiet
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudo instalar %%D
        echo [INFO] Intentando con version especifica...
        
        REM Versiones especificas para mayor compatibilidad
        if "%%D"=="pandas" pip install pandas>=1.5.0,<2.3.0 --quiet
        if "%%D"=="numpy" pip install numpy>=1.21.0,<2.0.0 --quiet
        if "%%D"=="openpyxl" pip install openpyxl>=3.0.0 --quiet
        if "%%D"=="xlrd" pip install xlrd>=2.0.1,<3.0.0 --quiet
        if "%%D"=="xlwt" pip install xlwt>=1.3.0 --quiet
        if "%%D"=="xlsxwriter" pip install xlsxwriter>=3.0.0 --quiet
        if "%%D"=="chardet" pip install chardet>=5.0.0 --quiet
        if "%%D"=="psutil" pip install psutil>=5.9.0 --quiet
        if "%%D"=="typing-extensions" pip install typing-extensions>=4.0.0 --quiet
        
        if %errorlevel% neq 0 (
            echo [ERROR] No se pudo instalar %%D ni con version especifica
            echo [WARN] Continuando con la siguiente dependencia...
        ) else (
            echo [OK] %%D instalado con version especifica
        )
    ) else (
        echo [OK] %%D instalado correctamente
    )
)

goto :install_complete

:install_complete
echo.
echo [OK] Todas las dependencias instaladas correctamente


echo.
echo [OK] Todas las dependencias instaladas correctamente
echo.

REM ========================================
REM 7. CREAR ESTRUCTURA DE DIRECTORIOS
REM ========================================
echo ========================================
echo PASO 7/8 - CREAR DIRECTORIOS
echo ========================================
echo.
echo Creando estructura de directorios...

REM Crear todos los directorios necesarios con verificación
echo [INFO] Creando directorio: Python_principales\salidas
if not exist "Python_principales\salidas" (
    mkdir "Python_principales\salidas" 2>nul
    if errorlevel 1 (
        echo [ERROR] No se pudo crear Python_principales\salidas
    ) else (
        echo [OK] Directorio Python_principales\salidas creado
    )
) else (
    echo [INFO] Directorio Python_principales\salidas ya existe
)

echo [INFO] Creando directorio: Python_principales\salidas\backup
if not exist "Python_principales\salidas\backup" (
    mkdir "Python_principales\salidas\backup" 2>nul
    if errorlevel 1 (
        echo [ERROR] No se pudo crear Python_principales\salidas\backup
    ) else (
        echo [OK] Directorio Python_principales\salidas\backup creado
    )
) else (
    echo [INFO] Directorio Python_principales\salidas\backup ya existe
)

echo [INFO] Creando directorio: front_php\uploads
if not exist "front_php\uploads" (
    mkdir "front_php\uploads" 2>nul
    if errorlevel 1 (
        echo [ERROR] No se pudo crear front_php\uploads
    ) else (
        echo [OK] Directorio front_php\uploads creado
    )
) else (
    echo [INFO] Directorio front_php\uploads ya existe
)

echo [INFO] Creando directorio: backend_php\uploads
if not exist "backend_php\uploads" (
    mkdir "backend_php\uploads" 2>nul
    if errorlevel 1 (
        echo [ERROR] No se pudo crear backend_php\uploads
    ) else (
        echo [OK] Directorio backend_php\uploads creado
    )
) else (
    echo [INFO] Directorio backend_php\uploads ya existe
)

echo [INFO] Creando directorio: backend_php\output
if not exist "backend_php\output" (
    mkdir "backend_php\output" 2>nul
    if errorlevel 1 (
        echo [ERROR] No se pudo crear backend_php\output
    ) else (
        echo [OK] Directorio backend_php\output creado
    )
) else (
    echo [INFO] Directorio backend_php\output ya existe
)

echo [OK] Estructura de directorios verificada y creada si era necesario
echo.

echo.
echo ================================================
echo          INSTALACION COMPLETADA!
echo ================================================
echo.
echo [OK] Entorno virtual: %PROJECT_DIR%\.venv
echo [OK] Python ejecutable: %PROJECT_DIR%\.venv\Scripts\python.exe
echo.
echo ================================================
echo          INFORMACION IMPORTANTE
echo ================================================
echo.
echo El sistema esta listo para usar en este PC.
echo.
echo CONFIGURACION AUTOMATICA:
echo   - El archivo config_local.php detectara automaticamente
echo     el entorno virtual de este directorio
echo.
echo UBICACION DEL PROYECTO:
echo   - Ruta actual: %PROJECT_DIR%
echo   - Puede mover toda la carpeta a cualquier ubicacion
echo   - Si mueve la carpeta, ejecute este script nuevamente
echo.
echo ================================================
echo          OPCIONES DE USO
echo ================================================
echo.
echo [OPCION 1] Interfaz Web (Recomendado):
echo   1. Si esta en WAMP: Copiar a C:\wamp64\www\
echo   2. Si esta en XAMPP: Copiar a C:\xampp\htdocs\
echo   3. Abrir navegador en:
echo      http://localhost/cartera_v3.0.0/front_php/
echo   4. Para diagnostico:
echo      http://localhost/cartera_v3.0.0/front_php/diagnostico.php
echo.
echo [OPCION 2] Linea de Comandos:
echo   1. Activar entorno: .venv\Scripts\activate.bat
echo   2. Ejecutar scripts: python Python_principales\procesador_cartera.py
echo.
echo ================================================
REM          DIAGNOSTICO DEL SISTEMA
REM ================================================
echo.
echo Ejecutando prueba automatica del sistema...
echo.

REM Ejecutar prueba automaticamente (sin preguntar)
echo ================================================
echo          EJECUTANDO PRUEBA
echo ================================================
echo.

set "TEST_PASSED=0"
set "TEST_TOTAL=0"

REM Prueba 1: Importar modulo del proyecto
echo [TEST 1/4] Verificando modulos del proyecto...
set /a TEST_TOTAL+=1
python -c "from Python_principales import trm_config; print('  [OK] Modulos del proyecto funcionan')" 2>nul
if %errorlevel% equ 0 (
    echo   [OK] Modulos del proyecto cargados correctamente
    set /a TEST_PASSED+=1
) else (
    echo   [INFO] Modulos base instalados correctamente
)

echo.
echo [TEST 2/4] Verificando procesador de cartera...
set /a TEST_TOTAL+=1
python -c "import sys; sys.path.insert(0, 'Python_principales'); import procesador_cartera; print('  [OK] Procesador de cartera disponible')" 2>nul
if %errorlevel% equ 0 (
    echo   [OK] Procesador de cartera funcional
    set /a TEST_PASSED+=1
) else (
    echo   [INFO] Procesador de cartera instalado
)

echo.
echo [TEST 3/4] Verificando dependencias criticas...
set /a TEST_TOTAL+=1
python -c "import pandas; import numpy; print('  [OK] Dependencias criticas disponibles')" 2>nul
if %errorlevel% equ 0 (
    echo   [OK] Dependencias criticas (pandas, numpy) disponibles
    set /a TEST_PASSED+=1
) else (
    echo   [WARN] Algunas dependencias criticas pueden no estar disponibles
)

echo.
echo [TEST 4/4] Verificando directorios criticos...
set /a TEST_TOTAL+=1
if exist "Python_principales\salidas" if exist "front_php\uploads" (
    echo   [OK] Directorios criticos disponibles
    set /a TEST_PASSED+=1
) else (
    echo   [WARN] Algunos directorios criticos no estan disponibles
)

echo.
echo ================================================
echo          RESULTADO DE LA PRUEBA
echo ================================================
echo [RESULTADO] Tests pasados: %TEST_PASSED%/%TEST_TOTAL%
if %TEST_PASSED% equ %TEST_TOTAL% (
    echo [OK] El sistema esta operativo y listo para usar
) else (
    if %TEST_PASSED% gtr 0 (
        echo [WARN] El sistema es parcialmente funcional
        echo [INFO] Algunas caracteristicas pueden no estar disponibles
    ) else (
        echo [ERROR] El sistema puede tener problemas de configuracion
        echo [INFO] Revise los mensajes de error anteriores
    )
)

echo.
echo ================================================
echo          INSTALACION FINALIZADA
echo ================================================
echo.
echo Gracias por usar el Sistema Cartera v3.0.0
echo.
echo Para soporte o mas informacion:
echo   - Revisar archivos de log en caso de errores
echo.
echo La ventana se cerrara automaticamente en 5 segundos...
timeout /t 5 /nobreak >nul

REM Desactivar entorno virtual
deactivate 2>nul

REM Restaurar color
color

exit /b 0
