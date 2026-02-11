# Backend PHP Documentation

This directory contains all the PHP components that serve as the backend for the Cartera v3.0.0 system. It handles API requests, file management, and Python script execution.

## 📁 Directory Structure

```
backend_php/
├── api/
│   └── process.php              # Main API endpoint for all processors
├── config/
│   └── config.php               # System configuration settings
├── includes/
│   ├── ProcessorHandler.php     # Processor management utilities
│   ├── PythonRunner.php         # Python script execution handler
│   ├── FileHandler.php          # File upload and management
│   └── Response.php             # Standardized JSON responses
├── uploads/                     # Temporary uploaded files
├── output/                      # Temporary output files
├── test_processors.php          # System verification script
└── download.php                 # File download handler
```

## 🔌 API Endpoints

### Main Processing Endpoint

**URL**: [/api/process.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/api/process.php)
**Method**: POST
**Description**: Handles all file processing requests

#### Actions:
1. `cartera` - Process portfolio files
2. `anticipos` - Process advance payments
3. `modelo_deuda` - Generate debt model
4. `focus` - Update FOCUS reports

### Download Endpoint

**URL**: [/download.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/download.php)
**Method**: GET
**Description**: Provides file downloads for processed results

## ⚙️ Core Components

### Processor Handler ([ProcessorHandler.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/includes/ProcessorHandler.php))

Manages the execution of Python processors:
- Validates input parameters
- Routes requests to appropriate processors
- Handles error management
- Coordinates file paths

### Python Runner ([PythonRunner.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/includes/PythonRunner.php))

Executes Python scripts securely:
- Sets execution timeouts
- Captures stdout/stderr
- Manages virtual environment activation
- Handles process termination

### File Handler ([FileHandler.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/includes/FileHandler.php))

Manages file operations:
- Validates uploaded files
- Enforces size limits (50MB default)
- Checks file extensions (.csv, .xlsx, .xls)
- Generates unique filenames
- Cleans temporary files

### Response Handler ([Response.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/includes/Response.php))

Standardizes API responses:
- Success/failure formatting
- JSON response generation
- HTTP status codes
- Error message structuring

## 🧪 System Verification

Run the system verification script to check installation:

```bash
php test_processors.php
```

This verifies:
- Python accessibility
- Script file existence
- Directory permissions
- TRM file validity

## ⚙️ Configuration

### System Settings ([config/config.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/config/config.php))

Key configuration parameters:
- `PYTHON_PATH`: Path to Python executable
- `MAX_FILE_SIZE`: Maximum upload size (bytes)
- `ALLOWED_EXTENSIONS`: Permitted file types
- `OUTPUT_DIR`: Processed files directory
- `UPLOAD_DIR`: Uploaded files directory

## 🔒 Security Features

- File extension validation
- Size limit enforcement
- Unique filename generation
- Directory isolation (uploads vs. outputs)
- Controlled Python script execution
- Standardized error handling

## 📊 Logging

System logs are captured in standardized JSON responses, including:
- Execution timestamps
- File processing details
- Error information with tracebacks
- Performance metrics

## 🛠️ Maintenance

Regular maintenance tasks:
- Clean [uploads/](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/uploads/) directory
- Clean [output/](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/backend_php/output/) directory
- Update TRM values in configuration
- Monitor disk space usage
