# Frontend PHP Documentation

This directory contains all the frontend components of the Cartera v3.0.0 system. It provides the web interface for users to interact with the backend processors.

## 📁 Directory Structure

```
front_php/
├── index.php                   # Main landing page
├── procesar.php                # File processing interface
├── trm.php                     # Exchange rate management
├── api.php                     # Frontend API connector
├── download.php                # File download handler
├── config.php                  # Frontend configuration
├── config_local.php            # Local environment settings
├── backend_config.php          # Backend connection settings
├── assets/
│   ├── css/
│   │   └── styles.css          # Main stylesheet
│   ├── js/
│   │   └── main.js             # Client-side JavaScript
│   └── img/                    # Image assets
├── uploads/                    # Temporary user uploads
└── includes/                   # Shared frontend components
```

## 🖥️ Main Pages

### Landing Page ([index.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/index.php))

Main entry point with navigation to all processing options:
- Cartera processing
- Anticipos processing
- Debt model generation
- FOCUS report updating

### Processing Interface ([procesar.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/procesar.php))

Dynamic form handler that adapts to different processing types:
- File upload controls
- Parameter inputs
- Progress indicators
- Results display

### TRM Management ([trm.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/trm.php))

Interface for managing exchange rates:
- View current USD/EUR rates
- Update rate values
- See last update timestamp

## 🎨 User Interface Components

### Styles ([assets/css/styles.css](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/assets/css/styles.css))

Custom styling for:
- Responsive layout
- Form elements
- Buttons and controls
- Tables and data display
- Loading animations

### JavaScript ([assets/js/main.js](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/assets/js/main.js))

Client-side functionality:
- Form validation
- AJAX requests
- File upload handling
- Dynamic UI updates
- Error messaging

### Images ([assets/img/](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/assets/img/))

Visual assets:
- Logo and branding
- Icons
- Loading animations
- Background elements

## ⚙️ Configuration Files

### Main Config ([config.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/config.php))

Frontend configuration settings:
- Backend API URLs
- Display preferences
- Localization settings

### Local Config ([config_local.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/config_local.php))

Environment-specific settings:
- Development vs. production flags
- Debug options
- Custom paths

### Backend Connection ([backend_config.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/backend_config.php))

Backend integration settings:
- API endpoint URLs
- Authentication credentials
- Timeout values

## 🌐 API Integration

### Frontend API Connector ([api.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/api.php))

Handles communication with backend:
- HTTP request formatting
- Response processing
- Error handling
- Retry logic

## 📤 File Handling

### Upload Management

The [uploads/](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/uploads/) directory temporarily stores user files during processing:
- Automatic cleanup after processing
- Unique filename generation
- Size validation
- Type restriction

### Download Handler ([download.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/download.php))

Manages delivery of processed files:
- MIME type setting
- Header configuration
- Streaming delivery
- Error handling

## 📱 Responsive Design

The frontend is designed to work on:
- Desktop browsers
- Tablet devices
- Mobile phones

Key responsive features:
- Flexible grid layout
- Media queries
- Touch-friendly controls
- Adaptive forms

## 🛠️ Debugging Tools

Development helper files:
- [debug_procesar.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/debug_procesar.php): Detailed processing debug view
- [procesar_debug.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/procesar_debug.php): Enhanced debugging interface
- [diagnostico.php](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/diagnostico.php): System diagnostics

## 🔧 Maintenance

Regular maintenance tasks:
- Clear [uploads/](file:///c%3A/wamp64/www/modelo-deuda-python/cartera_v3.0.0/front_php/uploads/) directory
- Update CSS/JS assets
- Check broken links
- Verify form validations
- Test browser compatibility

## 🎯 User Experience Features

- Intuitive navigation
- Clear error messaging
- Progress indicators
- Success confirmations
- Download automation
- Responsive feedback