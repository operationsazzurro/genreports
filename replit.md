# Overview

This is a Flask-based REST API service that generates Excel reports from JSON data. The application provides an endpoint to convert structured data into formatted Excel spreadsheet files, specifically designed for weekly site activity reports with photo URLs.

# User Preferences

Preferred communication style: Simple, everyday language.

# Recent Changes

**October 11, 2025** - Bug Fixes and Error Handling Improvements:
1. Fixed critical error handling to properly catch `BadRequest` and `UnsupportedMediaType` exceptions, returning appropriate HTTP status codes (400/415) instead of 500 errors
2. Fixed column mismatch in title row merge (changed from A1:F1 to A1:E1 to match 5 columns)
3. Added comprehensive input validation for JSON data structure
4. Fixed merged cell handling in auto-width calculation to prevent AttributeError
5. Properly close BytesIO resource with try/finally block
6. Changed port from 8080 to 5000 for proper Replit environment compatibility

# System Architecture

## Backend Architecture

**Framework**: Flask (Python web framework)
- Chosen for its lightweight nature and simplicity for API-only services
- Minimal overhead for single-purpose report generation endpoint
- Easy deployment and scaling for microservice architecture

**Excel Generation**: OpenPyXL library
- Selected for pure Python implementation (no external dependencies like MS Office)
- Provides comprehensive Excel formatting capabilities (fonts, alignment, cell merging, fills)
- Supports .xlsx format which is widely compatible

**Design Pattern**: RESTful API
- Single POST endpoint `/generate_excel` accepts JSON data
- Stateless request-response model
- Returns base64-encoded Excel file or error JSON

## Data Flow

1. Client sends POST request with JSON payload containing report data
2. Server validates JSON structure and data types
3. Excel workbook created in-memory using BytesIO (no file system writes)
4. Data formatted with styling (headers, title row, color fills)
5. Excel file returned as base64-encoded string or download

## Error Handling Strategy

- Input validation at multiple levels (JSON parsing, data type checking, list validation)
- HTTP-appropriate status codes (400 for bad requests)
- Structured JSON error responses for client consumption
- Graceful handling of malformed requests using Werkzeug exceptions

## Report Structure

**Fixed 5-column layout**:
- Activity ID
- Site ID  
- Date
- Before Photo URL
- After Photo URL

**Formatting decisions**:
- Merged title cell (A1:E1) for professional appearance
- Bold headers with light blue fill (#B7DEE8) for visual hierarchy
- Center-aligned title for improved readability

# External Dependencies

## Python Libraries

- **Flask**: Web framework for HTTP request handling and routing
- **Werkzeug**: WSGI utility library (bundled with Flask) for exception handling
- **OpenPyXL**: Excel file creation and manipulation (.xlsx format)
- **Base64** (Python stdlib): Encoding Excel files for API responses
- **DateTime** (Python stdlib): Date handling for report data
- **BytesIO** (Python stdlib): In-memory file handling to avoid disk I/O

## No External Services

The application currently operates standalone with no external API calls, databases, or third-party service integrations. All processing happens in-memory.