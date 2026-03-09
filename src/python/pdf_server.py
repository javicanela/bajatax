#!/usr/bin/env python3
"""
PDF Preview Server — BajaTax
Sirve los archivos de SALIDA_PDF en http://localhost:8080
"""
import http.server
import os

PORT = 8080
SERVE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SALIDA_PDF")

os.makedirs(SERVE_DIR, exist_ok=True)
os.chdir(SERVE_DIR)

handler = http.server.SimpleHTTPRequestHandler
handler.extensions_map.update({".pdf": "application/pdf"})

print(f"Sirviendo PDFs en http://localhost:{PORT}")
print(f"Directorio: {SERVE_DIR}")
http.server.HTTPServer(("", PORT), handler).serve_forever()
python3 --version

python3 --version 

