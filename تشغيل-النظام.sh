#!/usr/bin/env bash
set -e
cd "$(dirname "$0")/dist"
echo "مُتابع المعاملات يعمل على http://localhost:8000"
python3 -m http.server 8000
