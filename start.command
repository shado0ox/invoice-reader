#!/bin/bash
cd "$(dirname "$0")" || exit 1
export PORT=5050
open "http://127.0.0.1:5050" >/dev/null 2>&1 &
python3 app.py
