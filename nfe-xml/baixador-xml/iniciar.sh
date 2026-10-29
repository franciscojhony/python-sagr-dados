#!/usr/bin/env bash
cd "$(dirname "$0")" || exit 1
exec "$(command -v python3 || command -v python)" app.py
