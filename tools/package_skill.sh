#!/usr/bin/env bash
# Собирает dist/gsom-presentation.zip для загрузки в claude.ai (Customize → Skills → Upload a skill).
# В корне архива — папка gsom-presentation (так требует claude.ai). Запускать после изменений скилла.
set -euo pipefail
cd "$(dirname "$0")/../skills"
rm -f ../dist/gsom-presentation.zip
zip -qr ../dist/gsom-presentation.zip gsom-presentation -x '*__pycache__*' '*.pyc' '*.DS_Store'
ls -lh ../dist/gsom-presentation.zip
