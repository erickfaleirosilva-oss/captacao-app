#!/bin/bash
# Deploy automático para GitHub Pages
cd "$(dirname "$0")"
git add -A
git commit -m "deploy $(date '+%d/%m/%Y %H:%M')"
git push origin main
echo "✅ Deploy feito — https://erickfaleirosilva-oss.github.io/captacao-app/"
