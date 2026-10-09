#!/bin/sh
# Usage: ./set-host.sh https://yourname.github.io/hotcam-proposals-cc
# Points the manifest at wherever you've hosted these files (no trailing slash).
set -e
[ -n "$1" ] || { echo "Usage: $0 https://yourname.github.io/hotcam-proposals-cc"; exit 1; }
HOST="${1%/}"
sed -i.bak "s#https://HOSTURL#${HOST}#g" manifest.xml && rm -f manifest.xml.bak
echo "manifest.xml now points at ${HOST}"
grep -c "${HOST}" manifest.xml | xargs echo "URLs updated:"
