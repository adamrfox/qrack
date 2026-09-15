#!/bin/bash
# Always serve plain HTTP on 8000. If a cert/key pair is mounted at
# /certs, also serve HTTPS on 8443 -- this is what avoids the browser's
# "insecure download blocked" behavior for the generated .pptx files.
# Certs are never baked into the image; mount them read-only at runtime.
set -e

uvicorn web.app:app --host 0.0.0.0 --port 8000 &
pids="$!"

if [ -f /certs/qrack-cert.pem ] && [ -f /certs/qrack-key.pem ]; then
    uvicorn web.app:app --host 0.0.0.0 --port 8443 \
        --ssl-certfile /certs/qrack-cert.pem --ssl-keyfile /certs/qrack-key.pem &
    pids="$pids $!"
else
    echo "No cert/key found at /certs -- HTTPS on 8443 disabled, HTTP on 8000 only."
fi

# Exit (and let Docker's restart policy take over) if either server dies.
wait -n $pids
