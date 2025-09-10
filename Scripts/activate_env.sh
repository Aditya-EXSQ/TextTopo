#!/bin/bash
echo "Setting up TextTopo development environment..."
export PYTHONPYCACHEPREFIX="E:\Tech\Miscellaneous\TextTopo\.dev\pycache"
echo "[OK] Python cache centralized to: $PYTHONPYCACHEPREFIX"
echo "[OK] Environment ready! Run your Python commands now."
exec "$SHELL"
