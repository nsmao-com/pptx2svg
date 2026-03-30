#!/bin/sh
set -e

if command -v fc-cache >/dev/null 2>&1; then
  fc-cache -fv >/dev/null 2>&1 || true
fi

exec "$@"
