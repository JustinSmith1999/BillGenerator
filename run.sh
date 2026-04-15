#!/bin/bash
# Linux launcher — identical to run.command. Run with: ./run.sh
set -e
cd "$(dirname "$0")"
exec ./run.command
