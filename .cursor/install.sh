#!/usr/bin/env bash
# Idempotent Cloud Agent bootstrap for the İzlog ERP Yük/Sevk automation.
# Runs after the repository is checked out. Safe to run repeatedly.
set -euo pipefail

# Resolve repo root (this script lives in .cursor/).
cd "$(dirname "$0")/.."

# --- System packages --------------------------------------------------------
# python3-venv is required to create the virtual environment on Ubuntu.
sudo apt-get update -qq
sudo apt-get install -y -qq python3.12-venv

# --- Python virtual environment + dependencies ------------------------------
if [ ! -d venv ]; then
  python3 -m venv venv
fi
./venv/bin/python -m pip install --upgrade pip
./venv/bin/pip install -r requirements.txt

# --- Playwright browsers ----------------------------------------------------
# The automation drives real Google Chrome (channel="chrome"); chromium is
# also installed for generic Playwright work. install-deps needs root.
sudo ./venv/bin/playwright install-deps chromium
./venv/bin/playwright install chromium chrome

# --- Local settings file ----------------------------------------------------
# ayarlar.py holds real Oracle/ERP credentials and is gitignored. Seed it from
# the template so the code can be imported/linted; the user must fill in real
# secrets (DB_SIFRE, ERP_SIFRE) before any live run.
if [ ! -f ayarlar.py ]; then
  cp ayarlar.example.py ayarlar.py
fi

echo "İzlog ERP otomasyon geliştirme ortamı hazır."
