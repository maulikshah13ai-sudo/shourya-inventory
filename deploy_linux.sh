#!/bin/bash
# ==============================================================================
# One-Click Linux Setup Script for Bags & Footwear Inventory System
# Compatible with Ubuntu 20.04/22.04/24.04 and Debian Linux
# ==============================================================================

set -e

echo "=================================================================="
echo "  Setting up Bags & Footwear Inventory System on Linux Server...  "
echo "=================================================================="

# 1. Update packages and install Python3 & venv
echo "[1/5] Checking and installing Python 3 environment..."
sudo apt update -y
sudo apt install -y python3 python3-pip python3-venv sqlite3

# 2. Setup Virtual Environment
echo "[2/5] Creating Python virtual environment..."
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

source venv/bin/activate

# 3. Install requirements & Gunicorn
echo "[3/5] Installing Python libraries (Flask, Gunicorn, Waitress)..."
pip install --upgrade pip
pip install -r requirements.txt

# 4. Initialize Database
echo "[4/5] Initializing local database and permissions..."
python database.py

# 5. Create Systemd Service for 24/7 background running & auto-restart
echo "[5/5] Creating automatic system service (runs 24/7)..."
SERVICE_FILE="/etc/systemd/system/inventory.service"
CURRENT_USER=$(whoami)

sudo bash -c "cat > $SERVICE_FILE" <<EOL
[Unit]
Description=Bags and Footwear Inventory Web Application
After=network.target

[Service]
User=$CURRENT_USER
WorkingDirectory=$SCRIPT_DIR
ExecStart=$SCRIPT_DIR/venv/bin/gunicorn --workers 3 --bind 0.0.0.0:5000 app:app
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOL

sudo systemctl daemon-reload
sudo systemctl enable inventory
sudo systemctl restart inventory

PUBLIC_IP=$(curl -s ifconfig.me || hostname -I | awk '{print $1}')

echo ""
echo "=================================================================="
echo "  [SUCCESS] Application is now LIVE and running on your Linux host!"
echo "=================================================================="
echo "  Access the app in your browser at:"
echo "    http://$PUBLIC_IP:5000"
echo ""
echo "  Default Logins:"
echo "    Admin / Owner:        admin    / admin123"
echo "    Salesman:             sales1   / sales123"
echo "    Accounts:             accounts / accounts123"
echo "    Warehouse:            godown   / godown123"
echo "=================================================================="
