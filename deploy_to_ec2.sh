#!/bin/bash

# EC2 Deployment Script for Subtable Title Comparison System
# This script deploys the complete backend with subtable title comparison functionality

echo "🚀 Starting EC2 deployment for Subtable Title Comparison System..."

# Update system packages
echo "📦 Updating system packages..."
sudo yum update -y

# Install Python 3.9 and pip
echo "🐍 Installing Python 3.9..."
sudo yum install -y python3.9 python3.9-pip python3.9-devel

# Install system dependencies
echo "🔧 Installing system dependencies..."
sudo yum install -y gcc gcc-c++ make openssl-devel bzip2-devel libffi-devel zlib-devel

# Create virtual environment
echo "📁 Creating virtual environment..."
python3.9 -m venv venv
source venv/bin/activate

# Upgrade pip
echo "⬆️ Upgrading pip..."
pip install --upgrade pip

# Install Python dependencies
echo "📚 Installing Python dependencies..."
pip install fastapi uvicorn python-multipart pandas openpyxl pdfplumber python-dotenv

# Create necessary directories
echo "📂 Creating directories..."
mkdir -p logs
mkdir -p temp
mkdir -p uploads

# Set permissions
echo "🔐 Setting permissions..."
chmod +x *.py
chmod 755 logs temp uploads

# Create environment file
echo "⚙️ Creating environment configuration..."
cat > .env << EOF
# Server Configuration
HOST=0.0.0.0
PORT=8000
DEBUG=False

# File Upload Settings
MAX_FILE_SIZE=100MB
UPLOAD_DIR=./uploads
TEMP_DIR=./temp

# Logging
LOG_LEVEL=INFO
LOG_FILE=./logs/app.log

# CORS Settings
ALLOWED_ORIGINS=["*"]
ALLOWED_METHODS=["GET", "POST", "PUT", "DELETE"]
ALLOWED_HEADERS=["*"]

# Cache Settings
CACHE_TTL=3600
CACHE_MAX_SIZE=1000
EOF

# Create systemd service file
echo "🔧 Creating systemd service..."
sudo tee /etc/systemd/system/subtable-comparison.service > /dev/null << EOF
[Unit]
Description=Subtable Title Comparison API
After=network.target

[Service]
Type=simple
User=ec2-user
WorkingDirectory=$(pwd)
Environment=PATH=$(pwd)/venv/bin
ExecStart=$(pwd)/venv/bin/uvicorn server.main:app --host 0.0.0.0 --port 8000
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

# Create startup script
echo "📜 Creating startup script..."
cat > start_server.sh << 'EOF'
#!/bin/bash

# Activate virtual environment
source venv/bin/activate

# Start the server
uvicorn server.main:app --host 0.0.0.0 --port 8000 --reload
EOF

chmod +x start_server.sh

# Create health check script
echo "🏥 Creating health check script..."
cat > health_check.py << 'EOF'
#!/usr/bin/env python3
"""
Health check script for the subtable comparison API
"""

import requests
import sys
import time

def check_health():
    try:
        response = requests.get('http://localhost:8000/health', timeout=10)
        if response.status_code == 200:
            print("✅ API is healthy")
            return True
        else:
            print(f"❌ API returned status code: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Health check failed: {e}")
        return False

if __name__ == "__main__":
    if check_health():
        sys.exit(0)
    else:
        sys.exit(1)
EOF

# Create test script
echo "🧪 Creating test script..."
cat > test_api.py << 'EOF'
#!/usr/bin/env python3
"""
Test script for the subtable title comparison API
"""

import requests
import json
import os

def test_api():
    base_url = "http://localhost:8000"
    
    # Test health endpoint
    print("Testing health endpoint...")
    try:
        response = requests.get(f"{base_url}/health")
        print(f"Health check: {response.status_code}")
        print(f"Response: {response.json()}")
    except Exception as e:
        print(f"Health check failed: {e}")
    
    # Test subtable title comparison endpoint
    print("\nTesting subtable title comparison endpoint...")
    try:
        # This would require actual PDF and Excel files
        print("Note: Full testing requires PDF and Excel files")
        print("Endpoint: POST /tender/compare-subtable-titles")
    except Exception as e:
        print(f"Test failed: {e}")

if __name__ == "__main__":
    test_api()
EOF

# Create monitoring script
echo "📊 Creating monitoring script..."
cat > monitor.py << 'EOF'
#!/usr/bin/env python3
"""
Monitoring script for the subtable comparison API
"""

import psutil
import time
import requests
import json
from datetime import datetime

def get_system_stats():
    cpu_percent = psutil.cpu_percent(interval=1)
    memory = psutil.virtual_memory()
    disk = psutil.disk_usage('/')
    
    return {
        'timestamp': datetime.now().isoformat(),
        'cpu_percent': cpu_percent,
        'memory_percent': memory.percent,
        'memory_available': memory.available // (1024**3),  # GB
        'disk_percent': disk.percent,
        'disk_free': disk.free // (1024**3)  # GB
    }

def check_api_health():
    try:
        response = requests.get('http://localhost:8000/health', timeout=5)
        return response.status_code == 200
    except:
        return False

def main():
    while True:
        stats = get_system_stats()
        api_healthy = check_api_health()
        
        print(f"[{stats['timestamp']}] CPU: {stats['cpu_percent']}% | "
              f"Memory: {stats['memory_percent']}% | "
              f"Disk: {stats['disk_percent']}% | "
              f"API: {'✅' if api_healthy else '❌'}")
        
        time.sleep(60)  # Check every minute

if __name__ == "__main__":
    main()
EOF

# Create README for deployment
echo "📖 Creating deployment README..."
cat > DEPLOYMENT_README.md << 'EOF'
# Subtable Title Comparison System - EC2 Deployment

## Overview
This system provides subtable title comparison functionality between PDF and Excel files.

## Files Included
- `subtable_title_comparator.py` - Main comparison logic
- `subtable_pdf_extractor.py` - PDF subtable extraction
- `excel_subtable_extractor.py` - Excel subtable extraction
- `server/` - FastAPI server implementation
- `deploy_to_ec2.sh` - This deployment script

## Quick Start

### 1. Start the server
```bash
./start_server.sh
```

### 2. Test the API
```bash
python test_api.py
```

### 3. Monitor the system
```bash
python monitor.py
```

## API Endpoints

### Health Check
- `GET /health` - Check API health

### Subtable Title Comparison
- `POST /tender/compare-subtable-titles` - Compare subtable titles
- `POST /tender/compare-cached-subtables` - Compare using cached data

## Configuration
Edit `.env` file to modify server settings.

## Logs
Check `logs/app.log` for application logs.

## Troubleshooting
1. Check if Python 3.9 is installed: `python3.9 --version`
2. Verify virtual environment: `source venv/bin/activate`
3. Check server status: `python health_check.py`
4. View logs: `tail -f logs/app.log`
EOF

# Enable and start the service
echo "🚀 Enabling and starting service..."
sudo systemctl daemon-reload
sudo systemctl enable subtable-comparison.service
sudo systemctl start subtable-comparison.service

# Check service status
echo "📊 Checking service status..."
sudo systemctl status subtable-comparison.service

# Create firewall rules (if needed)
echo "🔥 Configuring firewall..."
sudo yum install -y firewalld
sudo systemctl start firewalld
sudo systemctl enable firewalld
sudo firewall-cmd --permanent --add-port=8000/tcp
sudo firewall-cmd --reload

echo "✅ Deployment completed successfully!"
echo ""
echo "📋 Next steps:"
echo "1. Test the API: python test_api.py"
echo "2. Monitor the system: python monitor.py"
echo "3. Check logs: tail -f logs/app.log"
echo "4. Access the API: http://your-ec2-ip:8000"
echo ""
echo "🔧 Useful commands:"
echo "- Start server: ./start_server.sh"
echo "- Stop service: sudo systemctl stop subtable-comparison.service"
echo "- Restart service: sudo systemctl restart subtable-comparison.service"
echo "- View logs: sudo journalctl -u subtable-comparison.service -f"
