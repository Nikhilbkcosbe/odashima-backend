#!/bin/bash

# Create a directory for the layer
mkdir -p tesseract_layer/python

# Install Tesseract OCR and its dependencies
yum install -y gcc gcc-c++ make autoconf automake libtool pkgconfig
yum install -y libpng-devel libjpeg-devel libtiff-devel zlib-devel
yum install -y tesseract tesseract-devel

# Create a temporary directory for building
mkdir -p /tmp/tesseract_build
cd /tmp/tesseract_build

# Download and install Tesseract
wget https://github.com/tesseract-ocr/tesseract/archive/refs/tags/5.3.3.tar.gz
tar xzf 5.3.3.tar.gz
cd tesseract-5.3.3
./autogen.sh
./configure
make
make install

# Install Python package
pip install tesserocr -t ../../tesseract_layer/python/

# Copy Tesseract libraries and data
mkdir -p ../../tesseract_layer/lib
cp /usr/local/lib/libtesseract.so* ../../tesseract_layer/lib/
cp -r /usr/local/share/tessdata ../../tesseract_layer/

# Create the layer zip file
cd ../../tesseract_layer
zip -r ../tesseract_layer.zip .

# Clean up
cd ..
rm -rf tesseract_build tesseract_layer 