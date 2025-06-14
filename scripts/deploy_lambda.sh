#!/bin/bash

# Configuration
FUNCTION_NAME="tender-comparison-api"
LAYER_NAME="tesseract-ocr-layer"
REGION="ap-northeast-1"  # Change this to your desired region

# Create the Tesseract layer
echo "Building Tesseract layer..."
./build_tesseract_layer.sh

# Create or update the Lambda layer
echo "Creating/updating Lambda layer..."
aws lambda publish-layer-version \
    --layer-name $LAYER_NAME \
    --description "Tesseract OCR layer for tender comparison" \
    --zip-file fileb://tesseract_layer.zip \
    --compatible-runtimes python3.10 \
    --region $REGION

# Get the latest layer version
LAYER_VERSION=$(aws lambda list-layer-versions \
    --layer-name $LAYER_NAME \
    --region $REGION \
    --query 'LayerVersions[0].Version' \
    --output text)

# Create deployment package
echo "Creating deployment package..."
zip -r function.zip . -x "*.git*" "*.pyc" "__pycache__/*" "tests/*" "scripts/*"

# Create or update the Lambda function
echo "Creating/updating Lambda function..."
aws lambda create-function \
    --function-name $FUNCTION_NAME \
    --runtime python3.10 \
    --handler main.handler \
    --role arn:aws:iam::YOUR_ACCOUNT_ID:role/lambda-tender-comparison-role \
    --zip-file fileb://function.zip \
    --timeout 30 \
    --memory-size 512 \
    --region $REGION \
    --layers arn:aws:lambda:$REGION:YOUR_ACCOUNT_ID:layer:$LAYER_NAME:$LAYER_VERSION

# Clean up
rm -f function.zip tesseract_layer.zip

echo "Deployment complete!" 