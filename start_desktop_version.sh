#!/bin/bash

echo "NSNA Mail Merge Tool - Desktop Storage Version"
echo "============================================="
echo ""
echo "This will start the NSNA Mail Merge Tool with desktop storage."
echo "Your files will be saved to: $HOME/Desktop/NSNA_Mail_Merge/"
echo ""
echo "Starting Docker containers..."
echo ""

# Check if docker-compose exists
if ! command -v docker-compose &> /dev/null; then
    echo "docker-compose not found. Trying docker compose..."
    docker compose up --build
else
    docker-compose up --build
fi

echo ""
echo "Application stopped. Your files remain saved on your desktop."
