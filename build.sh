#!/bin/bash
apt-get update
apt-get install -y tesseract-ocr tesseract-ocr-ara libzbar0 libzbar-dev
pip install -r requirements.txt
