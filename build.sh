#!/usr/bin/env bash
# exit on error
set -o errexit

pip install --upgrade pip
pip install -r requirements.txt
pip install opencv-python
pip install -U pip
pip install -U matplotlib
pip install numpy
pip install openpyxl