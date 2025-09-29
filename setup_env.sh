#!/bin/bash

ENV_DIR="file_list_env"

# Step 1: Create virtual environment if it doesn't exist
if [ ! -d "$ENV_DIR" ]; then
    echo "🔧 Creating virtual environment in ./$ENV_DIR ..."
    python3 -m venv "$ENV_DIR"
else
    echo "✅ Virtual environment already exists at ./$ENV_DIR"
fi

# Step 2: Activate the environment
source "$ENV_DIR/bin/activate"

# Step 3: Install required packages
echo "📦 Installing 'openpyxl' ..."
pip install --upgrade pip
pip install openpyxl

echo ""
echo "🎉 Setup complete. Virtual environment is active."
echo "To activate it again later, run:"
echo "    source $ENV_DIR/bin/activate"
