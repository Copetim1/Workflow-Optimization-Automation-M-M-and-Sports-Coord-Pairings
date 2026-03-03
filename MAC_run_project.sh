#!/bin/zsh

# 1. Check/Install Homebrew & Python
echo "\n--- [1/4] Checking Python ---"
if ! command -v brew &> /dev/null; then
    echo "Homebrew not found. Please install it from https://brew.sh/ or install Python manually."
    exit 1
fi

if ! command -v python3 &> /dev/null; then
    echo "Python3 not found. Installing via Homebrew..."
    brew install python
else
    echo "Python3 is ready: $(python3 --version)"
fi

# 2. Setup Virtual Environment
VENV_PATH="./.venv"
VENV_PYTHON="$VENV_PATH/bin/python"

echo "\n--- [2/4] Virtual Environment ---"
if [ ! -d "$VENV_PATH" ]; then
    echo "Creating .venv..."
    python3 -m venv .venv
else
    echo "Environment already exists."
fi

# 3. Verify Libraries inside Venv
LIBRARIES=("pandas" "openpyxl")
echo "\n--- [3/4] Checking Libraries ---"

# Upgrade pip inside venv
$VENV_PYTHON -m pip install --upgrade pip --quiet

for lib in "${LIBRARIES[@]}"; do
    if $VENV_PYTHON -c "import $lib" &> /dev/null; then
        echo "Library '$lib' is verified."
    else
        echo "Installing $lib..."
        $VENV_PYTHON -m pip install $lib --quiet
    fi
done

# 4. Run main.py
echo "\n--- [4/4] Executing Project ---"
if [ -f "./main.py" ]; then
    echo "Running main.py..."
    $VENV_PYTHON ./main.py
else
    echo "Warning: 'main.py' not found. Creating a starter file..."
    cat <<EOF > main.py
import pandas as pd
print("macOS Environment Check Successful!")
print(f"Pandas version: {pd.__version__}")
EOF
    echo "Starter 'main.py' created. Run the script again."
fi

echo "\nWorkflow Complete."