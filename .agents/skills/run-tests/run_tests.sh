#!/bin/bash

# Load NVM if it exists
export NVM_DIR="$HOME/.nvm"
if [ -s "$NVM_DIR/nvm.sh" ]; then
    echo "Sourcing NVM..."
    . "$NVM_DIR/nvm.sh"
    echo "Activating Node version..."
    nvm use default || nvm use node || true
fi

# Fallback to standard paths if node/npm are still not found
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

LINT_STATUS=0
TEST_STATUS=0

echo "=== 1. Running ESLint Code Checks ==="
if npm run lint; then
    echo "✅ ESLint passed!"
else
    echo "❌ ESLint failed!"
    LINT_STATUS=1
fi

echo ""
echo "=== 2. Running Unit Tests (Vitest) ==="
if npx vitest run; then
    echo "✅ Unit tests passed!"
else
    echo "❌ Unit tests failed!"
    TEST_STATUS=1
fi

echo ""
if [ $LINT_STATUS -eq 0 ] && [ $TEST_STATUS -eq 0 ]; then
    echo "✅ All code checks and unit tests passed successfully!"
    exit 0
else
    echo "❌ Code checks or unit tests failed! Please review the logs above."
    exit 1
fi
