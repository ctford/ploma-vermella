#!/bin/sh
cat > .git/hooks/pre-commit << 'EOF'
#!/bin/sh
set -e
source .venv/bin/activate
ruff check gdocs.py server.py tests/
pytest tests/ -q
EOF
chmod +x .git/hooks/pre-commit
echo "pre-commit hook installed."
