$ErrorActionPreference = 'Stop'

python -m ruff check .
python -m ruff format --check .
python -m mypy src
python -m pytest

if (Test-Path -LiteralPath 'tools/gates.py') {
    python -m tools.gates all
}

if (Test-Path -LiteralPath 'tools/eye.py') {
    python -m tools.eye check --all
}
