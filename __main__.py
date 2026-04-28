"""
__main__.py -- Fallback entry point for ProtocolFormatter.

Allows invocation via:
    python -m protocol_formatter

This is a safety net for environments where the console script entry
point (protocol-formatter) fails to resolve 'main:app'. The recommended
invocation remains the console script after pip install -e .
"""
from main import app

if __name__ == "__main__":
    app()