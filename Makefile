.PHONY: cli check-examples exe

PYTHON ?= python

cli:
	$(PYTHON) -c "import sys; sys.path.insert(0, 'lib'); from yibinthesis_cli.app import main; raise SystemExit(main(['--help']))"

check:
	$(PYTHON) -m unittest discover -s lib/tests -p 'test_*.py' -v
	$(PYTHON) lib/audit_format.py

check-examples:
	$(PYTHON) lib/audit_format.py

exe:
	powershell -NoProfile -ExecutionPolicy Bypass -File .\build_exe.ps1
