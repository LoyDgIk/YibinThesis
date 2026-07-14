.PHONY: all pdf word check clean distclean

MAIN ?= main.tex
PDF_OUT ?= build/pdf
WORD_OUT ?= build/word/main.docx
PYTHON ?= python3
PANDOC ?= pandoc

all: pdf word

pdf:
	latexmk -xelatex -outdir=$(PDF_OUT) $(MAIN)

word:
	$(PYTHON) tools/build_word.py --main $(MAIN) --output $(WORD_OUT) --pandoc $(PANDOC)

check:
	$(PYTHON) tests/audit_format.py
	$(MAKE) pdf MAIN=tests/smoke.tex PDF_OUT=build/pdf/tests
	$(MAKE) pdf MAIN=tests/smoke-science.tex PDF_OUT=build/pdf/tests

clean:
	latexmk -c -outdir=$(PDF_OUT) $(MAIN)

distclean:
	latexmk -C -outdir=$(PDF_OUT) $(MAIN)
