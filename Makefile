DOCTYPE = RTN
DOCNUMBER = 046
DOCNAME = $(DOCTYPE)-$(DOCNUMBER)

tex = $(filter-out $(wildcard *acronyms.tex) , $(wildcard *.tex))

GITVERSION := $(shell git log -1 --date=short --pretty=%h)
GITDATE := $(shell git log -1 --date=short --pretty=%ad)
GITSTATUS := $(shell git status --porcelain)
ifneq "$(GITSTATUS)" ""
	GITDIRTY = -dirty
endif

export TEXMFHOME ?= lsst-texmf/texmf

# Add aglossary.tex as a dependancy here if you want a glossary (and remove acronyms.tex)
$(DOCNAME).pdf: $(tex) meta.tex local.bib acronyms.tex 
	latexmk -bibtex -xelatex -f $(DOCNAME)
#	makeglossaries $(DOCNAME)
#	xelatex $(DOCNAME)
# For glossary uncomment the 2 lines above

gdepend:
	pip install --upgrade pip
	pip install --upgrade -r requirements.txt

# Acronym tool allows for selection of acronyms based on tags - you may want more than DM
acronyms.tex: $(tex) myacronyms.txt
	$(TEXMFHOME)/../bin/generateAcronyms.py -t "DM" $(tex)

# If you want a glossary you must manually run generateAcronyms.py  -gu to put the \gls in your files.
aglossary.tex :$(tex) myacronyms.txt
	generateAcronyms.py  -g $(tex)


.PHONY: clean
clean:
	latexmk -c
	rm -f $(DOCNAME).{bbl,glsdefs,pdf}
	rm -f meta.tex

.FORCE:

meta.tex: Makefile .FORCE
	rm -f $@
	touch $@
	printf '%% GENERATED FILE -- edit this in the Makefile\n' >>$@
	printf '\\newcommand{\\lsstDocType}{$(DOCTYPE)}\n' >>$@
	printf '\\newcommand{\\lsstDocNum}{$(DOCNUMBER)}\n' >>$@
	printf '\\newcommand{\\vcsRevision}{$(GITVERSION)$(GITDIRTY)}\n' >>$@
	printf '\\newcommand{\\vcsDate}{$(GITDATE)}\n' >>$@



tree:
	python3 bin/makeProductTree.py 1UZe1Mm5OeHxg-NS2SlSTLaNvRrppARIUGLYJaroXduY FEver\!A1:H --depth=2
	pdflatex ProductTree.tex
	mv ProductTree.pdf ProductTreeSmall.pdf
	python3 bin/makeProductTree.py 1UZe1Mm5OeHxg-NS2SlSTLaNvRrppARIUGLYJaroXduY FEver\!A1:H 
	pdflatex ProductTree.tex

# pick up this form the lsst-texmf/bin
tables: .FORCE
	cd tables; makeTablesFromGoogle.py 1KAmk2NcSFknXiqbwMekaHkTQOzZ6RUAgt7_tqvBi1HM DMops\!A1:F

