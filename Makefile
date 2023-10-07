COMPILER = xelatex
SOURCE = addressbook.tex

all:
	$(COMPILER) $(SOURCE)

clean:
ifneq ($(wildcard *.synctex.gz),)
	@echo 'Cleaning .synctex.gz files ....'
	rm *.synctex.gz
endif
ifneq ($(wildcard *.aux),)
	@echo 'Cleaning .aux files ....'
	rm *.aux
endif
ifneq ($(wildcard *.log),)
	@echo 'Cleaning .log files ....'
	rm *.log
endif