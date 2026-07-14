$pdf_mode = 5;
$xelatex = 'xelatex %O -interaction=nonstopmode -halt-on-error -file-line-error -synctex=1 %S';
$bibtex_use = 2;
$biber = 'biber %O %B';
@default_files = ('main.tex');
$clean_ext .= ' bbl bcf blg run.xml synctex.gz';
