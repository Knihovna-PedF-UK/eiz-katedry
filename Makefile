All: out

out: eiz.lua data/pez-katedry_I.xlsx
	mkdir  out
	texlua eiz.lua data/pez-katedry_I.xlsx out
	cp out/*.* ../pedf-web/src/
