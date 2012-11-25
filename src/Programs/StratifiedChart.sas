/* Create a simple bar graph for the data to show the rankings */
/* and relative values */
goptions ypixels=%eval(250 * &categorycount) xpixels=500;
proc gchart data=topn
;
	hbar &report / 
		sumvar=&measure
		group=&category
		descending
		nozero
		clipref
		frame	
		discrete
		type=&stat
		patternid=group
;
run;
quit;