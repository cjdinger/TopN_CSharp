/* Create a simple bar graph for the data to show the rankings */
/* and relative values */
goptions xpixels=600 ypixels=400;
proc gchart data=topn
;
	hbar &report / 
		sumvar=&measure
		descending
		nozero
		clipref
		frame	
		discrete
		type=&stat
;
run;
quit;
