### Macros for SwitchTask

These are in two parts. First is the before the scanning task and there's only one file for that (Create_imgseq).  The rest are parts of the analysis on the behavioural data that came out of the scanning task. 


From a readme I wrote for myself some time ago (Jul 2010)

	Current workflow for Macros is the following
	
	0. Take the files in folder 'ScanFilesOutput'
	1. EventsMacro (do not write anything out), Specifically, run the routine 'defineEvents'
	- Make new sheet called 'BehavData'
	- Run Macro 'countTrials'
	2. Behaviour Macro
	3. Make Event Files (to first calulate the mean latencies then to write files out)

