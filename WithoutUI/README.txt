Excel Archival Tool v 0.1
Developed at the University of Minnesota for the Data Repository for the U of M (DRUM)
Creator: John McGrory (mcgr0237@umn.edu) 
Released: October 23, 2014

This program can convert Excel files (.xls, .xlsx) to comma-separated value files (.csv), while
	also extracting charts/figures, cell formulas, and an HTML snapshot that captures 
	spreadsheet formatting. This creates an "archival information packet" suitable for the
	preservation of the information contained within a proprietary Excel file.

Program Requirements:
	- This program consists of one file: ExcelArchivalTool.vbs
	- This file requires a Windows environment (XP, 7)

Instructions for use:
	1. Close all instances of Excel that may be currently running on your computer
	2. Open ExcelArchivalTool.vbs to begin the program
	3. Select a "Source": this can be a single Excel file or a folder containing multiple Excel files
		- Note: if you select a folder, the program will recursively find and convert Excel files
			within subfolders
	4. Select a "Destination Folder" where the .csv files, charts/figures, cell formula logs, and/or 
		HTML snapshots will be stored
		- Note: It is recommended that your "Destination folder" be empty for the
			conversion process
		- You may want to call this folder "Archival Version of Data"
	5. If you selected a directory as your "Source", the program will count how many Excel files are 
		contained within the "Source" folder and prompt you to determine if you want to continue.
	6. The program will run in the background until completion.
	5. You will be presented a pop-up that reads "Conversion complete" when the program is 
		finished. 
	6. If you have more folders to convert, start again at Step 1. 


An output report will be generated and placed in the "Destination Folder". This contains the following information:
	- Name of file/folder that was converted
	- Date and time of conversion
	- Number of Excel files converted
	- Number of .csv files generated
	- Number of charts/figures exported as .png files
	- Number of formula files generated


The license (GNU GPLv3) text is found in the "License.html" file in the program folder.