This Workbook contains the Macros for consolidating the Scrap Mat Excel Files  
	
You should run DeleteRows150_1400 on each sheet on each workbook, at least in the BPCuffs Folder	
This will cut down on the blank spaces. The Macro deletes rows 150-1400, be sure not to delete data!	
	
If the Worksheet names are not in a column, then run wsheetnamecolumnadd.	
This will add the worksheet name to a column (hopefully by a msg box)	
	
Run ConsolidateSheets, by default, True, True, and then select the folder with the excel files to consolidate	
If this errors out, it will most likely be do to too many rows trying to be copied.	
If this is the case, try to prune some blank rows out of the files.	
	
At this point you can run deleteblankrows.	
However I would suggest putting Total or some data in the last row in column A	
and a total scrap in the same row and record the number.	
This will check Column A and B for data, and if both are empty, delete that row.	
This check starts at the last column with data in A, hence the adding of Total in last row in column A	
