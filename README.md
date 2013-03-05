JSONtoCSV
=========

This contains two files, first is a java .jar file which parses JSON (it has error handling that will default to my Facebook API rest url if you don't give it an input parameter). The .bat batch file should be saved in the same folder as the .jar and your excel file. The .bat file outputs the parsed json array into a .csv file, which can be opened with excel. To use this with VBA, use the shell() function to call the batch file (the batch file calls the .jar file and outputs the result). Example: Shell(toCsv.bat "json rest url"). Make sure within vba to also set your directory (before you call shell()), which can be done by writing the following: ChDir ThisWorkbook.path . Please cite Micah Fulton if you plan to use this source code.

In summary:
1. Put all files into same folder as excel macro file
2. Set your directory from vba (I recommend c drive to avoid admin issues)
3. Call the .bat file using Shell() function from vba
4. Cite Micah Fulton for saving lives.
