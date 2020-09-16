# Debug-and-Document-VB
This will eventually hold several Excel addins I wrote to help Document and Debug 
VB scripts.  I have left all these spreadsheets as regular macro enabled spreadsheets.
You can convert them to the addin format by saving them as an Excel Addin ie. *.xlam. 
In the excel addin directory AppData > Roaming > Microsoft > Addins
Below is a breif description of each spreadsheet.


1) Print VB Components Addin

This addin complies a list of the Modules in the active workbook and displays these in a dropdown menu. Once a module is selected the Print control menu comes up .
Here you can select how you want the code displayed. Options are:
Select colors and font size to print Module, Subroutines , Functions, Comments, Code, With , Select.
Then select if you want to map Loops and If Then’s. You can track up to 4 embedded if then sequences. And select color.
Then Line numbers you can select how frequently you want to print them and in what color.
Then you can force Page breaks at the end of Subroutines or Functions
Then you can select what pages you want to print
Then if you want to show the Module name and the date in the header or footer.
Then select the font type to use (not all fonts are available) and the orientation.
When you select Done these choices are saved so you will not have to repeat the process.

Now you are sent to the sub ContinueProcess which controls the flow from here on .
First it runs ProcessData which analyses the contents of the CodeLineArry (which contains the text contents of the Module selected. It creates the PrintArry this has 4 elements for each line 
1 - Type of line ie Comment Code ect.
2 - # of characters in the line (used to determine if you need a line wrap
3 -  Loop Count 
4 – If Then count
While this is running you will see a progress window . This process is generally pretty quick .
This data is then used by the sub CreatePrintMod to Format the lines and keep track of the loops and if then sequences to create the Print Module sheet.  This subroutine also has a Progress form to give you an idea of where it is.  This processing time for this module starts to slow down a lot once you start working on Modules with over 800 lines of code. 

Then 2 buttons are added to the sheet one to proceed with printing the other allows you to go back to the Print Control form to make changes .

Flow chart from ContinueProcess

			
ContinueProcess			
	              ProcessData		
		                      ProgressMessage	
	              ProcessData		
                          	      FuncCmtorBlank	
	              ProcessData		
ContinueProcess			
	             CreatePrintMod		
		                       ProgressMessage	
	             CreatePrintMod		
ContinueProcess			
	             Create_Print_Button	
ContinueProcess			
	             Create_Return_Button	
ContinueProcess			
			
			
			
Sub/Func		Located in		Line # in Component
ContinueProcess		Print_Controls		19	
Create_Print_Button	CreatePrintModSheet	678	
Create_Return_Button	CreatePrintModSheet	706	


2) Code Doc and follow the sub	
	
First the Code Documentation Menu comes up as you have seen.	
	
Select the location of the output of the Code Documentor	
either your current workbook or a new workbook.	
If you choose a new workbook a name will be created but you can change 	
this to what you like. Make sure to include the extention type. Most likely .XLS	
Not all extention types will work . For example you can't create a .XLSM file 	
as the new file has no Macro's in it.	
	
Next you are asked about Notify if No Code. 	
I originally put this in to make sure the addin was working correctly but left 	
it in to allow users a choice.	
This refers to components which have no code in them . The three type of 	
components I usually run across a s standard Modules , Worksheets and UserForms	
The other 2 I haven't adressed yet are Document Module and ActiveX Designer.	
	
Modules would rarely have no code in them unless you have commented it all out 	
and therefore may want to remove that Module . Usually mark this as Yes.	
More as a precaution if the addin doesn't think there is code in a Module I want to know.	
	
Worksheets usually have no code in them so I usually mark this as no 	
	
UserForms may have no code in them, especially very simple ones. I usually 	
mark this as yes again just to make sure I agrre with the addin.	
	
Next you are asked which worksheets do you want created as output 	
	
Macro List 	
This  is an Alphabetical List of all Subroutines and Functions in the Worksheet.  It gives:	
Macro Name	
The line it is located at in the Module 	
The number of lines of code it comprises	
The Module it is located in 	
The procedure Type ie. Subroutine or Function	
The Procdure it is called from and the line number of the procedure it is called from	
	
I didn't include the calling Procedure module as you can easily get that by looking down the list to find	
the procedure name.	
	
Module Map 	
This is a list by Module of all subroutines and functions located in the Module and 	
all subroutines and functions called by the module	
You may see under Calling Sub/Func Nothing.  This means that there are isn't a 	
Calling Procedure it is probably activated by a command button on the UserForm .	
	
ShapeList	
This is a list by sheet of the shapes located on it , quite often buttons 	
The name of the shape , The Macro called by it , and any Text associated with the shape	
	
Named Ranges 	
This is a list of all named ranges 	
Contents are Range Name , The sheet it references and the Range Referenced	
Note: Some names may have #REF associated with them. These are	
broken names most likely due to a sheet  deletion. You probably want to delete these.	
	
Forms Information	
This goes thru each form and lists the controls type  , control name , value and caption 	
	

CreatePrintMod		CreatePrintModSheet	298	
FuncCmtorBlank		CreatePrintModSheet	268	
ProcessData		CreatePrintModSheet	2	
ProgressMessage		Get_Workbook_Components	190	
			
			

