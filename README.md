# Debug-and-Document-VB

Print VB Components Addin

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
			
			
			
Sub/Func	Located in	Line # in Component
ContinueProcess	Print_Controls	19	
Create_Print_Button	CreatePrintModSheet	678	
Create_Return_Button	CreatePrintModSheet	706	
CreatePrintMod	CreatePrintModSheet	298	
FuncCmtorBlank	CreatePrintModSheet	268	
ProcessData	CreatePrintModSheet	2	
ProgressMessage	Get_Workbook_Components	190	
			
			

