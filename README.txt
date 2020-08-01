Directory Structure:

    /Forms
        AddLeadMeasure.frm 
        AddLeadMeasure.frx 
        AddWIG.frm 
        AddWIG.frx 
        AdminLogin.frm 
        AdminLogin.frx 
        Login.frm 
        Login.frx 
        ModifyLead.frm 
        ModifyLead.frx 
        ModifyWIG.frm 
        ModifyWIG.frx 
        Register.frm 
        Register.frx 
    /Modules
        TableFunctions.bas
        Team4DXButtons.bas
        Team4DXGenerator.bas
    /Program
        Applying4DXtoSeniorProjects.xlsm
    /Sheets
        InformationInput.xlsm
        Start.xlsm
        ThisWorkbook.cls

Forms - This directory contains the form files that were exported from the 
Applying4DXtoSeniorProjects.xlsm program. These forms are important for 
Login/Registration as well as allowing teams to literally apply 4DX to 
their project. If you are going to import these forms in a new macro-enabled 
workbook, you only need to import the *.frm files.

Modules - This directory contains the files for the exported code from the 
Applying4DXtoSeniorProjects.xlsm program that are seperate from the code
attached to the forms and sheets. TableFunctions.bas contains the code
for modifying the WIG, Lead, and Scoreboard tables. Team4DXButtons.bas and
Team4DXGenerator.bas contains the code for generating the 4DX project sheets.
If you are going to import the forms in a new macro-enabled workbook, you 
also need to import these Module files in the same workbook.

Program - This directory contains the Applying4DXtoSeniorProjects program
executable itself. The macro-enabled workbook can be used or developed further 
through Excel's VBA IDE in this state.

Sheets - This directory contains the main sheets that are not generated in 
the program's lifespan and are needed for the program to function properly.
If you are going to create a new macro-enabled workbook and import the individual
forms and modules, you also need to open these single-sheet workbooks and move a 
copy of the "Start" sheet in Start.xlsm and the "InformationInput" sheet in 
InformationInput.xlsm into the new workbook. Also, open ThisWorkbook.cls and
copy the code below the written comment and paste it into the ThisWorkbook 
object of your new workbook. 

The files in Forms, Modules, and Sheets are static, and do not change as the 
program changes since they are manually exported from the program. When finished
making changes to the program, export the changed files and replace the ones in
folders.

   