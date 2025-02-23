
IMPORTANT UPDATE
The preferred option is to NOT use the solver (i.e. and to select FALSE in the 'Install Solver Addin').
A solution has been found to improve the accuracy of the 'GoalSeek' tool and it yields results as accurate as the solver add-in.

FIX - TO IMPROVE GOALSEEK ACCURACY:
1. Select 'Options' from the 'File' tab
2. Choose 'Formulas'
3. On the top-right hand side, click on 'Enable iterative calculation' 
4. Reduce the 'Maximum Change' figure to the desired accuracy (i.e 0.0000000000001)




Excel Solver might be the cause of a few issues:


1. Solver Installation

If the Solver Add-in doesn't get installed automatically in Excel, please add it manually before clicking on the calibration buttons.

To install it:

Windows
a. Click the File tab, click Options, and then click the Add-ins category.
b. In the Manage box, click Excel Add-ins, and then click Go.
c. In the Add-ins available box, select the Solver Add-in check box.
   If you don't see this name in the list, click the Browse... button and navigate to the folder containing Solver.xlam.
   Then click OK.
d. Now on the Data tab, in the Analysis group, you should see the Solver command.

Mac
a. Click the Tools menu, then click the Add-ins command.
b. In the Add-ins available box, select the Solver.xlam check box.
   If you don't see this name in the list, click the Select... button and navigate to the folder containing Solver.xlam.
   Then click OK.
c. Now on the Tools menu, you should see the Solver command.




2. Not able to Install the Solver 

If you are not able to install the solver, you might still get some VBA error where the solver is called. 
If that is the case, you might need to comment those lines. 
We will aim to provide a fix for this in the next release



3. Solver stopping at every scenario
This is an issue we don't have solution for and which seems to be specific to Mac.
If you do find a solution for this, please contact us.

Normally the parameter 'UserFinish' when set to 'True' should deactivate this.

In the sub 'RootFindingIndividualIRQuote' in the module 'rngCalibration' you can try replacing the line:

	SolverSolve UserFinish:=True, ShowRef:="SolverDisplayFunction"

by 

	Call SolverSolve(True, "SolverDisplayFunction")


If this still doesn't work, you will have to set the parameter 'Install Solver' in the 'Configuration' tab to False, so that it uses GoalSeek instead of the Solver