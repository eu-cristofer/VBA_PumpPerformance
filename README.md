# VBA Pump Performance

VBA Pump Performance is a set of tools to help engineers and technicians to asses the condition of API 610 pumps during its Performance and Mechanical Running Test trials.

# Resources
* Performance computation
* Operational curves generation
* Speed correction
* Relative density correction
* Viscous effects correction
* Report issue automation

## Installation in Microsoft Excel

1. Go to the Developer tab.
2. Click the Add-ins Button.
3. Inside the Add-ins Dialog Box, click the Browse… ...
4. The Explorer Window should default to the Microsoft add-in folder location.
5. Navigate and select the file `VBA-Pump-Performance.xlam`, then click OK.

### Sample File

You can try the functionalities with the file `SampleData.xlsx`.

## Technical specifications

The solution was written in VBA and built in a Microsoft Office Add-in.

The performance acceptance criteria adopted in this project is the API 610 11th edition, Centrifugal pumps for petroleum, petrochemical and natural gas industries.

The computation of test parameters follows the ANSI/HI 14.6 Rotodynamic Pumps - Hydraulic Performance Acceptance Tests.

The computation of viscosity correction is accordingly the ANSI/HI 9.6 Rotodynamic Pumps -  Guideline for Effects of Liquid Viscosity on Performance.

## Special notes

The author of the code does not make any warranty or representation, either express or implied, with respect to the accuracy, completeness, or usefulness of the results contained herein, or assume any liability or responsibility for any use, or the results of such use, of any information or process disclosed in this software.

# Guidelines to build the solution as an Add-in

1. Open the file `00 - BaseSpreadSheet.xlsx` in Excel Application.
2. Through the developer tab, access Visual Basic and import the script `01 - ImportModules.bas`.
3. Inside the imported module, adjust the variable `strAddress` with the path to the source (`src`) folder, i.e. `"C:\Users\john\GitHub\vba-pump-performance/src/"`.
4. Run the sub `01 - ImportModules`.
