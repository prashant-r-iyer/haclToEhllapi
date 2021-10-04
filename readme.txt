Title: haclToEhllapi - Utility to convert IBM HACL Automation Objects to IBM EHLLAPI functions

Author- Prashant Iyer

Description-
------------------------------------------------------------------------------------------------------------------------
The input text file or .mac file contains the HACL code. The user can use a GUI by running the program, or use a CLI to
perform the conversion of HACL to EHLLAPI. After this, an Excel file is created with the user's desired name. This
Excel file will contain a button assigned to the converted macro.

Usage-
------------------------------------------------------------------------------------------------------------------------
a) GUI
Enter the name of the text file or .mac file containing the HACL code as the input file
Enter the name of the text file which will have the EHLLAPI output as the output file
Clicking "Convert" will perform the conversion and write the output to the output file
Clicking "Create Excel file" will ask for the names of the EHLLAPI output file and desired name of the output Excel file
and will create this file in the "Documents" folder, with a button assigned to the converted macro
Clicking "Exit" will close the GUI

b) CLI
The format to run the program using command line arguments is:
    converterMain.py -i <inputFile> -o <outputFile>
<inputFile> refers to the name of the text file or .mac file containing the HACL code as the input file
<outputFile> refers to the name of the text file which will have the EHLLAPI output as the output file
Running the program using CLI will will ask for the names of the EHLLAPI output file and desired name of the output
Excel file and will create this file in the "Documents", with a button assigned to the converted macro

Files in package-
------------------------------------------------------------------------------------------------------------------------
createExcelFile.vbs: creates the output Excel file containing a button assigned to the converted macro
ehllapiFunctions.py: contains functions for all API's that generates the EHLLAPI code
ehllapiHeaderDictionary.txt: contains mappings from EHLLAPI function number to the respective API's header
ehllapiOutput.txt: the default output file for the program that contains the converted EHLLAPI code
ehllapiVariableDeclarations.txt: contains default variable declarations for variables in the converted EHLLAPI code
fileNames.py: contains the names of all files accessed in the main file
haclInput.txt: contains a sample input HACL program, it is the default input file
haclToEhllapi.py: main fail for conversion to take place
inputDictionary.txt: contains mappings from HACL function names to EHLLAPI function numbers
specialCharacterDictionary.txt: contains mappings from special characters in HACL to their EHLLAPI counterparts

END OF README
