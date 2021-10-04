' ************************************************************************* '
'                                                                           '
'  Module Name           : createExcelFile.vbs                              '
'                                                                           '
'  Descriptive Name      : Creating Excel file containing macro and button  '
'                          to run the macro                                 '
'  Author                : Prashant Iyer                                    '
'                                                                           '
' ************************************************************************* '

'Defining variables
Dim excelObject
Dim outputWorkbook
Dim excelModule
Dim ehllapiOutputFileObject
Dim content
Dim fileObject
Dim ehllapiOutputFile
Dim excelOutputFile
Dim ehllapiOutputDirectory
Dim excelOutputDirectory
Dim button
Dim methodName

'Accepting command line script
Set args = Wscript.Arguments
methodName = WScript.Arguments.Item(0)

'Creating the Excel object
Set excelObject = CreateObject("Excel.Application")

'Creating a workbook in Excel
excelObject.Application.Visible = False
Set outputWorkbook = excelObject.Workbooks.add
excelObject.Application.DisplayAlerts = False

'Creating the Excel module
Set excelModule = outputWorkbook.VBProject.VBComponents.Add(1)

'Inputting the EHLLAPI input file containing EHLLAPI code and the name of the Excel output file
ehllapiOutputFile = InputBox("Enter the name of the EHLLAPI output file")
excelOutputFile = InputBox("Enter the name of the Excel output file")
'ehllapiOutputFile = "ehllapiOutput.txt"
'excelOutputFile = "a.xlsm"
ehllapiOutputDirectory = ".\" & ehllapiOutputFile
excelOutputDirectory = ".\" & excelOutputFile

'Reading the contents from the EHLLAPI output file
Set fileObject = CreateObject("Scripting.FileSystemObject")
Set ehllapiOutputFileObject = fileObject.OpenTextFile(ehllapiOutputDirectory, 1)
content = ehllapiOutputFileObject.ReadAll

'Injecting the macro into the Excel file
excelModule.CodeModule.AddFromString content

'Creating a button in the Excel file
Set button = excelObject.ActiveSheet.Buttons.Add(50, 50, 50, 50)
button.text = "Click to run"
button.OnAction = methodName

'Saving the Excel file
excelObject.Activeworkbook.SaveAs excelOutputDirectory, 52
excelObject.Activeworkbook.Close
