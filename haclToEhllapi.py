# ************************************************************************* #
#                                                                           #
#  Module Name           : haclToEhllapi.py                                 #
#                                                                           #
#  Descriptive Name      : Main file to convert HACL API's to EHLLAPI       #
#                                                                           #
#  Author                : Prashant Iyer                                    #
#                                                                           #
# ************************************************************************* #

# Import appropriate modules and files
import getopt
from ehllapiFunctions import *
from tkinter import *
from fileNames import *
import os

# Declare global variable
haclMethodName = ''

# Accepting command line script
commandLineInput = ''
commandLineOutput = ''
if __name__ == "__main__":
    argv = sys.argv[1:]
    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["ifile=", "ofile="])
    except getopt.GetoptError:
        print('test.py -i <commandLineInput> -o <commandLineOutput>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('test.py -i <commandLineInput> -o <commandLineOutput>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            commandLineInput = arg
        elif opt in ("-o", "--ofile"):
            commandLineOutput = arg

# ------------------------------------------------------------------------- #
#  Function    : converter                                                  #
#                                                                           #
#  Description : Converts HACL to EHLLAPI                                   #
#                                                                           #
#  Parameter(s): None                                                       #
#                                                                           #
#  Return value: None                                                       #
#                                                                           #
# ------------------------------------------------------------------------- #
def converter():
    # Access input and output files from GUI and CLI
    if commandLineInput == '' or commandLineOutput == '':
        haclInputFile = haclInputFileUserInput.get()
        ehllapiOutputFile = ehllapiOutputFileUserInput.get()
    else:
        haclInputFile = commandLineInput
        ehllapiOutputFile = commandLineOutput

    # Creating default files
    if haclInputFile == '':
        haclInputFile = haclInputDefaultFile
    if ehllapiOutputFile == '':
        ehllapiOutputFile = ehllapiOutputDefaultFile

    # Initialise all files
    haclInputFileObject = open(haclInputFile, "r")
    ehllapiOutputFileObject = open(ehllapiOutputFile, "w")
    ehllapiHeaderFileObject = open(ehllapiHeaderFile, "r")
    inputDictionaryFileObject = open(inputDictionaryFile, "r")
    ehllapiHeaderDictionaryFileObject = open(ehllapiHeaderDictionaryFile, "r")
    haclInputFileContents = haclInputFileObject.read()
    ehllapiHeaderFileContents = ehllapiHeaderFileObject.read()
    inputDictionaryFileContents = inputDictionaryFileObject.read()
    ehllapiHeaderDictionaryFileContents = ehllapiHeaderDictionaryFileObject.read()

    # Creating input dictionary from inputDictionary.txt
    inputMappingDictionary = ast.literal_eval(inputDictionaryFileContents)

    # Create EHLLAPI header dictionary from ehllapiHeaderDictionary.txt
    ehllapiHeaderDictionary = ast.literal_eval(ehllapiHeaderDictionaryFileContents)

    # Split lines of input files into a list
    haclCommands = haclInputFileContents.split('\n')

    # Validate HACL header
    haclInputHeader = ''
    haclCorrectHeader = "[PCOMM SCRIPT HEADER]LANGUAGE=VBSCRIPTDESCRIPTION=[PCOMM SCRIPT SOURCE]OPTION EXPLICIT"
    count = 0
    while count < 5:
        haclInputHeader = haclInputHeader + haclCommands[count].rstrip(' ')
        count = count + 1
    if haclInputHeader != haclCorrectHeader:
        print("Invalid header")
        exit()

    # Identifying the method name of the main method using regular expressions
    for haclCommandLine in haclCommands:
        if haclCommandLine[:4] == "sub ":
            haclMethodRE = re.search(r'^.*sub\s(.*)\(.*\)$', haclCommandLine, flags=0)
            global haclMethodName
            haclMethodName = haclMethodRE.group(1)

    # Outputting default headers and variable declarations, and beginning of main method
    ehllapiOutputFileObject.write(ehllapiHeaderFileContents)
    ehllapiOutputFileObject.write("Sub " + haclMethodName + "()" + '\n')

    # Reset EHLLAPI header list
    ehllapiHeaderList = []

    # Iterate through all the commands in the input file
    for haclCommandsIndex in range(len(haclCommands)):
        haclInputLine = haclCommands[haclCommandsIndex]

        # Removing all indents
        if haclInputLine[:3] == "   ":
            haclInputLine = haclInputLine.replace("   ", "")

        # Identifying lines that are HACL functions excluding connect and disconnect
        if haclInputLine[:3] == "aut" and '(' not in haclInputLine and ')' not in haclInputLine:

            # Use regular expressions to separate function and parameters from rest of the input
            haclFunctionAndParamRE = re.search(r'^.*\..*\.(.*$)', haclInputLine, flags=0)
            haclFunctionAndParams = haclFunctionAndParamRE.group(1)

            # Use regular expressions to separate function name from parameters
            haclFunctionRE = re.search(r'(^.*?)\s(.*$)', haclFunctionAndParams, flags=0)

            # Identifying whether the regular expression has a match, meaning the line is in the correct format
            if haclFunctionRE:
                # Remove whitespace from input line
                haclFunction = haclFunctionRE.group(1).replace(' ', '')
                haclParameters = haclFunctionRE.group(2)

                # Add headers to EHLLAPI header list
                if haclFunction in inputMappingDictionary:
                    ehllapiHeaderList.append(ehllapiHeaderDictionary[inputMappingDictionary[haclFunction]])

                # Place all parameters in list
                haclParametersList = haclParameters.split(',')

                # Remove comma and double quotes from all parameters
                for haclParametersIndex in range(len(haclParametersList)):
                    haclParametersList[haclParametersIndex] = haclParametersList[haclParametersIndex].replace('"', '')
                    haclParametersList[haclParametersIndex] = haclParametersList[haclParametersIndex].replace(' ', '')

                # Identifying if the function has a map in the input dictionary
                if haclFunction in inputMappingDictionary:
                    # Use dictionary to find function number
                    ehllapiFunctionNumber = inputMappingDictionary[haclFunction]
                    # Perform appropriate function to generate output
                    if ehllapiFunctionNumber == 3:
                        ehllapiOutputString = ehllapiFunctionSendKey(haclParametersList[0])
                    elif ehllapiFunctionNumber == 4:
                        ehllapiOutputString = ehllapiFunctionWaitParams(haclParametersList[0])

                    # Output to output file
                    ehllapiOutputLineList = ehllapiOutputString.split('\n')
                    for ehllapiOutputLine in ehllapiOutputLineList:
                        ehllapiOutputFileObject.write('\t' + ehllapiOutputLine + '\n')
                else:
                    # Writing the line to the output file as it is since it is not mapped
                    ehllapiOutputFileObject.write('\t' + "'Input copied without conversion-" + '\n')
                    ehllapiOutputFileObject.write('\t' + "'" + haclInputLine + '\n')

            # Identifying that the line does not match the regular expression, meaning it has no parameters
            else:
                haclFunction = haclFunctionAndParams

                # Add headers to EHLLAPI header list
                if haclFunction in inputMappingDictionary:
                    ehllapiHeaderList.append(ehllapiHeaderDictionary[inputMappingDictionary[haclFunction]])

                # Identifying if the function has a map in the input dictionary
                if haclFunction in inputMappingDictionary:
                    # Use dictionary to find function number
                    ehllapiFunctionNumber = inputMappingDictionary[haclFunction]
                    # Perform appropriate function to generate output
                    if ehllapiFunctionNumber == 4:
                        ehllapiOutputString = ehllapiFunctionWaitNoParams()

                    # Output to output file
                    ehllapiOutputLineList = ehllapiOutputString.split('\n')
                    for ehllapiOutputLine in ehllapiOutputLineList:
                        ehllapiOutputFileObject.write('\t' + ehllapiOutputLine + '\n')

                # Writing the line to the output file as it is since it is not mapped
                else:
                    ehllapiOutputFileObject.write('\t' + "'Input copied without conversion-" + '\n')
                    ehllapiOutputFileObject.write('\t' + "'" + haclInputLine + '\n')

        # Identifying that the line has a different format of the connect function, containing parentheses
        elif haclInputLine[:3] == "aut" and '(' in haclInputLine and ')' in haclInputLine:
            # Performing regular expressions to identify the function name
            haclConnectFunctionRE = re.search(r'^.*\.(.*)\(', haclInputLine, flags=0)
            haclConnectFunctionName = haclConnectFunctionRE.group(1)

            # Confirming that the line is the connect function
            if haclConnectFunctionName == "SetConnectionByName":
                haclFunction = "SetConnectionByName"

            # Add headers to EHLLAPI header list
            if haclFunction in inputMappingDictionary:
                ehllapiHeaderList.append(ehllapiHeaderDictionary[inputMappingDictionary[haclFunction]])

            # Perform connect function to generate output
            ehllapiConnectOutputString = ehllapiFunctionConnectPS()

            # Output to output file
            ehllapiConnectOutputLineList = ehllapiConnectOutputString.split('\n')
            for ehllapiConnectOutputLine in ehllapiConnectOutputLineList:
                ehllapiOutputFileObject.write('\t' + ehllapiConnectOutputLine + '\n')

        # Identifying that the line is the end of the main method
        elif haclInputLine == "end sub":
            # Performing disconnect function
            haclFunction = "DisconnectPS"

            # Add headers to EHLLAPI header list
            if haclFunction in inputMappingDictionary:
                ehllapiHeaderList.append(ehllapiHeaderDictionary[inputMappingDictionary[haclFunction]])

            # Perform disconnect function to generate output
            ehllapiDisconnectOutputString = ehllapiFunctionDisconnectPS()

            # Output to output file
            ehllapiDisconnectOutputLineList = ehllapiDisconnectOutputString.split('\n')
            for ehllapiDisconnectOutputLine in ehllapiDisconnectOutputLineList:
                ehllapiOutputFileObject.write('\t' + ehllapiDisconnectOutputLine + '\n')

        # Identifying exception lines like the main method name, method call, auto-generated comment and ignoring them
        elif haclInputLine == ("sub " + haclMethodName + "()") or haclInputLine == haclMethodName or haclInputLine == "REM This line calls the macro subroutine":
            continue

        # Identifying that the line is not a function or an exception
        else:
            # Writing the line to the output as long as it is not the header
            if haclCommandsIndex not in range(6) and haclInputLine != "":
                ehllapiOutputFileObject.write('\t' + "'Input copied without conversion-" + '\n')
                ehllapiOutputFileObject.write('\t' + haclInputLine + '\n')

    # Ending the method
    ehllapiOutputFileObject.write("End Sub \n")

    # Removing duplicate headers from list
    ehllapiHeaderList = list(dict.fromkeys(ehllapiHeaderList))

    # Adding EHLLAPI headers to the beginning of the output file
    for ehllapiHeaderLine in ehllapiHeaderList:
        ehllapiOutputFileObject = open(ehllapiOutputFile, 'r+')
        ehllapiOutputFileContents = ehllapiOutputFileObject.read()
        ehllapiOutputFileObject.seek(0, 0)
        ehllapiOutputFileObject.write(ehllapiHeaderLine.rstrip('\r\n') + '\n' + ehllapiOutputFileContents)

    print("Converted")


# ------------------------------------------------------------------------- #
#  Function    : createExcelFile                                            #
#                                                                           #
#  Description : Uses a command line argument to run a VBScript file that   #
#                generates an xlsm file with a macro assigned to a button   #
#  Parameter(s): None                                                       #
#                                                                           #
#  Return value: None                                                       #
#                                                                           #
# ------------------------------------------------------------------------- #
def createExcelFile():
    # Using a command line argument to run the VBScript file that generates the excel file
    try:
        commandLine = "cmd /c \"cscript createExcelFile.vbs " + haclMethodName + '"'
        os.system(commandLine)
        print("Created Excel file")
    except:
        print("File may be open")


# ------------------------------------------------------------------------- #
#  Function    : exit                                                       #
#                                                                           #
#  Description : Closes the GUI of the program                              #
#                                                                           #
#  Parameter(s): None                                                       #
#                                                                           #
#  Return value: None                                                       #
#                                                                           #
# ------------------------------------------------------------------------- #
def exit():
    # Destroys the GUI window
    master.destroy()
    print("Closed GUI")


# Running GUI if there are no command line arguments
if commandLineInput == '' or commandLineOutput == '':
    master = Tk()
    master.title("haclToEhllapiConverter")
    canvas = Canvas(master, width=600, height=260, relief="raised")
    canvas.pack()

    inputInfoLabel = Label(master, text="Enter the name of the input file that contains the VBScript code to convert from HACL")
    canvas.create_window(300, 20, window=inputInfoLabel)
    inputLabel = Label(master, text="Input file:")
    canvas.create_window(300, 40, window=inputLabel)
    haclInputFileUserInput = Entry(master)
    canvas.create_window(300, 60, window=haclInputFileUserInput)

    outputInfoLabel = Label(master, text="Enter the name of the output file that will contain the VBScript code converted to EHLLAPI")
    canvas.create_window(300, 80, window=outputInfoLabel)
    outputLabel = Label(master, text="Output file:")
    canvas.create_window(300, 100, window=outputLabel)
    ehllapiOutputFileUserInput = Entry(master)
    canvas.create_window(300, 120, window=ehllapiOutputFileUserInput)

    convertButton = Button(master, text="Convert", command=converter)
    canvas.create_window(300, 150, window=convertButton)

    createButton = Button(master, text="Create Excel file", command=createExcelFile)
    canvas.create_window(300, 180, window=createButton)

    exitButton = Button(master, text="Exit", command=exit)
    canvas.create_window(300, 210, window=exitButton)

    summaryLabel = Label(master, text="This program converts VBScript code generated using HACL API's to EHLLAPI API's")
    canvas.create_window(300, 240, window=summaryLabel)

    master.mainloop()
# Running converter function using command line arguments
else:
    converter()
    createExcelFile()
