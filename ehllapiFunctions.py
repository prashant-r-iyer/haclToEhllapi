# ************************************************************************* #
#                                                                           #
#  Module Name           : ehllapiFunctions.py                              #
#                                                                           #
#  Descriptive Name      : Contains all functions to generate EHLLAPI output#
#                                                                           #
#  Author                : Prashant Iyer                                    #
#                                                                           #
# ************************************************************************* #

# Import appropriate modules and files
from fileNames import *
import ast

# Open the text file containing contents of the special character dictionary and create it
specialCharacterDictionaryFileObject = open(specialCharacterDictionaryFile)
specialCharacterDictionaryFileContents = specialCharacterDictionaryFileObject.read()
specialCharacterDictionary = ast.literal_eval(specialCharacterDictionaryFileContents)

# ------------------------------------------------------------------------- #
#  Function    : specialCharacterConverter                                  #
#                                                                           #
#  Description : Converts special characters from HACL to EHLLAPI using a   #
#                mapping dictionary                                         #
#  Parameter(s): Text to replace characters in                              #
#                                                                           #
#  Return value: Converted text                                             #
#                                                                           #
# ------------------------------------------------------------------------- #
def specialCharacterConverter(arg1):
    returnString = arg1
    for specialCharacter in specialCharacterDictionary:
        if specialCharacter in returnString:
            returnString = returnString.replace(specialCharacter, specialCharacterDictionary[specialCharacter])
    return returnString

# ------------------------------------------------------------------------- #
#  Function    : ehllapiFunctionSendKey                                     #
#                                                                           #
#  Description : Converts from HACL to EHLLAPI for the SendKeys API         #
#                                                                           #
#  Parameter(s): The HACL argument text to be outputted to the screen       #
#                                                                           #
#  Return value: The EHLLAPI counterpart of the input                       #
#                                                                           #
# ------------------------------------------------------------------------- #
def ehllapiFunctionSendKey(arg1):
    arg1 = specialCharacterConverter(arg1)
    HllFunctionNumber = 3
    HllData = "\"" + arg1 + "\""
    HllLength = len(arg1)
    HllReturnCode = 0
    returnString = "'The SendKey function outputs its parameter to the screen" + '\n'
    returnString = returnString + "HllFunctionNumber = " + str(HllFunctionNumber) + '\n'
    returnString = returnString + "HllData = " + str(HllData) + '\n'
    returnString = returnString + "HllLength = " + str(HllLength) + '\n'
    returnString = returnString + "HllReturnCode = " + str(HllReturnCode) + '\n'
    returnString = returnString + "retVal = PCOMM_SendKey(" + str(HllFunctionNumber) + ", " + str(HllData) + ", " + str(HllLength) + ", 0)" + '\n'
    return returnString

# ------------------------------------------------------------------------- #
#  Function    : ehllapiFunctionConnectPS                                   #
#                                                                           #
#  Description : Converts from HACL to EHLLAPI for the SetConnectionByName  #
#                API                                                        #
#  Parameter(s): None                                                       #
#                                                                           #
#  Return value: The EHLLAPI counterpart of the input                       #
#                                                                           #
# ------------------------------------------------------------------------- #
def ehllapiFunctionConnectPS():
    HllFunctionNumber = 1
    HllData = "sessionName"
    HllLength = 1
    HllReturnCode = 0
    returnString = "'The ConnectPS function connects to the presentation space" + '\n'
    returnString = returnString + 'sessionName = InputBox("Enter session name")' + '\n'
    returnString = returnString + "HllFunctionNumber = " + str(HllFunctionNumber) + '\n'
    returnString = returnString + "HllData = " + str(HllData) + '\n'
    returnString = returnString + "HllLength = " + str(HllLength) + '\n'
    returnString = returnString + "HllReturnCode = " + str(HllReturnCode) + '\n'
    returnString = returnString + "retVal = PCOMM_ConnectPS(" + str(HllFunctionNumber) + ", " + str(HllData) + ", " + str(HllLength) + ", 0)" + '\n'
    return returnString

# ------------------------------------------------------------------------- #
#  Function    : ehllapiFunctionDisconnectPS                                #
#                                                                           #
#  Description : Converts from HACL to EHLLAPI for the Disconnect API       #
#                                                                           #
#  Parameter(s): None                                                       #
#                                                                           #
#  Return value: The EHLLAPI counterpart of the input                       #
#                                                                           #
# ------------------------------------------------------------------------- #
def ehllapiFunctionDisconnectPS():
    HllFunctionNumber = 2
    HllData = 0
    HllLength = 0
    HllReturnCode = 0
    returnString = "'The DisconnectPS function disconnects from the presentation space" + '\n'
    returnString = returnString + "HllFunctionNumber = " + str(HllFunctionNumber) + '\n'
    returnString = returnString + "HllData = " + str(HllData) + '\n'
    returnString = returnString + "HllLength = " + str(HllLength) + '\n'
    returnString = returnString + "HllReturnCode = " + str(HllReturnCode) + '\n'
    returnString = returnString + "retVal = PCOMM_DisconnectPS(" + str(HllFunctionNumber) + ", " + str(HllData) + ", " + str(HllLength) + ", 0)" + '\n'
    return returnString

# ------------------------------------------------------------------------- #
#  Function    : ehllapiFunctionWaitParams                                  #
#                                                                           #
#  Description : Converts from HACL to EHLLAPI for the Wait API             #
#                                                                           #
#  Parameter(s): The HACL argument time in milliseconds to wait             #
#                                                                           #
#  Return value: The EHLLAPI counterpart of the input                       #
#                                                                           #
# ------------------------------------------------------------------------- #
def ehllapiFunctionWaitParams(arg1):
    HllFunctionNumber = 4
    HllData = int(arg1)
    HllLength = 0
    HllReturnCode = 0
    returnString = "'The Wait function with parameters converts the input in seconds to minutes and for each minute it performs 1 loop" + '\n'
    returnString = returnString + "HllFunctionNumber = " + str(HllFunctionNumber) + '\n'
    returnString = returnString + "HllData = " + str(HllData) + '\n'
    returnString = returnString + "HllLength = " + str(HllLength) + '\n'
    returnString = returnString + "HllReturnCode = " + str(HllReturnCode) + '\n'
    returnString = returnString + "retVal = PCOMM_Wait(" + str(HllFunctionNumber) + ", " + str(HllData) + ", " + str(HllLength) + ", 0)" + '\n'
    returnString = returnString + "waitValueMinutes = " + str(HllData / (1000 * 60)) + '\n'
    returnString = returnString + "count = 0" + '\n'
    returnString = returnString + "breakFlag = False" + '\n'
    returnString = returnString + "While count < waitValueMinutes AND breakFlag = False" + '\n'
    returnString = returnString + '\t' + "If retVal = 1 Then" + '\n'
    returnString = returnString + '\t\t' + 'count = count + 1' + '\n'
    returnString = returnString + '\t' + "ElseIf retVal = 0 Then" + '\n'
    returnString = returnString + '\t\t' + "breakFlag = True" + '\n'
    returnString = returnString + '\t' + "End If" + '\n'
    returnString = returnString + "Wend" + '\n'
    return returnString

# ------------------------------------------------------------------------- #
#  Function    : ehllapiFunctionWaitNoParams                                #
#                                                                           #
#  Description : Converts from HACL to EHLLAPI for the WaitForAppAvailable  #
#                and WaitForInputReady API'                                 #
#  Parameter(s): None                                                       #
#                                                                           #
#  Return value: The EHLLAPI counterpart of the input                       #
#                                                                           #
# ------------------------------------------------------------------------- #
def ehllapiFunctionWaitNoParams():
    HllFunctionNumber = 4
    HllData = 0
    HllLength = 0
    HllReturnCode = 0
    returnString = "'The Wait function without parameters loops through the while loop until the return value is 0, 4 or 5, and exits with an error if it is 1 or 9" + '\n'
    returnString = returnString + "HllFunctionNumber = " + str(HllFunctionNumber) + '\n'
    returnString = returnString + "HllData = " + str(HllData) + '\n'
    returnString = returnString + "HllLength = " + str(HllLength) + '\n'
    returnString = returnString + "HllReturnCode = " + str(HllReturnCode) + '\n'
    returnString = returnString + "retVal = PCOMM_Wait(" + str(HllFunctionNumber) + ", " + str(HllData) + ", " + str(HllLength) + ", 0)" + '\n'
    returnString = returnString + "breakFlag = False" + '\n'
    returnString = returnString + "While breakFlag <> True And retVal <> 0" + '\n'
    returnString = returnString + '\t' + "If retVal = 1 Then" + '\n'
    returnString = returnString + '\t\t' + 'MsgBox("Your application program is not connected to a valid session.")' + '\n'
    returnString = returnString + '\t\t' + "breakFlag = True" + '\n'
    returnString = returnString + '\t' + "ElseIf retVal = 9 Then" + '\n'
    returnString = returnString + '\t\t' + 'MsgBox("A system error was encountered.")' + '\n'
    returnString = returnString + '\t\t' + "breakFlag = True" + '\n'
    returnString = returnString + '\t' + "End If" + '\n'
    returnString = returnString + "Wend" + '\n'
    return returnString
