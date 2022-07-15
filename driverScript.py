#pending development :

#Timestamp keyword check
#JS Fetch

import os, sys
sys.path.append(".")
sys.path.append("customLib")
import openpyxl
import logging, traceback
import SystemConfig,userConfig
import time
import customLib.Report as Report
import customLib.Config as Config
import ApiLib
import dynamicConfig
import testDataHelper
import ExcelHelper as eh
import re
import datetime
import json
import ast

def setColumnNumbersForFileValidations():
    eh.read_sheet("Structures", SystemConfig.lastColumnInSheetStructures)

    SystemConfig.col_ApiName = eh.get_column_number_of_string(SystemConfig.field_apiName)
    SystemConfig.col_API_Structure = eh.get_column_number_of_string(SystemConfig.field_API_Structure)
    SystemConfig.col_EndPoint = eh.get_column_number_of_string(SystemConfig.field_EndPoint)
    SystemConfig.col_Method = eh.get_column_number_of_string(SystemConfig.field_Method)
    SystemConfig.col_Headers = eh.get_column_number_of_string(SystemConfig.field_Headers)
    SystemConfig.col_Authentication = eh.get_column_number_of_string(SystemConfig.field_Authentication)

    ######################################################################################
    ######################################################################################

    eh.read_sheet("TCs", SystemConfig.lastColumnInSheetTCs)
    SystemConfig.col_API_to_trigger = eh.get_column_number_of_string(SystemConfig.field_API_to_trigger)
    SystemConfig.col_Automation_Reference = eh.get_column_number_of_string(SystemConfig.field_Automation_Reference)
    SystemConfig.col_Status_Code = eh.get_column_number_of_string(SystemConfig.field_Status_Code)
    SystemConfig.col_HeadersToValidate = eh.get_column_number_of_string(SystemConfig.field_HeadersToValidate)
    SystemConfig.col_Assignments = eh.get_column_number_of_string(SystemConfig.field_Assignments)
    SystemConfig.col_TestCaseNo = eh.get_column_number_of_string(SystemConfig.field_TestCaseNo)
    SystemConfig.col_TestCaseName = eh.get_column_number_of_string(SystemConfig.field_TestCaseName)
    SystemConfig.col_ResponseParametersToCapture = eh.get_column_number_of_string(SystemConfig.field_ResponseParametersToCapture)
    SystemConfig.col_Parameters = eh.get_column_number_of_string(SystemConfig.field_Parameters)
    SystemConfig.col_GlobalParametersToStore = eh.get_column_number_of_string(SystemConfig.field_GlobalParametersToStore)
    SystemConfig.col_ClearGlobalParameters = eh.get_column_number_of_string(SystemConfig.field_ClearGlobalParameters)
    SystemConfig.col_Assignments = eh.get_column_number_of_string(SystemConfig.field_Assignments)
    SystemConfig.col_isJsonAbsolutePath = eh.get_column_number_of_string(SystemConfig.field_isJsonAbsolutePath)
    SystemConfig.col_preCommands = eh.get_column_number_of_string(SystemConfig.field_preCommands)
    SystemConfig.col_postCommands = eh.get_column_number_of_string(SystemConfig.field_postCommands)

def parseHeader(requestParameters):
    if requestParameters is None:
        return

    dictHeader = {}
    allParams  = []

    requestParameters = replacePlaceHolders(requestParameters)
    requestParameters = requestParameters.strip()

    if "\n" in requestParameters.strip():
        allParams = requestParameters.split("\n")
    else:
        allParams.append(requestParameters)

    for eachParamValuePair in allParams:
        [paramName, paramValue] = eachParamValuePair.split(":", 1)
        dictHeader[paramName]   = paramValue

    return dictHeader

def parametrizeRequest(requestStructure, requestParameters):
    requestStructure = replacePlaceHolders(requestStructure)
    requestStructure = requestStructure.encode('ascii', 'ignore')

    if requestParameters is None:
        return requestStructure

    allParams=[]
    requestParameters = replacePlaceHolders(requestParameters)
    if "\n" in requestParameters.strip():
        allParams=requestParameters.split("\n")
    else:
        allParams.append(requestParameters)

    for eachParamValuePair in allParams:
        [paramName, paramValue] = eachParamValuePair.split(":", 1)
        SystemConfig.localRequestDict[paramName]=paramValue

        if "Y" == SystemConfig.currentisJsonAbsolutePath:
            data = ast.literal_eval(requestStructure)
            tempString = "data" + paramName
            if paramValue.startswith("ADD("):
                paramValue = paramValue.replace("ADD(", "").replace(")", "")
                exec(tempString + ".append(" + paramValue + ")")
            else:
                exec(tempString + " = " + paramValue)
            requestStructure = str(data).replace("'", '"')
        else:
            if requestStructure is not None and "<"+paramName.strip()+">" in requestStructure:
                #handle xml replacement
                regexString="<"+paramName+">"+".*"+r'</'+paramName+">"
                newString="<"+paramName+">"+paramValue+r'</'+paramName+">"

            else:
                regexString='"'+paramName+'" *:.*"(.*)"'
                result = re.search(regexString, requestStructure)

                if result is not None:
                    stringToReplace=result.group(1)
                    oldString='"'+paramName+'"'+":"+'"'+stringToReplace+'"'
                    if stringToReplace=='':
                            newString=oldString.replace(oldString,'"'+paramName+'":'+'"'+paramValue+'"',1)
                    else:
                        newString=oldString.replace(stringToReplace,paramValue,1)
                else: #result is None
                    regexString='"'+paramName+'" *:(.*)'
                    result = re.search(regexString, requestStructure)

                    if result is not None:
                        stringToReplace=result.group(1)
                        oldString='"'+paramName+'"'+":"+stringToReplace
                        if stringToReplace=='':
                            newString=oldString.replace(oldString,'"'+paramName+'":'+paramValue,1)
                        else:
                            newString=oldString.replace(stringToReplace,paramValue,1)
                    else: #if result is None
                        print "No matching substitution for param : {0}".format(paramName)
                        Report.WriteTestStep("Excel Error: No matching substitution for param : {0}".format(paramName),"NA","NA","Failed")
                        return

            print "regexString:",regexString
            print "Old structure : ",requestStructure

            if regexString.endswith(","):
                newString=newString+","

            print "newString:",newString


            requestStructure=re.sub(regexString,newString,requestStructure)

            print "New Structure:",requestStructure
    return requestStructure

def parseValue(fieldToFind, responseChunk):

    #if response chunk is a dict, {"result":{"accessToken":"eyJraWQiOiJabkc5"}}
    #takes in a dictionary and tries to match the keys to the desired key
    valueParsed=None
    valueFound=False

    if type(responseChunk) is dict:
        for key in responseChunk.keys():
            if str(key).lower()==str(fieldToFind).lower():
                valueParsed=responseChunk[key]
                valueFound=True
                return (valueFound,valueParsed)
            if type(responseChunk[key]) is dict:
                return parseValue(fieldToFind,responseChunk[key])
            elif type(responseChunk[key]) is list:
                for eachValue in responseChunk[key]:
                    return parseValue(fieldToFind,eachValue)

    return (valueParsed,valueFound)

def parse_json_recursively(json_object, target_key):
    #global retval
    if type(json_object) is dict and json_object:
        for key in json_object:
            if key.lower() == str(target_key.lower()):
                if type(json_object[key]) is float:
                    SystemConfig.responseField=str(format(json_object[key], SystemConfig.floatLimit))
                else:
                    SystemConfig.responseField=str(json_object[key])
                print("{}: {}".format(target_key, json_object[key]))
                return;
            parse_json_recursively(json_object[key], target_key)

    elif type(json_object) is list and json_object:
        for item in json_object:
            parse_json_recursively(item, target_key)

def isObjectFound(json_object, targetKey):
    jsonPath = "json_object" + targetKey
    print("jsonPath : " + jsonPath)
    try:
        if type(eval(jsonPath)) is float:
            SystemConfig.responseField=str(format(eval(jsonPath), SystemConfig.floatLimit))
        else:
            SystemConfig.responseField=str(format(eval(jsonPath)))
        return True
    except Exception as e:
        print "Failure. Param : {0} not found in the response".format(targetKey)
        return False

def extractParamValueFromResponse(param):
    #returns specific param value from response

    #always set response to None initially
    SystemConfig.responseField=None

    if dynamicConfig.responseText is not None:
        try:
            if "xml" in str(dynamicConfig.currentResponse.headers['Content-Type']):
                #soap xml parsing

                #print "Handling xml parsing"
                preString="<"+param+">"
                postString=r"</"+param+">"

                try:
                    afterPreSplit=dynamicConfig.responseText.split(preString)[1]
                    #print "\nafterPreSplit: ",afterPreSplit
                    paramValue=afterPreSplit.split(postString)[0]
                    #print "\nafterSecondSplit: ",paramValue
                    return paramValue
                except:
                    print "Failure. Param : {0} not found in the response".format(param)

            else:
                data = dynamicConfig.currentResponse.json()
                strData=str(dynamicConfig.responseText)

                if param.startswith("[") and param.endswith("]"):
                    if isObjectFound(data, param):
                        return SystemConfig.responseField
                else:
                    if param in strData:
                        #(paramFoundStatus,paramResponseValue)=parseValue(param,data)
                        parse_json_recursively(data, param)
                        return SystemConfig.responseField
                    else:
                        print "Failure. Param : {0} not found in the response".format(param)


        except Exception,e:
            traceback.print_exc()
            print "Failure. Param : {0} not found in the response".format(param)

    return None

def extractParamValueFromHeaders(param):
    #returns specific param value from response

    #print "Extracting value : {0} from response".format(param)
    #print "ResponseText is : ",dynamicConfig.responseText

    if dynamicConfig.responseHeaders is not None:
        try:
            if "xml" in str(dynamicConfig.currentResponse.headers['Content-Type']):
                #soap xml parsing

                #print "Handling xml parsing"
                preString="<"+param+">"
                postString=r"</"+param+">"

                try:
                    afterPreSplit=dynamicConfig.responseText.split(preString)[1]
                    #print "\nafterPreSplit: ",afterPreSplit
                    paramValue=afterPreSplit.split(postString)[0]
                    #print "\nafterSecondSplit: ",paramValue
                    return paramValue
                except:
                    print "Failure. Param : {0} not found in the response".format(param)

            else:
                #json parsing

                data = dict(dynamicConfig.responseHeaders)
                print("Type of data is : {0}".format(type(data)))


                strData=str(data)

                if param in strData:
                    (paramFoundStatus,paramResponseValue)=parseValue(param,data)
                    if paramFoundStatus is True:
                        return paramResponseValue
                    else:
                        return None
                else:
                    print "Failure. Param : {0} not found in the response".format(param)

        except Exception,e:
            traceback.print_exc()
            print "Failure. Param : {0} not found in the response".format(param)

    return None

def storeGlobalParameters(globalParams):
    #parse response and find the global parameter
    val=None
    if globalParams is None:
        return

    try:
        globalParams=globalParams.strip()
    except:
        pass

    allGlobalParams = []
    if "\n" in globalParams:
        allGlobalParams=globalParams.split("\n")
    else:
        allGlobalParams.append(globalParams)

    for eachParam in allGlobalParams:
        val=None
        if ":" in eachParam:
            [key, val]=eachParam.split(":", 1)
            eachParam = key
            if val.startswith("["):
                val = extractParamValueFromResponse(val)

        elif eachParam.startswith("HEADER_"):
            eachParam = eachParam.replace("HEADER_", "")
            val = extractParamValueFromHeaders(eachParam)
        else:
            if eachParam in SystemConfig.localRequestDict:
                val = SystemConfig.localRequestDict[eachParam]
            else:
                val = extractParamValueFromResponse(eachParam)

        if val is not None:
            SystemConfig.globalDict[eachParam]=val

def parseAndValidateResponse(userParams):
    if userParams is None:
        return

    userParams    = replacePlaceHolders(userParams.strip())
    allUserParams = []
    if "\n" in userParams:
        allUserParams=userParams.split("\n")
    else:
        allUserParams.append(userParams)

    for eachUserParam in allUserParams:
        shouldContain = False

        if eachUserParam.startswith("textMatch_"):
            val=eachUserParam
            val=val.replace("textMatch_","")
            if val.lower() in str(dynamicConfig.responseText).lower():
                Report.WriteTestStep("Check text match in Response body: {0}".format(val),"Expected Text : {0} should appear in response body".format(val),"Expected text appeared","Pass")
            else:
                Report.WriteTestStep("Check text match in Response body: {0}".format(val),"Expected Text : {0} should appear in response body".format(val),"Expected text did not appear in response body","Fail")

        elif ":" in eachUserParam:
            [key, val] = eachUserParam.split(":", 1)
            expectedValue=str(val).strip()

            if expectedValue.startswith("contains("):
                expectedValue = expectedValue.replace("contains(", "").replace(")", "")
                shouldContain = True

            paramValue = extractParamValueFromResponse(key)

            val=expectedValue
            #val is the expectedValue
            #paramValue is the actualValue
            if paramValue is not None:
                if shouldContain:
                    if expectedValue.lower() in val.lower():
                        print "Success : param : {0} found in response and Value : {1} is  contained in value : {2}".format(key,paramValue,val)
                        Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Actual value : {0}".format(paramValue),"Pass")
                    else:
                        print "Failure : param : {0} found in response BUT Value : {1} is NOT contained in value : {2}".format(key,paramValue,val)
                        Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Actual value : {0}".format(paramValue),"Fail")
                else:
                    if str(val.lower())==str(paramValue.lower()):
                        print "Success : param : {0} found in response and Value : {1} is same as expected : {2}".format(key,paramValue,val)
                        Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Actual value : {0}".format(paramValue),"Pass")
                    else:
                        print "Failure : param : {0} found in response BUT Value : {1} is NOT same as expected : {2}".format(key,paramValue,val)
                        Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Actual value : {0}".format(paramValue),"Fail")
            else:
                print "Falure : param : {0} not found in response".format(key)
                Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Parameter was not found in the response structure","Fail")


        else:
            #just make sure fields are available
            key=eachUserParam
            paramValue=extractParamValueFromResponse(eachUserParam)
            if paramValue is not None:
                print "Success : param : {0} found in response. Value : {1}".format(eachUserParam,paramValue)
                Report.WriteTestStep("Capture Response Parameter : [{0}]".format(key),"Parameter should be present in the Response","Parameter : [{0}] having value : [{1}] was found in the Response".format(key,paramValue),"Pass")

            else:
                print "Failure : param : {0} not found in response".format(eachUserParam)
                Report.WriteTestStep("Capture Response Parameter : [{0}]".format(key),"Parameter should be present in the Response","Parameter : [{0}] was not found in Response".format(key),"Fail")

def parseAndValidateHeaders(userParams):
    if userParams is None:
        return

    userParams    = replacePlaceHolders(userParams.strip())
    allUserParams = []
    if "\n" in userParams:
        allUserParams=userParams.split("\n")
    else:
        allUserParams.append(userParams)

    for eachUserParam in allUserParams:
        shouldContain = False

        if eachUserParam.startswith("textMatch_"):
            eachUserParam=eachUserParam.replace("textMatch_","")
            val=eachUserParam
            if val.lower() in str(dynamicConfig.responseHeaders).lower():
                Report.WriteTestStep("Check text match in Response Headers: {0}".format(val),"Expected Text : {0} should appear in Response Headers".format(val),"Expected text appeared","Pass")
            else:
                Report.WriteTestStep("Check text match in Response Headers: {0}".format(val),"Expected Text : {0} should appear in Response Headers".format(val),"Expected text did not appear in Response Headers","Fail")
        elif ":" in eachUserParam:
            [key, val] = eachUserParam.split(":", 1)
            expectedValue=str(val).strip()

            if expectedValue.startswith("contains("):
                expectedValue = expectedValue.replace("contains(", "").replace(")", "")
                shouldContain = True

            paramValue = extractParamValueFromHeaders(key)
            val        = expectedValue

            if paramValue is not None:
                if shouldContain:
                    if expectedValue.lower() in val.lower():
                        print "Success : param : {0} found in response and Value : {1} is  contained in value : {2}".format(key,paramValue,val)
                        Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Actual value : {0}".format(paramValue),"Pass")
                    else:
                        print "Failure : param : {0} found in response BUT Value : {1} is NOT contained in value : {2}".format(key,paramValue,val)
                        Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Actual value : {0}".format(paramValue),"Fail")
                else:
                    if str(val.lower())==str(paramValue.lower()):
                        print "Success : param : {0} found in response and Value : {1} is same as expected : {2}".format(key,paramValue,val)
                        Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Actual value : {0}".format(paramValue),"Pass")
                    else:
                        print "Failure : param : {0} found in response BUT Value : {1} is NOT same as expected : {2}".format(key,paramValue,val)
                        Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Actual value : {0}".format(paramValue),"Fail")
            else:
                print "Falure : param : {0} not found in response".format(key)
                Report.WriteTestStep("Response Parameter Validation : [{0}]".format(key),"Expected value : {0}".format(val),"Parameter was not found in the response structure","Fail")
        else:
            paramValue = extractParamValueFromHeaders(eachUserParam)
            if paramValue is not None:
                print "Success : param : {0} found in response. Value : {1}".format(eachUserParam,paramValue)
                Report.WriteTestStep("Capture Response Parameter : [{0}]".format(eachUserParam),"Parameter should be present in the Response","Parameter : [{0}] having value : [{1}] was found in the Response".format(eachUserParam,paramValue),"Pass")
            else:
                print "Failure : param : {0} not found in response".format(eachUserParam)
                Report.WriteTestStep("Capture Response Parameter : [{0}]".format(eachUserParam),"Parameter should be present in the Response","Parameter : [{0}] was not found in Response".format(key),"Fail")

def executeCommand(vars):
    if vars is None:
        return

    vars=vars.strip()
    allVars=[]
    if "\n" in vars:
        allVars=vars.split("\n")
    else:
        allVars.append(vars)

    for val in allVars:
        if val.lower().startswith("sleep"):
            try:
                val=int(str(val.replace("sleep(","").replace(")","")).strip())
                print("[User-Command] Will sleep for {0} seconds".format(val))
                time.sleep(val)
                Report.WriteTestStep("Wait before proceeding to next step for {0} seconds".format(val), "Should wait","Waited for: {0} seconds".format(val),"Passed")
            except:
                print("[ERROR] Invalid argument for Sleep command : {0}".format(val))
                Report.WriteTestStep("Wait before proceeding to next step for {0} seconds".format(val), "The argument passed should be an Integer","The argument passed is Invalid : {0}".format(val),"Passed")
        elif val.lower().startswith("terminateonfailure"):
            try:
                if dynamicConfig.testCaseHasFailed:
                    print("[User-Command] Terminate on failure")
                    val=str(val.replace("terminateonfailure(","").replace(")","")).strip()
                    if val.lower()=="true":
                        Report.WriteTestStep("Terminating flow since failure is encountered","NA","NA","Failed")
                        Report.evaluateIfTestCaseIsPassOrFail()
                        endProcessing()
            except:
                print("[ERROR] Invalid argument for TerminateOnFailure command : {0}".format(val))
                Report.WriteTestStep("Invalid argument for TerminateOnFailure command : {0}".format(val), "The argument passed should be either yes or no","Invalid argument","Failed")
        elif val.lower().startswith("validatefor1gie"):
            print("[User-Command] validateFor1Gie")
            pathForValidation=val.replace("validateFor1Gie_","")
        elif val.lower().startswith("skiponfailure"):
            try:
                if dynamicConfig.testCaseHasFailed:
                    print("[User-Command] SkipOnFailure")
                    val=int(str(val.replace("skiponfailure(","").replace(")","")).strip())
                    testCaseNumberToSkipTo=val

                    if(testCaseNumberToSkipTo<=SystemConfig.endRow):
                        rowNumberWhichFailed=SystemConfig.currentRow
                        SystemConfig.currentRow=testCaseNumberToSkipTo
                        Report.WriteTestStep("Skip to TC #{0} since failure is encountered".format(testCaseNumberToSkipTo),"NA","NA".format(SystemConfig.endRow),"Failed")
                        testCaseNumberToSkipTo=val
                        Report.evaluateIfTestCaseIsPassOrFail()
                    else:
                        Report.WriteTestStep("TC#:{0} to skip to does not exist.".format(testCaseNumberToSkipTo),"The TC# to skip to should be within the range of total # of TCs","Max TCs {0}: ".format(SystemConfig.endRow),"Failed")
            except:
                traceback.print_exc()
                print("[ERROR] Invalid argument for SkipOnFailure command : {0}".format(val))
                Report.WriteTestStep("Invalid argument for SkipOnFailure command : {0}".format(val), "The argument passed should be an Integer","Invalid argument","Failed")

        elif val.lower().startswith("skipalways"):
            try:
                print("[User-Command] Skip-Always")
                val=int(str(val.replace("skipalways(","").replace(")","")).strip())
                testCaseNumberToSkipTo=val

                if(testCaseNumberToSkipTo<=SystemConfig.endRow):
                    rowNumberWhichFailed=SystemConfig.currentRow
                    SystemConfig.currentRow=testCaseNumberToSkipTo
                    testCaseNumberToSkipTo=val
                    Report.evaluateIfTestCaseIsPassOrFail()
                else:
                    Report.WriteTestStep("TC#:{0} to skip to does not exist.".format(testCaseNumberToSkipTo),"The TC# to skip to should be within the range of total # of TCs","Max TCs {0}: ".format(SystemConfig.endRow),"Failed")
            except:
                traceback.print_exc()
                print("[ERROR] Invalid argument for SkipOnFailure command : {0}".format(val))
                Report.WriteTestStep("Invalid argument for SkipOnFailure command : {0}".format(val), "The argument passed should be an Integer","Invalid argument","Failed")

def replacePlaceHolders(var):
    if "#{" not in var and "}#" not in var:
        return var

    for key in SystemConfig.globalDict.keys():
        stringToMatch = "#{" + key + "}#"
        if stringToMatch in var:
            var = var.replace(stringToMatch,str(SystemConfig.globalDict[key]))

    for key in SystemConfig.localRequestDict.keys():
        stringToMatch = "#{" + key + "}#"
        if stringToMatch in var:
            var = var.replace(stringToMatch,str(SystemConfig.localRequestDict[key]))

    if "#{" in var and "}#" in var:
            print "Failure: Undefined variable usage in " + var
            Report.WriteTestStep("User-Input error","Undefined variable used","Only variables which are defined can be used","Fail")
            endProcessing()

    return var

def storeUserDefinedVariables(vars):
    if vars is None:
        return

    allVars=[]
    vars=vars.strip()

    if "\n" in vars:
        allVars = vars.split("\n")
    else:
        allVars.append(vars)


    for eachVar in allVars:
        [key,val]=eachVar.split(":", 1)

        val = replacePlaceHolders(val)

        #random("TR",12,"")
        if "(" in val:
            val = val.partition("(")[2][:-1]

        if val.lower().startswith("random("):
            val = val.split(",")
            prefix = numberOfChars = suffix = pool = exclusions = ""
            if (5 == len(val)):
                pool = val [4]
            if (3 < len(val)):
                exclusions = val[3]
            prefix, numberOfChars, suffix =  val[0:3]
            val = testDataHelper.random_value(prefix, numberOfChars,
                                             suffix, exclusions, pool)

        elif val.lower().startswith("randomint("):
            [minValue, maxValue] = val.split(",")
            val = testDataHelper.random_int(minValue, maxValue)

        elif val.lower().startswith("split("):
            [baseString, delimiter, index] = val.split(",")
            val = testDataHelper.split(baseString, delimiter, index)

        elif val.lower().startswith("timestamp"):
            val = testDataHelper.generate_timestamp(val)

        elif val.lower().startswith("theiaDoubleEncode"):
            #timestamp(DDMMYYYY)
            val = testDataHelper.theia_double_encode()

        elif val.lower().startswith("encode("):
            [baseString, encodeType] = val.split(",")
            val = testDataHelper.encode_string(baseString, encodeType)

        elif val.lower().startswith("createcookie"):
            # val = val + sys.argv[1]
            # val = val.partition("(")[2]
            val = testDataHelper.create_cookie(val)

        # elif val.lower().startswith("getcookie"):
        #     val = sys.argv[1]

        print "{", key, "=", val, "}"
        SystemConfig.localRequestDict[key]=val

def setAuthentication(authentication):
    if authentication is None:
        dynamicConfig.currentAuthentication = None
        return

    authentication = authentication.encode('ascii', 'ignore')
    allVars=[]

    authentication = replacePlaceHolders(authentication)
    if "\n" in authentication:
        allVars=authentication.split("\n")
    else:
        allVars.append(authentication)

    for eachVar in allVars:
        [key,val] = eachVar.split(":", 1)
        key       = key.lower()
        SystemConfig.authenticationDict[key] = val

    if "BASIC" == SystemConfig.authenticationDict["type"].upper():
        dynamicConfig.currentAuthentication = (SystemConfig.authenticationDict["username"],
                                               SystemConfig.authenticationDict["password"])

def setCookies():
    if 'COOKIE' in SystemConfig.localRequestDict:
        dynamicConfig.currentCookie = SystemConfig.localRequestDict['COOKIE']
    elif 'COOKIE' in SystemConfig.globalDict:
        dynamicConfig.currentCookie = SystemConfig.globalDict['COOKIE']
    else:
        dynamicConfig.currentCookie = None

def setFiles():
    if 'filetoupload' in SystemConfig.localRequestDict:
        file = SystemConfig.localRequestDict['filetoupload']
        file = {"uploaded_file": (file, open(file, "rb"),'text/csv')}
        dynamicConfig.currentfile = file
    else:
        dynamicConfig.currentfile = None

def getEndRow(currentRow):
    while SystemConfig.maxRows >= currentRow:
        testCaseNumber = eh.get_cell_value(currentRow, SystemConfig.col_TestCaseNo)
        if "(END)" in str(testCaseNumber).upper():
            return currentRow
        currentRow += 1

def main():
    startingPointisFound=False
    eh.read_sheet("TCs", SystemConfig.lastColumnInSheetTCs)
    setColumnNumbersForFileValidations()
    currentRow = eh.get_row_number_of_string("ResponseParametersToCapture")
    currentRow += 1
    SystemConfig.startingRowNumberForRecordProcessing=currentRow
    endRow = getEndRow(currentRow)
    #print "Max Rows : {0}\n".format(SystemConfig.maxRows)

    while currentRow <= endRow:
        #print("Row #: {0}\n".format(currentRow))
        dynamicConfig.responseHeaders    = None
        dynamicConfig.responseStatusCode = None
        dynamicConfig.responseText       = None
        dynamicConfig.restRequestType    = None
        dynamicConfig.currentRequest     = None
        dynamicConfig.currentResponse    = None
        dynamicConfig.currentUrl         = None
        dynamicConfig.currentException   = None
        dynamicConfig.currentHeader      = None
        dynamicConfig.currentContentType = None
        dynamicConfig.testCaseHasFailed  = False

        eh.read_sheet("TCs",SystemConfig.lastColumnInSheetTCs)
        testCaseNumber = str(eh.get_cell_value(currentRow, SystemConfig.col_TestCaseNo))

        testCaseName=eh.get_cell_value(currentRow,SystemConfig.col_TestCaseName)

        if testCaseName is None or str(testCaseName).strip()=="":
            break

        if not startingPointisFound:
            if "(START)" not in str(testCaseNumber).upper():
                currentRow+=1
                continue
            else:
                startingPointisFound=True
                SystemConfig.startTime = datetime.datetime.now()
        statusCode                  = eh.get_cell_value(currentRow, SystemConfig.col_Status_Code)
        headerFieldsToValidate      = eh.get_cell_value(currentRow, SystemConfig.col_HeadersToValidate)
        responseParametersToCapture = eh.get_cell_value(currentRow, SystemConfig.col_ResponseParametersToCapture)
        headerParametersToCapture   = eh.get_cell_value(currentRow, SystemConfig.col_HeadersToValidate)
        requestParameters           = eh.get_cell_value(currentRow, SystemConfig.col_Parameters)
        apiToTrigger                = eh.get_cell_value(currentRow, SystemConfig.col_API_to_trigger)
        globalParams                = eh.get_cell_value(currentRow, SystemConfig.col_GlobalParametersToStore)
        clearGlobalParams           = eh.get_cell_value(currentRow, SystemConfig.col_ClearGlobalParameters)
        userDefinedVars             = eh.get_cell_value(currentRow, SystemConfig.col_Assignments)
        isJsonAbsolutePath          = eh.get_cell_value(currentRow, SystemConfig.col_isJsonAbsolutePath)
        preCommands                 = eh.get_cell_value(currentRow, SystemConfig.col_preCommands)
        postCommands                = eh.get_cell_value(currentRow, SystemConfig.col_postCommands)

        eh.read_sheet("Structures", SystemConfig.lastColumnInSheetStructures)
        matchedRow = eh.get_row_number_of_string(apiToTrigger)
        endPoint             = eh.get_cell_value(matchedRow, SystemConfig.col_EndPoint)
        requestStructure     = eh.get_cell_value(matchedRow, SystemConfig.col_API_Structure)
        rawHeaderText        = eh.get_cell_value(matchedRow, SystemConfig.col_Headers)
        typeOfRequest        = eh.get_cell_value(matchedRow, SystemConfig.col_Method)
        authenticationParams = eh.get_cell_value(matchedRow, SystemConfig.col_Authentication)

        executeCommand(preCommands)

        if typeOfRequest is not None:
            if "<soap" in requestStructure:
                typeOfRequest+="(soap)" #POST(soap)
            else:
                typeOfRequest+="(rest)" #POST(rest)
            print "type of request is : ",typeOfRequest

        testCaseNumber = testCaseNumber.upper().replace("(START)", "")
        testCaseNumber = testCaseNumber.upper().replace("(END)", "")
        Report.WriteTestCase("TC_{0}".format(testCaseNumber), testCaseName)

        SystemConfig.currentTestCaseNumber=testCaseNumber
        SystemConfig.currentAPI=apiToTrigger

        if isJsonAbsolutePath is not None:
            SystemConfig.currentisJsonAbsolutePath = isJsonAbsolutePath.upper()

        storeUserDefinedVariables(userDefinedVars)
        setAuthentication(authenticationParams)
        setCookies()
        setFiles()

        endPoint = replacePlaceHolders(endPoint)
        headers  = parseHeader(rawHeaderText)

        if headers is not None:
            dynamicConfig.currentHeader=headers
            if "Content-Type" in headers.keys():
                dynamicConfig.currentContentType = headers["Content-Type"].encode('ascii', 'ignore')
        else:
            if "rest" in typeOfRequest.lower():
                dynamicConfig.currentHeader={}
            else:
                dynamicConfig.currentHeader=""

        requestStructure = parametrizeRequest(requestStructure, requestParameters)

        if requestStructure is not None:
            dynamicConfig.currentRequest=requestStructure
        else:
            if "rest" in typeOfRequest.lower():
                dynamicConfig.currentRequest={}
            else:
                dynamicConfig.currentRequest=""

        dynamicConfig.currentUrl=endPoint
        dynamicConfig.restRequestType=typeOfRequest.strip().lower()

        print "\nTC# : {0}".format(testCaseNumber)
        if requestStructure is not None and str(requestStructure).startswith("<soap"):
            print "\n[ Executing SOAP Request ]"
            print "\nWebservice : {0}".format(apiToTrigger)
            print "\nEndPoint : {0}".format(endPoint)
            print "\nHeader : {0}".format(headers)
            print "\nRequest : {0}".format(requestStructure)

            Report.WriteTestStep("SOAP Request details","Log request details","EndPoint : {0}\nHeader: {1}\nBody : {2}".format(endPoint,headers,requestStructure),"Pass")

            ApiLib.triggerSoapRequest()
        else:
            print "\n[ Executing Rest Request ]"
            print "\nAPI : {0}".format(apiToTrigger)
            print "\nEndPoint : {0}".format(endPoint)
            print "\nHeader : {0}".format(headers)
            print "\nRequest : {0}".format(requestStructure)

            Report.WriteTestStep("Rest Request details","Log request details","Request Type : {0}\n\nEndPoint : {1}\n\nHeader: {2}\n\nBody : {3}".format(typeOfRequest,endPoint,headers,requestStructure),"Pass")
            ApiLib.triggerRestRequest()

        if dynamicConfig.currentResponse is None:
            Report.WriteTestStep("Log Response","Log Response","No Response received from server within user-configured timeout : {0} seconds".format(userConfig.timeoutInSeconds),"Fail")

        else:
            if "application/pdf" not in str(dynamicConfig.responseHeaders):
                Report.WriteTestStep("Log Response","Log Response","Status Code : {0}\n\nHeaders: {1}\n\nBody: {2}".format(dynamicConfig.responseStatusCode,dynamicConfig.responseHeaders,dynamicConfig.responseText),"Pass")
            else:
                Report.WriteTestStep("Log Response","Log Response","Status Code : {0}\n\nHeaders: {1}".format(dynamicConfig.responseStatusCode,dynamicConfig.responseHeaders),"Pass")

        if statusCode is not None:
            statusCode = str(statusCode)
            dynamicConfig.responseStatusCode = statusCode
            if dynamicConfig.responseStatusCode in statusCode:
                Report.WriteTestStep("Validate Response Code","Expected Response Code(s) : {0}".format(statusCode),"Actual Response Code : {0}".format(dynamicConfig.responseStatusCode),"Pass")
                print "[INFO] Valid Status Code: " + dynamicConfig.responseStatusCode + " is received"
            else:
                Report.WriteTestStep("Validate Response Code","Expected Response Code(s) : {0}".format(statusCode),"Actual Response Code : {0}".format(dynamicConfig.responseStatusCode),"Fail")
                print "[ERR] " + dynamicConfig.responseStatusCode + " not in Expected Status Codes : " + statusCode
        else:
            Report.WriteTestStep("Skipping Response Validation since no Response Code is specified in Datasheet","NA","NA","Pass")

        #Globals will be executed Post Request
        storeGlobalParameters(globalParams)

        #Compare Response with expected
        parseAndValidateResponse(responseParametersToCapture)
        parseAndValidateHeaders(headerParametersToCapture)

        executeCommand(postCommands)
        if str(clearGlobalParams).upper().startswith("Y"):
            SystemConfig.globalDict={}
        SystemConfig.localRequestDict={}

        currentRow+=1

        time.sleep(1)
        Report.evaluateIfTestCaseIsPassOrFail()

def initiateLogging(resultFolder):
    Config.logsFolder=resultFolder+"\\logs"
    print "Logs will be created at : {0}".format(Config.logsFolder)

    if not os.path.exists(Config.logsFolder):
        os.mkdir(Config.logsFolder)

def endProcessing():
    Report.GeneratePDFReport()
    sys.exit(-1)

def superMain():
    resultFolder=Report.InitializeReporting()
    initiateLogging(resultFolder)
    try:
        main()
    except Exception,e:
        traceback.print_exc()
    finally:
        executionTime = datetime.datetime.now() - SystemConfig.startTime
        os.environ["EXECTIME"] = str(executionTime)

    try:
        Report.GeneratePDFReport()
    except:
        print "[ERR] Failure in Generating Pdf."


if __name__ == '__main__':

    resultFolder=Report.InitializeReporting()
    initiateLogging(resultFolder)
    try:
        main()
    except Exception,e:
        traceback.print_exc()

    Report.GeneratePDFReport()