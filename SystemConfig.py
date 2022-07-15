import datetime

col_Automation_Reference=-1
col_TestCaseNo=-1
col_TestCaseName=-1
col_Status_Code=-1
col_Method=-1
col_API_to_trigger=-1
col_Parameters=-1
col_ResponseParametersToCapture=-1
col_HeadersToValidate=-1
col_isJsonAbsolutePath=-1
col_preCommands=-1
col_postCommands=-1

field_Automation_Reference="Automation_Reference"
field_TestCaseNo="TestCaseNo"
field_TestCaseName="TestCaseName"
field_Status_Code="Status_Code"
field_Method="Method"
field_API_to_trigger="API_to_trigger"
field_Parameters="Parameters"
field_ResponseParametersToCapture="ResponseParametersToCapture"
field_GlobalParametersToStore="GlobalParametersToStore"
field_ClearGlobalParameters="ClearGlobalParameters"
field_HeadersToValidate="HeaderValidation"
field_isJsonAbsolutePath="isJSONAbsolutePath"
field_preCommands="Pre_Commands"
field_postCommands="Post_Commands"

col_ApiName=-1
col_EndPoint=-1
col_API_Structure=-1
col_Headers=-1
col_GlobalParametersToStore=-1
col_ClearGlobalParameters=-1
col_Assignments=-1
col_Controller=-1
col_Authentication=-1

field_apiName="API_Name"
field_EndPoint="EndPoint"
field_API_Structure="API_Structure"
field_Headers="Headers"
field_Assignments="Assignments"
field_Authentication="Authentication"

currentTestCaseNumber=None
currentAPI=None
currentisJsonAbsolutePath="N"

globalDict={}
localRequestDict={}
authenticationDict={}

successfulResponseCode=["200","201"]

logsFolder="logs"
logNumber=0

lastColumnInSheetTCs="ClearGlobalParameters"
lastColumnInSheetStructures="API_Structure"

splitterPrefix="#{"
splitterPostfix="}#"

responseField=None

fixedExclusions = "{}\"',[]"
floatLimit='.11f'
