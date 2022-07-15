rootFolder="Results"
outputExcelPath=None
sheetObj=None
workBookObj=None

(TestScenarioSrNo,TestStepNo,TestCaseStepDesc,TestObjectName,TestData,
ExpectedResult,ActualResult,Status,ScreenshotPath,StartTime,EndTime)=("Test Scenario Sr No","Test Step No","Test Case/Step Desc","Test Object Name","Test Data",
"Expected Result","Actual Result","Test Case/Step Status","Screenshot","StartTime","EndTime")

ExcelHeaders=[TestScenarioSrNo,TestStepNo,TestCaseStepDesc,TestObjectName,TestData,ExpectedResult,ActualResult,Status,ScreenshotPath,StartTime,EndTime]

columnNameToNumber={}
currentTestCaseRowNumber=0
currentTestStepNumber=0

testStepFailed=False
