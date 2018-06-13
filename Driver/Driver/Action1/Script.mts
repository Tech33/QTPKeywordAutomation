Option Explicit
Dim strFrmwFolderPath                              
Dim strTestCaseFolderPath
'Test Case Folder Path                                  
Dim strDriverWorkbookName
'Name of the driver sheet
Dim strDriverWorkbookpath                                 
Dim strExecutionSheetName
Dim AllSuitesExcelObject
Dim strDataPrefix

strDriverWorkbookName = "Driver.xls"
strExecutionSheetName = "Execution"

strFrmwFolderPath = "..\..\"
strFrmwFolderPath = PathFinder.Locate(strFrmwFolderPath)


strTestCaseFolderPath = strFrmwFolderPath & "Test Cases"
strDriverWorkBookPath = strTestCaseFolderPath &"\" & strDriverWorkbookName

'Path to Driver Excel

Environment.Value("ResultFolder") = strFrmwFolderPath & "Results"

strDataPrefix = "#get#"

Set AllSuitesExcelobject = Fn_NewExcelReader()    'Excel Utility object to read driver excel

If ifFileExists(strDriverWorkbookPath) Then 

	AllSuitesExcelObject.SetFile strDriverWorkbookPath,GetSelectQuery(strExecutionSheetName,"ExecutionSheet")
Else
	msgbox "Driver files does not exists [" & strDriverWorkbookPath &"]"
	ExitTest
End If

Dim iSuiteRowCounter

For iSuiteRowCounter = 0 to AllSuitesExcelObject.GetRowCount - 1
	Dim strCurrSuiteSheet,SuiteExecutionExcel,strSuiteWorkbookPath,iTcRowCounter

	Environment.Value("CurrSuiteName") = AllSuitesExcelObject.GetCurrentCellValue("Suite_Name")
	strCurrSuiteSheet = AllSuitesExcelObject.GetCurrentCellValue("Suite_Workbook")
	strSuiteWorkbookPath = strTestCaseFolderPath & "\" & strCurrSuiteSheet
	Print Environment.Value("CurrSuiteName")
	Print strCurrSuiteSheet

	If IfFileExists(strSuiteWorkbookPath) Then
		Set SuiteExecutionExcel = Fn_NewExcelReader()
		SuiteExecutionExcel.SetFile strSuiteWorkbookPath,GetSelectQuery(strExecutionSheetName,"ExecutionSheet")

	For iTcRowCounter = 0 to SuiteExecutionExcel.GetRowCount - 1
	Dim strModuleName,TestCaseExcel,iTsRowCounter,strTestCaseStatus

	Environment.Value("TestCaseID") = SuiteExecutionExcel.GetCurrentCellValue("TestCaseNo")
	Environment.Value("TestCaseName") = SuiteExecutionExcel.GetCurrentCellValue("TestCase_Name")
	Environment.value("TestCaseFullName") = Environment.Value("CurrSuiteName") & "." & Environment.Value("TestCaseName")
	cp_SetResultLog Environment.Value("TestCaseFullName"),SuiteExecutionExcel.GetCurrentCellValue("Description")
	strTestCaseStatus = "Pass"

	strModuleName = SuiteExecutionExcel.GetCurrentCellValue("Module")
	Set TestCaseExcel = Fn_NewExcelReader()
	TestCaseExcel.SetFile strSuiteWorkbookPath,GetSelectQuery(strModuleName,"TestCaseSheet")

	RemoveAllvaluesFromDict
				
				For iTsRowCounter = 0 to TestCaseExcel.GetRowCount-1
				Dim strTS_Id,strTestStepDetail,strKeyword,strField,strTestData_1,strTestData_2,DictKeyword,strStringToBeEvaluated
				Dim blnActionStatus, strKey

				strTS_Id = TestCaseExcel.GetCurrentCellValue("TSID")
				Print strTS_Id
				strTestStepDetail = TestCaseExcel.GetCurrentCellValue("TestStep")
				strKeyword = TestCaseExcel.GetCurrentCellValue("Keyword")
				strField = TestCaseExcel.GetCurrentCellValue("Field")
				strTestData_1 = TestCaseExcel.GetCurrentCellValue("TestData_1")
				strTestData_2 = TestCaseExcel.GetCurrentCellValue("TestData_2")

				If Lcase(Left(strTestData_1,Len(strDataPrefix))) = Lcase(strDataPrefix) Then
					strKey=Mid(strTestData_1,Len(strDataPrefix)+1)
					strTestData_1=GetDataValueFromDict(Lcase(strkey))	
				End If
				If not (Trim(strTestData_2) ="") And Lcase(Left(strTestData_2,Len(strDataPrefix))) = Lcase(strDataPrefix) Then
					strKey = Mid(strTestData_2,Len(strDataPrefix)+1)
					strTestData_2=GetDataValueFromDict(Lcase(strkey))	
				End If

				Set DictKeyword = CreateObject("Scripting.Dictionary")
				DictKeyword.Add "Field",strField
				DictKeyword.Add "TestData_1",strTestData_1
				DictKeyword.Add "TestData_2",strTestData_2
				DictKeyword.Add "TestStepID",strTS_Id
				DictKeyword.Add "TestStepDetails",strTestStepDetail

				strStringToBeEvaluated = "blnActionStatus  =" & " " & strKeyWord & "( DictKeyword )"
				Execute (strStringToBeEvaluated)

				If not Cbool(blnActionStatus) Then 
					strTestCaseStatus = "Fail"
					Msgbox "Fail"
					Exit For
					End If
					TestCaseExcel.MoveToNextRecord
				Next
				cp_EndReport Environment.Value("TestCaseFullName"),strTestCaseStatus
				SuiteExecutionExcel.MoveToNextRecord
			Next
		Else
		msgbox "Suit files does not exits ["& strSuiteWorkbookPath &"]"
	End If

	AllSuitesExcelObject.MoveToNextRecord
Next	




'
'
'strSheetName = "CreateTrade"
'strExecutionFlagColmnName = "TCID"
'Environment.Value("TestCaseID") = "TC01"
'
'    Dim objAdodbConnection, objAdodbRecordSet 
'    Dim sQuery
'
'    Const adOpenForwardOnly = 0 
'    Const adOpenKeyset      = 1
'    Const adOpenDynamic     = 2
'    Const adOpenStatic      = 3
'    Const adLockOptimistic = 3
'
'
'    Set objAdodbConnection = CreateObject("ADODB.Connection")
'    Set objAdodbRecordSet = CreateObject("ADODB.Recordset")
'    
'    objAdodbConnection.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=C:\Style_Automation\Test Cases\New Trade.xls;ReadOnly=False;" 
'    objAdodbConnection.Open
'    sQuery="Select * from [" & strSheetName & "$] Where " & strExecutionFlagColmnName & " = '" & Environment.Value("TestCaseID") & "'"
'	' "'TC01'"
'    objAdodbRecordSet.Open sQuery, objAdodbConnection, adOpenKeyset, adLockOptimistic                  
'    
'    GetrReqParameter=objAdodbRecordSet(0).Value
'    msgbox GetrReqParameter
'    objAdodbRecordSet.Close
'    objAdodbConnection.Close
'
'
'
'
'
'
'
'
'
'
'

'***************************To Test CreateUpdateQuery From Excel.Util ***********************************

'strSuitWorkBookPath = "C:\Style_Automation\Test Cases\Driver.xls"
'strSheetName = "Execution"
'strColumnName = "Execution_Flag"
'TheValue = "No"
'
'Set TestCaseExcel = Fn_NewExcelReader()
'
'TestCaseExcel.SetFile  strSuitWorkBookPath,Fn_CreateUpdateQuery(strSheetName,strColumnName,TheValue)
'
'Public Function Fn_CreateUpdateQuery(strSheetName,strColumnName,TheValue)
'	Fn_CreateUpdateQuery = "Update [" & strSheetName & "$] Set " & strColumnName & " = '" & TheValue & "'"
'End Function
'
'





