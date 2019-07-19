#include-once
#Include <Array.au3>
#include <GuiEdit.au3>
#include "cURL.au3"
#Region Header
#cs
	Title:   		Janison Insights Automation UDF Library for AutoIt3
	Filename:  		JanisonInsights.au3
	Description: 	A collection of functions for creating, attaching to, reading from and manipulating Janison Insights
	Author:   		seangriffin
	Version:  		V0.1
	Last Update: 	25/02/18
	Requirements: 	AutoIt3 3.2 or higher,
					Janison Insights Release x.xx,
					cURL xxx
	Changelog:		---------24/12/08---------- v0.1
					Initial release.
#ce
#EndRegion Header
#Region Global Variables and Constants
Global Const $sap_vkey[100] = [ "Enter", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", _
								"F11", _ ; NOTE - "F11" is the same as "CTRL+S"
								"F12", _ ; NOTE - "F12" is the same as "Esc"
								"Shift+F1", "Shift+F2", "Shift+F3", "Shift+F4", "Shift+F5", "Shift+F6", "Shift+F7", "Shift+F8", "Shift+F9", _
								"Shift+Ctrl+0", "Shift+F11", "Shift+F12", _
								"Ctrl+F1", "Ctrl+F2", "Ctrl+F3", "Ctrl+F4", "Ctrl+F5", "Ctrl+F6", "Ctrl+F7", "Ctrl+F8", "Ctrl+F9", "Ctrl+F10", _
								"Ctrl+F11", "Ctrl+F12", _
								"Ctrl+Shift+F1", "Ctrl+Shift+F2", "Ctrl+Shift+F3", "Ctrl+Shift+F4", "Ctrl+Shift+F5", _
								"Ctrl+Shift+F6", "Ctrl+Shift+F7", "Ctrl+Shift+F8", "Ctrl+Shift+F9", "Ctrl+Shift+F10", "Ctrl+Shift+F11", _
								"Ctrl+Shift+F12", _
								"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
								"Ctrl+E", "Ctrl+F", "Ctrl+A", "Ctrl+D", "Ctrl+N", "Ctrl+O", "Shift+D", "Ctrl+I", "Shift+I", "Alt+B", _
								"Ctrl+Page up", "Page up", "Page down", "Ctrl+Page down", "Ctrl+G", "Ctrl+R", "Ctrl+P", _
								"", "", "", "", "", "", "", "Shift+F10", "", "", "", "", "" ]
Global $testrail_domain = ""
Global $testrail_username = ""
Global $testrail_password = ""
Global $testrail_json = ""
Global $testrail_html = ""
#EndRegion Global Variables and Constants
#Region Core functions
; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsSetup()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsSetup()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailSetup()

	; Initialise cURL
	cURL_initialise()


EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsShutdown()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsShutdown()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailShutdown()

	; Clean up cURL
	cURL_cleanup()

EndFunc


; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsDomainSet()
; Description ...:	Sets the domain to use in all other functions.
; Syntax.........:	_InsightsDomainSet($domain)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailDomainSet($domain)

	$testrail_domain = $domain
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsLogin()
; Description ...:	Login a user to Janison Insights.
; Syntax.........:	_InsightsLogin($username, $password)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailLogin($username, $password)

	$testrail_username = $username
	$testrail_password = $password
EndFunc

; Authentication


Func _TestRailAuth()

;	$response = cURL_easy($testrail_domain, "cookies.txt", 2, 0, "", "Content-Type: text/html", "name=sgriffin@janison.com&password=Gri01ffo&rememberme=1", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	Exit

	Local $iPID = Run('curl.exe -k -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/auth/login -c cookies.txt -d "name=' & $testrail_username & '&password=' & $testrail_password & '&rememberme=1" -X POST', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

EndFunc


; Projects

Func _TestRailGetProjects()

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_projects", "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	$testrail_json = $response[2]

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_projects', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

EndFunc

Func _TestRailGetProjectsIDAndNameArray()

	Local $output[0][2]

	_TestRailGetProjects()

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"name":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $name = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$name = StringMid($name, StringInStr($name, ":", 0, -1) + 1)
		$name = StringReplace($name, """", "")
		Local $id_name = $id & "|" & $name
		_ArrayAdd($output, $id_name)
	Next

	Return $output
EndFunc

; Suites

Func _TestRailGetSuitesIdName($project_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_suites/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_suites/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"name":"(.*)",', 3)
	Return $rr


EndFunc

Func _TestRailGetCasesIdTitleReferences($project_id, $suite_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_cases/" & $project_id & "&suite_id=" & $suite_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_cases/' & $project_id & '&suite_id=' & $suite_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"title":"(.*)",.*"refs":"(.*)"', 3)
	Return $rr
EndFunc

Func _TestRailGetCase($case_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_case/" & $case_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _TestRailGetRunsIdName($project_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_runs/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_runs/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"name":"(.*)"', 3)
	Return $rr

EndFunc

Func _TestRailGetRun($run_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_run/" & $run_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _TestRailGetRunIDFromPlanIDAndRunName($plan_id, $run_name)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '"runs":\[{"id":\d+,.*"name":"' & $run_name & '"', 1)
	$rr[0] = StringLeft($rr[0], StringInStr($rr[0], ",") - 1)
	$rr[0] = StringMid($rr[0], StringInStr($rr[0], ":", 0, 2) + 1)

	Return $rr[0]


EndFunc

Func _TestRailGetPlanRunsID($plan_id)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":(\d+),"suite_id":\d+,"name":".*","description"', 3)

	return $rr
EndFunc

Func _TestRailGetPlanRunsIDAndNameArray($plan_id)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":(\d+),"suite_id":\d+,"name":"(.*)","description"', 3)

	return $rr
EndFunc


Func _TestRailGetResults($test_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_results/" & $test_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _TestRailGetResultsForCaseIdTestIdStatusIdDefects($run_id, $case_id)

;	Local $cmd = 'curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $cmd = ' & $cmd & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	$rr = StringRegExp($testrail_json, '(?U)"defects":"(.*)","id":(.*),"status_id":(.*),"test_id":(.*),', 3)
	Return $rr


EndFunc

Func _TestRailGetResultsIdStatusIdDefects($test_id)

;	Local $cmd = 'curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $cmd = ' & $cmd & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results/' & $test_id & '&limit=1', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	filedelete("D:\dwn\fred.txt")
	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"test_id":(.*),.*"status_id":(.*),.*"defects":(.*),"custom_step_results"', 3)
;	_ArrayDisplay($rr)
;	Exit
	Return $rr


EndFunc

Func _TestRailGetResultsForRunIdStatusIdCreatedOnDefects($run_id)

;	Local $cmd = 'curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $cmd = ' & $cmd & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_run/' & $run_id & '&limit=1', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_run/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	filedelete("D:\dwn\fred.txt")
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"test_id":(.*),.*"status_id":(.*),.*"created_on":(.*),.*"defects":(.*),"custom_step_results"', 3)
;	_ArrayDisplay($rr)
;	Exit
	Return $rr


EndFunc

Func _TestRailGetTestsIdTitleCaseIdRunId($run_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_tests/" & $run_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_tests/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"case_id":(.*),.*"run_id":(.*),.*"title":"(.*)",', 3)
	Return $rr


EndFunc

Func _TestRailGetTestsTitleAndIDFromRunID($run_id)

	Local $test_title_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetTests($run_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"title":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $title = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$title = StringMid($title, StringInStr($title, ":", 0, -1) + 1)
		$title = StringReplace($title, """", "")
		$test_title_and_id_dict.Add($title, $id)
	Next

	Return $test_title_and_id_dict

EndFunc

Func _TestRailGetTestsReferenceAndIDFromRunID($run_id)

	Local $test_refs_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetTests($run_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"refs":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $refs = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $id = ' & $id & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		$refs = StringMid($refs, StringInStr($refs, ":", 0, -1) + 1)
		$refs = StringReplace($refs, """", "")
;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $refs = ' & $refs & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		$test_refs_and_id_dict.Add($refs, $id)
	Next

	Return $test_refs_and_id_dict

EndFunc

Func _TestRailGetPlan($plan_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_plan/" & $plan_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_plan/' & $plan_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

EndFunc

Func _TestRailGetPlans($project_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_plans/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_plans/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


EndFunc

Func _TestRailGetPlansIDAndNameArray($project_id)

	Local $output[0][2]

	_TestRailGetPlans($project_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"name":".*"', 3)

	if IsArray($rr) Then

		for $each in $rr

			Local $id = $each
			Local $name = $each

			$id = StringLeft($id, StringInStr($id, ",") - 1)
			$id = StringMid($id, StringInStr($id, ":") + 1)
			$name = StringMid($name, StringInStr($name, ":", 0, -1) + 1)
			$name = StringReplace($name, """", "")
			Local $id_name = $id & "|" & $name
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $id_name = ' & $id_name & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			_ArrayAdd($output, $id_name)
		Next
	EndIf

	Return $output
EndFunc

Func _TestRailGetPlanIDByName($project_id, $plan_name)

	_TestRailGetPlans($project_id)

	$rr = StringRegExp($testrail_json, '"id":\d+,"name":"' & $plan_name & '"', 1)
	$rr[0] = StringLeft($rr[0], StringInStr($rr[0], ",") - 1)
	$rr[0] = StringMid($rr[0], StringInStr($rr[0], ":") + 1)

	Return $rr[0]
EndFunc

Func _TestRailAddResult($test_id, $status_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/add_result/" & $test_id, "", 0, 0, "", "Content-Type: application/json", '{"status_id":' & $status_id & '}', 0, 1, 0, $testrail_username & ":" & $testrail_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _TestRailAddResults($run_id, $results_arr)

	; create the results JSON to post from the results array

	Local $results_json = '{"results":['

	for $i = 0 to (UBound($results_arr) - 1) step 3

		if StringLen($results_json) > StringLen('{"results":[') Then

			$results_json = $results_json & ','
		EndIf

		$results_json = $results_json & '{"test_id":' & $results_arr[$i + 0] & ',"status_id":' & $results_arr[$i + 1] & ',"comment":"' & $results_arr[$i + 2] & '"}'
	Next

	$results_json = $results_json & ']}'
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $results_json = ' & $results_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	FileDelete(@ScriptDir & "\curl_in.json")
	FileWrite(@ScriptDir & "\curl_in.json", $results_json)
	Local $iPID = Run('curl.exe -s -k -H "Content-Type: application/json" --data @curl_in.json -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/add_results/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	; unfortunately the below is unreliable.  Working intermittently
;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/add_results/" & $run_id, "", 0, 0, "", "Content-Type: application/json", $results_json, 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc


Func _TestRailGetIdFromTitle($json, $title)

	$rr = StringRegExp($json, '"id":.*"title":"' & $title & '"', 1)
	$tt = StringMid($rr[0], StringLen('"id":') + 1, StringInStr($rr[0], ",") - (StringLen('"id":') + 1))
	Return $tt

EndFunc

Func _TestRailGetStatusesIdLabel()

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_statuses", "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_statuses/', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	filedelete("D:\dwn\fred.txt")
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"label":"(.*)"', 3)
	Return $rr


EndFunc

Func _TestRailGetStatusLabelAndID()

	Local $status_label_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetStatuses()

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"label":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $label = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$label = StringMid($label, StringInStr($label, ":", 0, -1) + 1)
		$label = StringReplace($label, """", "")

		$status_label_and_id_dict.Add($label, $id)
	Next

	Return $status_label_and_id_dict

;	_ArrayDisplay($rr)

EndFunc


Func _TestRailGetSections($project_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_sections/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	$testrail_json = $response[2]
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc


Func _TestRailGetSectionNameAndDepth($project_id)

	_TestRailGetSections($project_id)

	$rr = StringRegExp($testrail_json, '(?U)"name":"(.*)",.*"depth":(\d+)', 3)

	Return $rr

EndFunc

; Jira integration



Func _TestRailGetTestCases($key)

	$response = cURL_easy($testrail_domain & "/index.php?/ext/jira/render_panel&ae=connect&av=1&issue=" & $key & "&panel=references&login=button&frame=tr-frame-panel-references", "cookies.txt", 1, 0, "", "Content-Type: text/html; charset=UTF-8", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	$testrail_html = $response[2]
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_html = ' & $testrail_html & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc




