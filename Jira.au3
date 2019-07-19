#include-once
#Include <Array.au3>
#include <GuiEdit.au3>
#include "cURL.au3"
#Include "Json.au3"
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
Global $jira_domain = ""
Global $jira_username = ""
Global $jira_password = ""
Global $jira_json = ""
Global $jira_html = ""
Global $app_name = "Data Extractor"
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
Func _JiraSetup()

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
Func _JiraShutdown()

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
Func _JiraDomainSet($domain)

	$jira_domain = $domain
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
Func _JiraLogin($username, $password)

	$jira_username = $username
	$jira_password = $password
EndFunc


; Search

Func _JiraSearchIssues($fields, $jql, $changelog = False)

	Local $startAt = 0
	Local $response = ""
	$jira_json = ""
	Local $changelog_json = ""

	if $changelog = True Then

		$changelog_json = "expand=changelog&"
	EndIf

	Local $jql_encoded = EncodeUrl($jql)

	Do

	;	MsgBox(0,"",$startAt)
;		ControlSetText($app_name, "", "Static1",  "_JiraSearchIssues startAt=" & $startAt & "&jql=" & $jql_encoded)


;		$response = cURL_easy($jira_domain & "/rest/api/2/search?" & $changelog_json & "fields=" & $fields & "&maxResults=100&startAt=" & $startAt & "&jql=" & $jql_encoded, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $jira_username & ":" & $jira_password)

		$ww = 'curl.exe -k -H "Accept: application/json" -H "Content-Type: application/json" -u ' & $jira_username & ':' & $jira_password & ' ' & $jira_domain & "/rest/api/2/search?" & $changelog_json & "fields=" & $fields & "&maxResults=100&startAt=" & $startAt & "&jql=" & $jql_encoded
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $ww = ' & $ww & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		Local $iPID = Run('curl.exe -k -H "Accept: application/json" -H "Content-Type: application/json" -u ' & $jira_username & ':' & $jira_password & ' ' & $jira_domain & "/rest/api/2/search?" & $changelog_json & "fields=" & $fields & "&maxResults=100&startAt=" & $startAt & "&jql=" & $jql_encoded, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
		ProcessWaitClose($iPID)
		$response = StdoutRead($iPID)
;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response = ' & $response & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


;	msgbox(0,"cc",$response[2])

;		FileDelete("D:\dwn\fred2.txt")
;		FileWrite("D:\dwn\fred2.txt", $jira_domain & "/rest/api/2/search?" & $changelog_json & "fields=" & $fields & "&maxResults=100&startAt=" & $startAt & "&jql=" & $jql_encoded)
;		FileDelete("D:\dwn\fred3.txt")
;		FileWrite("D:\dwn\fred3.txt", $response)


;		if StringInStr($response[2], """issues"":[]") > 0 Then
		if StringInStr($response, """issues"":[]") > 0 Then

			ExitLoop
		EndIf

		if StringLen($jira_json) > 0 Then

			$jira_json = StringLeft($jira_json, StringLen($jira_json) - 2)
;			$response[2] = "," & StringMid($response[2], StringInStr($response[2], "[") + 1)
			$response = "," & StringMid($response, StringInStr($response, "[") + 1)
		EndIf

;		$jira_json = $jira_json & $response[2]
		$jira_json = $jira_json & $response
		$startAt = $startAt + 100

	Until False


EndFunc


Func _JiraGetSearchResultKeysAndIssueTypeNames($fields, $jql)

	_JiraSearchIssues($fields, $jql)

	$rr = StringRegExp($jira_json, '(?U)"key":"(.*)".*"name":"(.*)"', 3)

	Return $rr

EndFunc

Func _JiraGetSearchResultKeysSummariesAndIssueTypeNames($fields, $jql)

	_JiraSearchIssues($fields, $jql)

;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_json = ' & $jira_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit

	; "key":"SEAB-4681","fields":{"summary":"1.13 - Regression Requirements - upload organisation units to the tenant","issuetype":

	local $rr[0]
	Local $decoded_json = Json_Decode($jira_json)

	for $i = 0 to 99999

		Local $key = Json_Get($decoded_json, '.issues[' & $i & '].key')

		if @error > 0 Then ExitLoop

		Local $summary = Json_Get($decoded_json, '.issues[' & $i & '].fields.summary')
		Local $name = Json_Get($decoded_json, '.issues[' & $i & '].fields.issuetype.name')

		_ArrayAdd($rr, $key, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $summary, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $name, 0, "|", @CRLF, 1)
	Next

	Return $rr

EndFunc

Func _JiraGetSearchResultKeysSummariesIssueTypeNameEpicKey($fields, $jql)

	_JiraSearchIssues($fields, $jql)

	; "key":"SEAB-4681","fields":{"summary":"1.13 - Regression Requirements - upload organisation units to the tenant","issuetype":

	$rr = StringRegExp($jira_json, '(?U)"key":"(.*)".*"summary":"(.*)".*"name":"(.*)".*"customfield_10008":"(.*)"', 3)



	Return $rr

EndFunc

Func _JiraGetSearchResultBugs($fields, $jql)

	_JiraSearchIssues($fields, $jql)

	FileDelete("D:\dwn\fred.txt")
	FileWrite("D:\dwn\fred.txt", $jira_json)
;Exit

	; "key":"SEAB-4681","fields":{"summary":"1.13 - Regression Requirements - upload organisation units to the tenant","issuetype":
	; Key,Summary,Reporter,Assignee,Status,Priority,AffectsVersions,FixVersions,Resolution,Labels,Environment,ScrumTeam,Sprint

;	$rr = StringRegExp($jira_json, '(?U)"key":"(.*)",.*"fixVersions":.*"name":"(.*)",.*"resolution":.*,"customfield_12000":.*"value":"(.*)",.*"priority":.*"name":"(.*)",.*"labels":[.*],.*"versions":.*"name":"(.*)",.*"status":.*"name":"(.*)",.*"reporter":.*"displayName":"(.*)",.*"customfield_10007":.*"name=(.*),goal=.*"summary":"(.*)",.*"environment":(.*),"customfield_11602".*"assignee":.*"displayName":"(.*)",', 3)
;	$rr = StringRegExp($jira_json, '(?U)"key":"(.*)","fields".*"fixVersions":.*"name":"(.*)",', 3)
;	$rr = StringRegExp($jira_json, '(?U)"id":[0-9]*,"key":"(.*)","fields"', 3)
;	$rr = StringRegExp($jira_json, '(?U)"self":".*","key":"(.*)","fields":{"customfield', 3)
;	$rr = StringRegExp($jira_json, '(?U)"key":".*","f', 3)
;	$rr = StringRegExp($jira_json, '(?U)\{"expand":".*","id":".*","self":".*","key":"(.*)".*"summary":"(.*)",.*"reporter":".*".*"displayName":"(.*)"', 3)
;	$rr = StringRegExp($jira_json, '(?U)\{"expand":".*","id":".*","self":".*","key":"(.*)".*"summary":"(.*)",.*"versions":(.*),.*"customfield_12000".*"value":"(.*)".*"reporter".*"displayName":"(.*)".*,"fixVersions":(.*),.*"priority":.*"name":"(.*)",.*"resolution":(.*),"labels":(.*)\}', 3)
;	$rr = StringRegExp($jira_json, '(?U)"operations,versionedRepresentations,editmeta,changelog,renderedFields","id":".*","self":".*","key":"(.*)".*"summary":"(.*)","customfield_10007":(.*),"environment":"(.*)","versions":(.*),.*"customfield_12000".*"value":"(.*)".*"reporter".*"displayName":"(.*)".*,"fixVersions":(.*),.*"priority":.*"name":"(.*)",.*"resolution":(.*)\{"expand":', 3)
;	_ArrayDisplay($rr)
;	Exit


;	Return $rr

EndFunc

Func _JiraGetSearchResultBugStatusHistory($fields, $jql)

	_JiraSearchIssues($fields, $jql, True)

	FileDelete("D:\dwn\fred.txt")
	FileWrite("D:\dwn\fred.txt", $jira_json)
;Exit

;	ControlSetText($app_name, "", "Static1",  "Decoding json ...")


	local $rr[0]
	Local $decoded_json = Json_Decode($jira_json)

	for $i = 0 to 99999

		Local $key = Json_Get($decoded_json, '.issues[' & $i & '].key')
;MsgBox(0,"a",$key)

		if @error > 0 or StringLen($key) = 0 Then ExitLoop

;		ControlSetText($app_name, "", "Static1",  "Getting History of Bug #" & $i & " key " & $key)

		Local $affected_version = Json_Get($decoded_json, '.issues[' & $i & '].fields.versions[0].name')
		Local $priority = Json_Get($decoded_json, '.issues[' & $i & '].fields.priority.name')

		Local $history = ""

		for $j = 0 to 99999

			$created = Json_Get($decoded_json, '.issues[' & $i & '].changelog.histories[' & $j & '].created')
;MsgBox(0,"b",$created)

			if @error > 0 or StringLen($created) = 0 Then ExitLoop

			for $k = 0 to 99999

				Local $field = Json_Get($decoded_json, '.issues[' & $i & '].changelog.histories[' & $j & '].items[' & $k & '].field')

				if @error > 0 or StringLen($field) = 0 Then ExitLoop

				if StringCompare($field, "status") = 0 Then

					Local $from = Json_Get($decoded_json, '.issues[' & $i & '].changelog.histories[' & $j & '].items[' & $k & '].fromString')
					Local $to = Json_Get($decoded_json, '.issues[' & $i & '].changelog.histories[' & $j & '].items[' & $k & '].toString')

					if StringLen($history) > 0 Then

						$history = $history & "|"
					EndIf

					$history = $history & $created & "," & $from & "," & $to
				EndIf
			Next
		Next

		if StringLen($history) > 0 Then

			$history = $history & "|"
		EndIf

		Local $created = Json_Get($decoded_json, '.issues[' & $i & '].fields.created')
		$history = $history & $created & ",,Open"

		_ArrayAdd($rr, $key, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $affected_version, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $priority, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $history, 0, "|", @CRLF, 1)
	Next

;	_ArrayDisplay($rr)
;	Exit

	Return $rr


EndFunc

Func _JiraGetSearchResultKeysSummariesIssueTypeNameEpicKeyRequirements($fields, $jql)

	_JiraSearchIssues($fields, $jql)

;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_json = ' & $jira_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit

	local $rr[0]
	Local $decoded_json = Json_Decode($jira_json)

	for $i = 0 to 99999

		Local $key = Json_Get($decoded_json, '.issues[' & $i & '].key')

		if @error > 0 Then ExitLoop

		Local $summary = Json_Get($decoded_json, '.issues[' & $i & '].fields.summary')
		Local $name = Json_Get($decoded_json, '.issues[' & $i & '].fields.issuetype.name')
		Local $customfield_10008 = Json_Get($decoded_json, '.issues[' & $i & '].fields.customfield_10008')

		Local $labels = ""

		for $j = 0 to 99999

			Local $label = Json_Get($decoded_json, '.issues[' & $i & '].fields.labels[' & $j & ']')

			if @error > 0 Then ExitLoop

			if StringInStr($label, "Req-ID-") > 0 Then

				$label = StringReplace($label, "Req-ID-", "")

				if StringLen($labels) > 0 Then

					$labels = $labels & "<br>"
				EndIf

				$labels = $labels & $label
			EndIf
		Next

		Local $fixVersions = Json_Get($decoded_json, '.issues[' & $i & '].fields.fixVersions[0].name')
		Local $status = Json_Get($decoded_json, '.issues[' & $i & '].fields.status.name')
		Local $test_notes = Json_Get($decoded_json, '.issues[' & $i & '].fields.customfield_15255')

		$test_notes = StringReplace($test_notes, @CRLF, "<br>")
		$test_notes = StringReplace($test_notes, @LF, "<br>")
		$test_notes = StringReplace($test_notes, @CR, "<br>")
		$test_notes = StringReplace($test_notes, "	", " ")

		_ArrayAdd($rr, $key, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $summary, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $name, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $customfield_10008, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $labels, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $fixVersions, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $status, 0, "|", @CRLF, 1)
;		_ArrayAdd($rr, $test_notes, 0, "|", @CRLF, 1)
	Next

	Return $rr

EndFunc



Func _JiraGetSearchResultKeysSummariesIssueTypeNameStoryKeyRequirements($fields, $jql)

	_JiraSearchIssues($fields, $jql)

;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_json = ' & $jira_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit

	local $rr[0]
	Local $decoded_json = Json_Decode($jira_json)

	for $i = 0 to 99999

		Local $key = Json_Get($decoded_json, '.issues[' & $i & '].key')

		if @error > 0 Then ExitLoop

		Local $summary = Json_Get($decoded_json, '.issues[' & $i & '].fields.summary')
		Local $description = Json_Get($decoded_json, '.issues[' & $i & '].fields.description')
		Local $name = Json_Get($decoded_json, '.issues[' & $i & '].fields.issuetype.name')
		Local $parent_key = Json_Get($decoded_json, '.issues[' & $i & '].fields.parent.key')

		Local $labels = ""

		for $j = 0 to 99999

			Local $label = Json_Get($decoded_json, '.issues[' & $i & '].fields.labels[' & $j & ']')

			if @error > 0 Then ExitLoop

			if StringInStr($label, "Req-ID-") > 0 Then

				$label = StringReplace($label, "Req-ID-", "")

				if StringLen($labels) > 0 Then

					$labels = $labels & "<br>"
				EndIf

				$labels = $labels & $label
			EndIf
		Next

		Local $fixVersions = Json_Get($decoded_json, '.issues[' & $i & '].fields.fixVersions[0].name')
		Local $status = Json_Get($decoded_json, '.issues[' & $i & '].fields.status.name')
		Local $aggregateprogress = Json_Get($decoded_json, '.issues[' & $i & '].fields.aggregateprogress.progress')
		Local $aggregateprogresstotal = Json_Get($decoded_json, '.issues[' & $i & '].fields.aggregateprogress.total')
		Local $aggregateprogresspercent = Json_Get($decoded_json, '.issues[' & $i & '].fields.aggregateprogress.percent')
		Local $environment = Json_Get($decoded_json, '.issues[' & $i & '].fields.environment')
;		Local $test_notes = Json_Get($decoded_json, '.issues[' & $i & '].fields.customfield_15255')

;		$test_notes = StringReplace($test_notes, @CRLF, "<br>")
;		$test_notes = StringReplace($test_notes, @LF, "<br>")
;		$test_notes = StringReplace($test_notes, @CR, "<br>")
;		$test_notes = StringReplace($test_notes, "	", " ")

		_ArrayAdd($rr, $key, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $summary, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $description, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $name, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $parent_key, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $labels, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $fixVersions, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $status, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $aggregateprogresstotal, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $aggregateprogress, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $aggregateprogresspercent, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $environment, 0, "|", @CRLF, 1)
;		_ArrayAdd($rr, $test_notes, 0, "|", @CRLF, 1)
	Next

	Return $rr

EndFunc



Func _JiraGetSearchResultKeysSummariesIssueTypeNamesIssueLinks($fields, $jql)

	_JiraSearchIssues($fields, $jql)

;	FileDelete("D:\dwn\fred.txt")
;	FileWrite("D:\dwn\fred.txt", $jira_json)
;	Exit

	; "key":"SEAB-4681","fields":{"summary":"1.13 - Regression Requirements - upload organisation units to the tenant","issuetype":



	local $rr[0]
	Local $decoded_json = Json_Decode($jira_json)

	for $i = 0 to 99999

		Local $key = Json_Get($decoded_json, '.issues[' & $i & '].key')

		if @error > 0 Then ExitLoop

		Local $summary = Json_Get($decoded_json, '.issues[' & $i & '].fields.summary')
		Local $name = Json_Get($decoded_json, '.issues[' & $i & '].fields.issuetype.name')

		Local $issuelinks = ""

		for $j = 0 to 99999

			Local $issuelink_id = Json_Get($decoded_json, '.issues[' & $i & '].fields.issuelinks[' & $j & '].id')

			if @error > 0 Then ExitLoop

			Local $inwardissue_key = Json_Get($decoded_json, '.issues[' & $i & '].fields.issuelinks[' & $j & '].inwardIssue.key')

			if @error = 0 Then

				if StringLen($issuelinks) = 0 Then

					$issuelinks = ","
				EndIf

				$issuelinks = $issuelinks & $inwardissue_key & ","
			EndIf
		Next

		_ArrayAdd($rr, $key, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $summary, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $name, 0, "|", @CRLF, 1)
		_ArrayAdd($rr, $issuelinks, 0, "|", @CRLF, 1)
	Next

	Return $rr

EndFunc

Func _JiraBrowseIssue($key)

	$response = cURL_easy($jira_domain & "/browse/" & $key, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $jira_username & ":" & $jira_password)
	$jira_html = $response[2]
EndFunc


Func _JiraGetVersions($project_key)

	$response = cURL_easy($jira_domain & "/rest/api/2/project/" & $project_key & "/versions", "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $jira_username & ":" & $jira_password)
	$jira_json = $response[2]
	$jira_json = '{"versions":' & $jira_json & '}'
EndFunc



Func _JiraGetVersionNames($project_key)


	_JiraGetVersions($project_key)

;	FileDelete("D:\dwn\fred.txt")
;	FileWrite("D:\dwn\fred.txt", $jira_json)
;Exit

	local $rr[0]
	Local $decoded_json = Json_Decode($jira_json)


	for $i = 0 to 99999

		Local $name = Json_Get($decoded_json, '.versions[' & $i & '].name')

		if @error > 0 Then ExitLoop

		_ArrayAdd($rr, $name, 0, "|", @CRLF, 1)
	Next

	Return $rr

EndFunc


Func _JiraGetTestRailTestCasesFromIssue($key)

	;_JiraBrowseIssue($key)

;	Local $testrail_url = StringRegExp($jira_html, '(?U)"url":"(https://jira.testrail.net/issues/references.*)"', 3)

;	_TestRailGetTestCases($key) ;$testrail_url[0])


;	$rr = StringRegExp($jira_json, '(?U)"key":"(.*)"', 3)

;	Return $rr

EndFunc



Func _JiraGetCurrentUser()

	$response = cURL_easy($jira_domain & "/rest/api/3/myself", "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $jira_username & ":" & $jira_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_domain = ' & $jira_domain & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_password = ' & $jira_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_username = ' & $jira_username & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	$jira_json = $response[2]

EndFunc




Func EncodeUrl($src)
    Local $i
    Local $ch
    Local $NewChr
    Local $buff

    ;Init Counter
    $i = 1

    While ($i <= StringLen($src))
        ;Get byte code from string
        $ch = Asc(StringMid($src, $i, 1))

        ;Look for what bytes we have
        Switch $ch
            ;Looks ok here
            Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
                $buff &= Chr($ch)
                ;Space found
            Case 32
                $buff &= "+"
            Case Else
                ;Convert $ch to hexidecimal
                $buff &= "%" & Hex($ch, 2)
        EndSwitch
        ;INC Counter
        $i += 1
    WEnd

    Return $buff
EndFunc   ;==>EncodeUrl

Func cURL_easy_retry($url, $cookie_file = "", $cookie_action = 0, $output_type = 0, $output_file = "", $request_headers = "", $request_data = "", $ssl_verifypeer = 0, $noprogress = 1, $followlocation = 0, $num_of_retries = 10)

	Local $response

	for $i = 1 to $num_of_retries

		$response = cURL_easy($url, $cookie_file, $cookie_action, $output_type, $output_file, $request_headers, $request_data, $ssl_verifypeer, $noprogress, $followlocation)

		if $response[0] <> 500 Then

			Return $response
		EndIf

		ConsoleWrite("url failed with response code " & $response[0] & " - " & $url & @CRLF)
		Sleep(500)
	Next
EndFunc

