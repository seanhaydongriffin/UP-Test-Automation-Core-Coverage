#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <Date.au3>
#include <File.au3>
#include <Array.au3>
#Include "Json.au3"
#include "Jira.au3"
#include "Confluence.au3"
#include "TestRail.au3"
#include <WindowsConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Crypt.au3>
#include "Toast.au3"


;$rr = "*Done*" & @CRLF & "Name Search - [RMS C68651|https://janison.testrail.com/index.php?/cases/view/68651]*Not Done*Date Search"

;$rr = StringReplace($rr, @CRLF, "<br>")
;$rr = StringRegExpReplace($rr, "(?U)\*(.*)\*", "<b>$1</b>")
;$rr = StringRegExpReplace($rr, "(?U)\[(.*)\|(.*)\]", "<a href=""$2"" target=""_blank"">$1</a>")
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $rr = ' & $rr & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console




Global $app_name = "UP Test Automation Core Coverage"
Global $ini_filename = @ScriptDir & "\" & $app_name & ".ini"
Global $log_filepath = @ScriptDir & "\" & $app_name & ".log"
Global $html, $markup, $storage_format
Global $aResult, $iRows, $iColumns, $iRval, $run_name = "", $max_num_defects = 0, $max_num_days = 0, $version_name = ""

; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)

FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Exec(-1, "PRAGMA synchronous = OFF;")		; this should speed up DB transactions
_SQLite_Exec(-1, "CREATE TABLE ProjectVersion (Name);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE Epic (Key,Summary);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE Story (Key,Summary,EpicKey,ReqID,FixVersion,Status);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SubTask (Key,Summary,Description,StoryKey,FixVersion,Status,EstimatedTime,TimeSpent,ProgressPercent,Environment);") ; CREATE a Table


; Startup Jira & TestRail

;GUICtrlSetData($status_input, "Starting the Jira connection ... ")
_JiraSetup()
_JiraDomainSet("https://janisoncls.atlassian.net")

_Toast_Set(0, -1, -1, -1, -1, -1, "", 100, 100)
_Toast_Show(0, $app_name, "Login to Jira ...", -30, False, True)

Local $jira_username = IniRead($ini_filename, "main", "username", "")
Local $jira_encrypted_password = IniRead($ini_filename, "main", "password", "")
Global $jira_decrypted_password = ""
Global $username_input, $password_input

if stringlen($jira_encrypted_password) > 0 Then

	$jira_decrypted_password = _Crypt_DecryptData($jira_encrypted_password, @ComputerName & @UserName, $CALG_AES_256)
	$jira_decrypted_password = BinaryToString($jira_decrypted_password)
EndIf

if stringlen($jira_decrypted_password) > 0 Then

	_JiraLogin($jira_username, $jira_decrypted_password)
	_JiraGetCurrentUser()
EndIf

if stringlen($jira_decrypted_password) = 0 or StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) > 0 Then

	_Toast_Show(0, $app_name, "Username or password incorrect or not set.                       " & @CRLF & "Set your Jira login below." & @CRLF & @CRLF & @CRLF & @CRLF & @CRLF, -9999, False, True)
	GUICtrlCreateLabel("Username:", 10, 70, 80, 20)
	$username_input = GUICtrlCreateInput("", 80, 70, 200, 20)
	GUICtrlCreateLabel("Password:", 10, 90, 80, 20)
	$password_input = GUICtrlCreateInput("", 80, 90, 200, 20, $ES_PASSWORD)
	$done_button = GUICtrlCreateButton("Done", 80, 110, 80, 20)

	While 1

		$msg = GUIGetMsg()

		if $msg = $done_button Then

			Local $tmp_username = GUICtrlRead($username_input)
			Local $tmp_password = GUICtrlRead($password_input)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $tmp_password = ' & $tmp_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			_Toast_Show(0, $app_name, "Login to Jira ...", -30, False, True)
			_JiraLogin($tmp_username, $tmp_password)
			_JiraGetCurrentUser()

			if StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) = 0 Then

				IniWrite($ini_filename, "main", "username", $tmp_username)
				$jira_encrypted_password = _Crypt_EncryptData($tmp_password, @ComputerName & @UserName, $CALG_AES_256)
				IniWrite($ini_filename, "main", "password", $jira_encrypted_password)
			EndIf

			_Toast_Hide()
			ExitLoop
		EndIf

		if $hToast_Handle = 0 Then

			Exit
		EndIf
	WEnd
EndIf

if StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) > 0 Then

	_Toast_Show(0, $app_name, "Username or password incorrect or not set." & @CRLF & "Exiting ...", -5, true, True)
	Exit
EndIf

; get all epics for the project

_Toast_Show(0, $app_name, "get all epics for the project", -30, False, True)
$issue = _JiraGetSearchResultKeysSummariesAndIssueTypeNames("summary,issuetype", "project = QA AND issuetype = Epic AND labels in (Core) AND labels in (Automation) AND labels in (Assessments)")
;$issue = _JiraGetSearchResultKeysSummariesAndIssueTypeNames("", "project = QA AND issuetype = Epic AND labels in (Core) AND labels in (Automation)")

for $i = 0 to (UBound($issue) - 1) Step 3

	$issue[$i + 1] = StringReplace($issue[$i + 1], "'", "''")

	$query = "INSERT INTO Epic (Key,Summary) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "');"
	_FileWriteLog($log_filepath, "Epic " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

; get all stories for the project

_Toast_Show(0, $app_name, "get all stories for the project", -30, False, True)
$issue = _JiraGetSearchResultKeysSummariesIssueTypeNameEpicKeyRequirements("summary,issuetype,customfield_10008,labels,fixVersions,status", "project = QA AND issuetype = Story AND labels in (Core) AND labels in (Automation) AND labels in (Assessments)")

for $i = 0 to (UBound($issue) - 1) Step 7

	$issue[$i + 1] = StringReplace($issue[$i + 1], "'", "''")

	$query = "INSERT INTO Story (Key,Summary,EpicKey,ReqID,FixVersion,Status) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 3] & "','','" & $issue[$i + 5] & "','" & $issue[$i + 6] & "');"
	_FileWriteLog($log_filepath, "Story " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

; get all sub-tasks for the project

_Toast_Show(0, $app_name, "get all sub-tasks for the project", -30, False, True)
$issue = _JiraGetSearchResultKeysSummariesIssueTypeNameStoryKeyRequirements("summary,description,issuetype,parent,labels,fixVersions,status,aggregateprogress,environment", "project = QA AND issuetype = Sub-task AND labels in (Core) AND labels in (Automation) AND labels in (Assessments)")

for $i = 0 to (UBound($issue) - 1) Step 12

	$issue[$i + 1] = StringReplace($issue[$i + 1], "'", "''")
	$issue[$i + 2] = StringReplace($issue[$i + 2], "'", "''")

	$query = "INSERT INTO SubTask (Key,Summary,Description,StoryKey,FixVersion,Status,EstimatedTime,TimeSpent,ProgressPercent,Environment) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 2] & "','" & $issue[$i + 4] & "','" & $issue[$i + 6] & "','" & $issue[$i + 7] & "','" & $issue[$i + 8] & "','" & $issue[$i + 9] & "','" & $issue[$i + 10] & "','" & $issue[$i + 11] & "');"
	_FileWriteLog($log_filepath, "SubTask " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

; Shutdown Jira

_JiraShutdown()


; select Epic.Summary as "Sub Category", Story.Summary as "Pages", SubTask.Summary as "Processes", case when SubTask.Status = 'Done' then 'Skip' when SubTask.Status = 'Waiting For Build' then 'Yes' when SubTask.Status = 'In Progress' or SubTask.Status = 'Beta Testing' or SubTask.Status = 'Ready' or SubTask.Status = 'Closed' or SubTask.Status = 'Cancelled' or SubTask.Status = 'Resolved' or SubTask.Status = 'Test Run' then 'Yes' else 'No' end as "Coverage", '' as "Percentage Completed", '' as "Platform Coverage", '' as "Notes" from Epic left join Story on Epic.Key = Story.EpicKey left join SubTask on Story.Key = SubTask.StoryKey

_Toast_Show(0, $app_name, "creating report", -30, False, True)
Create_HTML_Report(True)
Local $html = FileRead(@ScriptDir & "\html_report.html")
_Toast_Show(0, $app_name, "uploading report to confluence", -30, False, True)
Update_Confluence_Page("https://janisoncls.atlassian.net", $jira_username, $jira_password, "JAST", "390496535", "390922241", "UP Test Automation Core Coverage", $html)







Func SQLite_to_HTML_table($query, $th_classes, $td_classes, $empty_message, $run_id, $merged_cell_for_column_numbers, $confluence_html = False)

	Local $double_quotes = """"

	if $confluence_html = true Then

		$double_quotes = "\"""
	EndIf

	Local $th_class = StringSplit($th_classes, ",", 2)
	Local $td_class = StringSplit($td_classes, ",", 2)

	Local $aResult, $iRows, $iColumns, $iRval, $run_name = ""

;	$xx = "SELECT RunName AS ""Run Name"" FROM report WHERE RunID = '" & $run_id & "';"
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $xx = ' & $xx & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	if StringLen($run_id) > 0 Then

		$iRval = _SQLite_GetTable2d(-1, "SELECT RunName AS ""Run Name"" FROM report WHERE RunID = '" & $run_id & "';", $aResult, $iRows, $iColumns)

		If $iRval = $SQLITE_OK Then

;			_SQLite_Display2DResult($aResult)

			$run_name = $aResult[1][0]
		EndIf

		$html = $html &	"<h3>Test Run " & $run_id & " - " & $run_name & "</h3>" & @CRLF
	EndIf

;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)

	If $iRval = $SQLITE_OK Then

;		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		Local $num_cols = UBound($aResult, 2)

		if $num_rows < 2 Then

			$html = $html &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			if $confluence_html = true Then

				$html = $html &	"<font size=\""1\""><table class=\""wrapped fixed-table\"">" & @CRLF
;				$html = $html &	"<table>" & @CRLF
			Else

				$html = $html &	"<table style=" & $double_quotes & "table-layout:fixed" & $double_quotes & ">" & @CRLF
			EndIf

			$html = $html & "<tr>"

			Local $merged_cell_for_column_number = StringSplit($merged_cell_for_column_numbers, ",", 3)
			Local $rowspan_rows_remaining_for_column[$num_cols]

			for $i = 0 to ($num_cols - 1)

				$rowspan_rows_remaining_for_column[$i] = 0

				if $confluence_html = true Then

					$html = $html & "<th width=" & $double_quotes & $th_class[$i] & $double_quotes & ">" & $aResult[0][$i] & "</th>" & @CRLF
				Else

					$html = $html & "<th class=" & $double_quotes & $th_class[$i] & $double_quotes & ">" & $aResult[0][$i] & "</th>" & @CRLF
				EndIf
			Next

			$html = $html & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$html = $html & "<tr>"

				for $j = 0 to ($num_cols - 1)

;					if $j = 2 Then

	;					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $aResult[$i][$j] = ' & $aResult[$i][$j] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;						Switch $aResult[$i][$j]

;							case "Passed"

;								$td_class[$j] = "trp"

;							case "Failed"

;								$td_class[$j] = "trf"

;							case "Untested"

;								$td_class[$j] = "tru"

;							case "Blocked"

;								$td_class[$j] = "trb"
;						EndSwitch
;					EndIf


					if $confluence_html = true Then

						Local $html_text = $aResult[$i][$j]

						$html_text = StringReplace($html_text, @CRLF, "<br/>")
						$html_text = StringReplace($html_text, @LF, "<br/>")
						$html_text = StringRegExpReplace($html_text, "(?U)\*(.*)\*", "<b>$1</b>")
						$html_text = StringRegExpReplace($html_text, "(?U)\[(.*)\|(.*)\]", "<a href=""$2"" target=""_blank"">$1</a>")

;						$html_text = StringReplace($html_text, " \</td>", " \\</td>")
						$html_text = StringRegExpReplace($html_text, "([^\\])\\$", "$1\\\\")
;						$a = StringRegExpReplace($a, "([^\\])\\$", "$1\\\\")
						$html_text = StringReplace($html_text, "<br>", "<br/>")
						$html_text = StringReplace($html_text, "&", "&amp;")
						$html_text = StringReplace($html_text, """", "\""")
						$html_text = StringReplace($html_text, "\\""", "\""")

						; determine if this should be a rowspan

						Local $num_rowspans = 0
						Local $rowspan_html = ""

						if Int($merged_cell_for_column_number[$j]) = 1 And $rowspan_rows_remaining_for_column[$j] <= 0 Then

							for $x = $i to ($num_rows - 2)

								if StringCompare($aResult[$x][$j], $aResult[$x + 1][$j]) <> 0 Then

									ExitLoop
								EndIf

								$num_rowspans = $num_rowspans + 1
							Next

							if $num_rowspans > 0 Then

								$num_rowspans = $num_rowspans + 1
;								$rowspan_html = " rowspan=""" & $num_rowspans & """"
								$html = $html & "<td rowspan=\""" & $num_rowspans & "\"">" & $html_text & "</td>" & @CRLF
								$rowspan_rows_remaining_for_column[$j] = $num_rowspans
							EndIf
						EndIf

						if $rowspan_rows_remaining_for_column[$j] > 0 Then

							$rowspan_rows_remaining_for_column[$j] = $rowspan_rows_remaining_for_column[$j] - 1
						Else

							$html = $html & "<td>" & $html_text & "</td>" & @CRLF
						EndIf
					Else

						$html = $html & "<th class=" & $double_quotes & $th_class[$i] & $double_quotes & ">" & $html_text & "</th>" & @CRLF
					EndIf
				Next

				$html = $html & "</tr>" & @CRLF
			Next

			$html = $html &	"</table></font>" & @CRLF
;			$html = $html &	"</table>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf
EndFunc


Func Create_HTML_Report($confluence_html = False)

	_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")

	$html = 				""

	if $confluence_html = False Then

		$html = $html &		"<!DOCTYPE html>" & @CRLF & _
							"<html>" & @CRLF & _
							"<head>" & @CRLF & _
							"<style>" & @CRLF & _
							"table, th, td {" & @CRLF & _
							"    border: 1px solid black;" & @CRLF & _
							"    border-collapse: collapse;" & @CRLF & _
							"    font-size: 12px;" & @CRLF & _
							"    font-family: Arial;" & @CRLF & _
							"}" & @CRLF & _
							".ds {min-width: 400px; text-align: left;}" & @CRLF & _
							".tes {min-width: 800px; text-align: left;}" & @CRLF & _
							".mti {min-width: 110px; text-align: center;}" & @CRLF & _
							".tt {min-width: 500px; text-align: left;}" & @CRLF & _
							".ati {min-width: 150px; text-align: center;}" & @CRLF & _
							".sd {min-width: 1000px; text-align: left;}" & @CRLF & _
							".tc {min-width: 300px; text-align: left;}" & @CRLF & _
							".tr {min-width: 110px; text-align: center;}" & @CRLF & _
							".ts {min-width: 90px; text-align: center;}" & @CRLF & _
							".trp {min-width: 110px; text-align: center; background-color: yellowgreen;}" & @CRLF & _
							".trf {min-width: 110px; text-align: center; background-color: lightcoral; color:white;}" & @CRLF & _
							".tru {min-width: 110px; text-align: center; background-color: lightgray;}" & @CRLF & _
							".trb {min-width: 110px; text-align: center; background-color: darkred; color: white;}" & @CRLF & _
							".pass {background-color: yellowgreen;}" & @CRLF & _
							".fail {background-color: lightcoral; color:white;}" & @CRLF & _
							".untested {background-color: lightgray;}" & @CRLF & _
							".mp {background-color: yellow;}" & @CRLF & _
							".rh {background-color: seagreen; color: white;}" & @CRLF & _
							".rhr {background-color: seagreen; color: white; text-align:center; white-space:nowrap; transform-origin:50% 50%; transform: rotate(-90deg);}" & @CRLF & _
							".rhr:before {background-color: seagreen; color: white; content:''; padding-top:100%; display:inline-block; vertical-align:middle;}" & @CRLF & _
							".i {background-color: deepskyblue;}" & @CRLF & _
							"</style>" & @CRLF & _
							"</head>" & @CRLF & _
							"<body>" & @CRLF
	EndIf

	$html = $html & 		"<br /><a href=\""https://janisoncls.atlassian.net/secure/RapidBoard.jspa?projectKey=QA&amp;rapidView=157\"">Click to open the Automation Core Kanban Board</a><br /><br />" & @CRLF

	Local $num_coverage = 0
	Local $num_coverage_skipped = 0
	Local $num_coverage_yes = 0
	Local $num_coverage_no = 0
	Local $pcnt_coverage_yes = 0

	if $confluence_html = True Then

		$iRval = _SQLite_GetTable2d(-1, "select count(*) from Epic left join Story on Epic.Key = Story.EpicKey left join SubTask on Story.Key = SubTask.StoryKey order by Epic.Summary, Story.Summary, SubTask.Summary;", $aResult, $iRows, $iColumns)

		If $iRval = $SQLITE_OK Then

			$num_coverage = $aResult[1][0]
		EndIf

		$iRval = _SQLite_GetTable2d(-1, "select count(*) from Epic left join Story on Epic.Key = Story.EpicKey left join SubTask on Story.Key = SubTask.StoryKey where SubTask.Status = 'Done' order by Epic.Summary, Story.Summary, SubTask.Summary;", $aResult, $iRows, $iColumns)

		If $iRval = $SQLITE_OK Then

			$num_coverage_skipped = $aResult[1][0]
		EndIf

		$iRval = _SQLite_GetTable2d(-1, "select count(*) from Epic left join Story on Epic.Key = Story.EpicKey left join SubTask on Story.Key = SubTask.StoryKey where SubTask.Status = 'Waiting For Build' or SubTask.Status = 'In Progress' or SubTask.Status = 'Beta Testing' or SubTask.Status = 'Ready' or SubTask.Status = 'Closed' or SubTask.Status = 'Cancelled' or SubTask.Status = 'Resolved' or SubTask.Status = 'Test Run' order by Epic.Summary, Story.Summary, SubTask.Summary;", $aResult, $iRows, $iColumns)

		If $iRval = $SQLITE_OK Then

			$num_coverage_yes = $aResult[1][0]
		EndIf

		$num_coverage = $num_coverage - $num_coverage_skipped
		$num_coverage_no = $num_coverage - $num_coverage_yes
		$pcnt_coverage_yes = Int(($num_coverage_yes / $num_coverage) * 100)

		$html = $html & 	"Percent Covered = " & $pcnt_coverage_yes & "%<br />Total Covered = " & $num_coverage_yes & "<br />Total Not Covered = " & $num_coverage_no & "<br />Total Skipped = " & $num_coverage_skipped & "<br />Total = " & $num_coverage & "<br /><br />" & @CRLF

		SQLite_to_HTML_table("select '<a href=""https://janisoncls.atlassian.net/browse/' || Epic.Key || '"" target=""_blank"">' || Epic.Summary || '</a>' as ""Sub Category"", '<a href=""https://janisoncls.atlassian.net/browse/' || Story.Key || '"" target=""_blank"">' || Story.Summary || '</a>' as ""Pages"", '<a href=""https://janisoncls.atlassian.net/browse/' || SubTask.Key || '"" target=""_blank"">' || SubTask.Summary || '</a>' as ""Processes"", case when SubTask.Status = 'Done' then 'Skip' when SubTask.Status = 'Waiting For Build' then 'Yes' when SubTask.Status = 'In Progress' or SubTask.Status = 'Beta Testing' or SubTask.Status = 'Ready' or SubTask.Status = 'Closed' or SubTask.Status = 'Cancelled' or SubTask.Status = 'Resolved' or SubTask.Status = 'Test Run' then 'Yes' else 'No' end as ""Coverage"", case when length(SubTask.ProgressPercent) > 0 then SubTask.ProgressPercent || '%' else '' end as ""Percentage Completed"", SubTask.Environment as ""Platform Coverage"", SubTask.Description as ""Notes"" from Epic left join Story on Epic.Key = Story.EpicKey left join SubTask on Story.Key = SubTask.StoryKey order by Epic.Summary, Story.Summary, SubTask.Summary;", "150,200,200,50,60,100,500", "ts,tc,tc,ts,tc,tc,tc", "", "", "1,1,0,0,0,0,0", $confluence_html)
	EndIf

	$html = $html & 		"<br /><br /><a href=\""https://github.com/seanhaydongriffin/UP-Test-Automation-Core-Coverage/releases/download/v0.1/UP.Test.Automation.Core.Coverage.portable.exe\"">Click to update this page</a>" & @CRLF

	if $confluence_html = False Then

		$html = $html &		"</body>" & @CRLF & _
							"</html>" & @CRLF
	EndIf

	FileDelete(@ScriptDir & "\html_report.html")
	FileWrite(@ScriptDir & "\html_report.html", $html)
EndFunc


Func Update_Confluence_Page($url, $jira_username, $jira_password, $space_key, $ancestor_key, $page_key, $page_title, $page_body)

	_ConfluenceSetup()
	_ConfluenceDomainSet($url)
	_ConfluenceLogin($jira_username, $jira_password)
	_ConfluenceUpdatePage($space_key, $ancestor_key, $page_key, $page_title, $page_body)
	_ConfluenceShutdown()

EndFunc
