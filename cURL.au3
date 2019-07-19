#include-once
#Region Header
#cs
	Title:   		cURL UDF Library for AutoIt3
	Filename:  		cURL.au3
	Description: 	A collection of functions for cURL - a tool for transferring data with URL syntax
	Author:   		seangriffin
	Version:  		V0.3
	Last Update: 	03/04/16
	Requirements: 	AutoIt3 3.2 or higher,
					libcurl v7.24.0 or higher (libcurl.dll, libeay32.dll, libssl32.dll).
	Changelog:		---------03/04/16---------- v0.3
					Enhanced cURL_easy to accept HTTP request headers and deliver HTTP response codes, headers and data.

					---------13/02/12---------- v0.2
					Made $url the first parameter for cURL_easy().
					Replaced all cookie parameters with $cookie_action and $cookie_file.
					Added $output_type and $output_file parameters.

					---------11/02/12---------- v0.1
					Initial release.

#ce
#EndRegion Header
#Region Global Variables and Constants
Global Const $CURLOPTTYPE_LONG                   = 0
Global Const $CURLOPTTYPE_OBJECTPOINT            = 10000
Global Const $CURLOPTTYPE_FUNCTIONPOINT          = 20000
Global Const $CURLOPTTYPE_OFF_T                  = 30000
Global Const $CURLOPT_USERPWD				= 0x2715
Global Const $CURLOPT_URL 					= 0x2712
Global Const $CURLOPT_WRITEDATA 			= 0x2711
Global Const $CURLOPT_WRITEFUNCTION 		= 0x4E2B
Global Const $CURLOPT_PROGRESSFUNCTION 		= 0x4E58
Global Const $CURLOPT_NOPROGRESS 			= 0x2B
Global Const $CURLOPT_ERRORBUFFER 			= 0x271A
Global Const $CURLOPT_TRANSFERTEXT 			= 0x35
Global Const $CURL_ERROR_SIZE 				= 0x100
Global Const $CURLOPT_SSL_VERIFYPEER 		= 0x40
Global Const $CURLOPT_COOKIEFILE 			= $CURLOPTTYPE_OBJECTPOINT +31
Global Const $CURLOPT_COOKIEJAR 			= $CURLOPTTYPE_OBJECTPOINT +82
Global Const $CURLOPT_FOLLOWLOCATION 		= $CURLOPTTYPE_LONG +52
Global Const $CURLOPT_POSTFIELDS 			= $CURLOPTTYPE_OBJECTPOINT + 15
Global Const $CURLOPT_POST 					= $CURLOPTTYPE_LONG +47
Global Const $CURLOPT_POSTFIELDSIZE 		= $CURLOPTTYPE_LONG +60
Global Const $CURLOPT_HTTPHEADER 			= $CURLOPTTYPE_OBJECTPOINT +23
Global Const $CURLOPT_HEADERFUNCTION        = $CURLOPTTYPE_FUNCTIONPOINT + 79

Global Const $CURLINFO_STRING                    = 0x100000
Global Const $CURLINFO_LONG                      = 0x200000
Global Const $CURLINFO_DOUBLE                    = 0x300000
Global Const $CURLINFO_SLIST                     = 0x400000
Global Const $CURLINFO_MASK                      = 0x0fffff
Global Const $CURLINFO_TYPEMASK                  = 0xf00000

Global Const $CURLINFO_EFFECTIVE_URL             =           $CURLINFO_STRING + 1
Global Const $CURLINFO_RESPONSE_CODE             =             $CURLINFO_LONG + 2
Global Const $CURLINFO_TOTAL_TIME                =           $CURLINFO_DOUBLE + 3
Global Const $CURLINFO_NAMELOOKUP_TIME           =           $CURLINFO_DOUBLE + 4
Global Const $CURLINFO_CONNECT_TIME              =           $CURLINFO_DOUBLE + 5
Global Const $CURLINFO_PRETRANSFER_TIME          =           $CURLINFO_DOUBLE + 6
Global Const $CURLINFO_SIZE_UPLOAD               =           $CURLINFO_DOUBLE + 7
Global Const $CURLINFO_SIZE_DOWNLOAD             =           $CURLINFO_DOUBLE + 8
Global Const $CURLINFO_SPEED_DOWNLOAD            =           $CURLINFO_DOUBLE + 9
Global Const $CURLINFO_SPEED_UPLOAD              =           $CURLINFO_DOUBLE + 10
Global Const $CURLINFO_HEADER_SIZE               =             $CURLINFO_LONG + 11
Global Const $CURLINFO_REQUEST_SIZE              =             $CURLINFO_LONG + 12
Global Const $CURLINFO_SSL_VERIFYRESULT          =             $CURLINFO_LONG + 13
Global Const $CURLINFO_FILETIME                  =             $CURLINFO_LONG + 14
Global Const $CURLINFO_CONTENT_LENGTH_DOWNLOAD   =           $CURLINFO_DOUBLE + 15
Global Const $CURLINFO_CONTENT_LENGTH_UPLOAD     =           $CURLINFO_DOUBLE + 16
Global Const $CURLINFO_STARTTRANSFER_TIME        =           $CURLINFO_DOUBLE + 17
Global Const $CURLINFO_CONTENT_TYPE              =           $CURLINFO_STRING + 18
Global Const $CURLINFO_REDIRECT_TIME             =           $CURLINFO_DOUBLE + 19
Global Const $CURLINFO_REDIRECT_COUNT            =             $CURLINFO_LONG + 20
Global Const $CURLINFO_PRIVATE                   =           $CURLINFO_STRING + 21
Global Const $CURLINFO_HTTP_CONNECTCODE          =             $CURLINFO_LONG + 22
Global Const $CURLINFO_HTTPAUTH_AVAIL            =             $CURLINFO_LONG + 23
Global Const $CURLINFO_PROXYAUTH_AVAIL           =             $CURLINFO_LONG + 24
Global Const $CURLINFO_OS_ERRNO                  =             $CURLINFO_LONG + 25
Global Const $CURLINFO_NUM_CONNECTS              =             $CURLINFO_LONG + 26
Global Const $CURLINFO_SSL_ENGINES               =            $CURLINFO_SLIST + 27
Global Const $CURLINFO_COOKIELIST                =            $CURLINFO_SLIST + 28
Global Const $CURLINFO_LASTSOCKET                =             $CURLINFO_LONG + 29
Global Const $CURLINFO_FTP_ENTRY_PATH            =           $CURLINFO_STRING + 30
Global Const $CURLINFO_REDIRECT_URL              =           $CURLINFO_STRING + 31
Global Const $CURLINFO_PRIMARY_IP                =           $CURLINFO_STRING + 32
Global Const $CURLINFO_APPCONNECT_TIME           =           $CURLINFO_DOUBLE + 33
Global Const $CURLINFO_CERTINFO                  =            $CURLINFO_SLIST + 34
Global Const $CURLINFO_CONDITION_UNMET           =             $CURLINFO_LONG + 35
Global Const $CURLINFO_RTSP_SESSION_ID           =           $CURLINFO_STRING + 36
Global Const $CURLINFO_RTSP_CLIENT_CSEQ          =             $CURLINFO_LONG + 37
Global Const $CURLINFO_RTSP_SERVER_CSEQ          =             $CURLINFO_LONG + 38
Global Const $CURLINFO_RTSP_CSEQ_RECV            =             $CURLINFO_LONG + 39
Global Const $CURLINFO_PRIMARY_PORT              =             $CURLINFO_LONG + 40
Global Const $CURLINFO_LOCAL_IP                  =           $CURLINFO_STRING + 41
Global Const $CURLINFO_LOCAL_PORT                =             $CURLINFO_LONG + 42

Global Const $CURLINFO_HTTP_CODE                 = $CURLINFO_RESPONSE_CODE
Global $hDll_Libcurl
Global $pWriteFunc
Global $pHeaderFunc
Global $binary_output 						= 0
Global $response_data
Global $response_headers
#EndRegion Global Variables and Constants
#Region Core functions
; #FUNCTION# ;===============================================================================
;
; Name...........:	cURL_initialise()
; Description ...:	Initialises cURL.
; Syntax.........:	cURL_initialise()
; Parameters ....:
; Return values .:
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	Must be executed prior to any other cURL functions.
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
func cURL_initialise()

	; Load and initialize curl
	$hDll_Libcurl = DllOpen(@ScriptDir & "\libcurl.dll")
	$pWriteFunc = DllCallbackRegister ("_WriteFunc", "uint", "ptr;uint;uint;ptr")
	$pHeaderFunc = DllCallbackRegister ("_HeaderFunc", "uint", "ptr;uint;uint;ptr")
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	cURL_easy()
; Description ...:	Executes a cURL easy-session on a URL and returns the downloaded data.
; Syntax.........:	cURL_easy($url, $cookie_file = "", $cookie_action = 0, $output_type = 0, $output_file = "", $request_headers = "", $request_data = "", $ssl_verifypeer = 0, $noprogress = 1, $followlocation = 0)
; Parameters ....:	$url				- the URL you are requesting.
;					$cookie_file		- Optional: The name of a cookie file to include with the request.
;											"" = do not include a cookie
;					$cookie_action		- Optional: The actions to perform with $cookie.
;											0 = do nothing
;											1 = read the cookie
;											2 = write to the cookie
;											4 = delete the cookie, if it exists, before writing
;					$output_type		- Optional: the type of output for the HTTP response.
;											0 = output as text
;											1 = output as binary
;					$output_file		- Optional: the name of a file to output to.
;					$request_headers	- Optional: a semi-colon separated list of HTTP request headers (i.e. "Content-Type: application/xml").
;					$request_data		- Optional: the full data to post in a HTTP POST operation (i.e. URL parameters, XML data, etc).
;					$ssl_verifypeer		- Optional: determines whether the authenticity of the peer's certificate is verified.
;											1 = verify
;											0 = do not verify
;					$noprogress			- Optional: output the cURL progress meter?
;											1 = no progress meter
;											0 = progress meter
;					$followlocation		- Optional: follow any Location: header the server sends as part of an HTTP header?
;											0 = don't follow the location header
;											1 = follow the location header
;					$user_password		- Optional: a user and password combination for the request
;											"" = no username and password
; Return values .: 	On Success			- Returns the HTTP response data as the following array:
;											$response[0] = the HTTP response code
;											$response[1] = the HTTP response headers
;											$response[2] = the HTTP response data
;                 	On Failure			- Returns nothing.
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that cURL_initialise() has been executed.
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
func cURL_easy($url, $cookie_file = "", $cookie_action = 0, $output_type = 0, $output_file = "", $request_headers = "", $request_data = "", $ssl_verifypeer = 0, $noprogress = 1, $followlocation = 0, $user_password = "")

	; curl_easy_init

	$hCurlHandle = DllCall($hDll_LibCurl, "ptr:cdecl", "curl_easy_init")
	$hCurlHandle = $hCurlHandle[0]

	; request headers

	Local $slist = cURL_slist_append(0, $request_headers);
;	$slist = cURL_slist_append($slist, "Content-Type: application/xml");

	; curl_easy_setopt

	$binary_output = $output_type

	if $cookie_action > 0 Then

		$cookie_struct = DllStructCreate("char[256]")
		DllStructSetData($cookie_struct, 1, $cookie_file)

		if $cookie_action = 4 or $cookie_action = 6 or $cookie_action = 7 Then

			FileDelete($cookie_file)
		EndIf

		; if read cookie action
		if $cookie_action = 1 or $cookie_action = 3 or $cookie_action = 7 Then

			cURL_easy_setopt($hCurlHandle, $CURLOPT_COOKIEFILE, DllStructGetPtr($cookie_struct))
		EndIf

		; if write cookie action
		if $cookie_action = 2 or $cookie_action = 3 or $cookie_action = 6 Then

			cURL_easy_setopt($hCurlHandle, $CURLOPT_COOKIEJAR, DllStructGetPtr($cookie_struct))
		EndIf
	EndIf

	if StringLen($request_data) > 0 Then

		$request_data_size = StringLen($request_data)
		$request_data_struct = DllStructCreate("char[" & $request_data_size & "]")
		DllStructSetData($request_data_struct, 1, $request_data)
		cURL_easy_setopt($hCurlHandle, $CURLOPT_POSTFIELDS, DllStructGetPtr($request_data_struct))
	EndIf

	$url_struct = DllStructCreate("char[256]")
	DllStructSetData($url_struct, 1, $url)
	cURL_easy_setopt($hCurlHandle, $CURLOPT_URL, DllStructGetPtr($url_struct))

	if StringLen($user_password) > 0 Then

		$user_password_struct = DllStructCreate("char[256]")
		DllStructSetData($user_password_struct, 1, $user_password)
		cURL_easy_setopt($hCurlHandle, $CURLOPT_USERPWD, DllStructGetPtr($user_password_struct))
	EndIf

	cURL_easy_setopt($hCurlHandle, $CURLOPT_NOPROGRESS, $noprogress)
	cURL_easy_setopt($hCurlHandle, $CURLOPT_WRITEFUNCTION, DllCallbackGetPtr($pWriteFunc))
	cURL_easy_setopt($hCurlHandle, $CURLOPT_HEADERFUNCTION, DllCallbackGetPtr($pHeaderFunc))

	$CURL_ERROR = DllStructCreate("char[" & $CURL_ERROR_SIZE + 1 & "]")
	cURL_easy_setopt($hCurlHandle, $CURLOPT_ERRORBUFFER, DllStructGetPtr($CURL_ERROR))

	cURL_easy_setopt($hCurlHandle, $CURLOPT_SSL_VERIFYPEER, $ssl_verifypeer)
	cURL_easy_setopt($hCurlHandle, $CURLOPT_FOLLOWLOCATION, $followlocation)

	if StringLen($request_headers) > 0 Then

		cURL_easy_setopt($hCurlHandle, $CURLOPT_HTTPHEADER, $slist)
	EndIf



;	if $request_type = 1 Then

;		DllCall($hDll_LibCurl, "uint:cdecl", "curl_easy_setopt", "ptr", $hCurlHandle, "uint", $CURLOPT_POST, "int", 1)
;	EndIf

	; curl_easy_perform

	; initialise $response_data to be either binary or string output

	if $binary_output = 0 Then

		$response_data = ""
	Else

		$response_data = Binary('')
	EndIf

	$response_headers = ""

	$nPerform = DllCall($hDll_LibCurl, "uint:cdecl", "curl_easy_perform", "ptr", $hCurlHandle)

	$nPerform = $nPerform[0]
	If $nPerform <> 0 Then
		; libcurl reported an error
		ConsoleWrite("! " & DllStructGetData($CURL_ERROR, 1) & @CRLF)
	EndIf

	dim $option = $CURLINFO_RESPONSE_CODE
	dim $parameter

    Local $asTypes[4] = ["ptr", "long", "double", "ptr"]
    Local $iParamType = Int($option / 0x100000) - 1
    Local $tCURLSTRUCT_INFO = DllStructCreate($asTypes[$iParamType])
    Local $aRes = DllCall($hDll_LibCurl, "int:cdecl", "curl_easy_getinfo", "ptr", $hCurlHandle, "int", $option, "ptr", DllStructGetPtr($tCURLSTRUCT_INFO))
    If @error Then Return SetError(2, @extended, -1)
    Switch $iParamType
        Case 0
            Local $tCURLSTRUCT_INFOBUFFER = DllStructCreate("char[256]", DllStructGetData($tCURLSTRUCT_INFO, 1))
            $parameter = DllStructGetData($tCURLSTRUCT_INFOBUFFER, 1)
        Case Else
            $parameter = DllStructGetData($tCURLSTRUCT_INFO, 1)
    EndSwitch


	; Cleanup
	DllCall($hDll_LibCurl, "none:cdecl", "curl_easy_cleanup", "ptr", $hCurlHandle)

	; If file output is required
	if StringLen($output_file) > 0 Then

		$output_file_handle = FileOpen($output_file,2+16)
		FileWrite($output_file_handle, $response_data)
		FileClose($output_file_handle)
	EndIf

	Local $response[3]

	$response[0] = $parameter
	$response[1] = $response_headers
	$response[2] = $response_data

	return $response

EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	cURL_cleanup()
; Description ...:	Cleans up cURL.
; Syntax.........:	cURL_cleanup()
; Parameters ....:
; Return values .:
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that cURL_initialise() has been executed.
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
func cURL_cleanup()

	DllCallbackFree ($pWriteFunc)
	DllCallbackFree ($pHeaderFunc)
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_WriteFunc()
; Description ...:	This Callback function recieves the data downloaded with cURL.
; Syntax.........:	_WriteFunc ($ptr,$nSize,$nMemb,$pStream)
; Parameters ....:	$ptr		- TBD.
;					$nSize		- TBD.
;					$nMemb		- TBD.
;					$pStream	- TBD.
; Return values .: 	$nSize * $nMemb.
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
;
; ;==========================================================================================
Func _WriteFunc($ptr,$nSize,$nMemb,$pStream)

	Local $vData = DllStructCreate ("byte[" & $nSize*$nMemb & "]",$ptr)

	if $binary_output = 1 Then

		$response_data = $response_data & DllStructGetData($vData,1)
	Else

		$response_data = $response_data & BinaryToString(DllStructGetData($vData,1))
	EndIf

	Return $nSize*$nMemb
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_HeaderFunc()
; Description ...:	This Callback function recieves the headers downloaded with cURL.
; Syntax.........:	_HeaderFunc ($ptr,$nSize,$nMemb,$pStream)
; Parameters ....:	$ptr		- TBD.
;					$nSize		- TBD.
;					$nMemb		- TBD.
;					$pStream	- TBD.
; Return values .: 	$nSize * $nMemb.
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
;
; ;==========================================================================================
Func _HeaderFunc($ptr,$nSize,$nMemb,$pStream)

	Local $vData = DllStructCreate ("byte[" & $nSize*$nMemb & "]",$ptr)

	if $binary_output = 1 Then

		$response_headers = $response_headers & DllStructGetData($vData,1)
	Else

		$response_headers = $response_headers & BinaryToString(DllStructGetData($vData,1))
	EndIf

	Return $nSize*$nMemb
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	cURL_slist_append()
; Description ...:	TBD.
; Syntax.........:	cURL_slist_append($slist, $append)
; Parameters ....:	$slist		- TBD.
;					$append		- TBD.
; Return values .: 	$aResult[0].
; Author ........:	ProgAndy
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
;
; ;==========================================================================================
Func cURL_slist_append($slist, $append)

	Local $aResult = DllCall($hDll_Libcurl, "ptr:cdecl", "curl_slist_append", 'ptr', $slist, 'str', $append)
	If @error Then Return SetError(1, 0, 0)
	Return $aResult[0]
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	cURL_slist_append()
; Description ...:	TBD.
; Syntax.........:	cURL_slist_append($slist, $append)
; Parameters ....:	$slist		- TBD.
;					$append		- TBD.
; Return values .: 	$aResult[0].
; Author ........:	smartee
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:	https://www.autoitscript.com/forum/topic/139325-libcurl-udf-the-multiprotocol-file-transfer-library/
; Example .......:
;
; ;==========================================================================================
Func cURL_easy_setopt($handle, $option, $parameter)

	Local $asTypes[4] = ["long", "ptr", "ptr", "int64"]
    Local $aRes = DllCall($hDll_Libcurl, "int:cdecl", "curl_easy_setopt", "ptr", $handle, "int", $option, $asTypes[Int($option / 10000)], $parameter)
	If @error Then Return SetError(1, 0, 0)
    Return $aRes[0]
EndFunc

