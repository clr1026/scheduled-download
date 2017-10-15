
#include <IE.au3>
#include <Misc.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>

;This program will open and login TWSE using given password then download files on the website at scheduled time 
 

; CLEAR FOLDER
$oDLPath = "D:\###Archive###\"

If DirGetSize($oDLPath) <> -1 Then
	DirRemove($oDLPath, 1)
EndIf
Sleep(2000)
DirCreate($oDLPath)
Sleep(2000)


; LogIn to TWSE
$oIE = _IECreate("http://dataeshop.twse.com.tw/frontend/cht/index.jsp")
$oFram = _IEFrameGetObjByName($oIE, "mainFrame")  ;_IEGetObjByName($oIE, "mainFrame")
$oElements = _IETagNameGetCollection($oFram, "input")

For $oElement in $oElements
	If $oElement.type = 'text' Then
		 _IEFormElementSetValue($oElement, "ID#######")
		 ;MsgBox(0, "setuserID", @error)
		 Sleep(1000)
	EndIf

	If $oElement.type = 'password' Then
		 _IEFormElementSetValue($oElement, "pwd#######")
		;MsgBox(0, "setPW", @error)
		Sleep(1000)
	EndIf

	If $oElement.type = 'submit' Then
		_IEAction($oElement, "click")
		Sleep(1000)
	EndIf
Next

; Go DownLoad Page
Sleep(3000)
$oIEnew = _IEAttach("台灣證券交易所"   )
$oFram2 = _IEFrameGetObjByName($oIE, "mainFrame")
$oForm = _IEFormGetObjByName($oFram2, "form1") ;_IETagNameGetCollection($oFram, "form")
$oElements = _IETagNameGetCollection($oForm, "a")

;$oPATH = "C:\###Archive###"
;$oPATH = "\\############## Archive ############\"

For $oElement in $oElements
	If StringInStr($oElement.href, "filename####1.ZIP") <> 0 Then
		_IEAction($oElement, "click")
		Sleep(2000)
		;Send("{TAB 2}")
		Send("{TAB 2}{DOWN 2}{ENTER}")
			;Send("" & $oFile & ".zip")
			;Send("" & $oFile)
		Sleep(1000)
		Send("{ENTER}")
		Sleep(3000)
		Send("{TAB 35}")
		Send("{ENTER}")
		Sleep(5000)
		ExitLoop
	EndIf
Next


For $oElement in $oElements
	If StringInStr($oElement.href, "filename####2.ZIP") <> 0 Then
		_IEAction($oElement, "click")
		Sleep(3000)
		Send("{TAB 33}{DOWN 2}{ENTER}")
		Sleep(1000)
			;Send("" & $oFile)
			;Sleep(1000)
		Send("{ENTER}")
		Sleep(3000)
		Send("{TAB 35}")
		Send("{ENTER}")
		Sleep(5000)
		ExitLoop
	EndIf
Next

For $oElement in $oElements
	If StringInStr($oElement.href, "filename####3.ZIP") <> 0 Then
		_IEAction($oElement, "click")
		Sleep(3000)
		Send("{TAB 33}{DOWN 2}{ENTER}")
		Sleep(1000)
		;Send("" & $oFile)
		;Sleep(1000)
		Send("{ENTER}")
		Sleep(3000)
		Send("{TAB 35}")
		Send("{ENTER}")
		Sleep(5000)
		ExitLoop
	EndIf
Next

For $oElement in $oElements
	If StringInStr($oElement.href, "filename####4.ZIP") <> 0 Then
		_IEAction($oElement, "click")
		Sleep(3000)
		Send("{TAB 33}{DOWN 2}{ENTER}")
		;Sleep(1000)
		;Send("" & $oFile)
		Sleep(1000)
		Send("{ENTER}")
		Sleep(3000)
		Send("{TAB 35}")
		Send("{ENTER}")
		Sleep(5000)
		ExitLoop
	EndIf
Next

$oElements = _IETagNameGetCollection($oFram2, "a")
For $oElement in $oElements
	If StringCompare( _IEPropertyGet($oElement, "innertext"), "登出") == 0 Then
		_IEAction($oElement, "click")
		Sleep(500)
		ExitLoop
	EndIf
Next
_IEQuit($oIEnew)



;Move File
$aFiles = DirGetSize($oDLPath, 1)
$oToPATH = "\\##############local Archive############\"
If $aFiles[1] <> 4 Then
	MsgBox(0,"", "Download Incomplete")
Else
	;MsgBox(0, "", "DownLoad Finished, exit in 10 sec", 10)
	$aFileList = _FileListToArray($oDLPath, "*.zip", 1)
	For $aFile in $aFileList
		If StringInStr($aFile, "ZIP") Then
			$oDate = StringReplace(StringMid($aFile, StringInStr($aFile, "-")+1),".zip","")
			$oYear = StringLeft($oDate, 4)
			$oMonth = StringMid($oDate, 5, 2)
			$oYMD = $oYear & "-" & $oMonth & "-" & StringRight($oDate,2)
			$nPath = $oToPATH & $oYear & "\" & StringLeft($oYMD, 7) & "\" & $oYMD & "\"

			If StringInStr($aFile, "filename####1.ZIP") <> 0 Then
				$nFile = "#######local filename 1########" & $oYMD & ".zip"
			ElseIf StringInStr($aFile, "filename####2.zip") <> 0 Then
				$nFile = "#######local filename 2########" & $oYMD & ".zip"
			ElseIf StringInStr($aFile, "filename####3.zip") <> 0 Then
				$nFile = "#######local filename 3########." & $oYMD & ".zip"
			ElseIf StringInStr($aFile, "filename####4.zip") <> 0 Then
				$nFile = "#######local filename 4########" & $oYMD & ".zip"
			EndIf

			FileMove($oDLPath & $aFile, $nPath & $nFile, $FC_OVERWRITE + $FC_CREATEPATH)
			Sleep(3000)
			;MsgBox(0,"", $nPath & @CRLF & $nFile & @CRLF & $oYMD)
		EndIf
	Next
	MsgBox(0, "", "DownLoad Finished, exit in 10 sec", 5)
EndIf

If	ProcessExists("test.exe") Then
	ProcessClose("test.exe")
EndIf



;Exit

		;Local $hIE = WinGetHandle("[Class:IEFrame]")
		;Local $hCtrl = ControlGetHandle($hIE, "","[Class:DirectUIHWND; INSTANCE:1]")
		;ControlFocus($hIE, "", $hCtrl)
