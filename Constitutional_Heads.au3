
#include <file.au3>
#include <ClipBoard.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include <DateTimeConstants.au3>
#include <ProgressConstants.au3>
#include <objDictonary.au3>
#include <word.au3>

Dim $yorno = 7
Dim $szDrive, $szDir, $szFName, $szExt, $aFile, $cInputFileName, $cInputFile, $cInputFileText

;~ Global $cInputFolderDefault = "U:\Constitutional Heads\L Files"
Global $cInputFolderDefault = "E:\CR\OC"
;~ Global $cOutputFolderDefault = "U:\Constitutional Heads\Output"
Global $cOutputFolderDefault = "E:\RECSCAN\TofA"
Global $cInputFolder, $cOutputFolder

Global $tipmsg = "PLEASE WAIT..."

Dim $Date, $DateSelected, $ValidDate, $msg, $LocalDate, $Button_1, $progressbar, $inputFolder, $outputFolder

; create GUI and tabs

GUICreate("Constitutional Heads Program v0.9", 350, 300)
$tab = GUICtrlCreateTab(5, 5, 340, 290)

; tab 0
$tab0 = GUICtrlCreateTabItem("Main")
GUICtrlCreateLabel("Choose Date Below:", 15, 40, 300)
$LocalDate = _Date_Time_GetLocalTime()
$Date = GUICtrlCreateMonthCal(_Date_Time_SystemTimeToDateStr($LocalDate, 1), 65, 70, 220, 140, $MCS_NOTODAY)
$DateSelected = GUICtrlCreateLabel("Date Selected: " & _Date_Time_SystemTimeToDateStr($LocalDate, 1), 15, 220, 300)
$Button_1 = GUICtrlCreateButton("Process Heads", 115, 240, 120)
$progressbar = GUICtrlCreateProgress(70, 275, 210, 10, $PBS_SMOOTH)

; tab 1
$tab1 = GUICtrlCreateTabItem("Settings")
GUICtrlCreateLabel("Input Folder:", 15, 40, 300)
$inputFolder = GUICtrlCreateInput("", 15, 65, 320, 20)
$cInputFolder = GetInputOutput("input", $cInputFolderDefault)
GUICtrlSetData($inputFolder, $cInputFolder)

GUICtrlCreateLabel("Output Folder:", 15, 100, 300)
$outputFolder = GUICtrlCreateInput("", 15, 125, 320, 20)
$cOutputFolder = GetInputOutput("output", $cOutputFolderDefault)
GUICtrlSetData($outputFolder, $cOutputFolder)

$Default_Button = GUICtrlCreateButton("Default", 15, 160, 75)
$Apply_Button = GUICtrlCreateButton("Apply", 100, 160, 75)


GUICtrlCreateTabItem("") ; end tabitem definition

GUISetState()

; Run the GUI until the dialog is closed

While 1
	$msg = GUIGetMsg()
	Switch $msg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Default_Button
			$cInputFolder = $cInputFolderDefault
			GUICtrlSetData($inputFolder, $cInputFolder)
			$cOutputFolder = $cOutputFolderDefault
			GUICtrlSetData($outputFolder, $cOutputFolder)
		Case $Apply_Button
			$cInputFolder = GUICtrlRead($inputFolder)
			$cInputFolder = StringRegExpReplace($cInputFolder, '\\* *$', '') ; strip trailing \ and spaces
			If Not FileExists($cInputFolder) Then
				MsgBox(16, "Input folder invalid", "Input folder does not exists. Enter a valid input folder.")
			Else
				If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\ConstitutionalHeads", "input", "REG_SZ", $cInputFolder) Then
					MsgBox(16, "Input folder could not be saved", "The input folder could not be saved, Error #" & @error)
				EndIf
			EndIf
			GUICtrlSetData($inputFolder, $cInputFolder)

			$cOutputFolder = GUICtrlRead($outputFolder)
			$cOutputFolder = StringRegExpReplace($cOutputFolder, '\\* *$', '') ; strip trailing \ and spaces
			If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\ConstitutionalHeads", "output", "REG_SZ", $cOutputFolder) Then
				MsgBox(16, "Output folder could not be saved", "The output folder could not be saved, Error #" & @error)
			EndIf
			GUICtrlSetData($outputFolder, $cOutputFolder)
		Case $Button_1
			Dim $aMonths[13] = ["00", "JA", "FE", "MR", "AP", "MY", "JN", "JY", "AU", "SE", "OC", "NO", "DE"]
			Dim $cDay = GUICtrlRead($Date)
			Dim $nMonth = Number(StringRegExpReplace($cDay, '(\d{4})/(\d{2})/(\d{2})', '$2'))
			Dim $cTempDay = StringRegExpReplace($cDay, '(\d{4})/(\d{2})/(\d{2})', '$3')
			Dim $cCaptureFileName = "L" & $cTempDay & $aMonths[$nMonth] & "7.100"
			$cInputFolder = StringRegExpReplace($cInputFolder, '\\$', '')
			$cInputFileName = $cInputFolder & "\" & $cCaptureFileName

			If FileExists($cInputFileName) Then

				$cInputFileText = FileRead($cInputFileName)

				; preprocessing

				$cInputFileText = StringRegExpReplace($cInputFileText, '\r?\n', @CRLF) ; make end of line consistent
				$cInputFileText = StringRegExpReplace($cInputFileText, '~', ChrW(0x07)) ; precedence code
				$cInputFileText = StringRegExpReplace($cInputFileText, '\x{1A}', '') ; remove any stray end of files
				$cInputFileText = StringRegExpReplace($cInputFileText, '(?<!\x{0A})\x{07}(I|S|F)', @CRLF & ChrW(0x07) & '$1') ; make sure each bell I, bell S and bell F is on its own line
				$cInputFileText = StringRegExpReplace($cInputFileText, '\x{AE}MD[0-9A-Z]{2,2}\x{AF}', '') ; strip Xywrite modes
				$cInputFileText = StringRegExpReplace($cInputFileText, '\x{AE}IP.*?\x{AF}', '') ; strip Xywrite modes
				$cInputFileText = StringRegExpReplace($cInputFileText, '\x{AE}PT.*?\x{AF}', '') ; strip Xywrite modes
				$cInputFileText = StringRegExpReplace($cInputFileText, "(\r\n)(\1)+", "\1") ; strip double newline

;~ 				   ConsoleWrite($cInputFileText)
				; split the variable into an array of lines

				Dim $aRecords = StringSplit($cInputFileText, @CRLF, $STR_ENTIRESPLIT)
;~ 				   _ArrayDisplay($aRecords)


				; temp location, change this folder in final script

				If Not FileExists($cOutputFolder) Then
					DirCreate($cOutputFolder)
				EndIf

				; initialize variables

				Local $hrByDict = returnHeadMultiArray($aRecords)
				createWordDoc($hrByDict)

				GUICtrlSetData($progressbar, 100)
				Sleep(2000)
				GUICtrlSetData($progressbar, 0)

			Else
				MsgBox(16, "File does not exist", $cInputFileName & ' does NOT exist. Try selecting another date.')
			EndIf
		Case $GUI_EVENT_PRIMARYUP
			$ValidDate = _DateIsValid(GUICtrlRead($Date))
			If $ValidDate Then
				GUICtrlSetData($DateSelected, "Date Selected: " & GUICtrlRead($Date))
			EndIf
	EndSwitch
WEnd

; function to get input or output values from registry if they exist
Func GetInputOutput($IorO, $DefaultFolder)

	Dim $inputreg, $outputreg

	If $IorO = "input" Then
		$inputreg = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\ConstitutionalHeads", "input")
		If $inputreg = "" Then
			RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\ConstitutionalHeads", "input", "REG_SZ", $DefaultFolder)
			Return $DefaultFolder
		Else
			Return $inputreg
		EndIf
	Else
		$outputreg = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\ConstitutionalHeads", "output")
		If $outputreg = "" Then
			RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\ConstitutionalHeads", "output", "REG_SZ", $DefaultFolder)
			Return $DefaultFolder
		Else
			Return $outputreg
		EndIf
	EndIf

EndFunc   ;==>GetInputOutput

; function to get filename for amendment
Func GetSAFilename($cText)
	Return StringRegExpReplace($cText, '^.+?SA (.+?)\..+$', 'SA $1')
EndFunc   ;==>GetSAFilename

;~ function to create dictionary of H.R.'s and By's
Func returnHeadMultiArray($lineArray)
	Local $head_dict = _ObjDictCreate()

	For $ii = 1 To $lineArray[0] Step 1
		Local $myHitArray = StringRegExp($lineArray[$ii], 'I21(H\.R\..*?\d+\.)', 1)
		If @error == 2 Then
			ExitLoop
		ElseIf @error == 0 Then
			$lineArray[$ii - 1] = StringRegExpReplace($lineArray[$ii - 1], '(?:I17|I16)(.*?)T4(.*?)T1.*', '$1$2:', 1)
;~ 			ConsoleWrite($myHitArray[0]&"     "&$lineArray[$ii-1]&@CRLF)
			_ObjDictAdd($head_dict, $myHitArray[0], $lineArray[$ii - 1])
		EndIf

	Next
;~ 	_ObjDictList($head_dict)
	Return $head_dict
EndFunc   ;==>returnHeadMultiArray

Func createWordDoc($recordHash)
	$oWordApp = _Word_Create(0, True)
	$oFinalWordApp = _Word_Create(0, True)
	$oDoc_2 = _Word_DocAdd($oFinalWordApp)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_DocAdd Final Doc", "Error creating a new Word document." _
				 & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	Dim $progpercent = 10
	Dim $progincrement = Round(_ObjDictCount($recordHash) / $progpercent)
	GUICtrlSetData($progressbar, 0)

	For $myHR In $recordHash
;~ 		ConsoleWrite("My key: "&$myHR&" My value: "&_ObjDictGetValue($recordHash, $myHR)&@CRLF)
		$oDoc = _Word_DocAdd($oWordApp, $wdNewBlankDocument, @ScriptDir & "\CASTemplate.dotx")
		If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_DocAdd Template", "Error creating a new Word document from template." _
				 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<ByText>", _ObjDictGetValue($recordHash, $myHR))
		_Word_DocFindReplace($oDoc, "<BillNo>", $myHR)

		$oRange = _Word_DocRangeSet($oDoc, -1, Default, Default, $wdStory, 1)
		$oRange.Copy

		$oFinalRange = _Word_DocRangeSet($oDoc_2, 0, $wdStory, 1, $wdCharacter, -1)
		$oFinalRange.PasteAndFormat(16)
		_ObjDictDeleteKey($recordHash, $myHR)
		If _ObjDictCount($recordHash) <> 0 Then
			$oFinalRange = _Word_DocRangeSet($oDoc_2, $oFinalRange, $wdStory, 1, -1, Default)
			$oFinalRange.InsertBreak()
		EndIf
		GUICtrlSetData($progressbar, (100 - ($progincrement * _ObjDictCount($recordHash))))
		_Word_DocClose($oDoc)
	Next

	$oFinalWordApp.Visible = True
;~ 	_Word_DocSaveAs($oDoc_2, @ScriptDir & "\_Word_Test2.doc")
;~ 	_Word_DocClose($oDoc_2)
	Return
EndFunc   ;==>createWordDoc

