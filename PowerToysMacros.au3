#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Assets\PowerToysMacros.ico
#AutoIt3Wrapper_Outfile_x64=PowerToysMacros.exe
#AutoIt3Wrapper_Compile_Both=N
#AutoIt3Wrapper_UseX64=Y
#AutoIt3Wrapper_Res_Comment=https://www.fcofix.org
#AutoIt3Wrapper_Res_Description=Definable Powertoys Macros
#AutoIt3Wrapper_Res_Fileversion=0.4.0.0
#AutoIt3Wrapper_Res_ProductName=PowerToysMacros
#AutoIt3Wrapper_Res_ProductVersion=0.4.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Robert Maehl, using LGPL 3 License
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Compatibility=Win10
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7 -v1 -v2 -v3
#AutoIt3Wrapper_Run_Tidy=n
#Tidy_Parameters=/tc 0 /serc /scec
#AutoIt3Wrapper_Run_Au3Stripper=Y
#Au3Stripper_Parameters=/so
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("TrayIconHide", 1)
Opt("TrayAutoPause", 0)

#include <Date.au3>
#include <Misc.au3>
#include <Array.au3>
#include <String.au3>
#include <AutoItConstants.au3>
#include <StringConstants.au3>

#include "Includes\_Translation.au3"

Global $sVersion

If @Compiled Then
	$sVersion = FileGetVersion(@ScriptFullPath)
Else
	$sVersion = "x.x.x.x"
EndIf

ProcessCMDLine()

Func About()
	MsgBox($MB_OK + $MB_ICONINFORMATION + $MB_TOPMOST, _
		"PowerToysMacros", _
		$sVersion, _
		10)
EndFunc

Func ProcessCMDLine()
	Local $iParams = $CmdLine[0]

	If $iParams > 0 Then
		Switch $CmdLine[1]
			Case "/uninstall"
				RunRemoval()
			Case Else
				HandleMacro($CmdLine)
				RunUpdateCheck()
		EndSwitch
	Else
		RunSetup()
		MsgBox($MB_OK + $MB_ICONINFORMATION + $MB_TOPMOST, _
			"PowerToysMacros", _
			_Translate($aMUI[1], "Install Completed Successfully. Uninstall using Programs and Features"), _
			10)
	EndIf
EndFunc

Func HandleMacro($aCmdLine)

	Local $sData
	Local $sKind
	Local $sMode
	Local $sTemp
	Local $sVerb
	Local $sAlias
	Local $aInput
	Local $vSpread
	Local $aMatches
	Local $sReceiver

	$aCmdLine[1] = StringReplace($aCmdLine[1], "macro:", "")
	If $aCmdLine[1] = "" Then
		About()
	Else
		$aInput = StringSplit($aCmdLine[1], " ", $STR_NOCOUNT)
		IniReadSection(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0])
		If @error Then
			MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
				_Translate($aMUI[1], "No Macro"), _
				_Translate($aMUI[1], "Missing Macro Definition for: " & $aInput[0]), _
				10)
			Return
		Else
			Do
				; Process Aliases
				$sAlias = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Alias", "")
				If $sAlias = "" Then
					;;;
				Else
					$aCmdLine[1] = "macro:" & StringReplace($aCmdLine[1], $aInput[0], $sAlias, 1)
					HandleMacro($aCmdLine) ; REWRITE?
					Return
				EndIf

				; Process Data Handling for a Macro, Prevent Directly Running Input
				$sData = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Data", "")
				If $sData = "" Then
					MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
						_Translate($aMUI[1], "No Data"), _
						_Translate($aMUI[1], "Missing Macro Data for: " & $aInput[0]), _
						10)
					Return
				ElseIf $sData = "{0}" Then
					MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
						_Translate($aMUI[1], "No."), _
						_Translate($aMUI[1], "This is dangerous. Don't do this."), _
						10)
					Return
				EndIf

				; Replace {#} with Appropriate Parameters
				$sData = StringReplace($sData, "{0}", _ArrayToString($aInput, " ", 1))
				If Ubound($aInput) > 1 Then
					For $iLoop = 1 To Ubound($aInput) - 1 Step 1
						$sData = StringReplace($sData, "{" & $iLoop & "}", $aInput[$iLoop])
					Next

					; Empty Invalid {#} Entries if they exist
					$aMatches = StringRegExp($sData, "{\d+}", $STR_REGEXPARRAYGLOBALMATCH)
					If Not @error Then
						For $iLoop = 0 To UBound($aMatches) - 1 Step 1
							$sData = StringReplace($sData, $aMatches[$iLoop], "")
						Next
					EndIf

					; Check if {#...#} Exists, Replace with Appropriate Parameters
					$aMatches = StringRegExp($sData, "{\d+...\d+}", $STR_REGEXPARRAYGLOBALMATCH)
					If Not @error Then
						For $iLoop = 0 To UBound($aMatches) - 1 Step 1
							$vSpread = $aMatches[$iLoop]
							$vSpread = StringReplace($vSpread, "{", "")
							$vSpread = StringReplace($vSpread, "}", "")
							$vSpread = StringSplit($vSpread, "...", $STR_ENTIRESPLIT+$STR_NOCOUNT)
							$sTemp = ""
							For $iLoop2 = $vSpread[0] To $vSpread[1] Step 1
								If $iLoop2 >= UBound($aInput) Then ContinueLoop
								$sTemp &= $aInput[$iLoop2]
							Next
							$sData = StringReplace($sData, $aMatches[$iLoop], $sTemp)
						Next
					EndIf
				EndIf

				; Handle Appropriate Macro Type
				Switch IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Type", "")
					Case "Command"
						$sVerb = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Verb", "")
						Switch $sVerb
							Case "edit", "find", "open", "print", "properties", "runas"
								;;;
							Case ""
								$sVerb = Default
							Case Else
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "Invalid Verb"), _
									_Translate($aMUI[1], "Invalid Verb Type for: " & $aInput[0]), _
								10)
								Return
						EndSwitch
						ShellExecute($sData, Default, Default, $sVerb)
					Case "RawText"
						$sMode = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Mode", 2)
						Switch $sMode
							Case -4 To 4
								Opt("WinTitleMatchMode", $sMode)
							Case Else
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "Invalid RawText Mode"), _
									_Translate($aMUI[1], "Invalid RawText Mode for: " & $aInput[0]), _
									10)
								Return
						EndSwitch
						$sReceiver = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Receiver", "")
						If $sReceiver <> "" Then
							WinActivate($sReceiver)
							Sleep(100)
						EndIf
						Send($sData, $SEND_RAW)
					Case "SpecialText"
						$sMode = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Mode", 2)
						Switch $sMode
							Case -4 To -1, 1 To 4
								Opt("WinTitleMatchMode", $sMode)
							Case Else
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "Invalid SpecialText Mode"), _
									_Translate($aMUI[1], "Invalid SpecialText Mode for: " & $aInput[0]), _
									10)
								Return
						EndSwitch
						$sReceiver = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Receiver", "")
						If $sReceiver <> "" Then
							WinActivate($sReceiver)
							Sleep(100)
						EndIf
						Send($sData)
					Case "WaitFor"
						$sKind = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Kind", "")
						Switch $sKind
							Case "Process"
								; TODO: Validate Process Name
								If StringLeft($sData, 4) <> ".exe" Then
									MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
										_Translate($aMUI[1], "Invalid WaitFor Data"), _
										_Translate($aMUI[1], "Invalid WaitFor Data for: " & $aInput[0]), _
										10)
									Return
								EndIf
								ProcessWait($sData)
							Case "Window"
								$sMode = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Mode", 2)
								Switch $sMode
									Case -4 To 4
										Opt("WinTitleMatchMode", $sMode)
									Case Else
										MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
											_Translate($aMUI[1], "Invalid WaitFor Mode"), _
											_Translate($aMUI[1], "Invalid WaitFor Mode for: " & $aInput[0]), _
											10)
										Return
								EndSwitch
								WinWait($sData)
							Case ""
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "No WaitFor Kind"), _
									_Translate($aMUI[1], "Missing WaitFor Kind for: " & $aInput[0]), _
									10)
								Return
							Case Else
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "Invalid WaitFor Kind"), _
									_Translate($aMUI[1], "Invalid WaitFor Kind for: " & $aInput[0]), _
									10)
								Return
						EndSwitch
					Case ""
						MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
							_Translate($aMUI[1], "No Type"), _
							_Translate($aMUI[1], "Missing Macro Type for: " & $aInput[0]), _
							10)
						Return
					Case Else
						MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
							_Translate($aMUI[1], "Invalid Type"), _
							_Translate($aMUI[1], "Invalid Macro Type for: " & $aInput[0]), _
							10)
						Return
				EndSwitch
				$aInput[0] = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "RunAfter", "None")
				; Cleanup Input
				$aInput[0] = StringReplace($aInput[0], "[", "")
				$aInput[0] = StringReplace($aInput[0], "]", "")
			Until $aInput[0] = "None"
		EndIf
	EndIf
EndFunc

Func RunSetup()

	If Not FileCopy(@ScriptFullPath, @LocalAppDataDir & "\PowerToysMacros\PowerToysMacros.exe", $FC_CREATEPATH+$FC_OVERWRITE) Then
		;FileWrite($hLogs[$AppFailures], _NowCalc() & " - [CRITICAL] Unable to copy application to '" & @LocalAppDataDir & "\PowerToysMacros\PowerToysMacros.exe'" & @CRLF)
		MsgBox(0, _NowCalc() & " - [CRITICAL]", "Unable to copy application to '" & @LocalAppDataDir & "\PowerToysMacros\PowerToysMacros.exe'" & @CRLF)
		Exit 29 ; ERROR_WRITE_FAULT
	EndIf
	SetAppRegistry()
	FileInstall("./LICENSE", @LocalAppDataDir & "\PowerToysMacros\License.txt")
	FileInstall("./EXAMPLES/Macros.ini", @LocalAppDataDir & "\PowerToysMacros\Macros.ini")

EndFunc

Func RunRemoval($bUpdate = False)

	Local $aPIDs

	$aPIDs = ProcessList("PowerToysMacros.exe")
	For $iLoop = 1 To $aPIDs[0][0] Step 1
		If $aPIDs[$iLoop][1] <> @AutoItPID Then ProcessClose($aPIDs[$iLoop][1])
	Next

	Local $sLocation = @LocalAppDataDir & "\PowerToysMacros\"
	Local $sHive = "HKCU"

	; App Paths
	RegDelete($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PowerToysMacros.exe")

	; App Settings
	RegDelete($sHive & "\SOFTWARE\Robert Maehl Software\PowerToysMacros")

	; URI Handler
	RegDelete($sHive & "\Software\Classes\macro")

	; Generic Program Info
	RegDelete($sHive & "\Software\Classes\PowerToysMacros")
	RegDelete($sHive & "\Software\Classes\Applications\PowerToysMacros.exe")

	; Uninstall Info
	RegDelete($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros")

	; Start Menu Shortcuts
	FileDelete(@StartupDir & "\PowerToysMacros.lnk")
	DirRemove(@ProgramsCommonDir & "\PowerToysMacros", $DIR_REMOVE)
	DirRemove(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\PowerToysMacros", $DIR_REMOVE)

	If $bUpdate Then
		FileDelete($sLocation & "*")
	Else
		Run(@ComSpec & " /c " & 'ping google.com && del /Q "' & $sLocation & '*"', "", @SW_HIDE)
		Exit
	EndIf

EndFunc

Func RunUpdateCheck($bFull = False)
	Switch _GetLatestRelease($sVersion)
		Case -1
			MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
				_Translate($aMUI[1], "PowerToysMacros Test Build?"), _
				_Translate($aMUI[1], "You're running a newer build than publicly Available!"), _
				10)
		Case 0
			If $bFull Then
				Switch @error
					Case 0
						MsgBox($MB_OK + $MB_ICONINFORMATION + $MB_TOPMOST, _
							_Translate($aMUI[1], "Up to Date"), _
							_Translate($aMUI[1], "You're running the latest build!"), _
							10)
					Case 1
						MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
							_Translate($aMUI[1], "Unable to Check for Updates"), _
							_Translate($aMUI[1], "Unable to load release data."), _
							10)
					Case 2
						MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
							_Translate($aMUI[1], "Unable to Check for Updates"), _
							_Translate($aMUI[1], "Invalid Data Received!"), _
							10)
					Case 3
						Switch @extended
							Case 0
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "Unable to Check for Updates"), _
									_Translate($aMUI[1], "Invalid Release Tags Received!"), _
									10)
							Case 1
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "Unable to Check for Updates"), _
									_Translate($aMUI[1], "Invalid Release Types Received!"), _
									10)
						EndSwitch
				EndSwitch
			EndIf
		Case 1
			If MsgBox($MB_YESNO + $MB_ICONINFORMATION + $MB_TOPMOST, _
				_Translate($aMUI[1], "PowerToysMacros Update Available"), _
				_Translate($aMUI[1], "An Update is Available, would you like to download it?"), _
				10) = $IDYES Then ShellExecute("https://fcofix.org/PowerToysMacros/releases")
	EndSwitch
EndFunc

Func SetAppRegistry()

	Local $sLocation = @LocalAppDataDir & "\PowerToysMacros\"
	Local $sHive = "HKCU"

	; App Paths
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PowerToysMacros.exe", "", "REG_SZ", $sLocation & "PowerToysMacros.exe")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PowerToysMacros.exe", "Path", "REG_SZ", $sLocation)

	; URI Handler
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PowerToysMacros.exe", "SupportedProtocols", "REG_SZ", "macro")
	RegWrite($sHive & "\Software\Classes\macro", "", "REG_SZ", "URL: macro")
	RegWrite($sHive & "\Software\Classes\macro", "URL Protocol", "REG_SZ", "")
	RegWrite($sHive & "\Software\Classes\macro\shell\open\command", "", "REG_SZ", '"' & $sLocation & 'PowerToysMacros.exe" "%1"')

	; Generic Program Info
	RegWrite($sHive & "\Software\Classes\PowerToysMacros\DefaultIcon", "", "REG_SZ", '"' & $sLocation & 'PowerToysMacros.exe",0')
	RegWrite($sHive & "\Software\Classes\PowerToysMacros\shell\open\command", "", "REG_SZ", '"' & $sLocation & 'PowerToysMacros.exe" "%1"')
	RegWrite($sHive & "\Software\Classes\Applications\PowerToysMacros.exe", "FriendlyAppName", "REG_SZ", "PowerToysMacros")
	RegWrite($sHive & "\Software\Classes\Applications\PowerToysMacros.exe\DefaultIcon", "", "REG_SZ", '"' & $sLocation & 'PowerToysMacros.exe",0')

	; Uninstall Info
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "DisplayIcon", "REG_SZ", '"' & $sLocation & 'PowerToysMacros.exe",0')
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "DisplayName", "REG_SZ", "PowerToysMacros")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "DisplayVersion", "REG_SZ", $sVersion)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "EstimatedSize", "REG_DWORD", 1536)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "InstallDate", "REG_SZ", StringReplace(_NowCalcDate(), "/", ""))
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "InstallLocation", "REG_SZ", $sLocation)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "Language", "REG_DWORD", 1033)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "NoModify", "REG_DWORD", 1)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "NoRepair", "REG_DWORD", 1)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "Publisher", "REG_SZ", "Robert Maehl Software")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "UninstallString", "REG_SZ", '"' & $sLocation & 'PowerToysMacros.exe" /uninstall')
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PowerToysMacros", "Version", "REG_SZ", $sVersion)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetLatestRelease
; Description ...: Checks GitHub for the Latest Release
; Syntax ........: _GetLatestRelease($sCurrent)
; Parameters ....: $sCurrent            - a string containing the current program version
; Return values .: Returns True if Update Available
; Author ........: rcmaehl
; Modified ......: 11/11/2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetLatestRelease($sCurrent)

	Local $dAPIBin
	Local $sAPIJSON

	$dAPIBin = InetRead("https://api.fcofix.org/repos/rcmaehl/PowerToysMacros/releases")
	If @error Then Return SetError(1, 0, 0)
	$sAPIJSON = BinaryToString($dAPIBin)
	If @error Then Return SetError(2, 0, 0)

	Local $aReleases = _StringBetween($sAPIJSON, '"tag_name":"', '",')
	If @error Then Return SetError(3, 0, 0)
	Local $aRelTypes = _StringBetween($sAPIJSON, '"prerelease":', ',')
	If @error Then Return SetError(3, 1, 0)
	Local $aCombined[UBound($aReleases)][2]

	For $iLoop = 0 To UBound($aReleases) - 1 Step 1
		$aCombined[$iLoop][0] = $aReleases[$iLoop]
		$aCombined[$iLoop][1] = $aRelTypes[$iLoop]
	Next

	Return _VersionCompare($aCombined[0][0], $sCurrent)

EndFunc   ;==>_GetLatestRelease
