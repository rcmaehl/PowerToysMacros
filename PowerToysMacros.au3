#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Assets\PowerToysMacros.ico
#AutoIt3Wrapper_Outfile_x64=PowerToysMacros.exe
#AutoIt3Wrapper_Compile_Both=N
#AutoIt3Wrapper_UseX64=Y
#AutoIt3Wrapper_Res_Comment=https://www.fcofix.org
#AutoIt3Wrapper_Res_Description=Definable Powertoys Macros
#AutoIt3Wrapper_Res_Fileversion=0.1.0.0
#AutoIt3Wrapper_Res_ProductName=PowerToysMacros
#AutoIt3Wrapper_Res_ProductVersion=0.1.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Robert Maehl, using LGPL 3 License
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Compatibility=Win8,Win81,Win10
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
	MsgBox(0, "PowerToysMacro", $sVersion)
EndFunc

Func ProcessCMDLine()
	Local $iParams = $CmdLine[0]

	If $iParams > 0 Then
		Switch $CmdLine[1]
			Case "install"
				RunSetup()
			Case "uninstall"
				RunRemoval()
			Case Else
				HandleMacro($CmdLine)
				RunUpdateCheck()
		EndSwitch
	EndIf
EndFunc

Func HandleMacro($aCmdLine)
	Local $aInput
	Local $sCommand

	$aCmdLine[1] = StringReplace($aCmdLine[1], "macro:", "")
	If $aCmdLine[1] = "" Then
		About()
	Else
		$aInput = StringSplit($aCmdLine[1], " ", $STR_NOCOUNT)
		IniReadSection(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0])
		If @error Then
			MsgBox(0, "No Macro", "Missing Macro Definition for: " & $aInput[0])
			Return
		Else
			$sCommand = IniRead(@LocalAppDataDir & "\PowerToysMacros\Macros.ini", $aInput[0], "Command", Null)
			If $sCommand = Null Then
				MsgBox(0, "No Command", "Missing Macro Command for: " & $aInput[0])
				Return
			ElseIf $sCommand = "%param0%" Then
				MsgBox(0, "No", "This is dangerous. Don't do this.")
				Return
			Else
				$sCommand = StringReplace($sCommand, "%param0%", _ArrayToString($aInput, " ", 1))
				MsgBox(0, "CMD", $sCommand)
			EndIf
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
				_Translate($aMUI[1], "Test Build?"), _
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