#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\icon.ico
#pragma compile(ProductVersion, 1.5.2)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Приложения для инфомата для самостоятельной отметки о посещении)
#pragma compile(LegalCopyright, )
#pragma compile(ProductName, InfomatSelfChecking)
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <FontConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <Array.au3>
#include <ColorConstants.au3>
#include <GuiListView.au3>
#include <ListviewConstants.au3>
#include <GDIPlus.au3>
#include <File.au3>
#include <Timers.au3>
#include <IE.au3>
#include <AVIConstants.au3>
#include <Date.au3>
#include "Excel.au3"
#include <GuiButton.au3>
#include <AutoItConstants.au3>
#include <ScreenCapture.au3>
#include <GuiAVI.au3>



#Region ====================== Variables ======================
Local $scriptDir = @ScriptDir
Local $resourcesPath = $scriptDir & "\Resources\"
Local $printedAppointmentListPath = $scriptDir & "\Printed Appointments List\"
Local $logsPath = $scriptDir & "\Logs\"

Local $errStr = "===ERROR=== "
Local $sMailDeveloperAddress = ""
Local $iniFile = $resourcesPath & "\InfomatSelfChecking.ini"
If Not FileExists($iniFile) Then
	MsgBox($MB_ICONERROR, "Critical error!", "Cannot find the settings file:" & @CRLF & $iniFile & _
			@CRLF & @CRLF & "Please contact to developer: " & @CRLF & "Mail: " & $sMailDeveloperAddress & @CRLF & _
			"Internal phone number: 31-555")
	ToLog($errStr & "Cannot find the settings file:" & $iniFile)
	Exit
EndIf

Local $oMyError = ObjEvent("AutoIt.Error", "HandleComError")
OnAutoItExitRegister("OnExit")

Local $generalSectionName = "general"
Local $infoclinicaDB = IniRead($iniFile, $generalSectionName, "infoclinica_database_address", "")
Local $formMaxTimeWait = IniRead($iniFile, $generalSectionName, "form_max_time_wait_in_seconds", 30) * 1000
Local $showAppointmentsForm = IniRead($iniFile, $generalSectionName, "show_appointments_form", 0) = 0 ? False : True
Local $showIconsDescription = IniRead($iniFile, $generalSectionName, "show_icons_description", 0) = 0 ? False : True
Local $bDebug = IniRead($iniFile, $generalSectionName, "debug", 0) = 0 ? False : True

Local $colorSectionName = "colors"
Local $colorHeader = IniRead($iniFile, $colorSectionName, "header", 0x4e9b44)
Local $colorOkButton = IniRead($iniFile, $colorSectionName, "button", 0x4e9b44)
Local $colorOkButtonPressed = IniRead($iniFile, $colorSectionName, "button_pessed", 0x43853a)
Local $colorMainButton = IniRead($iniFile, $colorSectionName, "main_button", 0xe0e0e0)
Local $colorMainButtonPressed = IniRead($iniFile, $colorSectionName, "main_button_pressed", 0xd6d6d6)
Local $colorNameButtonSelected = IniRead($iniFile, $colorSectionName, "name_button_selected", 0x6dcbde)
Local $colorDisabled = IniRead($iniFile, $colorSectionName, "disabled", 0xdfdfdf)
Local $colorDisabledText = IniRead($iniFile, $colorSectionName, "disabled_text", 0xa5a5a5)
Local $colorText = IniRead($iniFile, $colorSectionName, "text", 0x2c3d3f)
Local $colorAlternateText = IniRead($iniFile, $colorSectionName, "alternate_text", 0xffffff)
Local $colorMainBackground = IniRead($iniFile, $colorSectionName, "main_background", 0xffffff)
Local $colorErrorTitle = IniRead($iniFile, $colorSectionName, "error_title", 0xf98d3c)

Local $fontSectionName = "font"
Local $fontName = IniRead($iniFile, $fontSectionName, "main_font_name", "Franklin Gothic")
Local $fontWeight = IniRead($iniFile, $fontSectionName, "main_font_weight", $FW_BOLD)
Local $fontQuality = IniRead($iniFile, $fontSectionName, "quality", $CLEARTYPE_QUALITY)
Local $fontNameAppointments = IniRead($iniFile, $fontSectionName, "appointments_font_name", "Franklin Gothic Book")
Local $fontWeightAppointments = IniRead($iniFile, $fontSectionName, "appointments_font_weight", $FW_NORMAL)

Local $timeBoundariesSectionName = "available_time_to_set_mark_in_minutes"
Local $timeBoundariesPast = IniRead($iniFile, $timeBoundariesSectionName, "past", 10)
Local $timeBoundariesFuture = IniRead($iniFile, $timeBoundariesSectionName, "future", 180)
Local $timeBoundariesAcceptableDifferenceBetweenAppointments = IniRead($iniFile, $timeBoundariesSectionName, _
		"acceptable_difference_between_appointments", 120)

Local $textTitleDialer = GetTextFromIni("title_dialer")
Local $textTitleNameConfirm = GetTextFromIni("title_name_confirm")
Local $textTitleNameConfirmMultiple = GetTextFromIni("title_name_confirm_multiple")
Local $textTitleAppointments = GetTextFromIni("title_appointments")
Local $sTitleWelcome = GetTextFromIni("title_welcome")
Local $textTitleNotification = GetTextFromIni("title_notification")

Local $sWelcomeTop = GetTextFromIni("welcome_top")
Local $sWelcomeBottom = GetTextFromIni("welcome_bottom")

Local $textNotificationDbNotAvailable = GetTextFromIni("notification_db_not_available")
Local $textNotificationNothingFound = GetTextFromIni("notification_nothing_found")
Local $textNotificationWrongName = GetTextFromIni("notification_wrong_name")
Local $textNotificationNoAppointmetnsForNow = GetTextFromIni("notification_no_appointmetns_for_now")
Local $textNotificationFirstVisit = GetTextFromIni("notification_first_visit")
Local $textNotificationMultiplePatientsError = GetTextFromIni("notification_need_go_to_registry")

Local $textAppointmentsMarkOk = GetTextFromIni("appointments_mark_ok")
Local $textAppointmentsMarkProblem = GetTextFromIni("appointments_mark_problem")
Local $textAppointmentsPrintOk = GetTextFromIni("appointments_print_ok")
Local $textAppointmentsPrintProblem = GetTextFromIni("appointments_print_problem")
Local $textAppointmentsWarningGeneral = GetTextFromIni("appointments_warning_general")
Local $textAppointmentsWarningCash = GetTextFromIni("appointments_warning_cash")
Local $textAppointmentsWarningTime = GetTextFromIni("appointments_warning_time")
Local $textAppointmentsWarningXray = GetTextFromIni("appointments_warning_xray")

Local $textPrintNotificationCash = GetTextFromIni("print_notification_cash")
Local $textPrintNotificationTime = GetTextFromIni("print_notification_time")
Local $textPrintNotificationXray = GetTextFromIni("print_notification_xray")
Local $textPrintMessageTimeOk = GetTextFromIni("print_message_time_ok")
Local $textPrintMessageTimeLate = GetTextFromIni("print_message_time_late")
Local $textPrintMessageFinalOk = GetTextFromIni("print_message_final_ok")
Local $textPrintMessageFinalCash = GetTextFromIni("print_message_final_cash")
Local $textPrintMessageFinalTime = GetTextFromIni("print_message_final_time")
Local $textPrintMessageFinalXray = GetTextFromIni("print_message_final_xray")
Local $textPrintMessageFinalMultiple = GetTextFromIni("print_message_final_multiple")

Local $sqlCheckEnteredNumber = "Select Distinct Cl.PCode, Cl.FirstName, Cl.MidName, Cl.BDate, "
$sqlCheckEnteredNumber &= GetTextFromIni("sql_check_entered_number", True)

Local $sqlGetAppointments = "Select Sch.SchedId, Sch.WorkDate, Sch.BHour, Sch.BMin, D.DName, Dep.DepName, R.RNum, "
$sqlGetAppointments &= GetTextFromIni("sql_get_appointments", True)

Local $sqlSetMark = "Update Schedule Set ScreenVisit = 1, ClVisit = 1, VisitTime = 'now', "
$sqlSetMark &= GetTextFromIni("sql_set_mark", True)

Local $sMailSectionName = "mail"
Local $sMailBackupServer = ""
Local $sMailBackupLogin = ""
Local $sMailBackupPassword = ""
Local $sMailBackupTo = $sMailDeveloperAddress
Local $sMailBackupSend = True
Local $sMailServer = IniRead($iniFile, $sMailSectionName, "server", $sMailBackupServer)
Local $sMailLogin = IniRead($iniFile, $sMailSectionName, "login", $sMailBackupLogin)
Local $sMailPassword = IniRead($iniFile, $sMailSectionName, "password", $sMailBackupPassword)
Local $sMailTo = IniRead($iniFile, $sMailSectionName, "to", $sMailBackupTo)
Local $sMailTitle = IniRead($iniFile, $sMailSectionName, "title", "")
Local $sMailSend = IniRead($iniFile, $sMailSectionName, "send_email", $sMailBackupSend) = 0 ? False : True
Local $sMailWorkingHoursBegins = IniRead($iniFile, $sMailSectionName, "working_hours_begins", "")
Local $sMailWorkingHoursEnds = IniRead($iniFile, $sMailSectionName, "working_hours_ends", "")
Local $sMailRegistryAddress = IniRead($iniFile, $sMailSectionName, "registry", "")

Local $sPrinterName = IniRead($iniFile, "printer", "name", "")

Local $dX = @DesktopWidth
Local $dY = @DesktopHeight
If $bDebug Then
	$dX = 1280
	$dY = 1024
EndIf

Local $numButSize = Round($dY / 10)
Local $distBt = Round($numButSize / 3)
Local $headerHeight = Round($numButSize * 1.5)
Local $initX = Round($dX / 2 - $numButSize * 1.5 - $distBt)
Local $initY = Round($dY / 2 - $numButSize * 1.5 - $distBt)
Local $fontSize = Round($numButSize / 3)

Local $timeLabel = ""
Local $enteredCode = ""
If $bDebug Then $enteredCode = "9601811873"

Local $pressedButtonTimeCounter = 0
Local $previousButtonPressedID[] = [0, 0]

Local $prevMinute = @MIN
Local $timer = 0
Local $timeCounter = 0

Local $aNextButtonPosition = 0
Local $bt_next = 0
Local $inp_pincode = 0

Local $bottonLineHeight = 11

Local $bPrinterError = False

Local $oExcel = _Excel_Open(False, False, False, False, True)

Local $aPrinterStatusCodes[][] = [ _
	[0,		  "Printer ready"], _
	[1,		  "Printer paused"], _
	[2,		  "Printer error"], _
	[4,		  "Printer pending deletion"], _
	[8, 	  "Paper jam"], _
	[16,	  "Out of paper"], _
	[32, 	  "Manual feed"], _
	[64, 	  "Paper problem"], _
	[128,	  "Printer offline"], _
	[256,	  "IO active"], _
	[512, 	  "Printer busy"], _
	[1024,	  "Printing"], _
	[2048, 	  "Printer output bin full"], _
	[4096,    "Not available."], _
	[8192, 	  "Waiting"], _
	[16384,	  "Processing"], _
	[32768,   "Initializing"], _
	[65536,   "Warming up"], _
	[131072,  "Toner low"], _
	[262144,  "No toner"], _
	[524288,  "Page punt"], _
	[1048576, "User intervention"], _
	[2097152, "Out of memory"], _
	[4194304, "Door open"], _
	[8388608, "Server unknown"], _
	[6777216, "Power save"]]

Local Enum  $enRecordsNotFound, _
			$enMarkOk, _
			$enMarkFail, _
			$enMarkOkPrinterOk, _
			$enMarkOkPrinterFail, _
			$enMarkFailPrinterOk, _
			$enMarkFailPrinterFalil, _
			$enFirstTime, _
			$enMainScreen, _
			$enServiceUnavailable, _
			$enMultiplePatientError, _
			$enNoAppointmentsForNow, _
			$enWrongName
#EndRegion ====================== Variables ======================

If Not $bDebug Then _WinAPI_ShowCursor(False)

FormShowMessage("", $enMainScreen)


Func FormDialer()
	Local $hDialerGui = GUICreate("FormDialer", $dX, $dY, 0, 0, $WS_POPUP, $bDebug ? -1 : $WS_EX_TOPMOST)

	CreateStandardDesign($hDialerGui, $textTitleDialer, False, True)

	Local $bt_1 = CreateButton("1", $initX, $initY, $numButSize, $numButSize)

	Local $prevBt
	$prevBt = ControlGetPos($hDialerGui, "", $bt_1)
	Local $bt_2 = CreateButton("2", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_2)
	Local $bt_3 = CreateButton("3", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_1)
	Local $bt_4 = CreateButton("4", $prevBt[0], $prevBt[1] + $prevBt[3] + $distBt, $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_4)
	Local $bt_5 = CreateButton("5", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_5)
	Local $bt_6 = CreateButton("6", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_4)
	Local $bt_7 = CreateButton("7", $prevBt[0], $prevBt[1] + $prevBt[3] + $distBt, $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_7)
	Local $bt_8 = CreateButton("8", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_8)
	Local $bt_9 = CreateButton("9", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_7)
	Local $bt_clear = CreateButton("C", $prevBt[0], $prevBt[1] + $prevBt[3] + $distBt, $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_clear)
	Local $bt_0 = CreateButton("0", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_0)
	Local $bt_backspace = CreateButton("<", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_clear)
	$bt_next = CreateButton("Продолжить", $prevBt[0], $prevBt[3] + $distBt + $prevBt[1], _
			$numButSize * 3 + $distBt * 2, $numButSize, $colorDisabled)
	GUICtrlSetColor(-1, $colorAlternateText)

	$prevBt = ControlGetPos($hDialerGui, "", $bt_next)
	If Not IsArray($aNextButtonPosition) Then $aNextButtonPosition = $prevBt
	Local $prevBt2 = ControlGetPos($hDialerGui, "", $bt_1)
	$inp_pincode = GUICtrlCreateLabel($enteredCode, $dX / 2 - $prevBt[2] * 2.3 / 2, _
			$prevBt2[1] - $prevBt2[3] - $distBt, $prevBt[2] * 2.3, $prevBt[3], BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetFont(-1, $fontSize * 1.8)
	GUICtrlSetColor(-1, $colorText)

	UpdateTimeLabel()
	UpdateInput($hDialerGui)

	GUISetState(@SW_SHOW)

	ToLog("FormDialer")

	While 1
		$timer = _Timer_Init()

		$nMsg = GUIGetMsg()

		If $timeCounter > $formMaxTimeWait Then
			ToLog("FormDialer force clear")
			$nMsg = $bt_clear
			$timeCounter = 0
			$timer = 0
			$enteredCode = ""
			GUIDelete($hDialerGui)
			Return
		EndIf

		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $bt_0
				NumPressed(0, $bt_0, $hDialerGui)
			Case $bt_1
				NumPressed(1, $bt_1, $hDialerGui)
			Case $bt_2
				NumPressed(2, $bt_2, $hDialerGui)
			Case $bt_3
				NumPressed(3, $bt_3, $hDialerGui)
			Case $bt_4
				NumPressed(4, $bt_4, $hDialerGui)
			Case $bt_5
				NumPressed(5, $bt_5, $hDialerGui)
			Case $bt_6
				NumPressed(6, $bt_6, $hDialerGui)
			Case $bt_7
				NumPressed(7, $bt_7, $hDialerGui)
			Case $bt_8
				NumPressed(8, $bt_8, $hDialerGui)
			Case $bt_9
				NumPressed(9, $bt_9, $hDialerGui)

			Case $bt_next
;~ 				If StringLen($enteredCode) < 10 Then ContinueLoop
				_Timer_KillAllTimers($hDialerGui)
				$timeCounter = 0
				$timer = 0

				FormCheckEnteredNumber($hDialerGui, $enteredCode)

				$enteredCode = ""
				Return
			Case $bt_backspace
				UpdateButtonBackgroundColor($bt_backspace)
				If StringLen($enteredCode) > 0 Then
					$enteredCode = StringLeft($enteredCode, StringLen($enteredCode) - 1)
					UpdateInput($hDialerGui)
				EndIf

			Case $bt_clear
				UpdateButtonBackgroundColor($bt_clear)
				$enteredCode = ""
				UpdateInput($hDialerGui)
		EndSwitch

		Sleep(20)

		If $pressedButtonTimeCounter Then
			$pressedButtonTimeCounter += 10
			If $pressedButtonTimeCounter > 200 Then
				GUICtrlSetBkColor($previousButtonPressedID[0], $previousButtonPressedID[1])
				$previousButtonPressedID[0] = 0
				$previousButtonPressedID[1] = 0
				$pressedButtonTimeCounter = 0
			EndIf
		EndIf

		If $timer Then
			Local $timeDiff = _Timer_Diff($timer)
			$timeCounter += $timeDiff
			_Timer_KillAllTimers($hDialerGui)
			$timer = 0
		EndIf

		If @MIN <> $prevMinute Then
			UpdateTimeLabel()
			$prevMinute = @MIN
		EndIf
	WEnd
EndFunc   ;==>FormDialer


Func FormCheckEnteredNumber($guiToDelete, $code)
	ToLog("FormCheckEnteredNumber: " & $code)
	Local $phoneNumberPrefix = StringLeft($code, 3)
	Local $phoneNumber = StringRight($code, 7)

	Local $sqlQuery = StringReplace($sqlCheckEnteredNumber, "*", "@", 1)
	$sqlQuery = StringReplace($sqlQuery, "*", $phoneNumberPrefix, 1)
	$sqlQuery = StringReplace($sqlQuery, "*", $phoneNumber, 1)
	$sqlQuery = StringReplace($sqlQuery, "@", "*", 1)

	Local $res = ExecuteSQL($sqlQuery)
	; res[0][0] - patient card number
	; res[0][1] - patient name
	; res[0][2] - patient surname
	; res[0][3] - patient bdate
	; res[0][4] - patient first visit mark
	; res[0][5] - patient family head id
	; res[0][6] - patient card number=

	Local $textPhoneNumber = "+7 (" & $phoneNumberPrefix & ") " & StringLeft($phoneNumber, 3) & _
			"-" & StringMid($phoneNumber, 4, 2) & "-" & StringRight($phoneNumber, 2)
	Local $enumMember = -1
	Local $sReplacementText = ""

	Local $bMultiplePatients = false

	If $res = 0 Then
		$enumMember = $enRecordsNotFound
		$sReplacementText = $textPhoneNumber
	ElseIf $res = -1 Then
		$enumMember = $enServiceUnavailable
	ElseIf IsArray($res) And UBound($res, $UBOUND_ROWS) = 1 And $res[0][4] Then
		$enumMember = $enFirstTime
	ElseIf IsArray($res) And UBound($res, $UBOUND_ROWS) > 1 Then
		If UBound($res, $UBOUND_ROWS) > 4 Then _
			$enumMember = $enMultiplePatientError

		$bMultiplePatients = True
	EndIf

	If $enumMember > -1 Then
		FormShowMessage($guiToDelete, $enumMember, $sReplacementText)
		Return
	EndIf

	Local $fioForm = GUICreate("FIO", $dX, $dY, 0, 0, $WS_POPUP, $bDebug ? -1 : $WS_EX_TOPMOST)

	CreateStandardDesign($fioForm, $bMultiplePatients ? $textTitleNameConfirmMultiple : $textTitleNameConfirm, False)

	Local $bt_ok = CreateButton("Продолжить", $dX - $distBt - $aNextButtonPosition[2], $aNextButtonPosition[1], _
			$aNextButtonPosition[2], $aNextButtonPosition[3], $colorOkButton)
		GUICtrlSetColor(-1, $colorAlternateText)

	Local $bt_not = CreateButton($bMultiplePatients ? "Закрыть" : "Неверно", 0 + $distBt, $aNextButtonPosition[1], $aNextButtonPosition[2], _
			$aNextButtonPosition[3])

	Local $aNameButtons[4][4]
	; $aNameButtons[0][0] - button id
	; $aNameButtons[0][1] - button state (selected / not selected)
	; $aNameButtons[0][2] - button label id
	; $aNameButtons[0][3] - button is icons present

	If UBound($res, $UBOUND_ROWS) = 1 Then
		Local $date = GetBDayFromString($res[0][3])
		Local $fullName = $res[0][1] & " " & $res[0][2]
		Local $mainText = $fullName & @CRLF & @CRLF & "Дата рождения: " & $date
		CreateLabel($mainText, 0, $dY * 0.3, $dX, $dY * 0.4, $colorText, $GUI_BKCOLOR_TRANSPARENT, $fioForm, $fontSize * 1.2)
	Else
		Local $nButtonsDistanceBetween = $distBt * 1.5
		Local $nNameButtonTotalHeight = $aNextButtonPosition[3] * UBound($res, $UBOUND_ROWS) + _
										$nButtonsDistanceBetween * (UBound($res, $UBOUND_ROWS) - 1)
		Local $nStartY = $headerHeight + ($aNextButtonPosition[1] - $headerHeight - $nNameButtonTotalHeight) / 2

		For $i = 0 To UBound($res, $UBOUND_ROWS) - 1
			Local $sFullName = $res[$i][1] & " " & $res[$i][2] & @CRLF & "Дата рождения: " & GetBDayFromString($res[$i][3])
			$aNameButtons[$i][0] = CreateButton("", 0 + $nButtonsDistanceBetween, $nStartY, _
								$dX - $nButtonsDistanceBetween * 2, $aNextButtonPosition[3])
			$aNameButtons[$i][1] = False
			$aNameButtons[$i][2] = CreateLabel($sFullName, 0 + $nButtonsDistanceBetween, $nStartY, _
								$dX - $nButtonsDistanceBetween * 2, $aNextButtonPosition[3], $colorText, _
								$GUI_BKCOLOR_TRANSPARENT, $fioForm, $fontSize * 0.9)
			$nStartY += $aNextButtonPosition[3] + $nButtonsDistanceBetween
			$aNameButtons[$i][3] = -1
		Next

		SetButtonEnabled($bt_ok, False)
	EndIf

	UpdateTimeLabel()

	GUISetState()

	Sleep(50)
	If $guiToDelete Then GUIDelete($guiToDelete)

	$timeCounter = 0

	Local $nLastSelectedName = 0

	While 1
		$timer = _Timer_Init()

		If $timeCounter > $formMaxTimeWait Then
			ToLog("FormCheckEnteredNumber force close" & @CRLF)
			GUIDelete($fioForm)
			Return
		EndIf

		$nMsg = GUIGetMsg()

		If $nMsg = $bt_not And Not $bMultiplePatients Then
				ToLog("FormCheckEnteredNumber close, fullname not correct: " & $fullName)
				FormShowMessage($fioForm, $enWrongName)
				Return
		ElseIf $nMsg = $bt_not And $bMultiplePatients Then
				ToLog("FormCheckEnteredNumber close, multiple patients")
				GUIDelete($fioForm)
				Return
		ElseIf $nMsg = $bt_ok Then
				Local $result = -1

				If $res[$nLastSelectedName][4] Then
					$result = $enFirstTime
					FormShowMessage("", $result)
				Else
					$result = FormShowAppointments($bMultiplePatients ? -1 : $fioForm, _
						$res[$nLastSelectedName][0], _
						$res[$nLastSelectedName][1], _
						$res[$nLastSelectedName][2])
				EndIf

				$timeCounter = 0
				$timer = 0

				If Not $bMultiplePatients Then Return

				GUISwitch($fioForm)
				If $aNameButtons[$nLastSelectedName][3] = -1 Then
					$aNameButtons[$nLastSelectedName][3] = $result
					SetButtonIcon(ControlGetPos($fioForm, "", $aNameButtons[$nLastSelectedName][0]), $result)
				EndIf

				UpdateNameButtons($aNameButtons, -1)
				SetButtonEnabled($bt_ok, False)

		ElseIf $nMsg > 0 Then
			For $i = 0 To UBound($res, $UBOUND_ROWS) - 1
				If $nMsg = $aNameButtons[$i][0] Then
					$timeCounter = 0
					UpdateNameButtons($aNameButtons, $i)
					SetButtonEnabled($bt_ok, True)
					$nLastSelectedName = $i
				EndIf
			Next
		EndIf

		Sleep(20)

		If $timer <> 0 Then
			Local $timeDiff = _Timer_Diff($timer)
			$timeCounter += $timeDiff
			$timer = 0
		EndIf

		If @MIN <> $prevMinute Then
			UpdateTimeLabel()
			$prevMinute = @MIN
		EndIf
	WEnd
EndFunc   ;==>FormCheckEnteredNumber


Func FormShowAppointments($guiToDelete, $patientID, $name, $surname)
	Local $fullName = $name & " " & $surname
	ToLog("FormShowAppointments: " & $fullName)

	Local $sqlQuery = StringReplace($sqlGetAppointments, "*", $patientID)
	Local $res = ExecuteSQL($sqlQuery)

	$res = GetAppointmentsForCurrentTime($res)

	If Not IsArray($res) Or Not UBound($res, $UBOUND_ROWS) Then
		FormShowMessage($guiToDelete, $enNoAppointmentsForNow)
		Return
	EndIf

	Local $destForm = 0
	Local $bt_close = -666
	Local $bt_print = -667
	Local $needRegistry = _ArrayMax($res, Default, Default, Default, 6) + _
						  _ArrayMax($res, Default, Default, Default, 7) + _
						  _ArrayMax($res, Default, Default, Default, 8)

	If $showAppointmentsForm Then
		$destForm = GUICreate("FormShowAppointments", $dX, $dY, 0, 0, $WS_POPUP, $bDebug ? -1 : $WS_EX_TOPMOST)
		CreateStandardDesign($destForm, StringReplace($textTitleAppointments, "*", $fullName), False)


		$bt_close = CreateButton("Закрыть", _
				0 + $distBt, _
				$aNextButtonPosition[1], _
				$aNextButtonPosition[2], _
				$aNextButtonPosition[3])

		$bt_print = CreateButton("Распечатать", _
				$dX - $distBt - $aNextButtonPosition[2], _
				$aNextButtonPosition[1], _
				$aNextButtonPosition[2], _
				$aNextButtonPosition[3], _
				$colorOkButton, _
				$colorAlternateText)

		GUISetFont($fontSize * 0.9, $fontWeightAppointments, -1, $fontNameAppointments, $destForm, $fontQuality)
		CreateAppointmentsTable($res, $destForm)

		UpdateTimeLabel()
		GUISetState()

		Sleep(50)
		If $guiToDelete Then GUIDelete($guiToDelete)
	EndIf

	If Not $needRegistry Then
		For $i = 0 To UBound($res, $UBOUND_ROWS) - 1
			Local $idToUpdate = $res[$i][0]
			Local $updateSql = StringReplace($sqlSetMark, "*", $idToUpdate)

			ExecuteSQL($updateSql)
			ToLog("Setting visit mark for: " & $idToUpdate)
		Next
	EndIf

	Local $needToClose = False
	Local $bPrintResult = -1

	$timeCounter = 0

	While True
		$timer = _Timer_Init()

		If $timeCounter > $formMaxTimeWait Then $needToClose = True

		$nMsg = GUIGetMsg()
		If Not $showAppointmentsForm Then
			$nMsg = $bt_print
			$destForm = $guiToDelete
		EndIf

		Switch $nMsg
			Case $bt_close
				ToLog("FormShowAppointments close")
				$needToClose = True
			Case $bt_print
				Local $printResult = PrintAppontments($res, $name, $surname)
				If Not $printResult Then
					$bPrinterError = False
					$bPrintResult = 1
				Else
					If Not $bPrinterError Then
						SendEmail($sMailTitle & @CRLF & "Инфомату не удалось распечатать список назначений пациента " & $patientID & _
							" " & $name & " " & $surname & @CRLF & $printResult, "", True)
						$bPrinterError = True
					EndIf
					$bPrintResult = 0
				EndIf
				$needToClose = True
		EndSwitch

		If $needToClose Then
			Local $enumMember = -1

			If $needRegistry Then
				$enumMember = $enMarkFail
				If $bPrintResult = 1 Then
					$enumMember = $enMarkFailPrinterOk
				ElseIf $bPrintResult = 0 Then
					$enumMember = $enMarkFailPrinterFalil
				EndIf
			Else
				$enumMember = $enMarkOk
				If $bPrintResult = 1 Then
					$enumMember = $enMarkOkPrinterOk
				ElseIf $bPrintResult = 0 Then
					$enumMember = $enMarkOkPrinterFail
				EndIf
			EndIf

			FormShowMessage($destForm, $enumMember)
			Return $enumMember
		EndIf

		Sleep(20)

		Local $timeDiff = _Timer_Diff($timer)
		$timeCounter += $timeDiff
		$timer = 0

		If @MIN <> $prevMinute Then
			UpdateTimeLabel()
			$prevMinute = @MIN
		EndIf
	WEnd
EndFunc   ;==>FormShowAppointments


Func FormShowMessage($guiToDelete, $enumMember, $sReplacementText = "")
	ConsoleWrite("---FormShowMessage: " & $enumMember & @CRLF)
	Local $sMessageTotal = ""
	Local $sMessageTop = ""
	Local $sMessageBottom = ""
	Local $sImageTopName = ""
	Local $sImageBottomName = ""

	Local $showError = False
	Local $checkDb = False
	Local $bMainScreen = False

	Switch $enumMember
		Case $enRecordsNotFound
			ConsoleWrite("---$enRecordsNotFound" & @CRLF)
			$sMessageTotal = StringReplace($textNotificationNothingFound, "*", $sReplacementText)
			$showError = True
			$sImageTopName = "PicNotFound.jpg"
		Case $enMarkOk
			$sMessageTotal = $textAppointmentsMarkOk
			$sImageTopName = "AnimationCheckOk.avi"
		Case $enMarkFail
			$sMessageTotal = $textAppointmentsMarkProblem
			$sImageTopName = "PicRegistry.jpg"
		Case $enMarkOkPrinterOk
			$sMessageTotal = $textAppointmentsMarkOk & @CRLF & @CRLF & $textAppointmentsPrintOk
			$sImageTopName = "AnimationCheckOk.avi"
		Case $enMarkOkPrinterFail
			$sMessageTotal = $textAppointmentsMarkOk & @CRLF & @CRLF & $textAppointmentsPrintProblem
			$sImageTopName = "PicOk.jpg"
			$sImageBottomName = "PicPrinterError.jpg"
		Case $enMarkFailPrinterOk
			$sMessageTotal = $textAppointmentsMarkProblem & @CRLF & @CRLF & $textAppointmentsPrintOk
			$sImageTopName = "PicRegistry.jpg"
		Case $enMarkFailPrinterFalil
			$sMessageTotal = $textAppointmentsMarkProblem & @CRLF & @CRLF & $textAppointmentsPrintProblem
			$sImageTopName = "PicRegistry.jpg"
			$sImageBottomName = "PicPrinterError.jpg"
		Case $enFirstTime
			$sMessageTotal = $textNotificationFirstVisit
			$sImageTopName = "PicRegistry.jpg"
		Case $enMainScreen
			$sMessageTotal = $sWelcomeTop & @CRLF & @CRLF & $sWelcomeBottom
			$bMainScreen = True
			$sImageTopName = "AnimationCheck.avi"
		Case $enServiceUnavailable
			$sMessageTotal = $textNotificationDbNotAvailable
			$showError = True
			$checkDb = True
			$sImageTopName = "PicError.jpg"
		Case $enMultiplePatientError
			$sMessageTotal = $textNotificationMultiplePatientsError
			$sImageTopName = "PicRegistry.jpg"
		Case $enNoAppointmentsForNow
			$sMessageTotal = $textNotificationNoAppointmetnsForNow
			$sImageTopName = "PicRegistry.jpg"
		Case $enWrongName
			$sMessageTotal = $textNotificationWrongName
			$sImageTopName = "PicRegistry.jpg"
		Case Else
			ToLog("FormShowMessage wrong enum!!!")
			Return
	EndSwitch

	ToLog("FormShowMessage: " & StringReplace($sMessageTotal, @CRLF, " | ") & ($bMainScreen ? " mainScreenMessage" : ""))

	Local $iLabelTopWidth = $dX
	Local $iLabelTopHeight = $headerHeight
	Local $iLabelTopX = 0
	Local $iLabelTopY = $headerHeight

	Local $iLabelBottomWidth = $iLabelTopWidth
	Local $iLabelBottomHeight = $iLabelTopHeight
	Local $iLabelBottomX = $iLabelTopX
	Local $iLabelBottomY = 0

	Local $iImageWidth = 0
	Local $iImageHeight = 0
	Local $iImageX = 0
	Local $iImageY = $iLabelTopY + $iLabelTopHeight

	If StringInStr($sMessageTotal, @CRLF & @CRLF) Then
		Local $aSplited = StringSplit($sMessageTotal, @CRLF & @CRLF, BitOR($STR_NOCOUNT, $STR_ENTIRESPLIT))
		For $i = 0 To UBound($aSplited) - 2
			If Not $aSplited[$i] Then ContinueLoop
			$sMessageTop &= $aSplited[$i]
			If $i < (UBound($aSplited) - 2) Then $sMessageTop &= @CRLF
		Next
		$sMessageBottom = $aSplited[UBound($aSplited) - 1]
	Else
		$sMessageTop = $sMessageTotal
	EndIf

	Local $nanForm = GUICreate("FormShowMessage", $dX, $dY, 0, 0, $WS_POPUP, $bDebug ? -1 : $WS_EX_TOPMOST)
	CreateStandardDesign($nanForm, (($bMainScreen Or $checkDb) ? $sTitleWelcome : $textTitleNotification), $showError, True)

	Local $bt_close = 666
	If Not $checkDb And Not $bMainScreen Then
		$bt_close = CreateButton("Закрыть", $aNextButtonPosition[0], $aNextButtonPosition[1], _
			$aNextButtonPosition[2], $aNextButtonPosition[3])
		Local $aButtonPos = ControlGetPos($nanForm, "", $bt_close)
		$iLabelBottomY = $aButtonPos[1] - $iLabelBottomHeight
	Else
		$iLabelBottomY = $dY - $bottonLineHeight - $iLabelBottomHeight
	EndIf

	If $sImageBottomName Then
		$iLabelTopX = $dX * 0.3
		$iLabelTopWidth *= 0.7
		$iLabelTopHeight = ($iLabelBottomY + $iLabelBottomHeight - $iLabelTopY ) / 2

		$iImageWidth = $dX * 0.2
		$iImageHeight = $iImageWidth
		$iImageX = ($dX * 0.3 - $iImageWidth ) / 2
		$iImageY = ($iLabelTopY + ($iLabelTopHeight - $iImageHeight) / 2)

		$iLabelBottomWidth = $iLabelTopWidth
		$iLabelBottomHeight = $iLabelTopHeight
		$iLabelBottomX = $iLabelTopX
		$iLabelBottomY = $iLabelTopY + $iLabelTopHeight
	Else
		$iImageHeight = $iLabelBottomY - $iImageY
		If Not $sMessageBottom Then _
			$iImageHeight += $iLabelBottomHeight

		$iImageWidth = $iImageHeight
		$iImageX = ($dX - $iImageWidth ) / 2
	EndIf

	If $bMainScreen Then
		ToLog("-----MainGui started-----")
		SendEmail("-----MainGui started-----")
	EndIf

	CreateLabel($sMessageTop, $iLabelTopX, $iLabelTopY, $iLabelTopWidth, $iLabelTopHeight, _
		$colorText, $GUI_BKCOLOR_TRANSPARENT, $nanForm, $sImageBottomName ? $fontSize * 0.8 : $fontSize)

	CreateLabel($sMessageBottom, $iLabelBottomX, $iLabelBottomY, $iLabelBottomWidth, $iLabelBottomHeight, _
		$colorText, $GUI_BKCOLOR_TRANSPARENT, $nanForm, $fontSize * 0.8)



	Local $sFile = $resourcesPath & $sImageTopName
	If StringInStr($sFile, ".jpg") Then
		GUICtrlCreatePic($sFile, $iImageX, $iImageY, $iImageWidth, $iImageHeight)
		If $sImageBottomName Then
			$iImageY = ($iLabelBottomY + ($iLabelBottomHeight - $iImageHeight) / 2)
			GUICtrlCreatePic($resourcesPath & $sImageBottomName, $iImageX, $iImageY, $iImageWidth, $iImageHeight)
		EndIf
	ElseIf StringInStr($sFile, ".avi") Then
		If $iImageHeight > 410 Then
			$iImageX += ($iImageWidth - 410) / 2
			$iImageY += ($iImageWidth - 410) / 2
			$iImageHeight = 410
			$iImageWidth = 410
		EndIf

		Local $g_hAVI = _GUICtrlAVI_Create($nanForm, $sFile, -1, $iImageX, $iImageY, _
			$iImageWidth, $iImageHeight, BitOR($ACS_CENTER, $ACS_AUTOPLAY))
		_GUICtrlAVI_Play($g_hAVI, -1, -1, StringInStr($sImageTopName, "Ok.avi") ? 1 : -1)
	EndIf

	UpdateTimeLabel()
	GUISetState()

	Sleep(50)

	If $guiToDelete Then GUIDelete($guiToDelete)

	If $showError And $checkDb Then SendEmail("Не удалось подключиться к БД: " & $infoclinicaDB, "", True)

	$timeCounter = 0

	Local $nMaxtTimeWait = $formMaxTimeWait
	If Not $showError And Not StringInStr($sMessageTop, "регистратуру") Then _
		$nMaxtTimeWait /= 2

	While 1
		If Not $checkDb And Not $bMainScreen Then $timer = _Timer_Init()

		$nMsg = GUIGetMsg($GUI_EVENT_ARRAY)

		If $timeCounter > $nMaxtTimeWait Then
			ToLog("FormShowMessage force close" & @CRLF)
			$nMsg[0] = $bt_close
			$nMsg[1] = $nanForm
		EndIf

		If @MIN <> $prevMinute Then
			If $bMainScreen And Not GetDatabaseAvailabilityStatus() Then _
					FormShowMessage("", $enServiceUnavailable) ;$textNotificationDbNotAvailable, True, True)

			UpdateTimeLabel()
			$prevMinute = @MIN
			If $checkDb Then
				Local $res = GetDatabaseAvailabilityStatus()
				If $res Then
					$nMsg[0] = $bt_close
					$nMsg[1] = $nanForm
				EndIf
			EndIf
		EndIf

		If $nMsg[1] = $nanForm And $nMsg[0] = $bt_close Then
			ToLog("FormShowMessage close" & @CRLF)
			_Timer_KillAllTimers($nanForm)
			$timeCounter = 0
			$timer = 0
			GUIDelete($nanForm)
			Return
		EndIf

		If $nMsg[1] = $nanForm And _
			($nMsg[0] = $GUI_EVENT_PRIMARYDOWN OR _
			$nMsg[0] = $GUI_EVENT_MOUSEMOVE) Then
			If Not $bMainScreen Then ContinueLoop
			Local $tempTimeLabel = $timeLabel

			FormDialer()

			$timeLabel = $tempTimeLabel
			_Timer_KillAllTimers($nanForm)
			$timeCounter = 0
			$timer = 0
		EndIf

		If Not $checkDb And Not $bMainScreen Then
			Local $timeDiff = _Timer_Diff($timer)
			$timeCounter += $timeDiff
			_Timer_KillAllTimers($nanForm)
			$timer = 0
		EndIf
	WEnd
EndFunc   ;==>FormShowMessage




Func CreateStandardDesign($gui, $titleText, $isError, $trademark = False)
	GUISetBkColor($colorMainBackground)
	GUISetFont($fontSize, $fontWeight, 0, $fontName, $gui, $fontQuality)

	Local $titleColor = $colorHeader
	If $isError Then $titleColor = $colorErrorTitle


	Local $sTrademarkFileName = "PicButterfly.jpg"
	Local $trademarkWidth = Round($dX * 0.12)
	Local $trademarkHeight = Round($trademarkWidth * 1.07)

	Local $headerBackgroundColor = $titleColor

	If $titleText = $sTitleWelcome Then
		If (@MDAY > 24 And @MON = 12) Or (@MDAY < 10 And @MON = 1) Then
			$sTrademarkFileName = "PicChristmasTree.jpg"
			$trademarkWidth = Round($dX * 0.133)
			$trademarkHeight = Round($trademarkHeight * 1.26)
			$headerBackgroundColor = $GUI_BKCOLOR_TRANSPARENT

			Local $sHeaderFileName = "PicChristmasHeader.jpg"
			Local $iHeaderWidth = $dX
			Local $iHeaderHeight = Round($iHeaderWidth * 0.137)
			GUICtrlCreatePic($resourcesPath & $sHeaderFileName, 0, 0, $iHeaderWidth, $iHeaderHeight)
		EndIf
	EndIf

	CreateLabel($titleText, 0, 0, $dX, $headerHeight, $colorAlternateText, $headerBackgroundColor, $gui)

	Local $timeScopeWidth = $numButSize * 1.7
	Local $timeIconWidth = 20

	GUISetFont($fontSize * 0.7, $fontWeight, 0, $fontName, $gui, $fontQuality)
	$timeLabel = GUICtrlCreateLabel("13:14", 0, 0, -1, -1, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	Local $timeLabelPosition = ControlGetPos($gui, "", $timeLabel)
	GUICtrlSetPos($timeLabel, $dX - $timeLabelPosition[2] - $distBt / 4, $distBt / 4)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, $colorAlternateText)
	GUISetFont($fontSize, $fontWeight, 0, $fontName, $gui, $fontQuality)

	$timeLabelPosition = ControlGetPos($gui, "", $timeLabel)
	Local $timePic = CreatePngControl($resourcesPath & "TimeIcon.png", $timeIconWidth, $timeIconWidth)
	GUICtrlSetPos($timePic, $timeLabelPosition[0] - $distBt / 4 - $timeIconWidth, _
			$timeLabelPosition[1] + $timeLabelPosition[3] / 2 - $timeIconWidth / 2)

	GUICtrlCreatePic($resourcesPath & "PicBottomLine.jpg", 0, $dY - $bottonLineHeight, $dX, $bottonLineHeight)

	If $trademark Then _
		GUICtrlCreatePic($resourcesPath & $sTrademarkFileName, $dX - $trademarkWidth - $distBt / 2, _
				$dY - $trademarkHeight - $bottonLineHeight - $distBt / 2, $trademarkWidth, $trademarkHeight)
EndFunc   ;==>CreateStandardDesign


Func CreatePngControl($pngPath, $width, $height)
	Local $newControl = GUICtrlCreatePic("", 0, 0, -1, -1)
	_GDIPlus_Startup()
	Local $hImage = _GDIPlus_ImageLoadFromFile($pngPath)
	Local $resize = _GDIPlus_ImageResize($hImage, $width, $height)
	Local $bmp = _GDIPlus_BitmapCreateHBITMAPFromBitmap($resize)
	_WinAPI_DeleteObject(GUICtrlSendMsg($newControl, 0x0172, 0, $bmp))
	_WinAPI_DeleteObject($bmp)
	_GDIPlus_ImageDispose($hImage)
	_GDIPlus_Shutdown()
	Return $newControl
EndFunc   ;==>CreatePngControl


Func CreateLabel($text, $x, $y, $width, $height, $colorText, $backgroundColor, $gui, $fntSize = $fontSize, $fntWeight = $fontWeight)
	$Label1 = GUICtrlCreateLabel("", $x, $y, $width, $height)
	GUICtrlSetBkColor(-1, $backgroundColor)

	GUISetFont($fntSize, $fntWeight, 0, $fontName, $gui, $fontQuality)
	Local $label = GUICtrlCreateLabel($text, 0, 0, -1, -1, $SS_CENTER)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, $colorText)

	Local $position = ControlGetPos($gui, "", $label)
	If IsArray($position) Then
		Local $newX = $x + ($width - $position[2]) / 2
		Local $newY = $y + ($height - $position[3]) / 2
		GUICtrlSetPos($label, $newX, $newY)
	EndIf
	GUISetFont($fontSize, $fontWeight, 0, $fontName, $gui, $fontQuality)

	Return $label
EndFunc   ;==>CreateLabel


Func CreateButton($text, $x, $y, $width, $height, $bkColor = $colorMainButton, $color = $colorText)
	Local $offsetX = 6
	Local $offsetY = 4
	Local $sizeX = 12
	Local $sizeY = 12

	If $width <> $height Then
		$offsetX = 16
		$sizeX = 32
	EndIf

	GUICtrlCreatePic($resourcesPath & "PicShadow.jpg", $x - $offsetX, $y - $offsetY, $width + $sizeX, $height + $sizeY, $SS_BLACKRECT)
	GUICtrlSetState(-1, $GUI_DISABLE)

	Local $id = GUICtrlCreateLabel($text, $x, $y, $width, $height, BitOR($SS_CENTER, $SS_CENTERIMAGE, $SS_NOTIFY))
	GUICtrlSetBkColor(-1, $bkColor)
	GUICtrlSetColor(-1, $color)

	Return $id
EndFunc   ;==>CreateButton


Func CreateAppointmentsTable($res, $gui)
	Local $head[1][UBound($res, $UBOUND_COLUMNS)]
	$head[0][2] = "Время"
	$head[0][3] = "Специалист"
	$head[0][4] = "Отделение"
	$head[0][5] = "Кабинет"
	_ArrayConcatenate($head, $res)

	Local $startX = $distBt
	Local $startY = Round($numButSize * 1.5 + $distBt)
	Local $height = Round($numButSize * 0.7)
	Local $iconWidth = Round($height * 0.6)
	Local $distance = Round($distBt / 6)
	Local $totalWidth = $dX - $distBt * 2 - $distance * 3
	Local $currentX = $startX + Round($distBt / 2)
	Local $currentY = $startY

	Local $sizes[3]
	$sizes[0] = GetOptimalLabelWidth($head[0][2], $gui)
	$sizes[1] = GetOptimalLabelWidth($head[0][5], $gui)
	$sizes[2] = $totalWidth - $sizes[0] - $sizes[1] - $distBt * 2

	Local $maxSymbols = Round($sizes[2] / ($fontSize * 0.8))

	GUICtrlCreateLabel("", $startX, $startY, $totalWidth + $distance * 3, $height - $distance)
	GUICtrlSetBkColor(-1, $colorMainButton)

	Local $arraySize = UBound($head, $UBOUND_ROWS) - 1
	If $arraySize > 6 Then $arraySize = 6

	Local $showCashWarning = False
	Local $showOutOfTimeWarning = False
	Local $showXrayWarning = False

	For $i = 0 To $arraySize
		Local $currentRow[4]
		Local $cash = $head[$i][6]
		Local $time = $head[$i][8]
		Local $xray = $head[$i][7]

		$currentRow[0] = $head[$i][2]
		$currentRow[1] = $head[$i][5]

		Local $dept = StringLower($head[$i][4])

		Local $doc = $head[$i][3]
		If StringInStr($doc, "(") Then
			$doc = StringLeft($doc, StringInStr($doc, "(") - 1)
			$doc = StringStripWS($doc, BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING, $STR_STRIPSPACES))
		EndIf

		$currentRow[2] = $doc & " (" & $dept & ")"
		While GetOptimalLabelWidth($currentRow[2], $gui) > $sizes[2] - $iconWidth * ($cash + $time + $xray) * 1.2
			$currentRow[2] = StringLeft($currentRow[2], StringLen($currentRow[2]) - 5) & "...)"
		WEnd

		For $x = 0 To 2
			Local $attribute = BitOR($SS_CENTER, $SS_CENTERIMAGE)
			If $x = 2 Then $attribute = $SS_CENTERIMAGE

			GUICtrlCreateLabel($currentRow[$x], _
					$currentX, _
					$currentY, _
					$sizes[$x], _
					$height, _
					$attribute)
			GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
			If Not $i Then GUICtrlSetFont(-1, -1, $FW_SEMIBOLD)

			If $i < $arraySize Then
				GUICtrlCreateLabel("", _
						$startX, _
						$currentY + ($i = 0 ? $height - $distance - 1 : $height), _
						$totalWidth + $distance * 3, _
						$distance)
				GUICtrlSetBkColor(-1, $colorMainButton)
			EndIf

			$currentX += $sizes[$x] + $distance + Round($distBt / 2)

			Local $iconsArray[3][2]
			$iconsArray[0][0] = $time
			$iconsArray[0][1] = "OutOfTimeIcon.png"

			$iconsArray[1][0] = $cash
			$iconsArray[1][1] = "RubleIcon.png"

			$iconsArray[2][0] = $xray
			$iconsArray[2][1] = "XrayIcon.png"

			Local $initialX = $dX - Round($distBt * 1.5) - $iconWidth
			For $y = 0 To UBound($iconsArray, $UBOUND_ROWS) - 1
				If Not $iconsArray[$y][0] Then ContinueLoop

				Local $tmp = CreatePngControl($resourcesPath & $iconsArray[$y][1], $iconWidth, $iconWidth)
				GUICtrlSetPos(-1, $initialX, $currentY + Round($height * 0.2))
				$initialX -= $iconWidth * 1.2
			Next

			If $time Then $showOutOfTimeWarning = True
			If $cash Then $showCashWarning = True
			If $xray Then $showXrayWarning = True
		Next

		$currentY += $height + $distance
		$currentX = $startX + Round($distBt / 2)

		If $i = 0 Then $currentY -= $distance
	Next

	Local $showWarning = Int($showCashWarning) + Int($showOutOfTimeWarning) + Int($showXrayWarning)
	If $showWarning Then
		Local $message = ""
		Local $textHeight = 1.5
		If $showCashWarning And _
				Not $showOutOfTimeWarning And _
				Not $showXrayWarning Then
			$message = $textAppointmentsWarningCash
		ElseIf $showOutOfTimeWarning And _
				Not $showCashWarning And _
				Not $showXrayWarning Then
			$message = $textAppointmentsWarningTime
		ElseIf $showXrayWarning And _
				Not $showCashWarning And _
				Not $showOutOfTimeWarning Then
			$message = $textAppointmentsWarningXray
		Else
			$message = $textAppointmentsWarningGeneral
			If Not StringInStr($message, @CRLF) Then $textHeight = 1.0
		EndIf

		CreateLabel($message, _
				$startX, _
				$currentY, _
				$totalWidth + $distance * 3, _
				Round($height * $textHeight), _
				$colorAlternateText, _
				$colorErrorTitle, _
				$gui, _
				Round($fontSize * 0.8), _
				$fontWeightAppointments)

		$currentY += $height + $distance

		GUISetFont($fontSize * 0.7, $fontWeightAppointments, 0, $fontName, $gui, $fontQuality)

		If Not $showIconsDescription Then Return
		If $showWarning < 2 Then Return

		Local $labelsArray[3][3]
		$labelsArray[0][0] = $showOutOfTimeWarning
		$labelsArray[0][1] = "OutOfTimeIcon.png"
		$labelsArray[0][2] = " - пропущено время"

		$labelsArray[1][0] = $showCashWarning
		$labelsArray[1][1] = "RubleIcon.png"
		$labelsArray[1][2] = " - наличный расчет"

		$labelsArray[2][0] = $showXrayWarning
		$labelsArray[2][1] = "XrayIcon.png"
		$labelsArray[2][2] = " - отделение лучевой диагностики"

		Local $startX = $dX / 2 + $distance * 3
		Local $arraySize = UBound($labelsArray, $UBOUND_ROWS) - 1
		For $i = 0 To $arraySize
			If Not $labelsArray[$i][0] Then ContinueLoop

			$startX -= ($iconWidth + $distance * 3 + GetOptimalLabelWidth($labelsArray[$i][2], $gui)) / 2
		Next

		For $i = 0 To $arraySize
			If Not $labelsArray[$i][0] Then ContinueLoop

			Local $tmp = CreatePngControl($resourcesPath & $labelsArray[$i][1], $iconWidth, $iconWidth)
			GUICtrlSetPos(-1, $startX, $currentY)
			$tmp = GUICtrlCreateLabel($labelsArray[$i][2], _
					$startX + $iconWidth, _
					$currentY, _
					 - 1, _
					$iconWidth, _
					$SS_CENTERIMAGE)

			Local $tmp2 = ControlGetPos($gui, -1, $tmp)
			$startX = $tmp2[0] + $tmp2[2] + $distance * 3
		Next
	EndIf
EndFunc   ;==>CreateAppointmentsTable




Func SetButtonSelected($idButton, $idLabel, $selected)
	If Not $idButton Or Not $idLabel Then Return
	GUICtrlSetBkColor($idButton, $selected ? $colorNameButtonSelected : $colorMainButton)
	GUICtrlSetColor($idLabel, $selected ? $colorAlternateText : $colorText)
EndFunc


Func SetButtonEnabled($id, $enabled)
	If Not $id Then Return
	GUICtrlSetBkColor($id, $enabled ? $colorOkButton : $colorMainButton)
	GUICtrlSetState($id, $enabled ? $GUI_ENABLE : $GUI_DISABLE)
EndFunc


Func SetButtonIcon($aButtonPos, $enumMember)
	If Not IsArray($aButtonPos) Then Return
	If UBound($aButtonPos) < 4 Then Return

	Local $gap = Round($aButtonPos[3] * 0.2)

	Local $width = $aButtonPos[3] - $gap * 2
	Local $height = $width

	Local $x = $aButtonPos[0] + $aButtonPos[2] - $width - $gap
	Local $y = $aButtonPos[1] + $gap

	Local $sPicName = ""

	If $enumMember = $enMarkOk Or _
		$enumMember = $enMarkOkPrinterFail Or _
		$enumMember = $enMarkOkPrinterOk Then
		$sPicName = "PicOk.png"
	ElseIf $enumMember = $enMarkFail Or _
		$enumMember = $enMarkFailPrinterFalil Or _
		$enumMember = $enMarkFailPrinterOk Or _
		$enumMember = $enFirstTime Then
		$sPicName = "PicRegistry2.png"
	EndIf

	If $sPicName Then
		Local $sFullPath = $resourcesPath & $sPicName
		If Not FileExists($sFullPath) Then Return

		Local $hPng = CreatePngControl($sFullPath, $width, $height)
		GUICtrlSetPos($hPng, $x, $y)
	EndIf
EndFunc


Func UpdateNameButtons(ByRef $aButtonsArray, $nSelectedButton, $enumMember = -1)
	If Not IsArray($aButtonsArray) Then Return
	If $nSelectedButton > UBound($aButtonsArray, $UBOUND_ROWS) - 1 Then Return
	If UBound($aButtonsArray, $UBOUND_COLUMNS) < 3 Then Return

	For $i = 0 To UBound($aButtonsArray, $UBOUND_ROWS) - 1
		If $i <> $nSelectedButton Then
			If $aButtonsArray[$i][1] Then
				$aButtonsArray[$i][1] = False
				SetButtonSelected($aButtonsArray[$i][0], $aButtonsArray[$i][2], False)
			EndIf
		Else
			If Not $aButtonsArray[$i][1] Then
				$aButtonsArray[$i][1] = True
				SetButtonSelected($aButtonsArray[$i][0], $aButtonsArray[$i][2], True)
			EndIf
		EndIf
	Next
EndFunc


Func UpdateButtonBackgroundColor($id, $bkColor = $colorMainButton, $glowColor = $colorMainButtonPressed)
	If $enteredCode Then $timeCounter = 0

	If $previousButtonPressedID[0] Then
		GUICtrlSetBkColor($previousButtonPressedID[0], $previousButtonPressedID[1])
		$pressedButtonTimeCounter = 0
	EndIf

	GUICtrlSetBkColor($id, $glowColor)
	$previousButtonPressedID[0] = $id
	$previousButtonPressedID[1] = $bkColor
	$pressedButtonTimeCounter = 1
EndFunc   ;==>UpdateButtonBackgroundColor


Func UpdateInput($hGui)
	Local $format = "+7 (___) ___-__-__"

	Local $codeLenght = StringLen($enteredCode)
	If $codeLenght = 0 Or $codeLenght = 9 Then
		SetButtonEnabled($bt_next, False)
	ElseIf $codeLenght = 10 Then
		SetButtonEnabled($bt_next, True)
	EndIf

	For $i = 1 To $codeLenght
		$format = StringReplace($format, "_", StringMid($enteredCode, $i, 1), 1)
	Next

	ControlSetText($hGui, "", $inp_pincode, $format)
EndFunc   ;==>UpdateInput


Func UpdateTimeLabel()
	Local $newTime = @HOUR & ":" & @MIN
	GUICtrlSetData($timeLabel, $newTime)
EndFunc   ;==>UpdateTimeLabel




Func GetBDayFromString($string)
	Local $sReturn = ""
	If StringLen($string) >= 8 Then _
		$sReturn = StringMid($string, 7, 2) & "." & StringMid($string, 5, 2) & "." & StringLeft($string, 4)
	Return $sReturn
EndFunc


Func GetDatabaseAvailabilityStatus()
	Local $dbAvailable = False
	Local $sqlQuery = "select date 'Now' from rdb$database"
	Local $res = ExecuteSQL($sqlQuery)

	If Not IsArray($res) Or $res < 0 Then Return False

	Return True
EndFunc   ;==>GetDatabaseAvailabilityStatus


Func GetOptimalLabelWidth($text, $gui)
	Local $tempLabel = GUICtrlCreateLabel($text, 0, 0)
	Local $tempLabelPos = ControlGetPos($gui, "", $tempLabel)
	GUICtrlDelete($tempLabel)
	Return $tempLabelPos[2]
EndFunc   ;==>GetOptimalLabelWidth


Func GetFullDate($hour, $minute)
	Local $today = @YEAR & "/" & @MON & "/" & @MDAY
	Return $today & " " & $hour & ":" & $minute & ":00"
EndFunc   ;==>GetFullDate


Func GetAppointmentsForCurrentTime($array)
	If Not IsArray($array) Then Return

	Local $columnsQuantity = UBound($array, $UBOUND_COLUMNS)
	Local $rowsQuantity = UBound($array, $UBOUND_ROWS) - 1

	Local $retArray[0][$columnsQuantity]
	_ArrayColInsert($array, $columnsQuantity)

	For $i = 0 To $rowsQuantity
		Local $hour = $array[$i][2]
		If StringLen($hour) < 2 Then $hour = "0" & $hour

		Local $minute = $array[$i][3]
		If StringLen($minute) < 2 Then $minute = "0" & $minute

		Local $fullTime = GetFullDate($hour, $minute)
		$array[$i][$columnsQuantity] = _DateDiff('n', _NowCalc(), $fullTime)
		$array[$i][2] = $hour & ":" & $minute
	Next

	_ArraySort($array, 0, -1, -1, 2)
	_ArrayColDelete($array, 3)

	Local $previousAdded = False
	For $i = 0 To $rowsQuantity
		Local $currentRow = _ArrayExtract($array, $i, $i)
		Local $timeDiff = $currentRow[0][$columnsQuantity - 1]

		If $timeDiff < $timeBoundariesPast * -1 Then
			If $currentRow[0][7] Then ContinueLoop
		ElseIf $timeDiff > $timeBoundariesFuture And _
				$timeBoundariesFuture <> 0 Then
			If Not $i Or _
					Not $previousAdded Or _
					$timeDiff - $array[$i - 1][$columnsQuantity - 1] > _
					$timeBoundariesAcceptableDifferenceBetweenAppointments Then _
					ExitLoop
		EndIf

		_ArrayAdd($retArray, $currentRow)
		$previousAdded = True
	Next

	_ArrayColDelete($retArray, 7)

	For $i = 0 To UBound($retArray, $UBOUND_ROWS) - 1
		If $retArray[$i][$columnsQuantity - 2] < $timeBoundariesPast * -1 Then
			$retArray[$i][$columnsQuantity - 2] = 1
		Else
			$retArray[$i][$columnsQuantity - 2] = 0
		EndIf
	Next

	Return $retArray
EndFunc   ;==>GetAppointmentsForCurrentTime


Func GetTextFromIni($sectionName, $sql = False)
	Local $array = IniReadSection($iniFile, $sectionName)
	Local $tmp = ""
	Local $arrayRows = UBound($array, $UBOUND_ROWS) - 1

	For $i = 1 To $arrayRows
		$tmp &= $array[$i][1]
		If $i < $arrayRows Then $tmp &= $sql ? " " : @CRLF
	Next

	Return $tmp
EndFunc   ;==>GetTextFromIni




Func PrintAppontments($array, $name, $surname)
	ToLog("PrintAppontments")

	If Not IsArray($array) Or _
		Not UBound($array, $UBOUND_ROWS) Or _
		UBound($array, $UBOUND_COLUMNS) < 9 Then Return "wrong array format"

	Local $sPrinterStatus = GetPrinterStatus()
	If $sPrinterStatus Then Return $sPrinterStatus

	Local $dateRow = 4
	Local $nameRow = 5
	Local $familyRow = 6
	Local $formatStyle = 7
	Local $startRow = 9
	Local $worksheet = "Template"

	Local $templatePath = $resourcesPath & "PrintTemplate.xlsx"
	If Not FileExists($templatePath) Then Return "Template file not exist: " & $resourcesPath & "PrintTemplate.xlsx"

	If Not IsObj($oExcel) Then $oExcel = _Excel_Open(False, False, False, False, True)
	If Not IsObj($oExcel) Or @error Then Return "cannot connect to Excel instance, error code: " & @error

	Local $oBook = _Excel_BookOpen($oExcel, $templatePath)
	If Not IsObj($oBook) Or @error Then
		Local $tmp = ["$oExcel is not an object or not an application object", _
					  "Specified $sFilePath does not exist", _
					  "Unable to open $sFilePath. @extended is set to the COM error code " & _
					  "returned by the Open method"]
		Return "cannot open workbook " & $templatePath & ", " & $tmp[@error - 1] & ", error code: " & @error
	EndIf

	_Excel_RangeWrite($oBook, $worksheet, $name, "A" & $nameRow)
	If @error Then ExcelWriteErrorToLog(@error)

	_Excel_RangeWrite($oBook, $worksheet, $surname, "A" & $familyRow)
	If @error Then ExcelWriteErrorToLog(@error)

	_Excel_RangeWrite($oBook, $worksheet, @MDAY & "." & @MON & "." & @YEAR & _
			", " & @HOUR & ":" & @MIN, "A" & $dateRow)
	If @error Then ExcelWriteErrorToLog(@error)

	Local $needToPay = False
	Local $outOfTime = False
	Local $xray = False

	Local $currentRow = $startRow
	Local $maxElement = UBound($array, $UBOUND_ROWS) - 1
	If $maxElement > 5 Then $maxElement = 5

	For $i = 0 To $maxElement
		Local $timeAndCabinet = $array[$i][2] & ", кабинет " & $array[$i][5]
		Local $doc = $array[$i][3]
		Local $sTmp = $array[$i][4]
		Local $dept = StringLeft($sTmp, 1) & StringLower(StringRight($sTmp, StringLen($sTmp) - 1))

		_Excel_RangeWrite($oBook, $worksheet, $timeAndCabinet, "A" & $currentRow)
		If @error Then ExcelWriteErrorToLog(@error)

		_Excel_RangeWrite($oBook, $worksheet, $doc, "A" & $currentRow + 1)
		If @error Then ExcelWriteErrorToLog(@error)

		_Excel_RangeWrite($oBook, $worksheet, $dept, "A" & $currentRow + 2)
		If @error Then ExcelWriteErrorToLog(@error)

		Local $statusArray[3][2]
		$statusArray[0][0] = $array[$i][6]
		$statusArray[0][1] = $textPrintNotificationCash

		$statusArray[1][0] = $array[$i][7]
		$statusArray[1][1] = $textPrintNotificationXray

		$statusArray[2][0] = $array[$i][8]
		$statusArray[2][1] = $textPrintNotificationTime

		For $x = 0 To UBound($statusArray, $UBOUND_ROWS) - 1
			If Not $statusArray[$x][0] Or Not $statusArray[$x][1] Then ContinueLoop

			_Excel_RangeCopyPaste($oBook.ActiveSheet, _
								  $oBook.ActiveSheet.Range("A" & $formatStyle), _
								  $oBook.ActiveSheet.Range("A" & $currentRow + 3))
			If @error Then ExcelCopyPasteErrorToLog(@error)

			_Excel_RangeWrite($oBook, $worksheet, $statusArray[$x][1], "A" & $currentRow + 3)
			If @error Then ExcelWriteErrorToLog(@error)

			$currentRow += 1
		Next

		If $statusArray[0][0] Then $needToPay = True
		If $statusArray[1][0] Then $xray = True
		If $statusArray[2][0] Then $outOfTime = True

		If $i < $maxElement Then
			_Excel_RangeCopyPaste($oBook.ActiveSheet, _
					$oBook.ActiveSheet.Range("A" & $startRow - 1), _
					$oBook.ActiveSheet.Range("A" & $currentRow + 3))
			If @error Then ExcelCopyPasteErrorToLog(@error)

			_Excel_RangeCopyPaste($oBook.ActiveSheet, _
					$oBook.ActiveSheet.Range("A" & $startRow & ":A" & $startRow + 2), _
					$oBook.ActiveSheet.Range("A" & $currentRow + 4))
			If @error Then ExcelCopyPasteErrorToLog(@error)
		EndIf

		$currentRow += 4
	Next

	Local $finalText = ""

	If Not $outOfTime Then
		Local $hour = StringLeft($array[0][2], 2)
		Local $minute = StringRight($array[0][2], 2)
		Local $timeDiff = _DateDiff('n', _NowCalc(), GetFullDate($hour, $minute))

		If $timeDiff < 0 Then
			$finalText &= StringReplace($textPrintMessageTimeLate, "*", Abs($timeDiff)) & @CRLF
		Else
			$finalText &= StringReplace($textPrintMessageTimeOk, "*", $timeDiff) & @CRLF
		EndIf
	EndIf

	Local $nErrorsResult = Int($needToPay) + Int($outOfTime) + Int($xray)
	If $nErrorsResult > 1 Then
		$finalText &= $textPrintMessageFinalMultiple
	ElseIf $nErrorsResult = 1 Then
		If $needToPay Then $finalText &= $textPrintMessageFinalCash
		If $outOfTime Then $finalText &= $textPrintMessageFinalTime
		If $xray Then $finalText &= $textPrintMessageFinalXray
	Else
		$finalText &= $textPrintMessageFinalOk
	EndIf

	_Excel_RangeCopyPaste($oBook.ActiveSheet, _
			$oBook.ActiveSheet.Range("A" & $startRow - 1), _
			$oBook.ActiveSheet.Range("A" & $currentRow - 1))
	If @error Then ExcelCopyPasteErrorToLog(@error)

	_Excel_RangeCopyPaste($oBook.ActiveSheet, _
			$oBook.ActiveSheet.Range("A" & $startRow), _
			$oBook.ActiveSheet.Range("A" & $currentRow))
	If @error Then ExcelCopyPasteErrorToLog(@error)

	_Excel_RangeWrite($oBook, $worksheet, $finalText, "A" & $currentRow)
	If @error Then ExcelWriteErrorToLog(@error)

	_Excel_Print($oExcel, $oBook)
	If @error Then
		Local $tmp = ["$oExcel is not an object or not an application object", _
					  "$vObject is not an object or an invalid A1 range. @error is set to the COM error code", _
					  "Error printing the object. @extended is set to the COM error code"]
		Excel_BookClose($oBook)
		Return "cannot print workbook: " & $tmp[@error - 1] & ", error code: " & @error
	EndIf

	If Not FileExists($printedAppointmentListPath) Then DirCreate($printedAppointmentListPath)

	_Excel_BookSaveAs($oBook, $printedAppointmentListPath & $name & " " & $surname & " " & _
			@YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC)
	If @error Then
		Local $tmp = ["$oWorkbook is not an object", _
					  "$iFormat is not a number", _
					  "File exists, overwrite flag not set", _
					  "File exists but could not be deleted", _
					  "Error occurred when saving the workbook. @extended is set to the COM error " & _
					  "code returned by the SaveAs method."]

		Excel_BookClose($oBook)

		Return "cannot save workbook as: " & $printedAppointmentListPath & _
			", " & $tmp[@error - 1] & ", error code: " & @error
	EndIf

	Excel_BookClose($oBook)

	Return
EndFunc   ;==>PrintAppontments


Func Excel_Close()
	_Excel_Close($oExcel, False, True)
	If @error Then
		Local $tmp = ["$oExcel is not an object or not an application object", _
					  "Error returned by method Application.Quit. @extended is set to the COM error code", _
					  "Error returned by method Application.Save. @extended is set to the COM error code"]
		ToLog("!!! Error - cannot close excel application: " & $tmp[@error - 1] & ", error code: " & @error)
	EndIf
	If ProcessExists("EXCEL.exe") Then ProcessClose("EXCEL.exe")
EndFunc   ;==>Excel_Close


Func Excel_BookClose($oBook)
	_Excel_BookClose($oBook, False)
	If @error Then
		Local $tmp = ["$oWorkbook is not an object or not a workbook object", _
					  "Error occurred when saving the workbook. @extended is set to the COM error " & _
					  "code returned by the Save method", _
					  "Error occurred when closing the workbook. @extended is set to the COM error code " & _
					  "returned by the Close method"]
		ToLog("!!! Error - cannot close workbook: " & $tmp[@error - 1] & ", error code: " & @error)
	EndIf
EndFunc   ;==>Excel_BookClose


Func ExcelWriteErrorToLog($code)
	Local $tmp = ["$oWorkbook is not an object or not a workbook object", _
			      "$vWorksheet name or index are invalid or $vWorksheet is not a worksheet object. " & _
				  "@extended is set to the COM error code", _
				  "$vRange is invalid. @extended is set to the COM error code", _
				  "Error occurred when writing a single cell. @extended is set to the COM error code", _
				  "Error occurred when writing data using the _ArrayTranspose function. @extended is set " & _
				  "to the COM error code", _
				  "Error occurred when writing data using the transpose method. @extended is set to " & _
				  "the COM error code"]
	ToLog("!!! Error - " & $tmp[$code - 1] & ", error code: " & $code)
EndFunc   ;==>ExcelWriteErrorToLog


Func ExcelCopyPasteErrorToLog($code)
	Local $tmp = ["$oWorkbook is not an object or not a workbook object", _
				  "$vSourceRange is invalid. @extended is set to the COM error code", _
				  "$vTargetRange is invalid. @extended is set to the COM error code", _
				  "Error occurred when pasting cells. @extended is set to the COM error code", _
				  "Error occurred when cutting cells. @extended is set to the COM error code", _
				  "Error occurred when copying cells. @extended is set to the COM error code", _
				  "$vSourceRange and $vTargetRange can't be set to keyword Default at the same time"]
	ToLog("!!! Error - " & $tmp[$code - 1] & ", error code: " & $code)
EndFunc   ;==>ExcelCopyPasteErrorToLog


Func GetPrinterStatus()
	Local $wbemFlagReturnImmediately = 0x10
	Local $wbemFlagForwardOnly = 0x20
	Local $colItems = ""
	Local $strComputer = "localhost"

	Local $nPrinterState = 0
	Local $bPrinterWorkOffline = False

	Local $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	Local $sQuery = "SELECT * FROM Win32_Printer"
	Local $colItems = $objWMIService.ExecQuery($sQuery, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

	If Not IsObj($colItems) Then Return "Cannot get any printers from WMIService"

	For $objItem In $colItems
		If $objItem.Name <> $sPrinterName Then ContinueLoop
		$nPrinterState = $objItem.PrinterState
		$bPrinterWorkOffline = $objItem.WorkOffline
	Next

	If ($nPrinterState = 0 Or $nPrinterState = 131072 Or $nPrinterState = 131072 + 2048) And _
		Not $bPrinterWorkOffline Then Return

	Local $sPrinterStatus = ""

	If $bPrinterWorkOffline = True Then _
		$sPrinterStatus &= "Printer is working offline"

	If $nPrinterState > 0 Then
		For $i = UBound($aPrinterStatusCodes, $UBOUND_ROWS) - 1 To 1 Step -1
			Local $nCurrentState = $aPrinterStatusCodes[$i][0]
			Local $sCurrentState = $aPrinterStatusCodes[$i][1]

			If $nPrinterState - $nCurrentState < 0 Then ContinueLoop
			$nPrinterState -= $nCurrentState

			If $nCurrentState = 131072 Or $nCurrentState = 2048 Then ContinueLoop

			$sPrinterStatus &= @CRLF & $sCurrentState
		Next
	EndIf

	Return $sPrinterStatus
EndFunc




Func ExecuteSQL($sql)
;~ 	ConsoleWrite($sql & @CRLF)
	Local $sqlBD = "DRIVER=Firebird/InterBase(r) driver; UID=sysdba; PWD=masterkey; DBNAME=" & $infoclinicaDB & ";"
	Local $adoConnection = ObjCreate("ADODB.Connection")
	Local $adoRecords = ObjCreate("ADODB.Recordset")

	$adoConnection.Open($sqlBD)
	$adoRecords.CursorType = 2
	$adoRecords.LockType = 3

	If Not $adoConnection.State Then Return -1

	Local $result = 0

	If StringInStr(StringLower($sql), "update") Then
		$adoRecords = $adoConnection.Execute($sql)
	Else
		$adoRecords.Open($sql, $adoConnection)
		Local $result = $adoRecords.GetRows
;~ 		ConsoleWrite("array: " & _ArrayToString($result) & @CRLF)
;~ 		_ArrayDisplay($result)
		If $adoRecords.EOF = True And $adoRecords.BOF = True Then Return
	EndIf

	$adoConnection.Close
	$adoConnection = 0

	Return $result
EndFunc   ;==>ExecuteSQL


Func NumPressed($n, $id, $hGui)
	If StringLen($enteredCode) < 10 Then
		UpdateButtonBackgroundColor($id)
		$enteredCode &= $n
		UpdateInput($hGui)
	EndIf
EndFunc   ;==>NumPressed


Func ToLog($message)
	Local $logFilePath = $logsPath & @ScriptName & "_" & @YEAR & @MON & @MDAY & ".log"
	$message &= @CRLF
	ConsoleWrite($message)
	_FileWriteLog($logFilePath, $message)
EndFunc   ;==>ToLog


Func HandleComError()
;~ 	ToLog("error.description: " & @TAB & $oMyError.description & @CRLF & _
;~ 			"err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
;~ 			"err.number is: " & @TAB & Hex($oMyError.number, 8) & @CRLF & _
;~ 			"err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
;~ 			"err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
;~ 			"err.source is: " & @TAB & $oMyError.source & @CRLF & _
;~ 			"err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
;~ 			"err.helpcontext is: " & @TAB & $oMyError.helpcontext & @CRLF)
EndFunc   ;==>HandleComError


Func OnExit()
	Excel_Close()

	ToLog("-----Exit code: " & @exitCode & "-----")
	ToLog("-----Exit method: " & @exitMethod & "-----")
	Switch @exitMethod
		Case $EXITCLOSE_NORMAL
			ToLog("Natural closing.")
		Case $EXITCLOSE_BYEXIT
			ToLog("close by Exit function.")
		Case $EXITCLOSE_BYCLICK
			ToLog("close by clicking on exit of the systray.")
		Case $EXITCLOSE_BYLOGOFF
			ToLog("close by user logoff.")
		Case $EXITCLOSE_BYSUTDOWN
			ToLog("close by Windows shutdown.")
	EndSwitch

;~ 	Local $sFileName = $logsPath & "Exit_screenshot_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".jpg"
;~ 	_ScreenCapture_SetJPGQuality(30)
;~ 	Local $hScreenshot = _ScreenCapture_Capture($sFileName)
;~ 	_WinAPI_DeleteObject($hScreenshot)
	SendEmail("-----Exiting----- " & @exitMethod);, $sFileName)
EndFunc   ;==>OnExit




Func SendEmail($messageToSend, $sAttachments = "", $bError = False, $bRegistry = False)
	If $bDebug Then Return
	If Not $sMailSend Then Return

	Local $sCurrentTime = @HOUR & ":" & @MIN
	If $bError And ($sCurrentTime < $sMailWorkingHoursBegins Or _
		$sCurrentTime > $sMailWorkingHoursEnds) Then Return

	ToLog(@CRLF & "-----Sending email-----")
	ToLog($messageToSend)
	ToLog(@CRLF & "-----------------------")

	Local $sCurrentCompName = @ComputerName
	Local $title = "Уведомление от инфомата"
	$messageToSend &= @CRLF & @CRLF & _
			"---------------------------------------" & @CRLF & _
			"Это автоматическое сообщение." & @CRLF & _
			"Пожалуйста, не отвечайте на него." & @CRLF & _
			"Имя системы: " & $sCurrentCompName

	Local $sMailReceiver = $sMailDeveloperAddress
	If $bError Then $sMailReceiver = $sMailTo
	If $bRegistry Then $sMailReceiver = $sMailRegistryAddress

	If Not _INetSmtpMailCom($sMailServer, $sCurrentCompName, $sMailLogin, $sMailReceiver, _
			$title, $messageToSend, $sAttachments, $sMailDeveloperAddress, "", $sMailLogin, $sMailPassword) Then

			$messageToSend &= @CRLF & @CRLF & $errStr & "Using backed up email settings"
			_INetSmtpMailCom($sMailBackupServer, $sCurrentCompName, $sMailBackupLogin, $sMailBackupTo, _
			$title, $messageToSend, $sAttachments, $sMailDeveloperAddress, "", $sMailBackupLogin, $sMailBackupPassword)
	EndIf
EndFunc   ;==>SendEmail


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", _
	$as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", _
	$s_Password = "", $IPPort = 25, $ssl = 0)

	Local $objEmail = ObjCreate("CDO.Message")
	Local $i_Error = 0
	Local $i_Error_desciption = ""

	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress

	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

	$objEmail.Subject = $s_Subject

	If $s_AttachFiles <> "" Then
		Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
		For $x = 1 To $S_Files2Attach[0]
			$S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
			If FileExists($S_Files2Attach[$x]) Then
				$objEmail.AddAttachment($S_Files2Attach[$x])
			Else
				$i_Error_desciption = $i_Error_desciption & @LF & 'File not found to attach: ' & $S_Files2Attach[$x]
				$as_Body &= $i_Error_desciption & @CRLF
			EndIf
		Next
	EndIf

	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
		$objEmail.TextBodyPart.Charset = "utf-8"
	EndIf

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf

	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf

	$objEmail.Configuration.Fields.Update
	$objEmail.Send

	If @error Then Return False
	Return True
EndFunc   ;==>_INetSmtpMailCom
