#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\icon.ico
#pragma compile(ProductVersion, 1.1)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Приложения для инфомата для самостоятельной отметки о посещении)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555 - nn-admin@nnkk.budzdorov.su)
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
#include <Excel.au3>
#include <GuiButton.au3>
#include <AutoItConstants.au3>



#Region ====================== Variables ======================
Local $oMyError = ObjEvent("AutoIt.Error", "HandleComError")
OnAutoItExitRegister("OnExit")
Local $iniFile = @ScriptDir & "\InfomatSelfChecking.ini"
Local $generalSectionName = "general"

Local $error

Local $headerColor = 0x4e9b44
Local $okButtonColor = 0x4e9b44
Local $okButtonPressedColor = 0x43853a
Local $mainButtonColor = 0xe0e0e0
Local $mainButtonPressedColor = 0xd6d6d6
Local $disabledColor = 0xdfdfdf
Local $disabledTextColor = 0xa5a5a5
Local $textColor = 0x2c3d3f
Local $alternateTextColor = 0xffffff
Local $mainBackgroundColor = 0xffffff
Local $errorTitleColor = 0xf98d3c

Local $bottonLineHeight = 11

Local $mainFontName = "Franklin Gothic"

Local $dX = @DesktopWidth
Local $dY = @DesktopHeight
;~ Local $dX = 1280;1024
;~ Local $dY = 1024;819

Local $numButSize = $dY / 10
Local $distBt = $numButSize / 3

Local $btFontSize = $numButSize / 3
Local $btWeight = $FW_BOLD
Local $btQual = $CLEARTYPE_QUALITY

Local $headerHeight = $numButSize * 1.5
Local $headerLabelFontWeight = $FW_SEMIBOLD

Local $initX = $dX / 2 - $numButSize * 1.5 - $distBt
Local $initY = $dY / 2 - $numButSize * 1.5 - $distBt

Local $timeLabel = ""

Local $infoclinicaDB = IniRead($iniFile, $generalSectionName, "InfoclinicaDatabaseAddress", "")
Local $enteredCode = "";"9601811873"

Local $bottomAppointmentTimeBoundaries = -15
Local $topAppointmentTimeBoundaries = 180

Local $formMaxTimeWait = IniRead($iniFile, $generalSectionName, "FormMaxTimeWaitInSeconds", 30) * 1000

Local $resourcesPath = @ScriptDir & "\Resources\"
Local $printedAppointmentListPath = @ScriptDir & "\Printed Appointments List\"
Local $logsPath = @ScriptDir & "\Logs\"

Local $pressedButtonTimeCounter = 0
Local $previousButtonPressedID[] = [0, 0]

Local $prevMinute = @MIN
Local $timer = 0
Local $timeCounter = 0

Local $mainGui = 0
Local $bt_next = 0
Local $inp_pincode = 0
#EndRegion


FormMainGui()




Func FormMainGui()
	_WinAPI_ShowCursor(False)
	$mainGui = GUICreate("SelfChecking", $dX, $dY, 0, 0, $WS_POPUP, $WS_EX_TOPMOST)

	Local $text = "Для отметки о посещении" & @CRLF & "введите Ваш номер мобильного телефона"
	CreateStandardDesign($mainGui, $text, False, True)

	Local $bt_1 = CreateButton("1", $initX, $initY, $numButSize, $numButSize)

	Local $prevBt
	$prevBt = ControlGetPos($mainGui, "", $bt_1)
	Local $bt_2 = CreateButton("2", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_2)
	Local $bt_3 = CreateButton("3", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_1)
	Local $bt_4 = CreateButton("4", $prevBt[0], $prevBt[1] + $prevBt[3] + $distBt, $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_4)
	Local $bt_5 = CreateButton("5", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_5)
	Local $bt_6 = CreateButton("6", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_4)
	Local $bt_7 = CreateButton("7", $prevBt[0], $prevBt[1] + $prevBt[3] + $distBt, $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_7)
	Local $bt_8 = CreateButton("8", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_8)
	Local $bt_9 = CreateButton("9", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_7)
	Local $bt_clear = CreateButton("C", $prevBt[0], $prevBt[1] + $prevBt[3] + $distBt, $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_clear)
	Local $bt_0 = CreateButton("0", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_0)
	Local $bt_backspace = CreateButton("<", $prevBt[0] + $prevBt[2] + $distBt, $prevBt[1], $numButSize, $numButSize)

	$prevBt = ControlGetPos($mainGui, "", $bt_clear)
	$bt_next = CreateButton("Продолжить", $prevBt[0], $prevBt[3] + $distBt + $prevBt[1], _
		$numButSize * 3 + $distBt * 2, $numButSize, $disabledColor)
	GUICtrlSetColor(-1, $alternateTextColor)

	$prevBt = ControlGetPos($mainGui, "", $bt_next)
	Local $prevBt2 = ControlGetPos($mainGui, "", $bt_1)
	$inp_pincode = GUICtrlCreateLabel($enteredCode, $dX / 2 - $prevBt[2] * 2.3 / 2, _
		$prevBt2[1] - $prevBt2[3] - $distBt , $prevBt[2] * 2.3, $prevBt[3], BitOr($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetFont(-1, $btFontSize * 1.8)
	GUICtrlSetColor(-1, $textColor)

	UpdateTimeLabel()
	UpdateInput()

	GUISetState(@SW_SHOW)

	ToLog("-----MainGui started-----")
	SendEmail("-----MainGui started-----")

	While 1
		If $enteredCode Then $timer = _Timer_Init()

		$nMsg = GUIGetMsg()

		If $timeCounter > $formMaxTimeWait Then
			ToLog("MainGui force clear")
			$nMsg = $bt_clear
			$timeCounter = 0
			$timer = 0
		EndIf

		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $bt_0
				NumPressed(0, $bt_0)
			Case $bt_1
				NumPressed(1, $bt_1)
			Case $bt_2
				NumPressed(2, $bt_2)
			Case $bt_3
				NumPressed(3, $bt_3)
			Case $bt_4
				NumPressed(4, $bt_4)
			Case $bt_5
				NumPressed(5, $bt_5)
			Case $bt_6
				NumPressed(6, $bt_6)
			Case $bt_7
				NumPressed(7, $bt_7)
			Case $bt_8
				NumPressed(8, $bt_8)
			Case $bt_9
				NumPressed(9, $bt_9)

			Case $bt_next
				Local $tempTimeLabel = $timeLabel
				If StringLen($enteredCode) < 10 Then ContinueLoop
				_Timer_KillAllTimers($mainGui)
				$timeCounter = 0
				$timer = 0

				FormCheckEnteredNumber($enteredCode)

				$enteredCode = ""
				UpdateInput()
				_Timer_KillAllTimers($mainGui)
				$timeLabel = $tempTimeLabel
				UpdateTimeLabel()

			Case $bt_backspace
				UpdateButtonBackgroundColor($bt_backspace)
				If StringLen($enteredCode) > 0 Then
					$enteredCode = StringLeft($enteredCode, StringLen($enteredCode) - 1)
					UpdateInput()
				EndIf

			Case $bt_clear
				UpdateButtonBackgroundColor($bt_clear)
				$enteredCode = ""
				UpdateInput()
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

		If $enteredCode Then
			If $timer Then
				Local $timeDiff = _Timer_Diff($timer)
				$timeCounter += $timeDiff
				_Timer_KillAllTimers($mainGui)
				$timer = 0
			EndIf
		EndIf

		If @MIN <> $prevMinute Then
			Local $res = GetDatabaseAvailabilityStatus()
			If Not $res Then
				Local $errorMessage = "К сожалению," & @CRLF & _
									  "cервис временно недоступен" & @CRLF & @CRLF & _
									  "Попробуйте позднее" & @CRLF & _
									  "или обратитесь на регистратуру"
				FormShowMessage("", $errorMessage, True, True)
			EndIf

			UpdateTimeLabel()
			$prevMinute = @MIN
		EndIf
	WEnd
EndFunc


Func FormCheckEnteredNumber($code)
	ToLog("FormCheckEnteredNumber: " & $code)
	Local $phoneNumberPrefix = StringLeft($code, 3)
	Local $phoneNumber = StringRight($code, 7)
	Local $sqlQuery = 	"Select Distinct Cl.PCode, Cl.FirstName, Cl.MidName, Cl.BDate " & _
						"From Schedule S " & _
						"Join Clients Cl On Cl.PCode = S.PCode " & _
						"Join ClPhones Clp On Clp.PCode = S.PCode " & _
						"Where S.WorkDate = 'today' " & _
						"And Clp.PhPrefix Containing '" & $phoneNumberPrefix & "' " & _
						"And Clp.Phone Containing '" & $phoneNumber & "'"

	Local $res = ExecuteSql($sqlQuery)

	Local $textPhoneNumber = "+7 (" & $phoneNumberPrefix & ") " & StringLeft($phoneNumber, 3) & _
		"-" & StringMid($phoneNumber, 4, 2) & "-" & StringRight($phoneNumber, 2)
	Local $errorMessage = ""
	Local $checkDb = False

	If $res = 0 Or UBound($res, $UBOUND_ROWS) > 1 Then
		$errorMessage &= "К сожалению, по номеру " & $textPhoneNumber & @CRLF & _
						 "не найдено записей на ближайшее время" & @CRLF & @CRLF & @CRLF & _
						 "Возможно указан неверный номер" & @CRLF & @CRLF & _
						 "Попробуйте снова" & @CRLF & _
						 "или обратитесь на регистратуру"
	ElseIf $res = -1 Then
		$errorMessage &= "К сожалению," & @CRLF & _
						 "в данный момент сервис недоступен" & @CRLF & @CRLF & @CRLF & _
						 "Попробуйте позднее" & @CRLF & _
						 "или обратитесь на регистратуру"
		$checkDb = True
	EndIf

	If $errorMessage Then
		FormShowMessage(0, $errorMessage, True, $checkDb)
		Return
	EndIf

	Local $fioForm = GUICreate("FIO", $dX, $dY, 0, 0, $WS_POPUP, $WS_EX_TOPMOST)

	Local $titleText = "Пожалуйста," & @CRLF & "убедитесь в соответствии Ваших данных:"
	CreateStandardDesign($fioForm, $titleText, False)

	Local $date = StringMid($res[0][3], 7, 2) & "." & StringMid($res[0][3], 5, 2) & "." & StringLeft($res[0][3], 4)
	Local $fullName = $res[0][1] & " " & $res[0][2]
	Local $mainText = $fullName & @CRLF & @CRLF & "Дата рождения: " & $date

	CreateLabel($mainText, 0, $dY * 0.3, $dX, $dY * 0.4, $textColor, $GUI_BKCOLOR_TRANSPARENT, $fioForm, $btFontSize * 1.2)

	$prevBt = ControlGetPos($mainGui, "", $bt_next)
	Local $bt_ok = CreateButton("Продолжить", $dx - $distBt - $prevBt[2], $prevBt[1], $prevBt[2], $prevBt[3], $okButtonColor)
	GUICtrlSetColor(-1, $alternateTextColor)

	Local $bt_not = CreateButton("Неверно", 0 + $distBt, $prevBt[1], $prevBt[2], $prevBt[3])

	UpdateTimeLabel()

	GUISetState()

	$timeCounter = 0

	While 1
		$timer = _Timer_Init()

		If $timeCounter > $formMaxTimeWait Then
			ToLog("FormCheckEnteredNumber force close" & @CRLF)
			GUIDelete($fioForm)
			Return
		EndIf

		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $bt_not
				ToLog("Fullname not correct: " & $fullName)
				$errorMessage = "Возможно указан неверный номер" & @CRLF & @CRLF & _
								"Попробуйте снова" & @CRLF & _
								"или обратитесь на регистратуру"
				FormShowMessage($fioForm, $errorMessage)
				Return
			Case $bt_ok
				FormShowAppointments($fioForm, $res[0][0], $res[0][1], $res[0][2])
				Return
		EndSwitch

		Sleep(20)

		Local $timeDiff = _Timer_Diff($timer)
		$timeCounter += $timeDiff
		_Timer_KillAllTimers($mainGui)
		$timer = 0

		If @MIN <> $prevMinute Then
			UpdateTimeLabel()
			$prevMinute = @MIN
		EndIf
	WEnd
EndFunc


Func FormShowAppointments($guiToDelete, $patientID, $name, $surname)
	Local $fullName = $name & " " & $surname
	ToLog("FormShowAppointments: " & $fullName)

	Local $sqlQuery = "Select Sch.SchedId, Sch.WorkDate, Sch.BHour, Sch.BMin, D.DName, Dep.DepName, R.RNum, " & _
					"Case (Case " & _
					"  When Sch.SectId Is Not Null And Sch.SectId != 0 Then Sch.SectId " & _
					"  Else (Select SectId From Clients Where PCode = Sch.PCode) " & _
					"End) " & _
					"When 4363 Then 1 " & _
					"When 991139394 Then 1 " & _
					"Else 0 " & _
					"End As Kateg, Iif(CoalEsce(Sch.ClVisit,0)=0,0,1) " & _
					"From Schedule Sch " & _
					"Join Doctor D On D.DCode = Sch.DCode " & _
					"Join DoctShedule Ds On Ds.DCode = Sch.DCode " & _
					"And Ds.SchedIdent = Sch.SchedIdent " & _
					"Join Departments Dep On Dep.DepNum = Ds.DepNum " & _
					"Join Chairs Ch On Ch.ChId = Ds.Chair " & _
					"Join Rooms R On R.RId = Ch.RId Where Sch.WorkDate = 'today' " & _
					"And Sch.PCode = " & $patientID


	Local $res = ExecuteSQL($sqlQuery)

	$res = GetAppointmentsForCurrentTime($res)

	If Not IsArray($res) Or Not UBound($res, $UBOUND_ROWS) Then
		Local $errorMessage = "На ближайшее время у Вас нет назначений" & @CRLF & @CRLF & _
							  "За подробной информацией Вы можете" & @CRLF & _
							  "обратиться на регистратуру"
		FormShowMessage($guiToDelete, $errorMessage)
		Return
	EndIf

	Local $destForm = GUICreate("FormShowAppointments", $dX, $dY, 0, 0, $WS_POPUP, $WS_EX_TOPMOST)
	Local $text = $fullName & "," & @CRLF & "Ваши записи на ближайшее время:"
	CreateStandardDesign($destForm, $text, False)

	$prevBt = ControlGetPos($mainGui, "", $bt_next)

	Local $bt_close = CreateButton("Закрыть", _
		0 + $distBt, _
		$prevBt[1], _
		$prevBt[2], _
		$prevBt[3])

	Local $bt_print = CreateButton("Распечатать", _
		$dX - $distBt - $prevBt[2], _
		$prevBt[1], _
		$prevBt[2], _
		$prevBt[3], _
		$okButtonColor, _
		$alternateTextColor)

	GUISetFont($btFontSize * 0.9, $FW_NORMAL, -1, "Franklin Gothic Book", $destForm, $btQual)
	Local $needRegistry = CreateAppointmentsTable($res, $destForm)

	UpdateTimeLabel()
	GUISetState()

	Sleep(10)

	If $guiToDelete Then GUIDelete($guiToDelete)

	Local $needToClose = False
	Local $textToShow = ""

	If $needRegistry Then
		$textToShow &= "Для отметки о посещении" & @CRLF & "просьба пройти на регистратуру"
	Else
		$textToShow &= "Отметка о посещении успешно проставлена" & @CRLF & @CRLF & _
					   "Пожалуйста пройдите на прием"
	EndIf

	$timeCounter = 0

	While True
		$timer = _Timer_Init()

		If $timeCounter > $formMaxTimeWait Then
			ToLog("FormShowAppointments force close")
			$needToClose = True
		EndIf

		$nMsg = GUIGetMsg()

		Switch $nMsg
			Case $bt_close
				ToLog("FormShowAppointments close")
				$needToClose = True
			Case $bt_print
				Local $printResult = PrintAppontments($res, $name, $surname)
				If $printResult Then
					$textToShow = "Список назначений успешно распечатан" & @CRLF & @CRLF & $textToShow
				Else
					$textToShow = "К сожалению, по техническим причинам" & @CRLF & _
						"не удалось распечатать список назначений" & @CRLF & _
						"Информация о проблеме будет передана" & @CRLF & _
						"ответственным лицам" & @CRLF & @CRLF & $textToShow
					SendEmail($textToShow)
				EndIf
				$needToClose = True
		EndSwitch

		If $needToClose Then
			FormShowMessage($destForm, $textToShow, False)
			Return
		EndIf

		Sleep(20)

		Local $timeDiff = _Timer_Diff($timer)
		$timeCounter += $timeDiff
		_Timer_KillAllTimers($mainGui)
		$timer = 0

		If @MIN <> $prevMinute Then
			UpdateTimeLabel()
			$prevMinute = @MIN
		EndIf
   WEnd
EndFunc


Func FormShowMessage($guiToDelete, $message, $showError = True, $checkDb = False)
	ToLog("FormShowMessage: " & StringReplace($message, @CRLF, " | "))

	Local $nanForm = GUICreate("FormShowMessage", $dX, $dY, 0, 0, $WS_POPUP, $WS_EX_TOPMOST)
	Local $text = "Уважаемый пациент!"
	CreateStandardDesign($nanForm, $text, $showError, True)

	$prevBt = ControlGetPos($mainGui, "", $bt_next)
	Local $bt_close = 666
	If Not $checkDb Then $bt_close = CreateButton("Закрыть", $prevBt[0], $prevBt[1], $prevBt[2], $prevBt[3])

	Local $x = 0
	Local $y = $dY * 0.3
	Local $sizeX = $dX
	Local $sizeY = $dY * 0.4
	CreateLabel($message, $x, $y, $sizeX, $sizeY, $textColor, $GUI_BKCOLOR_TRANSPARENT, $nanForm, $btFontSize * 1.2)

	UpdateTimeLabel()
	GUISetState()

	Sleep(10)

	If $guiToDelete Then GUIDelete($guiToDelete)

	$timeCounter = 0

	While 1
		If Not $checkDb Then $timer = _Timer_Init()

		If $timeCounter > $formMaxTimeWait Then
			ToLog("FormShowMessage force close" & @CRLF)
			GUIDelete($nanForm)
			Return
		EndIf

		$nMsg = GUIGetMsg()

		If @MIN <> $prevMinute Then
			UpdateTimeLabel()
			$prevMinute = @MIN
			If $checkDb Then
				Local $res = GetDatabaseAvailabilityStatus()
				If $res Then $nMsg = $bt_close
			EndIf
		EndIf

		Switch $nMsg
		Case $bt_close
			ToLog("FormShowMessage close" & @CRLF)
			GUIDelete($nanForm)
			Return
		EndSwitch

		If Not $checkDb Then
			Local $timeDiff = _Timer_Diff($timer)
			$timeCounter += $timeDiff
			_Timer_KillAllTimers($mainGui)
			$timer = 0
		EndIf
	WEnd
EndFunc




Func CreateStandardDesign($gui, $titleText, $isError, $trademark = False)
	GUISetBkColor($mainBackgroundColor)
	GUISetFont($btFontSize, $btWeight, 0, $mainFontName, $gui, $btQual)

	Local $titleColor = $headerColor
	If $isError Then $titleColor = $errorTitleColor

	CreateLabel($titleText, 0, 0, $dX, $headerHeight, $alternateTextColor, $titleColor, $gui)

	Local $timeScopeWidth = $numButSize * 1.7
	Local $timeIconWidth = 20

	GUISetFont($btFontSize * 0.7)
	$timeLabel = GuiCtrlCreateLabel("13:14", 0, 0, -1, -1, BitOr($SS_CENTER, $SS_CENTERIMAGE))
	Local $timeLabelPosition = ControlGetPos($gui, "", $timeLabel)
	GUICtrlSetPos($timeLabel, $dX - $timeLabelPosition[2] - $distBt / 4, $distBt / 4)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, $alternateTextColor)
	GUISetFont($btFontSize)

	$timeLabelPosition = ControlGetPos($gui, "", $timeLabel)
	Local $timePic = CreatePngControl($resourcesPath & "TimeIcon.png", $timeIconWidth, $timeIconWidth)
	GUICtrlSetPos($timePic, $timeLabelPosition[0] - $distBt / 4 - $timeIconWidth, _
		$timeLabelPosition[1] + $timeLabelPosition[3] / 2 - $timeIconWidth / 2)

	GUICtrlCreatePic($resourcesPath & "PicBottomLine.jpg", 0, $dY - $bottonLineHeight, $dX, $bottonLineHeight)

	If $trademark Then
		Local $trademarkWidth = 159
		Local $trademarkHeight = 170
		GUICtrlCreatePic($resourcesPath & "PicButterfly.jpg", $dX - $trademarkWidth - $distBt / 2, _
			$dY - $trademarkHeight - $bottonLineHeight - $distBt / 2, $trademarkWidth, $trademarkHeight)
	EndIf
EndFunc


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
EndFunc


Func CreateLabel($text, $x, $y, $width, $height, $textColor, $backgroundColor, $gui, $fontSize = $btFontSize)
	$Label1 = GUICtrlCreateLabel("", $x, $y, $width, $height)
	GUICtrlSetBkColor(-1, $backgroundColor)

	GUISetFont($fontSize)
	Local $label = GUICtrlCreateLabel($text, 0, 0, -1, -1, $SS_CENTER)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, $textColor)

	Local $position = ControlGetPos($gui, "", $label)
	If IsArray($position) Then
		Local $newX = $x + ($width - $position[2]) / 2
		Local $newY = $y + ($height - $position[3]) / 2
		GUICtrlSetPos($label, $newX, $newY)
	EndIf
	GUISetFont($btFontSize)
EndFunc


Func CreateButton($text, $x, $y, $width, $height, $bkColor = $mainButtonColor, $color = $textColor)
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
EndFunc


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

	Local $maxSymbols = Round($sizes[2] / ($btFontSize * 0.8))

	GUICtrlCreateLabel("", $startX, $startY, $totalWidth + $distance * 3, $height - $distance)
	GUICtrlSetBkColor(-1, $mainButtonColor)

	Local $arraySize = UBound($head, $UBOUND_ROWS) - 1
	If $arraySize > 6 Then $arraySize = 6

	Local $showCashWarning = False
	Local $showOutOfTimeWarning = False
	Local $showXrayWarning = False

	For $i = 0 To $arraySize
		Local $currentRow[4]
		Local $cash = $head[$i][6]
		Local $time = $head[$i][7]
		Local $xray = 0
;~ 		If $i < 2 Then $xray = False

;~ 		ToLog($cash & $time & $xray)

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
			Local $attribute = BitOr($SS_CENTER, $SS_CENTERIMAGE)
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
				GUICtrlSetBkColor(-1, $mainButtonColor)
			EndIf

			$currentX += $sizes[$x] + $distance + Round($distBt / 2)

			If $time Then
				$rubleIcon = CreatePngControl($resourcesPath & "OutOfTimeIcon.png", $iconWidth, $iconWidth)
				GUICtrlSetPos(-1, $dX - Round($distBt * 1.5) - $iconWidth, $currentY + Round($height * 0.2))
				$showOutOfTimeWarning = True
			EndIf

			If $cash Then
				$rubleIcon = CreatePngControl($resourcesPath & "RubleIcon.png", $iconWidth, $iconWidth)
				GUICtrlSetPos(-1, $dX - Round($distBt * 1.5) - $iconWidth - $iconWidth * $time * 1.2, _
					$currentY + Round($height * 0.2))
				$showCashWarning = True
			EndIf

			If $xray Then
				$rubleIcon = CreatePngControl($resourcesPath & "XrayIcon.png", $iconWidth, $iconWidth)
				GUICtrlSetPos(-1, $dX - Round($distBt * 1.5) - $iconWidth - $iconWidth * ($time + $cash) * 1.2, _
					$currentY + Round($height * 0.2))
				$showXrayWarning = True
			EndIf
		Next

		$currentY += $height + $distance
		$currentX = $startX + Round($distBt / 2)

		If $i = 0 Then
			$currentY -= $distance
		EndIf
	Next

	If $showCashWarning Or $showOutOfTimeWarning Or $showXrayWarning Then
		Local $message = ""
		Local $textHeight = 1.5
		If $showCashWarning And Not $showOutOfTimeWarning Then
			$message = "Имеются записи, запланированные за наличный расчет" & @CRLF & _
					   "Необходимо заранее оплатить данные приемы"
		ElseIf $showOutOfTimeWarning And Not $showCashWarning Then
			$message = "Имеются записи с пропущенным временем начала" & @CRLF & _
					   "Необходимо согласовать перенос на другое время"
		Else
			$message = "Для отметки о посещении необходимо обратиться в регистратуру"
			$textHeight = 1.0
		EndIf


		CreateLabel($message, _
			$startX, _
			$currentY, _
			$totalWidth + $distance * 3, _
			Round($height * $textHeight), _
			$alternateTextColor, _
			$errorTitleColor, _
			$gui, _
			Round($btFontSize * 0.8))

		 $currentY += $height + $distance

		 GUISetFont($btFontSize * 0.8)

;~ 		If $showOutOfTimeWarning Then
;~ 			$rubleIcon = CreatePngControl($resourcesPath & "OutOfTimeIcon.png", $iconWidth, $iconWidth)
;~ 			GUICtrlSetPos(-1, $startX, $currentY)
;~ 			Local $tmp = GUICtrlCreateLabel(" - пропущено время", _
;~ 				$startX + $iconWidth, _
;~ 				$currentY, _
;~ 				-1, _
;~ 				$iconWidth, _
;~ 				$SS_CENTERIMAGE)
;~ 			Local $tmp2 = ControlGetPos($gui, -1, $tmp)
;~ 			$startX = $tmp2[0] + $tmp2[2] + $distance * 3
;~ 		EndIf

;~ 		If $showCashWarning Then
;~ 			$rubleIcon = CreatePngControl($resourcesPath & "RubleIcon.png", $iconWidth, $iconWidth)
;~ 			GUICtrlSetPos(-1, $startX, $currentY)
;~ 			Local $tmp = GUICtrlCreateLabel(" - наличный расчет", _
;~ 				$startX + $iconWidth, _
;~ 				$currentY, _
;~ 				-1, _
;~ 				$iconWidth, _
;~ 				$SS_CENTERIMAGE)
;~ 			Local $tmp2 = ControlGetPos($gui, -1, $tmp)
;~ 			$startX = $tmp2[0] + $tmp2[2] + $distance * 3
;~ 		EndIf

;~ 		If $showXrayWarning Then
;~ 			$rubleIcon = CreatePngControl($resourcesPath & "XrayIcon.png", $iconWidth, $iconWidth)
;~ 			GUICtrlSetPos(-1, $startX, $currentY)
;~ 			Local $tmp = GUICtrlCreateLabel(" - лучевое отделение", _
;~ 				$startX + $iconWidth, _
;~ 				$currentY, _
;~ 				-1, _
;~ 				$iconWidth, _
;~ 				$SS_CENTERIMAGE)
;~ 			Local $tmp2 = ControlGetPos($gui, -1, $tmp)
;~ 			$startX = $tmp2[0] + $tmp2[2] + $distance * 3
;~ 		EndIf
	Else
		For $i = 1 To $arraySize
			Local $idToUpdate = $head[$i][0]
			Local $updateSql = "Update Schedule Set ScreenVisit = 1, ClVisit = 1, VisitTime = 'now', " & _
							   "Comment=(Select CoalEsce(Comment,'') From Schedule Where SchedId = " & $idToUpdate & ")||' " & _
							   "ИНФОМАТ' Where CoalEsce(ClVisit,0) = 0 And SchedId = " & $idToUpdate

			ExecuteSQL($updateSql)
			ToLog("Setting visit mark for: " & $idToUpdate)
		Next
	EndIf

	Return $showCashWarning Or $showOutOfTimeWarning Or $showXrayWarning
EndFunc





Func UpdateButtonBackgroundColor($id, $bkColor = $mainButtonColor, $glowColor = $mainButtonPressedColor)
	If $enteredCode Then $timeCounter = 0

	If $previousButtonPressedID[0] Then
		GUICtrlSetBkColor($previousButtonPressedID[0], $previousButtonPressedID[1])
		$pressedButtonTimeCounter = 0
	EndIf

	GUICtrlSetBkColor($id, $glowColor)
	$previousButtonPressedID[0] = $id
	$previousButtonPressedID[1] = $bkColor
	$pressedButtonTimeCounter = 1
EndFunc


Func UpdateInput()
	Local $format = "+7 (___) ___-__-__"

	Local $codeLenght = StringLen($enteredCode)
	If $codeLenght = 0 Or $codeLenght = 9 Then
		GUICtrlSetColor($bt_next, $disabledTextColor)
		GUICtrlSetBkColor($bt_next, $disabledColor)
	ElseIf $codeLenght = 10 Then
		GUICtrlSetColor($bt_next, $alternateTextColor)
		GUICtrlSetBkColor($bt_next, $okButtonColor)
	EndIf

	For $i = 1 To $codeLenght
		$format = StringReplace($format, "_", StringMid($enteredCode, $i, 1), 1)
	Next

	ControlSetText($mainGui, "", $inp_pincode, $format)
EndFunc


Func UpdateTimeLabel()
	Local $newTime = @HOUR & ":" & @MIN
	GUICtrlSetData($timeLabel, $newTime)
EndFunc





Func GetDatabaseAvailabilityStatus()
	Local $dbAvailable = False
	Local $sqlQuery = "select date 'Now' from rdb$database"
	Local $res = ExecuteSql($sqlQuery)

	If Not IsArray($res) Or $res < 0 Then Return False

	Return True
EndFunc


Func GetOptimalLabelWidth($text, $gui)
	Local $tempLabel = GUICtrlCreateLabel($text, 0, 0)
	Local $tempLabelPos = ControlGetPos($gui, "", $tempLabel)
	GUICtrlDelete($tempLabel)
	Return $tempLabelPos[2]
EndFunc


Func GetFullDate($hour, $minute)
	If $hour < 10 Then $hour = "0" & $hour
	If $minute < 10 Then $minute = "0" & $minute
	Local $today = @YEAR & "/" & @MON & "/" & @MDAY
	Return $today & " " & $hour & ":" & $minute & ":00"
EndFunc


Func GetAppointmentsForCurrentTime($array)
	If Not IsArray($array) Then Return

	Local $retArray[0][UBound($array, $UBOUND_COLUMNS)]

	For $i = 0 To UBound($array, $UBOUND_ROWS) - 1
		Local $currentRow = _ArrayExtract($array, $i, $i)

		Local $hour = $currentRow[0][2]
		If StringLen($hour) < 2 Then $hour = "0" & $hour

		Local $minute = $currentRow[0][3]
		If StringLen($minute) < 2 Then $minute = "0" & $minute

		Local $fullTime = GetFullDate($hour, $minute)
		Local $timeDiff = _DateDiff('n', _NowCalc(), $fullTime)

		If $timeDiff < $bottomAppointmentTimeBoundaries Then
			If $array[$i][8] Then
				ContinueLoop
			Else
				$currentRow[0][8] = 1
			EndIf
		ElseIf $timeDiff > $topAppointmentTimeBoundaries Then
			ContinueLoop
		EndIf

		$currentRow[0][2] = $hour & ":" & $minute
		_ArrayAdd($retArray, $currentRow)
	Next

	_ArraySort($retArray, 0, -1, -1, 2)
	_ArrayColDelete($retArray, 3)

	Return $retArray
EndFunc




Func PrintAppontments($array, $name, $surname)
	ToLog("PrintAppontments")
	Local $err = "!!! Error: "
	If Not IsArray($array) Or _
		Not UBound($array, $UBOUND_ROWS) Or _
		UBound($array, $UBOUND_COLUMNS) < 5 Then
		ToLog($err & "wrong array format")
		Return
	EndIf

	Local $dateRow = 4
	Local $nameRow = 5
	Local $familyRow = 6
	Local $formatStyle = 7
	Local $startRow = 9
	Local $worksheet = "Template"

	Local $templatePath = $resourcesPath & "PrintTemplate.xlsx"
	If Not FileExists($templatePath) Then
		ToLog($err & "template file not exist: " & $resourcesPath & "PrintTemplate.xlsx")
		Return
	EndIf

	Local $excel = _Excel_Open(False, False, False, False, True)
	If @error Then
		ToLog($err & "cannot connect to Excel instance, error code: " & @error)
		Return
	EndIf

	Local $book = _Excel_BookOpen($excel, $templatePath)
	If @error Then
		Local $tmp = ""
		Switch @error
			Case 1
				$tmp = "$oExcel is not an object or not an application object"
			Case 2
				$tmp = "Specified $sFilePath does not exist"
			Case 3
				$tmp = "Unable to open $sFilePath. @extended is set to the COM error code returned by the Open method"
		EndSwitch
		ToLog($err & "cannot open workbook " & $templatePath & ", " & $tmp & ", error code: " & @error)
		Excel_Close($excel)
		Return
	EndIf

	_Excel_RangeWrite($book, $worksheet, $name, "A" & $nameRow)
	If @error Then ExcelWriteErrorToLog(@error)

	_Excel_RangeWrite($book, $worksheet, $surname, "A" & $familyRow)
	If @error Then ExcelWriteErrorToLog(@error)

	_Excel_RangeWrite($book, $worksheet, @MDAY & "." & @MON & "." & @YEAR & _
		", " & @HOUR & ":" & @MIN, "A" & $dateRow)
	If @error Then ExcelWriteErrorToLog(@error)

	Local $needToPay = False
	Local $outOfTime = False
	Local $xray = False

	Local $currentRow = $startRow
	Local $maxElement = UBound($array, $UBOUND_ROWS) -1
	If $maxElement > 5 Then $maxElement = 5

	For $i = 0 To $maxElement
		Local $timeAndCabinet = $array[$i][2] & ", кабинет " & $array[$i][5]
		Local $doc = $array[$i][3]
		Local $dept = StringLeft($array[$i][4], 1) & StringLower(StringRight($array[$i][4], StringLen($array[$i][4]) - 1))

		_Excel_RangeWrite($book, $worksheet, $timeAndCabinet, "A" & $currentRow)
		If @error Then ExcelWriteErrorToLog(@error)

		_Excel_RangeWrite($book, $worksheet, $doc, "A" & $currentRow + 1)
		If @error Then ExcelWriteErrorToLog(@error)

		_Excel_RangeWrite($book, $worksheet, $dept, "A" & $currentRow + 2)
		If @error Then ExcelWriteErrorToLog(@error)

		If $array[$i][6] Then
			$needToPay = True
			_Excel_RangeCopyPaste($book.ActiveSheet, _
				$book.ActiveSheet.Range("A" & $formatStyle), _
				$book.ActiveSheet.Range("A" & $currentRow + 3))
			If @error Then ExcelCopyPasteErrorToLog(@error)

			_Excel_RangeWrite($book, $worksheet, "Прием запланирован за наличный расчет", "A" & $currentRow + 3)
			If @error Then ExcelWriteErrorToLog(@error)

			$currentRow += 1
		EndIf

		If $array[$i][7] Then
			$outOfTime = True
			_Excel_RangeCopyPaste($book.ActiveSheet, _
				$book.ActiveSheet.Range("A" & $formatStyle), _
				$book.ActiveSheet.Range("A" & $currentRow + 3))
			If @error Then ExcelCopyPasteErrorToLog(@error)

			_Excel_RangeWrite($book, $worksheet, "Пропущено время начала приема", "A" & $currentRow + 3)
			If @error Then ExcelWriteErrorToLog(@error)

			$currentRow += 1
		EndIf

		If $i < $maxElement Then
			_Excel_RangeCopyPaste($book.ActiveSheet, _
				$book.ActiveSheet.Range("A" & $startRow - 1), _
				$book.ActiveSheet.Range("A" & $currentRow + 3))
			If @error Then ExcelCopyPasteErrorToLog(@error)

			_Excel_RangeCopyPaste($book.ActiveSheet, _
				$book.ActiveSheet.Range("A" & $startRow & ":A" & $startRow + 2), _
				$book.ActiveSheet.Range("A" & $currentRow + 4))
			If @error Then ExcelCopyPasteErrorToLog(@error)
		EndIf

		$currentRow += 4
	Next

	Local $finalText = ""
	If StringInStr($array[0][2], ":") Then
		Local $tmp = StringSplit($array[0][2], ":", $STR_NOCOUNT)
		Local $hour = $tmp[0]
		Local $minute = $tmp[1]

		Local $timeDiff = _DateDiff('n', _NowCalc(), GetFullDate($hour, $minute))

		If Not $outOfTime Then
			If $timeDiff < 0 Then
				$finalText &= "Вы опаздываете на ближайший прием, прошло минут: " & Abs($timeDiff) & @CRLF
			Else
				$finalText &= "До начала ближайщего приема осталось минут: " & $timeDiff & @CRLF
			EndIf
		EndIf
	EndIf


	If $needToPay And Not $outOfTime Then
		$finalText &= "У Вас имеются назначения за наличный расчет" & @CRLF & _
					  "Просьба пройти в регистратуру для оплаты приема"
	ElseIf $outOfTime And Not $needToPay Then
		$finalText &= "У Вас имеются назначения у которых пропущено время начала" & @CRLF & _
					  "Просьба пройти в регистратуру для согласования переноса"
	ElseIf $needToPay And $outOfTime Then
		$finalText &= "Для отметки о посещении" & @CRLF & _
					  "Просьба пройти в регистратуру"
	Else
		$finalText &= "Просьба проходить на прием"
	EndIf

	_Excel_RangeCopyPaste($book.ActiveSheet, _
		$book.ActiveSheet.Range("A" & $startRow - 1), _
		$book.ActiveSheet.Range("A" & $currentRow - 1))
	If @error Then ExcelCopyPasteErrorToLog(@error)

	_Excel_RangeCopyPaste($book.ActiveSheet, _
		$book.ActiveSheet.Range("A" & $startRow), _
		$book.ActiveSheet.Range("A" & $currentRow))
	If @error Then ExcelCopyPasteErrorToLog(@error)

	_Excel_RangeWrite($book, $worksheet, $finalText, "A" & $currentRow)
	If @error Then ExcelWriteErrorToLog(@error)

	_Excel_Print($excel, $book)
	If @error Then
		Local $tmp = ""
		Switch @error
			Case 1
				$tmp = "$oExcel is not an object or not an application object"
			Case 2
				$tmp = "$vObject is not an object or an invalid A1 range. @error is set to the COM error code"
			Case 3
				$tmp = "Error printing the object. @extended is set to the COM error code"
		EndSwitch
		ToLog($err & "cannot print workbook: " & $tmp & ", error code: " & @error)
		Excel_BookClose($book)
		Excel_Close($excel)
		Return
	EndIf

	If Not FileExists($printedAppointmentListPath) Then _
		DirCreate($printedAppointmentListPath)

	_Excel_BookSaveAs($book, $printedAppointmentListPath & $name & " " & $surname & " " & _
		@YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC)
	If @error Then
		Local $tmp = ""
		Switch @error
			Case 1
				$tmp = "$oWorkbook is not an object"
			Case 2
				$tmp = "$iFormat is not a number"
			Case 3
				$tmp = "File exists, overwrite flag not set"
			Case 4
				$tmp = "File exists but could not be deleted"
			Case 5
				$tmp = "Error occurred when saving the workbook. @extended is set to the COM error code returned by the SaveAs method."
		EndSwitch
		ToLog($err & "cannot save workbook as: " & $printedAppointmentListPath & ", " & $tmp & ", error code: " & @error)
		Excel_BookClose($book)
		Excel_Close($excel)
		Return
	EndIf

	Excel_BookClose($book)
	Excel_Close($excel)

	Return 1
EndFunc


Func Excel_Close($excel)
	_Excel_Close($excel, False, True)
	If @error Then
		Local $tmp = ""
		Switch @error
			Case 1
				$tmp = "$oExcel is not an object or not an application object"
			Case 2
				$tmp = "Error returned by method Application.Quit. @extended is set to the COM error code"
			Case 3
				$tmp = "Error returned by method Application.Save. @extended is set to the COM error code"
		EndSwitch
		ToLog("!!! Error - cannot close excel application: " & $tmp & ", error code: " & @error)
	EndIf
	If ProcessExists("EXCEL.exe") Then ProcessClose("EXCEL.exe")
EndFunc


Func Excel_BookClose($book)
	_Excel_BookClose($book, False)
	If @error Then
		Local $tmp = ""
		Switch @error
			Case 1
				$tmp = "$oWorkbook is not an object or not a workbook object"
			Case 2
				$tmp = "Error occurred when saving the workbook. @extended is set to the COM error code returned by the Save method"
			Case 3
				$tmp = "Error occurred when closing the workbook. @extended is set to the COM error code returned by the Close method"
		EndSwitch
		ToLog("!!! Error - cannot close workbook: " & $tmp & ", error code: " & @error)
	EndIf
EndFunc


Func ExcelWriteErrorToLog($code)
	Local $tmp = ""
	Switch $code
		Case 1
			$tmp = "$oWorkbook is not an object or not a workbook object"
		Case 2
			$tmp = "$vWorksheet name or index are invalid or $vWorksheet is not a worksheet object. @extended is set to the COM error code"
		Case 3
			$tmp = "$vRange is invalid. @extended is set to the COM error code"
		Case 4
			$tmp = "Error occurred when writing a single cell. @extended is set to the COM error code"
		Case 5
			$tmp = "Error occurred when writing data using the _ArrayTranspose function. @extended is set to the COM error code"
		Case 6
			$tmp = "Error occurred when writing data using the transpose method. @extended is set to the COM error code"
	EndSwitch
	ToLog("!!! Error - " & $tmp & ", error code: " & $code)
EndFunc


Func ExcelCopyPasteErrorToLog($code)
	Local $tmp = ""
	Switch $code
		Case 1
			$tmp = "$oWorkbook is not an object or not a workbook object"
		Case 2
			$tmp = "$vSourceRange is invalid. @extended is set to the COM error code"
		Case 3
			$tmp = "$vTargetRange is invalid. @extended is set to the COM error code"
		Case 4
			$tmp = "Error occurred when pasting cells. @extended is set to the COM error code"
		Case 5
			$tmp = "Error occurred when cutting cells. @extended is set to the COM error code"
		Case 6
			$tmp = "Error occurred when copying cells. @extended is set to the COM error code"
		Case 7
			$tmp = "$vSourceRange and $vTargetRange can't be set to keyword Default at the same time"
	EndSwitch
	ToLog("!!! Error - " & $tmp & ", error code: " & $code)
EndFunc




Func ExecuteSQL($sql)
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
		If $adoRecords.EOF = True And $adoRecords.BOF = True Then Return
	EndIf

	$adoConnection.Close
	$adoConnection = 0

	Return $result
EndFunc


Func NumPressed($n, $id)
	If StringLen($enteredCode) < 10 Then
		UpdateButtonBackgroundColor($id)
		$enteredCode &= $n
		UpdateInput()
	EndIf
EndFunc


Func ToLog($message)
	Local $logFilePath = $logsPath & @ScriptName & "_" & @YEAR & @MON & @MDAY & ".log"
	$message &= @CRLF
	ConsoleWrite($message)
	_FileWriteLog($logFilePath, $message)
EndFunc


Func HandleComError()
	  ConsoleWrite("error.description: " & @TAB & $oMyError.description  & @CRLF & _
				  "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
				  "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
				  "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
				  "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
				  "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
				  "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
				  "err.helpcontext is: " & @TAB & $oMyError.helpcontext & @CRLF)
Endfunc


Func OnExit()
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
	ToLog("-----Exiting-----")
	SendEmail("-----Exiting----- " & @exitMethod)
EndFunc


Func SendEmail($messageToSend)
	Local $send_email = True
	Local $current_pc_name = @ComputerName
	If Not $send_email Then Exit

	Local $title = "Infomat notification"
	$messageToSend &= @CRLF & @CRLF & _
		"---------------------------------------" & @CRLF & _
		"This is automatically generated message" & @CRLF & _
		"Sended from: " & $current_pc_name & @CRLF & _
		"Please do not reply"

	ToLog(@CRLF & "-----Sending email-----")
	Local $login = "infomat_notification@nnkk.budzdorov.su"
	Local $password = "fnpxmagr"
	Local $server = "smtp.budzdorov.ru"
	Local $from = $current_pc_name
	Local $to = "nn-admin@bzklinika.ru"

	_INetSmtpMailCom($server, $from, $login, $to, _
		$title, $messageToSend, "", "nn-admin@bzklinika.ru", "", $login, $password)
EndFunc


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", $s_Password = "",$IPPort=25, $ssl=0)

    Local $objEmail = ObjCreate("CDO.Message")
    Local $i_Error = 0
    Local $i_Error_desciption = ""

    $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
    $objEmail.To = $s_ToAddress

    If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
    If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

    $objEmail.Subject = $s_Subject

    If StringInStr($as_Body,"<") and StringInStr($as_Body,">") Then
        $objEmail.HTMLBody = $as_Body
    Else
        $objEmail.Textbody = $as_Body & @CRLF
	 EndIf

	 ProgressSet(50)

;~    ConsoleWrite($s_AttachFiles)
    If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
;~ 		_ArrayDisplay($S_Files2Attach)
        For $x = 1 To $S_Files2Attach[0] - 1
            $S_Files2Attach[$x] = _PathFull ($S_Files2Attach[$x])
            If FileExists($S_Files2Attach[$x]) Then
                $objEmail.AddAttachment ($S_Files2Attach[$x])
            Else
                $i_Error_desciption = $i_Error_desciption & @lf & 'File not found to attach: ' & $S_Files2Attach[$x]
				ConsoleWriteError("file not found")
                SetError(1)
                return 0
            EndIf
        Next
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
   ProgressSet(60)
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $Ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
   ProgressSet(70)
    ;Update settings
    $objEmail.Configuration.Fields.Update
   ProgressSet(80)
    ; Sent the Message
    $objEmail.Send
	ProgressSet(90)
    if @error then
        SetError(2)
		ProgressOff
        ;return $oMyRet[1]
    EndIf
EndFunc