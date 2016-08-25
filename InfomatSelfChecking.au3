#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\icon.ico
#pragma compile(ProductVersion, 0.9)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Приложения для инфомата для самостоятельной отметки о посещении)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - )
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


#Region ====================== Variables ======================
Local $oMyError = ObjEvent("AutoIt.Error", "HandleComError")
Local $iniFile = @ScriptDir & "\InfomatSelfChecking.ini"
Local $generalSectionName = "general"

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

;~ Local $dX = @DesktopWidth
;~ Local $dY = @DesktopHeight
Local $dX = 1024
Local $dY = 819

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
Local $enteredCode = "0000000000"

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

	ToLog("MainGui started")

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
			ToLog("FormCheckEnteredNumber force close")
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
				FormShowAppointments($fioForm, $res[0][0], $res[0][1] & " " & $res[0][2])
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


Func FormShowAppointments($guiToDelete, $patientID, $name)
	ToLog("FormShowAppointments: " & $name)

	Local $sqlQuery = "Select Sch.SchedId, Sch.WorkDate, Sch.BHour, Sch.BMin, D.DName, Dep.DepName, R.RNum, " & _
					"Case (Case " & _
					"  When Sch.SectId Is Not Null And Sch.SectId != 0 Then Sch.SectId " & _
					"  Else (Select SectId From Clients Where PCode = Sch.PCode) " & _
					"End) " & _
					"When 4363 Then 1 " & _
					"When 991139394 Then 1 " & _
					"Else 0 " & _
					"End As Kateg " & _
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
	Local $text = $name & "," & @CRLF & "Ваши записи на ближайшее время:"
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
	Local $needToPay = CreateAppointmentsTable($res, $destForm)

	UpdateTimeLabel()
	GUISetState()

	Sleep(10)

	If $guiToDelete Then GUIDelete($guiToDelete)

	Local $needToClose = False
	Local $textToShow = ""


	If $needToPay Then
		$textToShow &= "Просьба пройти на регистратуру" & @CRLF & _
					   "для оплаты приемов" & @CRLF & _
					   "запланированных за наличный расчет"
	Else
		$textToShow &= "Отметка о посещении успешно проставлена" & @CRLF & @CRLF & _
					   "Просьба проходить на прием"
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
				PrintAppontments($res, $name)
				$textToShow &= @CRLF & @CRLF & "Ваш список назначений успешно распечатан"
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
			ToLog("FormShowMessage force close")
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
	Local $head[1][7]
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
	For $i = 0 To $arraySize
		Local $currentRow[4]

		$currentRow[0] = $head[$i][2]
		$currentRow[1] = $head[$i][5]

		Local $dept = StringLower($head[$i][4])

		Local $doc = $head[$i][3]
		If StringInStr($doc, "(") Then
			$doc = StringLeft($doc, StringInStr($doc, "(") - 1)
			$doc = StringStripWS($doc, BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING, $STR_STRIPSPACES))
		EndIf

		$currentRow[2] = $doc & " (" & $dept & ")"
		While GetOptimalLabelWidth($currentRow[2], $gui) > $sizes[2] - ($head[$i][6] = 1 ? $iconWidth : 0)
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

			If $head[$i][6] Then
				$rubleIcon = CreatePngControl($resourcesPath & "RubleIcon.png", $iconWidth, $iconWidth)
				GUICtrlSetPos(-1, $dX - Round($distBt * 1.5) - $iconWidth, $currentY + Round($height * 0.2))
				$showCashWarning = True
			EndIf
		Next

		$currentY += $height + $distance
		$currentX = $startX + Round($distBt / 2)

		If $i = 0 Then
			$currentY -= $distance
		EndIf
	Next

	If $showCashWarning Then
		CreateLabel("Назначения со знаком рубля запланированы за наличный расчет" & @CRLF & _
			"Необходимо подойти на регистратуру для оплаты данных приемов", _
			$startX, _
			$currentY, _
			$totalWidth + $distance * 3, _
			Round($height * 1.5), _
			$alternateTextColor, _
			$errorTitleColor, _
			$gui, _
			Round($btFontSize * 0.8))
	Else
		Local $idToUpdate = ""

		For $i = 1 To $arraySize
			$idToUpdate &= $head[$i][0] & ","
		Next

		$idToUpdate = StringLeft($idToUpdate, StringLen($idToUpdate) - 1)

		Local $resUp1 = "update schedule set clvisit = 1 where schedid in (" & $idToUpdate & ")"
		Local $resUp2 = "update schedule set screenvisit = 1 where schedid in (" & $idToUpdate & ")"

		ExecuteSQL($resUp1)
		ExecuteSQL($resUp2)

		ToLog("Setting visit mark for: " & $idToUpdate)
	EndIf

	Return $showCashWarning
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

		If $timeDiff >= $bottomAppointmentTimeBoundaries And _
			$timeDiff <= $topAppointmentTimeBoundaries Then
			$currentRow[0][2] = $hour & ":" & $minute
			_ArrayAdd($retArray, $currentRow)
		EndIf
	Next

	_ArraySort($retArray, 0, -1, -1, 2)
	_ArrayColDelete($retArray, 3)

	Return $retArray
EndFunc






Func PrintAppontments($array, $fullName)
	ToLog("PrintAppontments")
	If Not IsArray($array) Or Not UBound($array, $UBOUND_ROWS) Then Return

	Local $dateRow = 4
	Local $nameRow = 5
	Local $familyRow = 6
	Local $formatStyle = 7
	Local $startRow = 9
	Local $worksheet = "Template"
	Local $excel = _Excel_Open(False, False, False, False, True)
	Local $book = _Excel_BookOpen($excel, $resourcesPath & "PrintTemplate.xlsx")

	Local $name = StringSplit($fullName, " ", $STR_NOCOUNT)[0]
	Local $family = StringSplit($fullName, " ", $STR_NOCOUNT)[1] & ","

	_Excel_RangeWrite($book, $worksheet, $name, "A" & $nameRow)
	_Excel_RangeWrite($book, $worksheet, $family, "A" & $familyRow)
	_Excel_RangeWrite($book, $worksheet, @MDAY & "." & @MON & "." & @YEAR & _
		", " & @HOUR & ":" & @MIN, "A" & $dateRow)

;~ 	_ArrayDisplay($array)

	Local $needToPay = False
	Local $currentRow = $startRow
	Local $maxElement = UBound($array, $UBOUND_ROWS) -1
	If $maxElement > 5 Then $maxElement = 5

	For $i = 0 To $maxElement
		Local $timeAndCabinet = $array[$i][2] & ", кабинет " & $array[$i][5]
		Local $doc = $array[$i][3]
		Local $dept = StringLeft($array[$i][4], 1) & StringLower(StringRight($array[$i][4], StringLen($array[$i][4]) - 1))

		_Excel_RangeWrite($book, $worksheet, $timeAndCabinet, "A" & $currentRow)
		_Excel_RangeWrite($book, $worksheet, $doc, "A" & $currentRow + 1)
		_Excel_RangeWrite($book, $worksheet, $dept, "A" & $currentRow + 2)

		If $array[$i][6] Then
			$needToPay = True
			_Excel_RangeCopyPaste($book.ActiveSheet, _
				$book.ActiveSheet.Range("A" & $formatStyle), _
				$book.ActiveSheet.Range("A" & $currentRow + 3))
			_Excel_RangeWrite($book, $worksheet, "Прием запланирован за наличный расчет", "A" & $currentRow + 3)
			$currentRow += 1
		EndIf

		If $i < $maxElement Then
			_Excel_RangeCopyPaste($book.ActiveSheet, _
				$book.ActiveSheet.Range("A" & $startRow - 1), _
				$book.ActiveSheet.Range("A" & $currentRow + 3))

			_Excel_RangeCopyPaste($book.ActiveSheet, _
				$book.ActiveSheet.Range("A" & $startRow & ":A" & $startRow + 2), _
				$book.ActiveSheet.Range("A" & $currentRow + 4))
		EndIf

		$currentRow += 4
	Next

	Local $hour = StringSplit($array[0][2], ":", $STR_NOCOUNT)[0]
	Local $minute = StringSplit($array[0][2], ":", $STR_NOCOUNT)[1]
	Local $timeDiff = _DateDiff('n', _NowCalc(), GetFullDate($hour, $minute))
	Local $finalText = ""

	If $timeDiff < 0 Then
		$finalText &= "Вы опаздываете на ближайший прием, прошло минут: " & Abs($timeDiff) & @CRLF
	Else
		$finalText &= "До начала ближайщего приема осталось минут: " & $timeDiff & @CRLF
	EndIf


	If $needToPay Then
		$finalText &= "У Вас имеются назначения за наличный расчет" & @CRLF & _
					  "Просьба пройти на регистратуру для оплаты приема"
	Else
		$finalText &= "Просьба проходить на прием"
	EndIf

	_Excel_RangeCopyPaste($book.ActiveSheet, _
		$book.ActiveSheet.Range("A" & $startRow - 1), _
		$book.ActiveSheet.Range("A" & $currentRow - 1))
	_Excel_RangeCopyPaste($book.ActiveSheet, _
		$book.ActiveSheet.Range("A" & $startRow), _
		$book.ActiveSheet.Range("A" & $currentRow))
	_Excel_RangeWrite($book, $worksheet, $finalText, "A" & $currentRow)
	_Excel_Print($excel, $book)

	If Not FileExists($printedAppointmentListPath) Then _
		DirCreate($printedAppointmentListPath)

	_Excel_BookSaveAs($book, $printedAppointmentListPath & $fullName & " " & @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC)
	_Excel_BookClose($book)
	_Excel_Close($excel, False)
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