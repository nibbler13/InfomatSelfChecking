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

Local $oMyError = ObjEvent("AutoIt.Error", "HandleComError")

Local $headerColor = 0x4e9b44
Local $okButtonColor = 0x4e9b44
Local $okButtonPressedColor = 0x43853a
Local $mainButtonColor = 0xe0e0e0
Local $mainButtonPressedColor = 0xd6d6d6
Local $disabledColor = 0xdfdfdf;0xf5f5f5
Local $disabledTextColor = 0xa5a5a5
Local $textColor = 0x2c3d3f
Local $alternateTextColor = 0xffffff
Local $mainBackgroundColor = 0xffffff
Local $errorTitleColor = 0xf98d3c

Local $bottonLineHeight = 11

Local $mainFontName = "Franklin Gothic"

Local $dX = @DesktopWidth
Local $dY = @DesktopHeight
;~ Local $dX = 1280
;~ Local $dY = 1024

Local $numButSize = $dY / 10
Local $distBt = $numButSize / 3

Local $initX = $dX / 2 - $numButSize * 1.5 - $distBt
Local $initY = $dy / 2 - $numButSize * 1.5 - $distBt

Local $btFontSize = $numButSize / 3
Local $btWeight = $FW_BOLD
Local $btQual = $CLEARTYPE_QUALITY

Local $headerHeight = $numButSize * 1.5
Local $headerLabelFontWeight = $FW_SEMIBOLD

Local $enteredCode = ""

#Region ====================== MainGUI ======================
Local $mainGui = GUICreate("SelfChecking", $dX, $dY, 0, 0, $WS_POPUP);, $WS_EX_TOPMOST)

Local $text = "Для отметки о посещении введите Ваш номер" & @CRLF & "мобильного телефона и нажмите кнопку «Продолжить»"
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
Local $bt_next = CreateButton("Продолжить", $prevBt[0], $dY - $prevBt[3] - $bottonLineHeight - $distBt, _
	$numButSize * 3 + $distBt * 2, $numButSize, $disabledColor)
GUICtrlSetColor(-1, $alternateTextColor)

$prevBt = ControlGetPos($mainGui, "", $bt_next)
Local $prevBt2 = ControlGetPos($mainGui, "", $bt_1)
Local $inp_pincode = GUICtrlCreateLabel($enteredCode, $dX / 2 - $prevBt[2] * 2.3 / 2, $headerHeight + $distBt, _
	$prevBt[2] * 2.3, $prevBt[3], BitOr($SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetFont(-1, $btFontSize * 1.8)
GUICtrlSetColor(-1, $textColor)

UpdateInput()

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


While 1
	$nMsg = GUIGetMsg()
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
			If StringLen($enteredCode) < 10 Then ContinueLoop
			ColorButton($bt_next, $okButtonColor, $okButtonPressedColor)
			CheckCode($enteredCode)
			ConsoleWrite("Next" & @CRLF)
		Case $bt_backspace
			ColorButton($bt_backspace)
			If StringLen($enteredCode) > 0 Then
				$enteredCode = StringLeft($enteredCode, StringLen($enteredCode) - 1)
				UpdateInput()
			EndIf
			ConsoleWrite("Backspace" & @CRLF)
		Case $bt_clear
		   ColorButton($bt_clear)
			$enteredCode = ""
			UpdateInput()
			ConsoleWrite("Clear" & @CRLF)
	EndSwitch
;~ 	Sleep(200)
WEnd


Func CreateStandardDesign($gui, $titleText, $isError, $trademark = False)
	GUISetBkColor($mainBackgroundColor)
	GUISetFont($btFontSize, $btWeight, 0, $mainFontName, $gui, $btQual)
	Local $titleColor = $headerColor
	If $isError Then $titleColor = $errorTitleColor
	CreateLabel($titleText, 0, 0, $dX, $headerHeight, $alternateTextColor, $titleColor, $gui)
	GUICtrlCreatePic(@ScriptDir & "\picBottomLine.jpg", 0, $dY - $bottonLineHeight, $dX, $bottonLineHeight)
	If $trademark Then
		Local $trademarkWidth = 159
		Local $trademarkHeight = 170
		GUICtrlCreatePic(@ScriptDir & "\picButterfly.jpg", $dX - $trademarkWidth - $distBt / 2, _
			$dY - $trademarkHeight - 11 - $distBt / 2, $trademarkWidth, $trademarkHeight)
	EndIf
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
		Local $newX = $x + ($width - $position[2] ) / 2
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

	GUICtrlCreatePic(@ScriptDir & "\picShadow.jpg", $x - $offsetX, $y - $offsetY, $width + $sizeX, $height + $sizeY, $SS_BLACKRECT)
	GUICtrlSetState(-1, $GUI_DISABLE)
	Local $id = GUICtrlCreateLabel($text, $x, $y, $width, $height, BitOR($SS_CENTER, $SS_CENTERIMAGE, $SS_NOTIFY))
	GUICtrlSetBkColor(-1, $bkColor)
	GUICtrlSetColor(-1, $color)
	Return $id
EndFunc


Func CheckCode($code)
	ConsoleWrite("CheckCode: " & $code & @CRLF)
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
;~ 	_ArrayDisplay($res)

	Local $textPhoneNumber = "+7 (" & $phoneNumberPrefix & ") " & StringLeft($phoneNumber, 3) & _
		"-" & StringMid($phoneNumber, 4, 2) & "-" & StringRight($phoneNumber, 2)
	Local $errorMessage = "К сожалению, по номеру " & $textPhoneNumber & @CRLF & "не найдено записей на ближайшее время" & _
		@CRLF & @CRLF & @CRLF & "Возможно указан неверный номер" & @CRLF & @CRLF & "Попробуйте снова" & @CRLF & "или обратитесь на регистратуру"

	If $res = False Or UBound($res, $UBOUND_ROWS) > 1 Then
		ShowErrorMessage($errorMessage)
		Return
	EndIf

	Local $fioForm = GUICreate("FIO", $dX, $dY, 0, 0, $WS_POPUP)

	Local $titleText = "Пожалуйста, убедитесь в соответствии Ваших данных:"
	CreateStandardDesign($fioForm, $titleText, False)

	Local $date = StringMid($res[0][3], 7, 2) & "." & StringMid($res[0][3], 5, 2) & "." & StringLeft($res[0][3], 4)
	Local $mainText = $res[0][1] & " " & $res[0][2] & @CRLF & @CRLF & "Дата рождения: " & $date

	CreateLabel($mainText, 0, $dY * 0.3, $dX, $dY * 0.4, $textColor, $GUI_BKCOLOR_TRANSPARENT, $fioForm, $btFontSize * 1.2)

	$prevBt = ControlGetPos($mainGui, "", $bt_next)
	Local $bt_ok = CreateButton("Продолжить", $dx - $distBt - $prevBt[2], $prevBt[1], $prevBt[2], $prevBt[3], $okButtonColor)
	GUICtrlSetColor(-1, $alternateTextColor)

	Local $bt_not = CreateButton("Неверно", 0 + $distBt, $prevBt[1], $prevBt[2], $prevBt[3])

	GUISetState()

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $bt_not
			ConsoleWrite("Not correct FIO" & @CRLF)

			$errorMessage = "Возможно указан неверный номер" & @CRLF & @CRLF & "Попробуйте снова" & @CRLF & "или обратитесь на регистратуру"
			ShowErrorMessage($errorMessage)
			GUIDelete($fioForm)
			Return
		Case $bt_ok
			ConsoleWrite("Correct FIO" & @CRLF)
			ShowDestinations($res[0][0], $errorMessage, $res[0][1] & " " & $res[0][2])
			GUIDelete($fioForm)
			$enteredCode = ""
			UpdateInput()
			Return
		EndSwitch
	WEnd
EndFunc


Func ShowDestinations($patientID, $errorMessage, $name)
	ConsoleWrite("ShowDestinations" & @CRLF)

	Local $sqlQuery = "Select Sch.SchedId, Sch.WorkDate, Sch.BHour, Sch.BMin, D.DName, Dep.DepName, R.RNum" & _
					" From Schedule Sch" & _
					" Join Doctor D On D.DCode = Sch.DCode" & _
					" Join DoctShedule Ds On Ds.DCode = Sch.DCode" & _
					" And Ds.SchedIdent = Sch.SchedIdent" & _
					" Join Departments Dep On Dep.DepNum = Ds.DepNum" & _
					" Join Chairs Ch On Ch.ChId = Ds.Chair" & _
					" Join Rooms R On R.RId = Ch.RId Where Sch.WorkDate = 'today'" & _
					" And Sch.PCode = " & $patientID

	Local $res = ExecuteSQL($sqlQuery)

	If $res = False Then
		ShowErrorMessage($errorMessage)
		Return
	EndIf

	For $i = 0 To UBound($res, $UBOUND_ROWS) - 1
		Local $hour = $res[$i][2]
		If StringLen($hour) < 2 Then $hour = "0" & $hour
		Local $minute = $res[$i][3]
		If StringLen($minute) < 2 Then $minute = "0" & $minute
		$res[$i][2] = $hour & ":" & $minute
	Next

	_ArraySort($res, 0, -1, -1, 2)
	_ArrayColDelete($res, 3)

	Local $head[1][6]
	$head[0][2] = "Время"
	$head[0][3] = "Специалист"
	$head[0][4] = "Отделение"
	$head[0][5] = "Кабинет"
	_ArrayConcatenate($head, $res)

	Local $destForm = GUICreate("ShowDestinations", $dX, $dY, 0, 0, $WS_POPUP)
	Local $text = $name & "," & @CRLF & "Ваши записи к специалистам на ближайшее время:"
	CreateStandardDesign($destForm, $text, False)

	$prevBt = ControlGetPos($mainGui, "", $bt_next)
	Local $bt_close = CreateButton("Закрыть", 0 + $distBt, $prevBt[1], $prevBt[2], $prevBt[3])
	Local $bt_print = CreateButton("Распечатать", $dX - $distBt - $prevBt[2], $prevBt[1], $prevBt[2], $prevBt[3], $okButtonColor, $alternateTextColor)

	Local $startX = $distBt
	Local $startY = $numButSize * 1.5 + $distBt
	Local $height = $numButSize * 0.7
	Local $distance = $distBt / 6
	Local $totalWidth = $dX - $distBt * 2 - $distance * 3
	Local $currentX = $startX
	Local $currentY = $startY

	Local $sizes[4]
	$sizes[0] = $totalWidth * 0.13
	$sizes[1] = $totalWidth * 0.16
	$sizes[2] = $totalWidth * 0.355
	$sizes[3] = $totalWidth * 0.355

	ConsoleWrite($sizes[2] & @CRLF)

	Local $maxSymbols = Round($sizes[2] / 23.5)

	GUICtrlCreateLabel("", $currentX, $currentY, $totalWidth + $distance * 3, $height - $distance)
	GUICtrlSetBkColor(-1, $mainButtonColor)

;~ 	_ArrayDisplay($head)

	Local $arraySize = UBound($head, $UBOUND_ROWS) - 1
	If $arraySize > 6 Then $arraySize = 6
	For $i = 0 To $arraySize
		Local $currentRow[4]

		$currentRow[0] = $head[$i][2]
		$currentRow[1] = $head[$i][5]
		Local $dept = StringLeft($head[$i][4], 1) & StringLower(StringRight($head[$i][4], StringLen($head[$i][4]) - 1))
		If StringLen($dept) > $maxSymbols Then $dept = StringLeft($dept, $maxSymbols - 3) & "..."
		$currentRow[2] = $dept

		Local $doc = $head[$i][3]
		If StringInStr($doc, "(") Then
			ConsoleWrite($doc & @CRLF)
			$doc = StringLeft($doc, StringInStr($doc, "(") - 1)
			$doc = StringStripWS($doc, BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING, $STR_STRIPSPACES))
		EndIf
		If StringLen($doc) > $maxSymbols Then $doc = StringLeft($doc, $maxSymbols - 3) & "..."
		$currentRow[3] = $doc

		For $x = 0 To 3
			GUICtrlCreateLabel($currentRow[$x], $currentX, $currentY, $sizes[$x], $height, BitOr($SS_CENTER, $SS_CENTERIMAGE))
			GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
			If $i < $arraySize Then
				GUICtrlCreateLabel("", $startX, $currentY + ($i = 0 ? $height - $distance - 1 : $height), $totalWidth + $distance * 3, $distance)
				GUICtrlSetBkColor(-1, $mainButtonColor)
			EndIf
			$currentX += $sizes[$x] + $distance
		Next

		$currentY += $height + $distance
		$currentX = $startX

		If $i = 0 Then
			$currentY -= $distance
			GUISetFont($btFontSize, $FW_NORMAL, -1, "Franklin Gothic Book", $destForm, $btQual)
		EndIf
	Next

	CreateLabel("Текущее время:" & @CRLF & @HOUR & ":" & @MIN, $prevBt[0], $prevBt[1], _
		$prevBt[2], $prevBt[3], $textColor, $GUI_BKCOLOR_TRANSPARENT, $destForm, $btFontSize * 0.8)

	GUISetState()

;~    Local $resUp1 = "update schedule set clvisit = 1 where schedid in " & $idToUpdate
;~    Local $resUp2 = "update schedule set screenvisit = 1 where schedid in " & $idToUpdate

;~    ExecuteSQL($resUp1)
;~    ExecuteSQL($resUp2)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
		Case $bt_close
			ConsoleWrite("Close" & @CRLF)
			$enteredCode = ""
			UpdateInput()
			GUIDelete($destForm)
			Return
		Case $bt_print
			ConsoleWrite("Print" & @CRLF)
	  EndSwitch
   WEnd
EndFunc


Func ExecuteSQL($sql)
	ConsoleWrite("ExecuteSQL: " & $sql & @CRLF)
	Local $sqlBD = "DRIVER=Firebird/InterBase(r) driver; UID=sysdba; PWD=masterkey; DBNAME=172.16.166.2:nnkk;"
	Local $adoConnection = ObjCreate("ADODB.Connection")
	Local $adoRecords = ObjCreate("ADODB.Recordset")

	$adoConnection.Open($sqlBD)
	$adoRecords.CursorType = 2
	$adoRecords.LockType = 3 ;3 - locks 1 - readonly

	Local $result = ""

	If StringInStr(StringLower($sql), "update") Then
		$adoRecords = $adoConnection.Execute($sql)
	Else
		$adoRecords.Open($sql, $adoConnection)
		Local $result = $adoRecords.GetRows

		If $adoRecords.EOF = True And $adoRecords.BOF = True Then
			ConsoleWrite("SQL EOF OR BOF" & @CRLF)
			Return ""
		EndIf

		$adoRecords.Close
		$adoRecords = 0
	EndIf

	$adoConnection.Close
	$adoConnection = 0

	Return $result
EndFunc


Func ShowErrorMessage($message)
	ConsoleWrite("ShowErrorMessage" & @CRLF)

	Local $nanForm = GUICreate("NothingFounded", $dX, $dY, 0, 0, $WS_POPUP)
	Local $text = "Уважаемый пациент!"
	CreateStandardDesign($nanForm, $text, True, True)

	$prevBt = ControlGetPos($mainGui, "", $bt_next)
	Local $bt_close = CreateButton("Закрыть", $prevBt[0], $prevBt[1], $prevBt[2], $prevBt[3])

	Local $x = 0
	Local $y = $dY * 0.3
	Local $sizeX = $dX
	Local $sizeY = $dY * 0.4
	CreateLabel($message, $x, $y, $sizeX, $sizeY, $textColor, $GUI_BKCOLOR_TRANSPARENT, $nanForm, $btFontSize * 1.2)

	GUISetState()

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
		Case $bt_close
			ConsoleWrite("Close" & @CRLF)
			$enteredCode = ""
			UpdateInput()
			GUIDelete($nanForm)
			Return
		EndSwitch
	WEnd
EndFunc


Func NumPressed($n, $id)
	If StringLen($enteredCode) < 10 Then
		ColorButton($id)
		$enteredCode &= $n
		UpdateInput()
	EndIf
EndFunc


Func ColorButton($id, $bkColor = $mainButtonColor, $glowColor = $mainButtonPressedColor)
	GUICtrlSetBkColor($id, $glowColor)
	Sleep(100)
	GUICtrlSetBkColor($id, $bkColor)
EndFunc


Func UpdateInput()
	Local $format = "+7 (___) ___-__-__"
	If StringLen($enteredCode) < 10 Then
		GUICtrlSetColor($bt_next, $disabledTextColor)
		GUICtrlSetBkColor($bt_next, $disabledColor)
	Else
		GUICtrlSetColor($bt_next, $alternateTextColor)
		GUICtrlSetBkColor($bt_next, $okButtonColor)
	EndIf

	For $i = 1 To StringLen($enteredCode)
		$format = StringReplace($format, "_", StringMid($enteredCode, $i, 1), 1)
;~ 		ConsoleWrite($i & @CRLF)
	Next

	ControlSetText($mainGui, "", $inp_pincode, $format)

;~ 	_WinAPI_SetWindowText($inp_pincode, $format)
;~ 	GUICtrlSetData($inp_pincode, $format)
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