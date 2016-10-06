#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\icon2.ico
#pragma compile(ProductVersion, 1.0)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Приложения для перезапуска процесса инфомата)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555)
#pragma compile(ProductName, InfomatAutoRestart)
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Local $sProcessName = "InfomatSelfChecking.exe"
Local $sErrorWindowTitle = "AutoIt Error"
Local $sErrorWindowText = 'Error: Variable must be of type: "Object".'

While 1
	If Not ProcessExists($sProcessName) Then ShellExecute($sProcessName)

	If WinExists($sErrorWindowTitle) Then
		WinClose($sErrorWindowTitle)
		ProcessClose($sProcessName)
		ShellExecute($sProcessName)
	EndIf

	Sleep(10 * 1000)
WEnd