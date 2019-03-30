#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>


;User Choose Button
Func Choice()
	Local $UserChoice 	;记录用户的选择

	Local $mGUI = GUICreate("Sodexo China GuestPrinter", 320, 200)
	Local $idButton_install = GUICtrlCreateButton("Install",85,20,150,40)
	Local $idButton_uninstall = GUICtrlCreateButton("UnInstall", 85,80,150,40)
	Local $idButton_Close = GUICtrlCreateButton("Close", 85, 140,150,40)

	;Display the GUI
	GUISetState(@SW_SHOW,$mGUI)

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $idButton_Close
				;取消
				$UserChoice = $CANCEL
				ExitLoop

			Case $idButton_install
				;选择安装
				if ping($PrintServiceIP) then
				   ;确认连接到内网才执行安装进程
				   $UserChoice = $INSTALL
				   ExitLoop
				Else
				   MsgBox(0,"Connection failed",$MsgToUser)
				EndIf

			Case $idButton_uninstall
				;选择卸载
				$UserChoice = $UNINSTALL
				ExitLoop
		EndSwitch
	WEnd

	;Hide the GUI
	GUISetState(@SW_HIDE,$mGUI)

	;Return the choice
	Return $UserChoice

	;Delete the previous GUI and all controls.
	GUIDelete($mGui)
EndFunc ;==> Choice






;InstallPrinter
Func InstallPrinter()

	If GetInfo($Account,$Pwd) Then
		;连接并设置打印机

		ConnectPrint($PrintService,$PrintName,$PrintServiceUser,$PrintServiceUserPwd)


		SetPrint($PrintName,$DomainName,$Account,$Pwd)


		;Disconnect
		DisConnect()

		Sleep(3000)


		 Return 1
		;msgbox(0,"Succeed","Install Printer Succeed!")
		else
		;用户选择取消，关闭窗口
			;msgbox(0,"Cancel","Cancel!")
	EndIF

EndFunc ;==>InstallPrinter






;Getting account and password
Func GetInfo(ByRef $Account,ByRef $Pwd)
	Local $hGUI = GUICreate("Please Login!", 320, 200)

	GUICtrlCreateLabel("Account:",40,30,70)
	Local $A = GUICtrlCreateInput("",100, 26, 180,25)

	GUICtrlCreateLabel("Password:",40,65,70)
	Local $P = GUICtrlCreateInput("",100, 60, 180,25,$ES_PASSWORD)

	GUICtrlCreateLabel($MsgToUser,40,100,280,30)


	Local $idButton_OK = GUICtrlCreateButton("OK", 80, 150, 60)
	Local $idButton_Close = GUICtrlCreateButton("Cancel", 180, 150, 60)

   ;默认ok按钮状态
    GUICtrlSetState($idButton_OK ,$GUI_DISABLE)

	;Display the GUI
	GUISetState(@SW_SHOW,$hGUI)


	While 1
	  ;更新ok按钮状态
	  ;#comments-start
	   IF GuiCtrlRead($A) <> "" AND GuiCtrlRead($P) <> "" Then
		  if GUICtrlGetState($idButton_OK) <> $GUI_ENABLE Then
			 GUICtrlSetState($idButton_OK ,$GUI_ENABLE)
		  EndIf
	   Else
		  if GUICtrlGetState($idButton_OK) <> $GUI_DISABLE Then
			 GUICtrlSetState($idButton_OK ,$GUI_DISABLE)
		  EndIf
	   EndIf
	   ;#comments-end

		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $idButton_Close

				Local $Flag = False
				ExitLoop
			Case $idButton_OK
				$Account = GuiCtrlRead($A)
				$Pwd = GuiCtrlRead($P)

				Local $Flag = True
				ExitLoop
		EndSwitch
	WEnd

	;Hide the GUI
	GUISetState(@SW_HIDE,$hGUI)

	;Return User's choose
	Return $Flag

	;Delete the previous GUI and all controls.
	GUIDelete($hGui)

EndFunc	 ;==>GetInfo







;Connecting GuestPrint
Func ConnectPrint($PrintService,$PrintName,$PrintServiceUser,$PrintServiceUserPwd)
	run($PSEXE & "net use "&$PrintService&"\ipc$ /user:"&$PrintServiceUser&" "&$PrintServiceUserPwd)
    WinWait($PSTitle,"",5000)
	WinSetState($PSTitle,"",@SW_HIDE)
	WinWaitClose($PSTitle,"",10000)

	Local $PSCmd = "(New-Object -ComObject WScript.Network).AddWindowsPrinterConnection('" & $PrintName & "')"
	run($PSEXE & $PSCmd )

	;等待POWERSHELL进程关闭
	WinWait($PSTitle,"",5000)
	WinSetState($PSTitle,"",@SW_HIDE)
	WinWaitClose($PSTitle,"",20000)
 EndFunc ;==>ConnectPrint




 Func DisConnect()
 ;关闭连接
	run($PSEXE & "net use /delete "&$PrintService&"\ipc$" )
    WinSetState($PSTitle,"",@SW_HIDE)
 EndFunc ;==>DisConnect




;Setting the printer preference
Func SetPrint($PrintName,$DomainName,$Account,$Pwd)
	run ("rundll32 printui.dll,PrintUIEntry /e /n"&$PrintName)
	;Sleep(5000)
	;切换到详细设置
	WinWait($TitlePrinterGUI,"",5000)
	WinSetState($TitlePrinterGUI,"",@SW_HIDE)
	WinActivate($TitlePrinterGUI)
	Send ("^{TAB}{TAB 4}{UP}")
	;进入验证设置
	;Sleep(2000)

	$WaitTime = 0
	While 1
		if ControlCommand($TitlePrinterGUI,"","[CLASS:Button; INSTANCE:15]","IsVisible") then
			exitloop
		elseif $WaitTime > 10 then
			;加入超时错误
			Exitloop
		else
			$Temp = $WaitTime
			$WaitTime = $Temp +1
			Sleep(1000)
		EndIF
	WEnd
   ;Click Button Authentication
    WinActivate($TitlePrinterGUI)
	ControlClick($TitlePrinterGUI,"","[CLASS:Button; INSTANCE:15]")


	WinWait($TitlePrinterAuth)
	WinSetState($TitlePrinterAuth,"",@SW_HIDE)
	WinActivate($TitlePrinterAuth)

	ControlClick($TitlePrinterAuth,"","[CLASS:Edit; INSTANCE:1]")
	Send("^a{del}{CAPSLOCK ON}"&$Account)
	Sleep(500)

    WinActivate($TitlePrinterAuth)
	ControlClick($TitlePrinterAuth,"","[CLASS:Edit; INSTANCE:2]")
	Send("^{a}{del}{CAPSLOCK OFF}" & $Pwd)
	Sleep(500)


    WinActivate($TitlePrinterAuth)
	ControlClick($TitlePrinterAuth,"","[CLASS:Edit; INSTANCE:3]")
	Send("{CAPSLOCK OFF}" & $Pwd)
	Sleep(500)

    WinActivate($TitlePrinterAuth)
	ControlClick($TitlePrinterAuth,"","[CLASS:Edit; INSTANCE:4]")
	Send("^a{DELETE}{CAPSLOCK ON}"&$DomainName)
	Sleep(500)


	;Click OK Button on Authentication interface
    WinActivate($TitlePrinterAuth)
	ControlClick($TitlePrinterAuth,"","[CLASS:Button; INSTANCE:6]")

   ;Click OK Button
	WinWaitClose($TitlePrinterAuth)
    WinActivate($TitlePrinterGUI)
	ControlClick($TitlePrinterGUI,"","[CLASS:Button; INSTANCE:47]")
	WinWaitClose($TitlePrinterGUI)


EndFunc ;==>SetPrint





;UnInstall the printer
Func UnInstall($PrintName)

	Local $PSCmd = "(New-Object -ComObject WScript.Network).RemovePrinterConnection('" & $PrintName & "')"
	run($PSEXE & $PSCmd )

	;等待POWERSHELL进程关闭
	WinWait($PSTitle,"",5000)
	WinSetState($PSTitle,"",@SW_HIDE)
	WinWaitClose($PSTitle,"",10000)

	msgbox(0,"Succeed","Uninstall Printer Succeed!")
EndFunc ;==>UninstallPrinter