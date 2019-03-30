#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GuestPrintFunc.au3>


Const $DomainName = "CE-SDXCORP"
Const $PrintService = "\\Server"
Const $PrintServiceIP="192.168.1.2"
Const $PrintName = "\\Server\GuestPrint"
Const $PrintServiceUser = "username"
Const $PrintServiceUserPwd = "Password"

GloBal Const $PSEXE = "C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe -command "

Dim $Account ,$Pwd

;Default OS Version
GloBal $PSTitle = "C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe"

;Default Language
GloBal $TitlePrinterGUI = "Server 上的 GuestPrint 打印首选项"
GloBal $TitlePrinterAuth= "验证"
GloBal $MsgToUser="请务必确认你的电脑已连接到办公室网络！"







main()



Func main()
	Global $CANCEL = 0
	Global Const $INSTALL= 1
	Global $UNINSTALL = -1

   ;setting language of title
   if @OSLang = 0804 Then
	  ;简体中文
	  $TitlePrinterGUI = "Server 上的 GuestPrint 打印首选项"
	  $TitlePrinterAuth= "验证"
	  $MsgToUser="请务必确认你的电脑已连接到办公室网络！"
   Else
	  ;非中文（英语）
	  $TitlePrinterGUI = "GuestPrint on 192.168.1.2 Printing Preferences"
	  $TitlePrinterAuth= "Authentication"
	  $MsgToUser="Please make sure your computer is connected to Office network!"
   EndIf

   ;Setting Title
   If @OSVersion = "WIN_10" Then
	  $PSTitle = "Windows PowerShell"
   Else
	  $PSTitle = "C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe"
   EndIf


	While 1
		Switch Choice()
			Case $CANCEL
				;msgbox(0,"Cancel","Bye")
				ExitLoop

			Case $INSTALL
				InstallPrinter()
				;msgbox(0,"Succeed","Install Printer Succeed!")

				ExitLoop

			Case $UNINSTALL
				UnInstall($PrintName)
				ExitLoop
		EndSwitch
	WEnd



EndFunc ;==> main













