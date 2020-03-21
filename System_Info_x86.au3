#RequireAdmin
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5 + file SciTEUser.properties in your UserProfile e.g. C:\Users\User-10

 Author:        WIMB  -  Oct 24, 2019

 Program:       System_Info_x86.exe - Version 4.2

 Script Function:

 Credits and Thanks to:

	- Uwe Sieber for making ListUsbDrives - http://www.uwe-sieber.de/english.html
	- Nir Sofer for making produkey - https://www.nirsoft.net/utils/product_cd_key_viewer.html
	- Nir Sofer for making serviwin - https://www.nirsoft.net/utils/serviwin.html
	- Franck Delattre for making CPU-Z - https://www.cpuid.com/softwares/cpu-z.html
	- JFX for making AutoIt Function to determine Windows + Office Key - https://www.autoitscript.com/forum/topic/131797-windows-and-office-key/
	- Terenz for making AutoIt Functions to determine Partition Style and Firmware -https://www.autoitscript.com/forum/topic/186012-detect-an-uefi-windows-and-gpt-disk-type/

	The program is released "as is" and is free for redistribution, use or changes as long as original author,
	credits part and link to the reboot.pro support forum are clearly mentioned
	System_Info - http://reboot.pro/files/file/611-system-info/ and http://reboot.pro/topic/22053-system-info/

	Author does not take any responsibility for use or misuse of the program.

#ce ----------------------------------------------------------------------------

#include <guiconstants.au3>
#include <ProgressConstants.au3>
#include <GuiConstantsEx.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#Include <GuiStatusBar.au3>
#include <Array.au3>
#Include <String.au3>
#include <Process.au3>
#include <Date.au3>
#include <Constants.au3>
#include <WinAPIDlg.au3>
#include <MemoryConstants.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
; ------------------------------------------------------------------------------

Opt('MustDeclareVars', 1)
Opt("TrayIconHide", 1)

Global $str = "", $bt_files[10] = ["\HP_Sys_Diag\HPSysDig.exe", "\cpu-z\cpuz_x32.exe", "\cpu-z\cpuz_x64.exe", "\makebt\devcon\x64\devcon.exe", "\makebt\devcon\x86\devcon.exe", _
"\makebt\listusbdrives\ListUsbDrives.exe", "\produkey\ProduKey.exe", "\produkey-x64\ProduKey.exe", "\serviwin\serviwin.exe", "\serviwin-x64\serviwin.exe"]

If @OSArch <> "X86" Then
   MsgBox(48, "ERROR - Environment", "In x64 environment use System_Info_x64.exe ")
   Exit
EndIf

For $str In $bt_files
	If Not FileExists(@ScriptDir & $str) Then
		MsgBox(48, "ERROR - Missing File", "File " & $str & " NOT Found ")
		Exit
	EndIf
Next

Global $Firmware = "UEFI", $PartStyle = "MBR", $PartStyle_ScriptDrive = "MBR", $WinLang = "en-US", $ReleaseID = "", $RegOwner = "", $OSEdition = "", $ProductName = ""
Global $GUI_Start, $button1, $button2, $button3, $button4, $button5, $button6, $button7, $button8, $button9, $button10, $button11, $button12, $nMsg, $WindowsDir = @WindowsDir
Global $button13, $button14, $button15, $button16, $button17, $button18, $button19, $button20, $hStatus
Global $Mobo_Manufacturer = "", $Mobo_Product = "", $cpu, $Number_Cores, $gpu = "",  $RAM = "", $PE_flag = 0, $UserName = @UserName, $ComputerName = @ComputerName
Global $Software = "SOFTWARE", $OS_DriveLetter = StringLeft(@SystemDir, 1)  ; $LastLoggedOnSAMUser, $pos
Global $OSBuild, $System_Reg, $OSArch = @OSArch, $UUID = "", $TargetSelect = "", $pos_TS, $len_TS, $path_folder = "", $drv = ""

Global $bKey_Id, $bKey_Id4, $bKey_Def, $bKey_Def4

SystemFileRedirect("On")

If StringLeft(@SystemDir, 1) = "X" Then
	$PE_flag = 1
	$OS_DriveLetter = "C"
	; _Windows_drive()
Else
	$PE_flag = 0
	$OS_DriveLetter = StringLeft(@SystemDir, 1)
EndIf

$Firmware = _WinAPI_GetFirmwareEnvironmentVariable()
$PartStyle_ScriptDrive = _GetDrivePartitionStyle(StringLeft(@ScriptDir, 1))

If $PE_flag = 1 And FileExists($OS_DriveLetter & ":\Windows\System32\config\software") Then
	$Software = "SOFTWARE_" & $OS_DriveLetter
	$System_Reg = "SYSTEM_" & $OS_DriveLetter

	$PartStyle = _GetDrivePartitionStyle($OS_DriveLetter)

	If FileExists($OS_DriveLetter & ":\Windows\SysWOW64") Then
		$OSArch = "X64"
	Else
		$OSArch = "X86"
	EndIf

	RunWait(@ComSpec & " /c reg load HKLM\" & $Software & " " & $OS_DriveLetter & ":\Windows\System32\config\software", @ScriptDir, @SW_HIDE)

	$ReleaseID = RegRead("HKLM\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "ReleaseId")
	$RegOwner = RegRead("HKLM\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
	$OSEdition = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "EditionID")
	$ProductName = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "ProductName")
	$OSBuild = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "CurrentBuild")

	If StringInStr(@OSArch, "64") Then
		$gpu = RegRead("HKEY_LOCAL_MACHINE64\" & $Software & "\Microsoft\Windows NT\CurrentVersion\Winsat", "PrimaryAdapterString")
	Else
		$gpu = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion\Winsat", "PrimaryAdapterString")
	EndIf

;~ 		$ComputerName = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows\CurrentVersion\Group Policy\DataStore\Machine\0", "szName")
;~ 		If $ComputerName = "" Then
;~ 			$LastLoggedOnSAMUser = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows\CurrentVersion\Authentication\LogonUI", "LastLoggedOnSAMUser")
;~ 			$pos = StringInStr($LastLoggedOnSAMUser, "\", 0, -1)
;~ 			$ComputerName = StringLeft($LastLoggedOnSAMUser, $pos-1)
;~ 		EndIf

	$UserName = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion\Winlogon", "LastUserName")
	If $UserName = "" Then
		$UserName = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName")
	EndIf
	$WindowsDir = $OS_DriveLetter & ":\Windows"

	; Windows Key
	$bKey_Id = RegRead("HKLM64\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "DigitalProductId")
	$bKey_Id4 = RegRead("HKLM64\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "DigitalProductId4")
	$bKey_Def = RegRead("HKLM64\" & $Software & "\Microsoft\Windows NT\CurrentVersion\DefaultProductKey", "DigitalProductId")
	$bKey_Def4 = RegRead("HKLM64\" & $Software & "\Microsoft\Windows NT\CurrentVersion\DefaultProductKey", "DigitalProductId4")

	RunWait(@ComSpec & " /c reg unload HKLM\" & $Software, @ScriptDir, @SW_HIDE)

	RunWait(@ComSpec & " /c reg load HKLM\" & $System_Reg & " " & $OS_DriveLetter & ":\Windows\System32\config\system", @ScriptDir, @SW_HIDE)
	$ComputerName = RegRead("HKEY_LOCAL_MACHINE\" & $System_Reg & "\ControlSet001\Control\ComputerName\ComputerName", "ComputerName")
	RunWait(@ComSpec & " /c reg unload HKLM\" & $System_Reg, @ScriptDir, @SW_HIDE)

Else
	$OS_DriveLetter = StringLeft(@SystemDir, 1)
	$Software = "SOFTWARE"
	$PartStyle = _GetDrivePartitionStyle($OS_DriveLetter)

	$ReleaseID = RegRead("HKLM\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "ReleaseId")
	$RegOwner = RegRead("HKLM\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
	$OSEdition = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "EditionID")
	$ProductName = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "ProductName")
	$OSBuild = @OSBuild

	If StringInStr(@OSArch, "64") Then
		$gpu = RegRead("HKEY_LOCAL_MACHINE64\" & $Software & "\Microsoft\Windows NT\CurrentVersion\Winsat", "PrimaryAdapterString")
	Else
		$gpu = RegRead("HKEY_LOCAL_MACHINE\" & $Software & "\Microsoft\Windows NT\CurrentVersion\Winsat", "PrimaryAdapterString")
	EndIf

	$ComputerName = @ComputerName
	$UserName = @UserName
	$WindowsDir = @WindowsDir
	$OSArch = @OSArch

	; Windows Key
	$bKey_Id = RegRead("HKLM64\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "DigitalProductId")
	$bKey_Id4 = RegRead("HKLM64\" & $Software & "\Microsoft\Windows NT\CurrentVersion", "DigitalProductId4")
	$bKey_Def = RegRead("HKLM64\" & $Software & "\Microsoft\Windows NT\CurrentVersion\DefaultProductKey", "DigitalProductId")
	$bKey_Def4 = RegRead("HKLM64\" & $Software & "\Microsoft\Windows NT\CurrentVersion\DefaultProductKey", "DigitalProductId4")

EndIf

$cpu = RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString")
$Number_Cores = EnvGet("NUMBER_OF_PROCESSORS")
$Mobo_Manufacturer = RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "BaseBoardManufacturer")
$Mobo_Product = RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "BaseBoardProduct")

$UUID = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\HardwareConfig", "LastConfig")

$RAM = Round(_GetPhysicallyInstalledSystemMemory()/1024/1024)

_WinLang()
SystemFileRedirect("Off")

If $gpu = "" Then
	Dim $Obj_WMIService = ObjGet("winmgmts:\\" & "localhost" & "\root\cimv2")
	Global $Obj_Item, $VideoController
	Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from Win32_VideoController")
	For $Obj_Item In $Obj_Services
		$VideoController &= $Obj_Item.Name
	Next
	$gpu = $VideoController

;~ 			Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from Win32_ComputerSystem")
;~ 			For $Obj_Item In $Obj_Services
;~ 				$manufacturer =  $Obj_Item.Manufacturer
;~ 				$model = $Obj_Item.Model
;~ 			Next
EndIf

If $UUID = "" Then
	Global $strIdentifyingNumber, $strName, $strVersion, $strWMIQuery, $objItem
	Dim $uuItem, $objSWbemObject, $uiDitem

	Dim $objWMIService = ObjGet("winmgmts:\\" & "localhost" & "\root\cimv2")

	$uuItem = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
	If IsObj($uuItem) Then
		For $objSWbemObject IN $uuItem
			$strIdentifyingNumber = $objSWbemObject.IdentifyingNumber
			$strName = $objSWbemObject.Name
			$strVersion = $objSWbemObject.Version
		Next
	EndIf
	$strWMIQuery = ":Win32_ComputerSystemProduct.IdentifyingNumber='" & $strIdentifyingNumber & "',Name='" & $strName & "',Version='" & $strVersion & chr(39)

	$uiDitem = ObjGet("winmgmts" & $strWMIQuery)

	If IsObj($uiDitem) Then
		For $objItem in $uiDitem.Properties_
			If $objItem.name =  "UUID" Then
				$UUID = $objItem.value
			EndIf
		Next
	EndIf
	; MsgBox(0, "UUID", "UUID = " & $UUID)
EndIf

$GUI_Start = GUICreate("System Info - x86 Version 4.2", 390, 410, -1, -1, BitXOR($GUI_SS_DEFAULT_GUI, $WS_MINIMIZEBOX))
$button19 = GUICtrlCreateButton(" Sys Info ", 30, 25, 70)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button1 = GUICtrlCreateButton(" MS Info ", 110, 25, 70)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button2 = GUICtrlCreateButton(" Device Manager ", 30, 65, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button3 = GUICtrlCreateButton(" HP Diag ", 30, 105, 70)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button20 = GUICtrlCreateButton(" DX Diag ", 110, 105, 70)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button4 = GUICtrlCreateButton("Disk Manager ", 30, 145, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button5 = GUICtrlCreateButton(" Task Manager ", 30, 185, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button17 = GUICtrlCreateButton(" Folder List ", 30, 225, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")

$button6 = GUICtrlCreateButton(" Drivers + Services ", 30, 265, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button7 = GUICtrlCreateButton(" Windows Settings ", 30, 305, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button8 = GUICtrlCreateButton(" Windows + Office Key ", 30, 345, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button9 = GUICtrlCreateButton(" System Info ", 210, 25, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button10 = GUICtrlCreateButton(" Hardware Info ", 210, 65, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button11 = GUICtrlCreateButton(" Save HWIDs ", 210, 105, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button12 = GUICtrlCreateButton(" Drive Info ", 210, 145, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button13 = GUICtrlCreateButton(" Directories ", 210, 185, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button18 = GUICtrlCreateButton(" File List ", 210, 225, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")


$button14 = GUICtrlCreateButton(" CPU + Memory Info ", 210, 265, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button15 = GUICtrlCreateButton(" Control Panel ", 210, 305, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")
$button16 = GUICtrlCreateButton(" Save System Info ", 210, 345, 150)
GUICtrlSetFont(-1, 10, "", "", "Tahoma")

$hStatus = _GUICtrlStatusBar_Create($GUI_Start, -1, "", $SBARS_TOOLTIPS)

_GUICtrlStatusBar_SetText($hStatus,"", 0)

DisableMenus(0)

GUISetState(@SW_SHOW)

While 1
    $nMsg = GUIGetMsg()
	If $nMsg = $GUI_EVENT_CLOSE Then Exit
	If $nMsg > 0 Then
		DisableMenus(1)
		Switch $nMsg
			Case $button19
				_GUICtrlStatusBar_SetText($hStatus," systeminfo.exe is Running - Wait ....", 0)
				If FileExists(@ScriptDir & "\" & $ComputerName & "_MS_System_Info.txt") Then
					FileCopy(@ScriptDir & "\" & $ComputerName & "_MS_System_Info.txt", @ScriptDir & "\" & $ComputerName & "_MS_System_Info_bak.txt", 1)
					FileDelete(@ScriptDir & "\" & $ComputerName & "_MS_System_Info.txt")
				EndIf
				RunWait(@ComSpec & " /c systeminfo.exe > " & $ComputerName & "_MS_System_Info.txt", @ScriptDir, @SW_HIDE)

				ShellExecute("notepad.exe", @ScriptDir & "\" & $ComputerName & "_MS_System_Info.txt", @ScriptDir)
				_GUICtrlStatusBar_SetText($hStatus,"", 0)

			Case $button1
				ShellExecute("msinfo32.exe")
				MsgBox(0,"Timeout", "", 0.3)
			Case $button2
				ShellExecute("devmgmt.msc")
				MsgBox(0,"Timeout", "", 0.3)
			Case $button3
				ShellExecute(@ScriptDir & "\HP_Sys_Diag\HPSysDig.exe")
				MsgBox(0,"Timeout", "", 0.3)
			Case $button20
				ShellExecute("dxdiag.exe")
				MsgBox(0," DX Diag ", "DX Diag is Running - Wait", 1)
			Case $button4
				ShellExecute("diskmgmt.msc")
				MsgBox(0,"Timeout", "", 0.3)
			Case $button5
				ShellExecute("Taskmgr.exe")
				MsgBox(0,"Timeout", "", 0.3)

			Case $button17
				_GUICtrlStatusBar_SetText($hStatus," Select Path to make Folder List from Path", 0)
				$TargetSelect = FileSelectFolder("Select Path to make Folder List from Path ", "")
				If @error Then
					_GUICtrlStatusBar_SetText($hStatus,"", 0)
					MsgBox(48,"ERROR - Path Invalid", "Error - Path Invalid")
				Else
					If StringInStr($TargetSelect, "\", 0, -1) = 0 Or StringInStr($TargetSelect, ":", 0, 1) <> 2 Then
						_GUICtrlStatusBar_SetText($hStatus,"", 0)
						MsgBox(48,"ERROR - Path Invalid", "Drive Invalid  :  Or \ Not found" & @CRLF & @CRLF & "Selected Path = " & $TargetSelect)
					Else
						_GUICtrlStatusBar_SetText($hStatus," Using Dir Command to Make Folder List - Wait ...", 0)
						$len_TS = StringLen($TargetSelect)
						$drv = StringLeft($TargetSelect, 1)
						If $len_TS > 3 Then
							$pos_TS = StringInStr($TargetSelect, "\", 0, -1)
							$path_folder = StringRight($TargetSelect, $len_TS - $pos_TS)
							RunWait(@ComSpec & " /u /c dir /b /s /ad-l /on > " & '"' & @ScriptDir & "\Folder_List_" & $drv & "_" & $path_folder & ".txt" & '"', $TargetSelect, @SW_HIDE)
							ShellExecute("notepad.exe", "Folder_List_" & $drv & "_" & $path_folder & ".txt", @ScriptDir)
						Else
							RunWait(@ComSpec & " /u /c dir /b /s /ad-l /on > " & '"' & @ScriptDir & "\Folder_List_" & $drv & ".txt" & '"', $TargetSelect, @SW_HIDE)
							ShellExecute("notepad.exe", "Folder_List_" & $drv & ".txt", @ScriptDir)
						EndIf
					EndIf
				EndIf
				_GUICtrlStatusBar_SetText($hStatus,"", 0)
			Case $button6
				If @OSArch = "X86" Then
					ShellExecute(@ScriptDir & "\serviwin\serviwin.exe")
					MsgBox(0,"Timeout", "", 0.3)
				Else
					ShellExecute(@ScriptDir & "\serviwin-x64\serviwin.exe")
					MsgBox(0,"Timeout", "", 0.3)
				EndIf
			Case $button7
				RunWait(@ComSpec & " /c start ms-settings:", @ScriptDir, @SW_HIDE)
			Case $button8
				If @OSArch = "X86" Then
					ShellExecute(@ScriptDir & "\produkey\ProduKey.exe")
					MsgBox(0,"Timeout", "", 0.3)
				Else
					ShellExecute(@ScriptDir & "\produkey-x64\ProduKey.exe")
					MsgBox(0,"Timeout", "", 0.3)
				EndIf

			Case $button9
				If $PE_flag = 1 Then
					MsgBox(0, "  System Info", "  Firmware = " & $Firmware & @CRLF _
					& @CRLF & "  Partition Style System Drive " & $OS_DriveLetter & ":  = " & $PartStyle & @CRLF _
					& @CRLF & "  Partition Style   App  Drive " & StringLeft(@ScriptDir, 2) & "  = " & $PartStyle_ScriptDrive & @CRLF _
					& @CRLF & "  OS Architecture = " & $OSArch & @CRLF & "  OS Build = " & $OSBuild _
					& @CRLF & "  Release ID = " & $ReleaseID & @CRLF & "  OS Edition = " & $OSEdition & @CRLF & "  Product Name = " & $ProductName & @CRLF _
					& @CRLF & "  OS Lang = " & $WinLang _
					& @CRLF & "  CPU Arch = " & @CPUArch & @CRLF & "  KB Layout = " & @KBLayout & @CRLF _
					& @CRLF & "  Computer Name = " & $ComputerName & @CRLF & @CRLF & "  User Name = " & $UserName & @CRLF & @CRLF & "  IP Address = " & @IPAddress1 _
					& @CRLF & @CRLF & "  Registered Owner = " & $RegOwner)
				Else
					MsgBox(0, "  System Info", "  Firmware = " & $Firmware & @CRLF _
					& @CRLF & "  Partition Style System Drive " & $OS_DriveLetter & ":  = " & $PartStyle & @CRLF _
					& @CRLF & "  Partition Style   App  Drive " & StringLeft(@ScriptDir, 2) & "  = " & $PartStyle_ScriptDrive & @CRLF _
					& @CRLF & "  OS Version = " & @OSVersion & @CRLF & "  OS Architecture = " & @OSArch & @CRLF & "  OS Build = " & $OSBuild _
					& @CRLF & "  Release ID = " & $ReleaseID & @CRLF & "  OS Edition = " & $OSEdition & @CRLF & "  Product Name = " & $ProductName & @CRLF _
					& @CRLF & "  OS Lang = " & $WinLang & @CRLF & "  OS SP = " & @OSServicePack & @CRLF & "  OS Type = " & @OSTYPE _
					& @CRLF & "  CPU Arch = " & @CPUArch & @CRLF & "  KB Layout = " & @KBLayout & @CRLF & "  MUI Lang = " & @MUILang & @CRLF _
					& @CRLF & "  Computer Name = " & $ComputerName & @CRLF & @CRLF & "  User Name = " & $UserName & @CRLF & @CRLF & "  IP Address = " & @IPAddress1 _
					& @CRLF & @CRLF & "  Registered Owner = " & $RegOwner)
				EndIf

			Case $button10
				MsgBox(0, "  Hardware Info ", "  UUID = " & $UUID & @CRLF _
				& @CRLF & "  Motherboard = " & $Mobo_Manufacturer & @CRLF & "  Model = " & $Mobo_Product & @CRLF _
				& @CRLF & "  Installed RAM = " & $RAM & " GB" & @CRLF _
				& @CRLF & "  CPU Name = " & $cpu & @CRLF & "  Number of Cores = " & $Number_Cores & @CRLF _
				& @CRLF & "  GPU Name = " & $gpu)

			Case $button11
				If FileExists(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt") Then
					FileCopy(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", @ScriptDir & "\" & $ComputerName & "_HWIDs_Info_bak.txt", 1)
					FileDelete(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt")
				EndIf
				If @OSArch = "X86" Then
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "  " & $ComputerName & "_HWIDs_Info.txt - System Info x86 - Version 4.2 - " & @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF _
					& @CRLF & "=========== " & @CRLF & "PCI Devices " & @CRLF & "=========== " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x86\devcon find pci* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "=========== " & @CRLF & "USB Devices " & @CRLF & "=========== " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x86\devcon find usb* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "============= " & @CRLF & "Input Devices " & @CRLF & "============= " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x86\devcon find hid* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "============ " & @CRLF & "ACPI Devices " & @CRLF & "============ " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x86\devcon find acpi* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "============ " & @CRLF & " HDA Audio " & @CRLF & "============ " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x86\devcon find hdaudio* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "============ " & @CRLF & "RAID Devices " & @CRLF & "============ " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x86\devcon hwids *CC_01* *Raid* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
				Else
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "  " & $ComputerName & "_HWIDs_Info.txt - System Info x64 - Version 4.2 - " & @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF _
					& @CRLF & "=========== " & @CRLF & "PCI Devices " & @CRLF & "=========== " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x64\devcon find pci* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "=========== " & @CRLF & "USB Devices " & @CRLF & "=========== " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x64\devcon find usb* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "============= " & @CRLF & "Input Devices " & @CRLF & "============= " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x64\devcon find hid* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "============ " & @CRLF & "ACPI Devices " & @CRLF & "============ " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x64\devcon find acpi* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "============ " & @CRLF & " HDA Audio " & @CRLF & "============ " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x64\devcon find hdaudio* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", _
					@CRLF & "============ " & @CRLF & "RAID Devices " & @CRLF & "============ " & @CRLF)
					RunWait(@ComSpec & " /c makebt\devcon\x64\devcon hwids *CC_01* *Raid* >> " & $ComputerName & "_HWIDs_Info.txt", @ScriptDir, @SW_HIDE)
				EndIf
				ShellExecute("notepad.exe", @ScriptDir & "\" & $ComputerName & "_HWIDs_Info.txt", @ScriptDir)

			Case $button12
				If FileExists(@ScriptDir & "\" & $ComputerName & "_Drive_Info.txt") Then
					FileCopy(@ScriptDir & "\" & $ComputerName & "_Drive_Info.txt", @ScriptDir & "\" & $ComputerName & "_Drive_Info_bak.txt", 1)
					FileDelete(@ScriptDir & "\" & $ComputerName & "_Drive_Info.txt")
				EndIf
				 RunWait(@ComSpec & " /c makebt\listusbdrives\ListUsbDrives.exe -a > " & $ComputerName & "_Drive_Info.txt", @ScriptDir, @SW_HIDE)

				ShellExecute("notepad.exe", @ScriptDir & "\" & $ComputerName & "_Drive_Info.txt", @ScriptDir)

			Case $button13
				MsgBox(0, "  Directories", "  HomeDrive = " & @HomeDrive & @CRLF & @CRLF & "  Windows Dir = " & @WindowsDir _
				& @CRLF & @CRLF & "  System Dir  = " & @SystemDir & @CRLF & @CRLF & "  Command Spec = " & @ComSpec _
				& @CRLF & @CRLF & "  Temp Dir    = " & @TempDir & @CRLF & @CRLF & "  UserProfileDir  = " & @UserProfileDir _
				& @CRLF & @CRLF & "  Local App Data  = " & @LocalAppDataDir _
				& @CRLF & @CRLF & "  MyDocumentsDir  = " & @MyDocumentsDir & @CRLF & @CRLF & "  ProgramFilesDir = " & @ProgramFilesDir & @CRLF _
				& @CRLF & "  ProgramsDir  = " & @ProgramsDir & @CRLF & @CRLF & "  StartMenuDir = " & @StartMenuDir)

			Case $button18
				_GUICtrlStatusBar_SetText($hStatus," Select Path to make File List from Path", 0)
				$TargetSelect = FileSelectFolder("Select Path to make File List from Path ", "")
				If @error Then
					_GUICtrlStatusBar_SetText($hStatus,"", 0)
					MsgBox(48,"ERROR - Path Invalid", "Error - Path Invalid")
				Else
					If StringInStr($TargetSelect, "\", 0, -1) = 0 Or StringInStr($TargetSelect, ":", 0, 1) <> 2 Then
						_GUICtrlStatusBar_SetText($hStatus,"", 0)
						MsgBox(48,"ERROR - Path Invalid", "Drive Invalid  :  Or \ Not found" & @CRLF & @CRLF & "Selected Path = " & $TargetSelect)
					Else
						_GUICtrlStatusBar_SetText($hStatus," Using Dir Command to Make File List - Wait ...", 0)
						$len_TS = StringLen($TargetSelect)
						$drv = StringLeft($TargetSelect, 1)
						If $len_TS > 3 Then
							$pos_TS = StringInStr($TargetSelect, "\", 0, -1)
							$path_folder = StringRight($TargetSelect, $len_TS - $pos_TS)
							RunWait(@ComSpec & " /u /c dir /b /s /a-d-l /on > " & '"' & @ScriptDir & "\File_List_" & $drv & "_" & $path_folder & ".txt" & '"', $TargetSelect, @SW_HIDE)
							ShellExecute("notepad.exe", "File_List_" & $drv & "_" & $path_folder & ".txt", @ScriptDir)
						Else
							RunWait(@ComSpec & " /u /c dir /b /s /a-d-l /on > " & '"' & @ScriptDir & "\File_List_" & $drv & ".txt" & '"', $TargetSelect, @SW_HIDE)
							ShellExecute("notepad.exe", "File_List_" & $drv & ".txt", @ScriptDir)
						EndIf
					EndIf
				EndIf
				_GUICtrlStatusBar_SetText($hStatus,"", 0)

			Case $button14
				If @OSArch = "X86" Then
					ShellExecute(@ScriptDir & "\cpu-z\cpuz_x32.exe")
					MsgBox(0,"Timeout", "", 0.3)
				Else
					ShellExecute(@ScriptDir & "\cpu-z\cpuz_x64.exe")
					MsgBox(0,"Timeout", "", 0.3)
				EndIf

			Case $button15
				RunWait(@ComSpec & " /c control panel", @ScriptDir, @SW_HIDE)

			Case $button16
				If FileExists(@ScriptDir & "\" & $ComputerName & "_System_Info.txt") Then
					FileCopy(@ScriptDir & "\" & $ComputerName & "_System_Info.txt", @ScriptDir & "\" & $ComputerName & "_System_Info_bak.txt", 1)
					FileDelete(@ScriptDir & "\" & $ComputerName & "_System_Info.txt")
				EndIf
				If $PE_flag = 1 Then
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_System_Info.txt", _
					@CRLF & "  " & $ComputerName & "_System_Info.txt - System Info x86 - Version 4.2 - " & @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF _
					& @CRLF & "  Firmware = " & $Firmware & @CRLF _
					& @CRLF & "  Partition Style System Drive " & $OS_DriveLetter & ":  = " & $PartStyle & @CRLF _
					& @CRLF & "  Partition Style   App  Drive " & StringLeft(@ScriptDir, 2) & "  = " & $PartStyle_ScriptDrive & @CRLF _
					& @CRLF & "  OS Architecture = " & $OSArch & @CRLF & "  OS Build = " & $OSBuild _
					& @CRLF & "  Release ID = " & $ReleaseID & @CRLF & "  OS Edition = " & $OSEdition & @CRLF & "  Product Name = " & $ProductName & @CRLF _
					& @CRLF & "  OS Lang = " & $WinLang _
					& @CRLF & "  CPU Arch = " & @CPUArch & @CRLF & "  KB Layout = " & @KBLayout & @CRLF _
					& @CRLF & "  Computer Name = " & $ComputerName & @CRLF & "  User Name = " & $UserName & @CRLF & "  IP Address = " & @IPAddress1 _
					& @CRLF & "  Registered Owner = " & $RegOwner & @CRLF _
					& @CRLF & "  Hardware UUID = " & $UUID & @CRLF _
					& @CRLF & "  Motherboard = " & $Mobo_Manufacturer & @CRLF & "  Model = " & $Mobo_Product & @CRLF _
					& @CRLF & "  Installed RAM = " & $RAM & " GB" & @CRLF _
					& @CRLF & "  CPU Name = " & $cpu & @CRLF & "  Number of Cores = " & $Number_Cores & @CRLF _
					& @CRLF & "  GPU Name = " & $gpu & @CRLF _
					& @CRLF & "  Windows Key       = " & _Decode_ProductKey("Windows") & @CRLF _
					& @CRLF & "  Windows Key DPid4 = " & _Decode_ProductKey("Windows_DPid4") & @CRLF _
					& @CRLF & "  Windows Default   = " & _Decode_ProductKey("Windows_Def") & @CRLF _
					& @CRLF & "  Windows Def DPid4 = " & _Decode_ProductKey("Windows_Def_DPid4"))
				Else
					FileWriteLine(@ScriptDir & "\" & $ComputerName & "_System_Info.txt", _
					@CRLF & "  " & $ComputerName & "_System_Info.txt - System Info x86 - Version 4.2 - " & @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF _
					& @CRLF & "  Firmware = " & $Firmware & @CRLF _
					& @CRLF & "  Partition Style System Drive " & $OS_DriveLetter & ":  = " & $PartStyle & @CRLF _
					& @CRLF & "  Partition Style   App  Drive " & StringLeft(@ScriptDir, 2) & "  = " & $PartStyle_ScriptDrive & @CRLF _
					& @CRLF & "  OS Version = " & @OSVersion & @CRLF & "  OS Architecture = " & @OSArch & @CRLF & "  OS Build = " & $OSBuild _
					& @CRLF & "  Release ID = " & $ReleaseID & @CRLF & "  OS Edition = " & $OSEdition & @CRLF & "  Product Name = " & $ProductName & @CRLF _
					& @CRLF & "  OS Lang = " & $WinLang & @CRLF & "  OS SP = " & @OSServicePack & @CRLF & "  OS Type = " & @OSTYPE _
					& @CRLF & "  CPU Arch = " & @CPUArch & @CRLF & "  KB Layout = " & @KBLayout & @CRLF & "  MUI Lang = " & @MUILang & @CRLF _
					& @CRLF & "  Computer Name = " & $ComputerName & @CRLF & "  User Name = " & $UserName & @CRLF & "  IP Address = " & @IPAddress1 _
					& @CRLF & "  Registered Owner = " & $RegOwner & @CRLF _
					& @CRLF & "  HomeDrive = " & @HomeDrive & @CRLF & "  Windows Dir = " & @WindowsDir & @CRLF & "  System Dir  = " & @SystemDir  _
					& @CRLF & "  Temp Dir    = " & @TempDir & @CRLF & "  UserProfileDir  = " & @UserProfileDir & @CRLF & "  Local App Data  = " & @LocalAppDataDir _
					& @CRLF & "  MyDocumentsDir  = " & @MyDocumentsDir & @CRLF & "  ProgramFilesDir = " & @ProgramFilesDir & @CRLF _
					& @CRLF & "  ProgramsDir  = " & @ProgramsDir & @CRLF & "  StartMenuDir = " & @StartMenuDir & @CRLF _
					& @CRLF & "  Hardware UUID = " & $UUID & @CRLF _
					& @CRLF & "  Motherboard = " & $Mobo_Manufacturer & @CRLF & "  Model = " & $Mobo_Product & @CRLF _
					& @CRLF & "  Installed RAM = " & $RAM & " GB" & @CRLF _
					& @CRLF & "  CPU Name = " & $cpu & @CRLF & "  Number of Cores = " & $Number_Cores & @CRLF _
					& @CRLF & "  GPU Name = " & $gpu & @CRLF _
					& @CRLF & "  Windows Key       = " & _Decode_ProductKey("Windows") & @CRLF _
					& @CRLF & "  Windows Key DPid4 = " & _Decode_ProductKey("Windows_DPid4") & @CRLF _
					& @CRLF & "  Windows Default   = " & _Decode_ProductKey("Windows_Def") & @CRLF _
					& @CRLF & "  Windows Def DPid4 = " & _Decode_ProductKey("Windows_Def_DPid4") & @CRLF _
					& @CRLF & "  Office   XP Key   = " & _Decode_ProductKey("Office XP") & @CRLF _
					& @CRLF & "  Office 2003 Key   = " & _Decode_ProductKey("Office 2003") & @CRLF _
					& @CRLF & "  Office 2007 Key   = " & _Decode_ProductKey("Office 2007") & @CRLF _
					& @CRLF & "  Office 2010 x86   = " & _Decode_ProductKey("Office 2010 x86") & @CRLF _
					& @CRLF & "  Office 2010 x64   = " & _Decode_ProductKey("Office 2010 x64") & @CRLF _
					& @CRLF & "  Office 2013 x86   = " & _Decode_ProductKey("Office 2013 x86") & @CRLF _
					& @CRLF & "  Office 2013 x64   = " & _Decode_ProductKey("Office 2013 x64") & @CRLF _
					& @CRLF & "  Office 365 2016 2019 x86   = " & _Decode_ProductKey("Office 2016 x86") & @CRLF _
					& @CRLF & "  Office 365 2016 2019 x64   = " & _Decode_ProductKey("Office 2016 x64"))
				EndIf

				ShellExecute("notepad.exe", @ScriptDir & "\" & $ComputerName & "_System_Info.txt", @ScriptDir)
				MsgBox(0,"Timeout", "", 0.3)

		EndSwitch
		DisableMenus(0)
	EndIf
WEnd

;===================================================================================================
Func _GetPhysicallyInstalledSystemMemory()
	Local $aRet = DllCall("Kernel32.dll", "int", "GetPhysicallyInstalledSystemMemory", "int*", "")
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[1]
EndFunc   ;==>  _GetPhysicallyInstalledSystemMemory
;===================================================================================================
Func _GetDrivePartitionStyle($sDrive = "C")
    Local $tDriveLayout = DllStructCreate('dword PartitionStyle;' & _
            'dword PartitionCount;' & _
            'byte union[40];' & _
            'byte PartitionEntry[8192]')
    Local $hDrive = DllCall("kernel32.dll", "handle", "CreateFileW", _
            "wstr", "\\.\" & $sDrive & ":", _
            "dword", 0, _
            "dword", 0, _
            "ptr", 0, _
            "dword", 3, _ ; OPEN_EXISTING
            "dword", 0, _
            "ptr", 0)
    If @error Or $hDrive[0] = Ptr(-1) Then Return SetError(@error, @extended, 0) ; INVALID_HANDLE_VALUE
    DllCall("kernel32.dll", "int", "DeviceIoControl", _
            "hwnd", $hDrive[0], _
            "dword", 0x00070050, _
            "ptr", 0, _
            "dword", 0, _
            "ptr", DllStructGetPtr($tDriveLayout), _
            "dword", DllStructGetSize($tDriveLayout), _
            "dword*", 0, _
            "ptr", 0)
    DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hDrive[0])
    Switch DllStructGetData($tDriveLayout, "PartitionStyle")
        Case 0
            Return "MBR"
        Case 1
            Return "GPT"
        Case 2
            Return "RAW"
        Case Else
            Return "UNKNOWN"
    EndSwitch
EndFunc   ;==>_GetDrivePartitionStyle
;===================================================================================================
Func _WinAPI_GetFirmwareEnvironmentVariable()
    DllCall("kernel32.dll", "dword", _
            "GetFirmwareEnvironmentVariableW", "wstr", "", _
            "wstr", "{00000000-0000-0000-0000-000000000000}", "wstr", "", "dword", 4096)
    Local $iError = DllCall("kernel32.dll", "dword", "GetLastError")
    Switch $iError[0]
        Case 1
            Return "LEGACY"
        Case 998
            Return "UEFI"
        Case Else
            Return "UNKNOWN"
    EndSwitch
EndFunc   ;==>_WinAPI_GetFirmwareEnvironmentVariable
;===================================================================================================
Func SystemFileRedirect($Wow64Number)
	If @OSArch = "X64" Then
		Local $WOW64_CHECK = DllCall("kernel32.dll", "int", "Wow64DisableWow64FsRedirection", "ptr*", 0)
		If Not @error Then
			If $Wow64Number = "On" And $WOW64_CHECK[1] <> 1 Then
				DllCall("kernel32.dll", "int", "Wow64DisableWow64FsRedirection", "int", 1)
			ElseIf $Wow64Number = "Off" And $WOW64_CHECK[1] <> 0 Then
				DllCall("kernel32.dll", "int", "Wow64EnableWow64FsRedirection", "int", 1)
			EndIf
		EndIf
	EndIf
EndFunc   ;==> SystemFileRedirect
;===================================================================================================
Func _WinLang()
	If FileExists($WindowsDir & "\System32\en-US\ieframe.dll.mui") Then $WinLang = "en-US"
	If FileExists($WindowsDir & "\System32\ar-SA\ieframe.dll.mui") Then $WinLang = "ar-SA"
	If FileExists($WindowsDir & "\System32\bg-BG\ieframe.dll.mui") Then $WinLang = "bg-BG"
	If FileExists($WindowsDir & "\System32\cs-CZ\ieframe.dll.mui") Then $WinLang = "cs-CZ"
	If FileExists($WindowsDir & "\System32\da-DK\ieframe.dll.mui") Then $WinLang = "da-DK"
	If FileExists($WindowsDir & "\System32\de-DE\ieframe.dll.mui") Then $WinLang = "de-DE"
	If FileExists($WindowsDir & "\System32\el-GR\ieframe.dll.mui") Then $WinLang = "el-GR"
	If FileExists($WindowsDir & "\System32\es-ES\ieframe.dll.mui") Then $WinLang = "es-ES"
	If FileExists($WindowsDir & "\System32\es-MX\ieframe.dll.mui") Then $WinLang = "es-MX"
	If FileExists($WindowsDir & "\System32\et-EE\ieframe.dll.mui") Then $WinLang = "et-EE"
	If FileExists($WindowsDir & "\System32\fi-FI\ieframe.dll.mui") Then $WinLang = "fi-FI"
	If FileExists($WindowsDir & "\System32\fr-FR\ieframe.dll.mui") Then $WinLang = "fr-FR"
	If FileExists($WindowsDir & "\System32\he-IL\ieframe.dll.mui") Then $WinLang = "he-IL"
	If FileExists($WindowsDir & "\System32\hr-HR\ieframe.dll.mui") Then $WinLang = "hr-HR"
	If FileExists($WindowsDir & "\System32\hu-HU\ieframe.dll.mui") Then $WinLang = "hu-HU"
	If FileExists($WindowsDir & "\System32\it-IT\ieframe.dll.mui") Then $WinLang = "it-IT"
	If FileExists($WindowsDir & "\System32\ja-JP\ieframe.dll.mui") Then $WinLang = "ja-JP"
	If FileExists($WindowsDir & "\System32\ko-KR\ieframe.dll.mui") Then $WinLang = "ko-KR"
	If FileExists($WindowsDir & "\System32\lt-LT\ieframe.dll.mui") Then $WinLang = "lt-LT"
	If FileExists($WindowsDir & "\System32\lv-LV\ieframe.dll.mui") Then $WinLang = "lv-LV"
	If FileExists($WindowsDir & "\System32\nb-NO\ieframe.dll.mui") Then $WinLang = "nb-NO"
	If FileExists($WindowsDir & "\System32\nl-NL\ieframe.dll.mui") Then $WinLang = "nl-NL"
	If FileExists($WindowsDir & "\System32\pl-PL\ieframe.dll.mui") Then $WinLang = "pl-PL"
	If FileExists($WindowsDir & "\System32\pt-BR\ieframe.dll.mui") Then $WinLang = "pt-BR"
	If FileExists($WindowsDir & "\System32\pt-PT\ieframe.dll.mui") Then $WinLang = "pt-PT"
	If FileExists($WindowsDir & "\System32\ro-RO\ieframe.dll.mui") Then $WinLang = "ro-RO"
	If FileExists($WindowsDir & "\System32\ru-RU\ieframe.dll.mui") Then $WinLang = "ru-RU"
	If FileExists($WindowsDir & "\System32\sk-SK\ieframe.dll.mui") Then $WinLang = "sk-SK"
	If FileExists($WindowsDir & "\System32\sl-SI\ieframe.dll.mui") Then $WinLang = "sl-SI"
	If FileExists($WindowsDir & "\System32\sr-Latn-CS\ieframe.dll.mui") Then $WinLang = "sr-Latn-CS"
	If FileExists($WindowsDir & "\System32\sv-SE\ieframe.dll.mui") Then $WinLang = "sv-SE"
	If FileExists($WindowsDir & "\System32\th-TH\ieframe.dll.mui") Then $WinLang = "th-TH"
	If FileExists($WindowsDir & "\System32\tr-TR\ieframe.dll.mui") Then $WinLang = "tr-TR"
	If FileExists($WindowsDir & "\System32\uk-UA\ieframe.dll.mui") Then $WinLang = "uk-UA"
	If FileExists($WindowsDir & "\System32\zh-CN\ieframe.dll.mui") Then $WinLang = "zh-CN"
	If FileExists($WindowsDir & "\System32\zh-HK\ieframe.dll.mui") Then $WinLang = "zh-HK"
	If FileExists($WindowsDir & "\System32\zh-TW\ieframe.dll.mui") Then $WinLang = "zh-TW"
EndFunc   ;==> _WinLang
;===================================================================================================
; modified for PE Offline Keys
Func _Decode_ProductKey($Product, $Offset = 0)
    Local $sKey[29], $Value = 0, $hi = 0, $n = 0, $i = 0, $dlen = 29, $slen = 15, $Result, $bKey, $iKeyOffset = 52, $RegKey, $var

    Switch $Product
        Case "Windows"
            $bKey = $bKey_Id

        Case "Windows_DPid4"
            $bKey = $bKey_Id4
            $iKeyOffset = 0x328

		Case "Windows_Def"
            $bKey = $bKey_Def

        Case "Windows_Def_DPid4"
            $bKey = $bKey_Def4
            $iKeyOffset = 0x328

        Case "Office XP"
            $RegKey = 'HKLM\SOFTWARE\Microsoft\Office\10.0\Registration'
            If @OSArch = 'x64' Then $RegKey = 'HKLM64\SOFTWARE\Wow6432Node\Microsoft\Office\10.0\Registration'
            For $i = 1 To 100
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next

        Case "Office 2003"
            $RegKey = 'HKLM\SOFTWARE\Microsoft\Office\11.0\Registration'
            If @OSArch = 'x64' Then $RegKey = 'HKLM64\SOFTWARE\Wow6432Node\Microsoft\Office\11.0\Registration'
            For $i = 1 To 100
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next

        Case "Office 2007"
            $RegKey = 'HKLM\SOFTWARE\Microsoft\Office\12.0\Registration'
            If @OSArch = 'x64' Then $RegKey = 'HKLM64\SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Registration'
            For $i = 1 To 100
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next

        Case "Office 2010 x86"
            $RegKey = 'HKLM\SOFTWARE\Microsoft\Office\14.0\Registration'
            If @OSArch = 'x64' Then $RegKey = 'HKLM64\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Registration'
            For $i = 1 To 100
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next
            $iKeyOffset = 0x328

        Case "Office 2010 x64"
            If @OSArch <> 'x64' Then Return SetError(1, 0, "Product not found")
            $RegKey = 'HKLM64\SOFTWARE\Microsoft\Office\14.0\Registration'
            For $i = 1 To 100
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next
            $iKeyOffset = 0x328

		Case "Office 2013 x86"
            $RegKey = 'HKLM\SOFTWARE\Microsoft\Office\15.0\Registration'
            If @OSArch = 'x64' Then $RegKey = 'HKLM64\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Registration'
            For $i = 1 To 1024
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next
            $iKeyOffset = 0x328

        Case "Office 2013 x64"
            If @OSArch <> 'x64' Then Return SetError(1, 0, "Product not found")
            $RegKey = 'HKLM64\SOFTWARE\Microsoft\Office\15.0\Registration'
            For $i = 1 To 1024
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next
            $iKeyOffset = 0x328

		Case "Office 2016 x86"
            $RegKey = 'HKLM\SOFTWARE\Microsoft\Office\16.0\Registration'
            If @OSArch = 'x64' Then $RegKey = 'HKLM64\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Registration'
            For $i = 1 To 1024
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next
            $iKeyOffset = 0x328

        Case "Office 2016 x64"
            If @OSArch <> 'x64' Then Return SetError(1, 0, "Product not found")
            $RegKey = 'HKLM64\SOFTWARE\Microsoft\Office\16.0\Registration'
            For $i = 1 To 1024
                $var = RegEnumKey($RegKey, $i)
                If @error <> 0 Then ExitLoop
                $bKey = RegRead($RegKey & '\' & $var, 'DigitalProductId')
                If Not @error Then ExitLoop
            Next
            $iKeyOffset = 0x328

		Case Else
        Return SetError(1, 0, "Product not supported")
    EndSwitch

    If Not BinaryLen($bKey) Then Return ""

    Local $aKeys[BinaryLen($bKey)]
    For $i = 0 To UBound($aKeys) - 1
        $aKeys[$i] = Int(BinaryMid($bKey, $i + 1, 1))
    Next

	Local Const $isWin8 = BitAND(BitShift($aKeys[$iKeyOffset + 14], 3), 1)
	$aKeys[$iKeyOffset + 14] = BitOR(BitAND($aKeys[$iKeyOffset + 14], 0xF7), BitShift(BitAND($isWin8, 2), -2))

    $i = 24
    Local $sChars = "BCDFGHJKMPQRTVWXY2346789", $iCur, $iX, $sKeyOutput, $iLast
    While $i > -1
        $iCur = 0
        $iX = 14
        While $iX > -1
			$iCur = BitShift($iCur, -8)
			$iCur = $aKeys[$iX + $iKeyOffset] + $iCur
            $aKeys[$iX + $iKeyOffset] = Int($iCur / 24)
            $iCur = Mod($iCur, 24)
            $iX -= 1
        WEnd
        $i -= 1
        $sKeyOutput = StringMid($sChars, $iCur + 1, 1) & $sKeyOutput
        $iLast = $iCur
    WEnd

    If $isWin8 Then
        $sKeyOutput = StringMid($sKeyOutput, 2, $iLast) & "N" & StringTrimLeft($sKeyOutput, $iLast + 1)
    EndIf

    Return StringRegExpReplace($sKeyOutput, '(\w{5})(\w{5})(\w{5})(\w{5})(\w{5})', '\1-\2-\3-\4-\5')

EndFunc   ;==>_Decode_ProductKey
;===================================================================================================
Func DisableMenus($endis)
	If $endis = 0 Then
		$endis = $GUI_ENABLE
	Else
		$endis = $GUI_DISABLE
	EndIf
	GUICtrlSetState($button1, $endis)
	GUICtrlSetState($button2, $endis)
	GUICtrlSetState($button4, $endis)
	GUICtrlSetState($button5, $endis)
	GUICtrlSetState($button6, $endis)
	GUICtrlSetState($button8, $endis)
	GUICtrlSetState($button9, $endis)
	GUICtrlSetState($button10, $endis)
	GUICtrlSetState($button11, $endis)
	GUICtrlSetState($button12, $endis)
	GUICtrlSetState($button13, $endis)
	GUICtrlSetState($button14, $endis)
	GUICtrlSetState($button15, $endis)
	GUICtrlSetState($button16, $endis + $GUI_FOCUS)
	GUICtrlSetState($button17, $endis)
	GUICtrlSetState($button18, $endis)
	If $PE_flag = 1 Then
		GUICtrlSetState($button3, $GUI_DISABLE)
		GUICtrlSetState($button7, $GUI_DISABLE)
		GUICtrlSetState($button19, $GUI_DISABLE)
		GUICtrlSetState($button20, $GUI_DISABLE)
	Else
		GUICtrlSetState($button3, $endis)
		If @OSVersion = "WIN_10" Or @OSVersion = "WIN_81" Or @OSVersion = "WIN_8" Then
			GUICtrlSetState($button7, $endis)
		Else
			GUICtrlSetState($button7, $GUI_DISABLE)
		EndIf
		GUICtrlSetState($button19, $endis)
		GUICtrlSetState($button20, $endis)
	EndIf

EndFunc ;==>DisableMenus
;===================================================================================================
;~ 		Func _Windows_drive()
;~ 			Local $WinDrvSelect, $Tdrive, $FSvar, $valid = 0, $ValidDrives, $RemDrives, $DriveSysType
;~ 			Local $NoDrive[3] = ["A:", "B:", "X:"], $FileSys[1] = ["NTFS"]
;~ 			Local $pos

;~ 			; DisableMenus(1)
;~ 			; $WIM_Path = ""
;~ 			$ValidDrives = DriveGetDrive( "FIXED" )
;~ 			_ArrayPush($ValidDrives, "")
;~ 			_ArrayPop($ValidDrives)
;~ 			$RemDrives = DriveGetDrive( "REMOVABLE" )
;~ 			_ArrayPush($RemDrives, "")
;~ 			_ArrayPop($RemDrives)
;~ 			_ArrayConcatenate($ValidDrives, $RemDrives)
;~ 			; _ArrayDisplay($ValidDrives)

;~ 			$OS_DriveLetter = "C"

;~ 			$WinDrvSelect = FileSelectFolder("Select Windows Folder on System Drive ", "")
;~ 			If @error Then
;~ 				Return
;~ 			EndIf

;~ 			If StringRight($WinDrvSelect, 9) <> ":\Windows" Then
;~ 				MsgBox(48,"ERROR - Path Invalid", ":\Windows Folder Not Selected" & @CRLF & @CRLF & "Selected Path = " & $WinDrvSelect)
;~ 				Return
;~ 			EndIf

;~ 		;~ 		$pos = StringInStr($WinDrvSelect, "\", 0, -1)
;~ 		;~ 		If $pos = 0 Then
;~ 		;~ 			MsgBox(48,"ERROR - Path Invalid", "Path Invalid - No Backslash Found" & @CRLF & @CRLF & "Selected Path = " & $WinDrvSelect)
;~ 		;~ 			Return
;~ 		;~ 		EndIf

;~ 		;~ 		$pos = StringInStr($WinDrvSelect, " ", 0, -1)
;~ 		;~ 		If $pos Then
;~ 		;~ 			MsgBox(48,"ERROR - Path Invalid", "Path Invalid - Space Found" & @CRLF & @CRLF & "Selected Path = " & $WinDrvSelect & @CRLF & @CRLF _
;~ 		;~ 			& "Solution - Use simple Path without Spaces ")
;~ 		;~ 			Return
;~ 		;~ 		EndIf

;~ 		;~ 		$pos = StringInStr($WinDrvSelect, ":", 0, 1)
;~ 		;~ 		If $pos <> 2 Then
;~ 		;~ 			MsgBox(48,"ERROR - Path Invalid", "Drive Invalid - : Not found" & @CRLF & @CRLF & "Selected Path = " & $WinDrvSelect)
;~ 		;~ 			Return
;~ 		;~ 		EndIf

;~ 			$Tdrive = StringLeft($WinDrvSelect, 2)
;~ 			FOR $d IN $ValidDrives
;~ 				If $d = $Tdrive Then
;~ 					$valid = 1
;~ 					ExitLoop
;~ 				EndIf
;~ 			NEXT
;~ 			FOR $d IN $NoDrive
;~ 				If $d = $Tdrive Then
;~ 					$valid = 0
;~ 					MsgBox(48, "ERROR - Drive NOT Valid", " Drive A: B: and X: ", 3)
;~ 					Return
;~ 				EndIf
;~ 			NEXT
;~ 			If $valid And DriveStatus($Tdrive) <> "READY" Then
;~ 				$valid = 0
;~ 				MsgBox(48, "ERROR - Drive NOT Ready", "Drive NOT READY", 3)
;~ 				DisableMenus(0)
;~ 				Return
;~ 			EndIf
;~ 			If $valid Then
;~ 				$FSvar = DriveGetFileSystem( $Tdrive )
;~ 				FOR $d IN $FileSys
;~ 					If $d = $FSvar Then
;~ 						$valid = 1
;~ 						ExitLoop
;~ 					Else
;~ 						$valid = 0
;~ 					EndIf
;~ 				NEXT
;~ 				IF Not $valid Then
;~ 					MsgBox(48, "ERROR - Invalid FileSystem", " NTFS FileSystem NOT Found ", 3)
;~ 					Return
;~ 				EndIf
;~ 			EndIf

;~ 			$DriveSysType=DriveGetType($Tdrive)

;~ 			If $DriveSysType="Removable" Or $DriveSysType="Fixed" Then
;~ 			Else
;~ 				MsgBox(48, "ERROR - Target System Drive NOT Valid", "Target System Drive = " & $Tdrive & " Not Valid " & @CRLF & @CRLF & _
;~ 				" Only Removable Or Fixed Drive allowed ", 0)
;~ 				Return
;~ 			EndIf

;~ 			If Not FileExists(StringLeft($WinDrvSelect, 2) & "\Windows\System32\config\software") Then
;~ 				MsgBox(48,"ERROR - Path Invalid", "Windows\System32\config\software Not Found" & @CRLF & @CRLF & "Selected Path = " & $WinDrvSelect)
;~ 				Return
;~ 			EndIf
;~ 			If Not FileExists(StringLeft($WinDrvSelect, 2) & "\Windows\System32\config\system") Then
;~ 				MsgBox(48,"ERROR - Path Invalid", "Windows\System32\config\system Not Found" & @CRLF & @CRLF & "Selected Path = " & $WinDrvSelect)
;~ 				Return
;~ 			EndIf

;~ 			If $valid Then

;~ 				$OS_DriveLetter = StringLeft($WinDrvSelect, 1)

;~ 			EndIf
;~ 		EndFunc   ;==> _Windows_drive
;~ 		;===================================================================================================
