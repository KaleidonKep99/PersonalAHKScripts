; Microsoft Office Keyboard RT9450
; written by pondermatic 2015-06-03
; updated by KaleidonKep99 2018-11-22
; This AutoHotkey 1.1 script makes the scroll wheel work with Windows 7, 8 and 10.
; The wheel has four speeds: slowest, slow, fast, fastest.

; requires AHKHID library from https://github.com/jleb/AHKHID

;  0  1  2  3  4  5  6  7 
; -- -- -- -- -- -- -- -- -----
; 01 01 00 00 00 00 00 00 Help
; 01 00 00 00 01 00 00 00 Office Home
; 01 00 10 00 00 00 00 00 Task Pane
; 01 00 80 00 00 00 00 00 New
; 01 00 00 01 00 00 00 00 Open
; 01 00 00 02 00 00 00 00 Close
; 01 00 00 00 20 00 00 00 Reply
; 01 00 00 00 40 00 00 00 Fwd
; 01 00 00 00 80 00 00 00 Send
; 01 00 20 00 00 00 00 00 Spell
; 01 00 00 04 00 00 00 00 Save
; 01 00 00 08 00 00 00 00 Print
; 01 00 00 10 00 00 00 00 Undo
; 01 00 00 00 10 00 00 00 Redo
; 01 00 00 00 00 01 00 00 =
; 01 00 00 00 00 02 00 00 (
; 01 00 00 00 00 04 00 00 )
;                         Back space
; 
; 01 10 00 00 00 00 00 00 Word
; 01 20 00 00 00 00 00 00 Excel
; 01 00 00 00 02 00 00 00 Web/Home
; 01 40 00 00 00 00 00 00 Mail
; 01 80 00 00 00 00 00 00 Calendar
; 01 00 40 00 00 00 00 00 Files
; 01 00 01 00 00 00 00 00 Calculator
; 01 02 00 00 00 00 00 00 Mute
; 01 04 00 00 00 00 00 00 Volume -
; 01 08 00 00 00 00 00 00 Volume +
; 01 00 02 00 00 00 00 00 Log Off
;                         Sleep
; 
; 01 00 00 00 04 00 00 00 Back
; 01 00 00 00 08 00 00 00 Forward
; 01 00 00 00 00 00 01 00 scroll up slowest
; 01 00 00 00 00 00 21 00 scroll up slow
; 01 00 00 00 00 00 41 00 scroll up fast
; 01 00 00 00 00 00 61 00 scroll up fastest
; 01 00 00 00 00 00 1F 00 scroll down slowest
; 01 00 00 00 00 00 3F 00 scroll down slow
; 01 00 00 00 00 00 5F 00 scroll down fast
; 01 00 00 00 00 00 7F 00 scroll down fastest
; 01 00 00 40 00 00 00 00 Cut
; 01 00 00 20 00 00 00 00 Copy
; 01 00 00 80 00 00 00 00 Paste
; 01 00 08 00 00 00 00 00 Application left (alt-tab)
; 01 00 04 00 00 00 00 00 Application right
; 
;  7 Qualifier Key
; -- -------------
; 00 none
; 01 left control
; 02 left shift
; 04 left alt	
; 08 left windows
; 10 right control
; 20 right shift	
; 40 right alt


#SingleInstance force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Set up the AHKHID constants
AHKHID_UseConstants()

; Create GUI to receive messages (won't be displayed)
Gui, +LastFound
hGui := WinExist()

; Intercept WM_INPUT messages
WM_INPUT := 0xFF
OnMessage(WM_INPUT, "InputMsg")

; Register with RIDEV_INPUTSINK (so that data is received even in the background)
result := AHKHID_Register(12, 1, hGui, RIDEV_INPUTSINK)

Return


InputMsg(wParam, lParam)
{
	Local devh, bData, bSize, wheel, direction, speed
    
	Critical
    
	; Check for the correct HID device
	devh := AHKHID_GetInputInfo(lParam, II_DEVHANDLE)
	vn := AHKHID_GetDevInfo(devh, DI_HID_VERSIONNUMBER, True)
	if (devh == -1)
		or (AHKHID_GetDevInfo(devh, DI_DEVTYPE, True) != RIM_TYPEHID)
		or (AHKHID_GetDevInfo(devh, DI_HID_VENDORID, True) != 1118) ; Microsoft 0x045E
		or (AHKHID_GetDevInfo(devh, DI_HID_PRODUCTID, True) != 72) ; Office keyboard 0x0048
		or (AHKHID_GetDevInfo(devh, DI_HID_VERSIONNUMBER, True) != 292) ; 0x0124
	{
		return
	}
    
	bSize := AHKHID_GetInputData(lParam, bData)
;	hex := Bin2Hex(&bData, bSize)
;	OutputDebug Data: %hex%
;	return
    
	if (bSize != 8) {
		return ; we are expecting exactly 8 bytes
	}
    
	b0 := NumGet(bData, 0, "UChar")
	b1 := NumGet(bData, 1, "UChar")
	b2 := NumGet(bData, 2, "UChar")
	b3 := NumGet(bData, 3, "UChar")
	b4 := NumGet(bData, 4, "UChar")
	b5 := NumGet(bData, 5, "UChar")
	b6 := NumGet(bData, 6, "UChar")
	b7 := NumGet(bData, 7, "UChar")
    
;	Help
	if (b1 == 0x01) and (b2 = 0x00) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send {F1}
	}

;	New
	if (b1 == 0x00) and (b2 = 0x80) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^N
	}

;	Open
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x01) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^{F12}
	}

;	Print
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x08) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^P
	}

;	Save
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x04) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^S
	}

;	Undo
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x10) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^Z
	}

;	Redo
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x00) and (b4 = 0x10) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^Y
	}

;	Word
	if (b1 == 0x00) and (b2 = 0x40) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Run, %A_MyDocuments%
	}

;	Excel
	if (b1 == 0x20) and (b2 = 0x00) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		run "excel.exe"
	}

;	Calendar
	if (b1 == 0x80) and (b2 = 0x00) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		run "Outlook.exe" /select outlook:calendar
	}

;	Windows Explorer
	if (b1 == 0x10) and (b2 = 0x00) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		run "winword.exe"
	}

;	Cut to clipboard
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x40) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^x
	}

;	Copy to clipboard
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x20) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^c
	}

;	Paste from clipboard
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x80) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send ^v
	}

;	Disconnect, repurposed as "Play/Pause media"
	if (b1 == 0x00) and (b2 = 0x02) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send {Media_Play_Pause}
	}

;	Previous app, repurposed as "Previous media"
	if (b1 == 0x00) and (b2 = 0x08) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send {Media_Prev}
	}

;	Next app, repurposed as "Next media"
	if (b1 == 0x00) and (b2 = 0x04) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x00) and (b6 = 0x00) and (b7 = 0x00) {
		Send {Media_Next}
	}

;	Various miscellaneous keys
	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x01) and (b6 = 0x00) and (b7 = 0x00) {
		Send {=}
	}

	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x02) and (b6 = 0x00) and (b7 = 0x00) {
		Send (
	}

	if (b1 == 0x00) and (b2 = 0x00) and (b3 = 0x00) and (b4 = 0x00) and (b5 = 0x04) and (b6 = 0x00) and (b7 = 0x00) {
		Send )
	}
;	Various miscellaneous keys
    
	wheelData := NumGet(bData, 6, "UChar")
	if (wheelData) {
		HandleWheel(wheelData)
	}
}

HandleWheel(data) {
	direction := (data & 0x0F) >> 3
	speed := (data>>4) // 2 * 2 + 1
    
	if (direction == 1) {
		Send {WheelDown %speed%}
	} else {
		Send {WheelUp %speed%}
	}
}

; By Laszlo, adapted by TheGood
; http://www.autohotkey.com/forum/viewtopic.php?p=377086#377086
Bin2Hex(addr,len) {
	Static fun, ptr 
	If (fun = "") {
		If A_IsUnicode
			If (A_PtrSize = 8)
				h=4533c94c8bd14585c07e63458bd86690440fb60248ffc2418bc9410fb6c0c0e8043c090fb6c00f97c14180e00f66f7d96683e1076603c8410fb6c06683c1304180f8096641890a418bc90f97c166f7d94983c2046683e1076603c86683c13049ffcb6641894afe75a76645890ac366448909c3
			Else h=558B6C241085ED7E5F568B74240C578B7C24148A078AC8C0E90447BA090000003AD11BD2F7DA66F7DA0FB6C96683E2076603D16683C230668916240FB2093AD01BC9F7D966F7D96683E1070FB6D06603CA6683C13066894E0283C6044D75B433C05F6689065E5DC38B54240833C966890A5DC3
		Else h=558B6C241085ED7E45568B74240C578B7C24148A078AC8C0E9044780F9090F97C2F6DA80E20702D1240F80C2303C090F97C1F6D980E10702C880C1308816884E0183C6024D75CC5FC606005E5DC38B542408C602005DC3
		VarSetCapacity(fun, StrLen(h) // 2)
		Loop % StrLen(h) // 2
			NumPut("0x" . SubStr(h, 2 * A_Index - 1, 2), fun, A_Index - 1, "Char")
		ptr := A_PtrSize ? "Ptr" : "UInt"
		DllCall("VirtualProtect", ptr, &fun, ptr, VarSetCapacity(fun), "UInt", 0x40, "UInt*", 0)
	}
	VarSetCapacity(hex, A_IsUnicode ? 4 * len + 2 : 2 * len + 1)
	DllCall(&fun, ptr, &hex, ptr, addr, "UInt", len, "CDecl")
	VarSetCapacity(hex, -1) ; update StrLen
	Return hex
}
