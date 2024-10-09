#AutoIt3Wrapper_Change2CUI=Y
#NoTrayIcon

#include <WinAPIHObj.au3>
#include <WinAPIError.au3>
#include <WinAPISysWin.au3>
#include <WinAPIProc.au3>
#include <WinAPIFiles.au3>
#include <String.au3>

#include "DllStructEx.au3"
#include "DllStructEx.debug.au3"

#Region Global variables
Global Const $tagCONSOLE_CURSOR_INFO = "dword dwSize;int bVisible"

Global Const $tagINPUT_RECORD = _
"WORD  EventType;"& _
"union {"& _
"    KEY_EVENT_RECORD          KeyEvent;"& _
"    MOUSE_EVENT_RECORD        MouseEvent;"& _
"    WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeEvent;"& _
"    MENU_EVENT_RECORD         MenuEvent;"& _
"    FOCUS_EVENT_RECORD        FocusEvent;"& _
"} Event;"

Global Const $EventType_FOCUS_EVENT = 0x0010, $EventType_KEY_EVENT = 0x0001, $EventType_MENU_EVENT = 0x0008, $EventType_MOUSE_EVENT = 0x0002, $EventType_WINDOW_BUFFER_SIZE_EVENT = 0x0004

Global Const $tagKEY_EVENT_RECORD = _
"BOOL  bKeyDown;"& _
"WORD  wRepeatCount;"& _
"WORD  wVirtualKeyCode;"& _
"WORD  wVirtualScanCode;"& _
"union {"& _
"  WCHAR UnicodeChar;"& _
"  CHAR  AsciiChar;"& _
"} uChar;"& _
"DWORD dwControlKeyState;"

Global Const $tagMOUSE_EVENT_RECORD = _
"COORD dwMousePosition;" & _
"DWORD dwButtonState;" & _
"DWORD dwControlKeyState;" & _
"DWORD dwEventFlags;"

Global Const $tagCOORD = _
"SHORT X;"& _
"SHORT Y;"

Global Const $tagWINDOW_BUFFER_SIZE_RECORD = _
"COORD dwSize;"

Global Const $tagMENU_EVENT_RECORD = _
"UINT dwCommandId;"

Global Const $tagFOCUS_EVENT_RECORD = _
"BOOL bSetFocus;"

Global Const $tagCONSOLE_SCREEN_BUFFER_INFO = _
"COORD dwSize;"& _
"COORD dwCursorPosition;"& _
"WORD wAttributes;"& _
"SMALL_RECT srWindow;"& _
"COORD dwMaximumWindowSize;"

Global Const $tagSMALL_RECT = _
"SHORT Left;"& _
"SHORT Top;"& _
"SHORT Right;"& _
"SHORT Bottom;"
#EndRegion Global variables

#Region Functions
Func AllocConsole()
    Local $aRet = DllCall("kernel32.dll", "BOOL", "AllocConsole")
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func FreeConsole()
    Local $aRet = DllCall("kernel32.dll", "BOOL", "FreeConsole")
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func IsDebuggerPresent()
    Local $aRet = DllCall("kernel32.dll", "BOOL", "IsDebuggerPresent")
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func SetConsoleMode($hConsoleHandle, $dwMode)
    Local $aRet = DllCall("kernel32.dll", "BOOL", "SetConsoleMode", "HANDLE", $hConsoleHandle, "DWORD", $dwMode)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func SetStdHandle($nStdHandle, $hHandle) ;returns BOOL
	;http://msdn.microsoft.com/en-us/library/windows/desktop/ms686244%28v=vs.85%29.aspx
    Local $aRet = DllCall("kernel32.dll", "BOOL", "SetStdHandle", "DWORD", $nStdHandle, "HANDLE", $hHandle)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func CreateFile($lpFileName, $dwDesiredAccess, $dwShareMode, $lpSecurityAttributes, $dwCreationDisposition, $dwFlagsAndAttributes)
    Local $aResult = DllCall("kernel32.dll", "handle", "CreateFileW", "wstr", $lpFileName, "dword", $dwDesiredAccess, "dword", $dwShareMode, "struct*", $lpSecurityAttributes, "dword", $dwCreationDisposition, "dword", $dwFlagsAndAttributes, "ptr", 0)
	If @error Or ($aResult[0] = Ptr(-1)) Then Return SetError(@error, @extended, 0) ; $INVALID_HANDLE_VALUE
	Return $aResult[0]
EndFunc

Func GetConsoleCursorInfo($hConsole, $lpConsoleCursorInfo)
    Local $aRet = DllCall("kernel32.dll", "BOOL", "GetConsoleCursorInfo", "hwnd", $hConsole, "STRUCT*", $lpConsoleCursorInfo)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func SetConsoleCursorInfo($hConsole, $lpConsoleCursorInfo)
    Local $aRet = DllCall("kernel32.dll", "BOOL", "SetConsoleCursorInfo", "hwnd", $hConsole, "STRUCT*", $lpConsoleCursorInfo)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func SetConsoleCtrlHandler($HandlerRoutine, $Add)
    Local $aRet = DllCall("kernel32.dll", "BOOL", "SetConsoleCtrlHandler", "ptr", $HandlerRoutine, "BOOL", $Add)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func GetConsoleScreenBufferInfo($hConsole, $tScreenBufferInfo)
    $aRet = DllCall("kernel32.dll", "bool", "GetConsoleScreenBufferInfo", "hwnd", $hConsole, "STRUCT*", $tScreenBufferInfo)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func GetConsoleCursorPosition($hConsoleOutput, $dwCursorPosition, $tScreenBufferInfo)
    GetConsoleScreenBufferInfo($hConsoleOutput, $tScreenBufferInfo)
    If @error <> 0 Then Return SetError(@error, @extended, False)
    $dwCursorPosition = DllStructCreate($tagCOORD, DllStructGetPtr($tScreenBufferInfo, 2))
    $dwCursorPosition.X = $dwCursorPosition.X
    $dwCursorPosition.Y = $dwCursorPosition.Y
    Return True
EndFunc

Func SetConsoleCursorPosition($hConsoleOutput, $dwCursorPosition)
    Local $aRet = DllCall("kernel32.dll", "BOOL", "SetConsoleCursorPosition", "HANDLE", $hConsoleOutput, "STRUCT", $dwCursorPosition)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func ReadConsoleInput($hConsoleInput, $lpBuffer, $nLength)
    Local $aRet = DllCall("kernel32.dll", "BOOL", "ReadConsoleInput", "HANDLE", $hConsoleInput, "STRUCT*", $lpBuffer, "DWORD", $nLength, "DWORD*", 0)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func FillConsoleOutputCharacter($hConsoleInput, $cCharacter, $nLength, $dwWriteCoord)
    Local $aRet = DllCall("kernel32.dll", "BOOL", "FillConsoleOutputCharacterW", "HANDLE", $hConsoleInput, "BYTE", $cCharacter, "DWORD", $nLength, "STRUCT", $dwWriteCoord, "DWORD*", 0)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func SetConsoleScreenBufferSize($hConsoleOutput, $dwSize)
    Local $aRet = DllCall("kernel32.dll", "BOOL", "SetConsoleScreenBufferSize", "HANDLE", $hConsoleOutput, "STRUCT", $dwSize)
    If @error <> 0 Then Return SetError(@error, 0, False)
    Return $aRet[0] <> 0
EndFunc

Func _ConsoleWriteCenter($columns, $rows, $hStdOut, $sText)
    Local $_tCOORD = DllStructCreate($tagCOORD)
    Local $tCOORD = DllStructCreate($tagCOORD)

    ;GetConsoleCursorPosition($hStdOut, $_tCOORD, $tScreenBufferInfo)

    Local $s = $sText
    Local $l = StringLen($s)
    $tCOORD.X = Round($columns / 2) - Round($l / 2)
    $tCOORD.Y = Round($rows / 2)
    SetConsoleCursorPosition($hStdOut, $tCOORD)
    _WinAPI_WriteConsole($hStdOut, $s)

    SetConsoleCursorPosition($hStdOut, $_tCOORD)
EndFunc

Func _ConsoleHideCursor($hStdOut)
    Local $tConsoleCursorInfo = DllStructCreate($tagCONSOLE_CURSOR_INFO)
    GetConsoleCursorInfo($hStdOut, $tConsoleCursorInfo)
    $tConsoleCursorInfo.bVisible = 0
    SetConsoleCursorInfo($hStdOut, $tConsoleCursorInfo)
EndFunc
#EndRegion Functions

$a = AllocConsole()
$aRet = DllCall("kernel32.dll", "ptr", "GetConsoleWindow")
_WinAPI_ShowWindow($aRet[0], @SW_SHOWNORMAL)

$hConOut = CreateFile("CONOUT$", BitOR($GENERIC_READ, $GENERIC_WRITE), BitOR($FILE_SHARE_READ, $FILE_SHARE_WRITE), Null, $OPEN_EXISTING, $FILE_ATTRIBUTE_NORMAL)
$hConIn = CreateFile("CONIN$", BitOR($GENERIC_READ, $GENERIC_WRITE), BitOR($FILE_SHARE_READ, $FILE_SHARE_WRITE), Null, $OPEN_EXISTING, $FILE_ATTRIBUTE_NORMAL)

Global Const $STD_INPUT_HANDLE = -10
Global Const $STD_OUTPUT_HANDLE = -11
Global Const $STD_ERROR_HANDLE = -12

SetStdHandle($STD_OUTPUT_HANDLE, $hConOut)
SetStdHandle($STD_INPUT_HANDLE, $hConIn)

$hStdIn = _WinAPI_GetStdHandle(0)
$hStdOut = _WinAPI_GetStdHandle(1)

_ConsoleHideCursor($hStdOut)

Global Const $ENABLE_EXTENDED_FLAGS = 0x0080

;FIXME: store the original mode and restore on exit, to avoid parent console being affected.
SetConsoleMode($hStdIn, $ENABLE_EXTENDED_FLAGS)

Global $iRun = 1

Global Const $CTRL_C_EVENT = 0
Global Const $CTRL_BREAK_EVENT = 1
Global Const $CTRL_CLOSE_EVENT = 2
Global Const $CTRL_LOGOFF_EVENT = 5
Global Const $CTRL_SHUTDOWN_EVENT = 6

Func HandlerRoutine($dwCtrlType)
    ;ConsoleWrite("HandlerRoutine"&@CRLF)
    $iRun = 0
    ;https://stackoverflow.com/a/48190051
    DllCall("kernel32.dll", "NONE", "ExitThread", "DWORD", 0)
    Return True
EndFunc

$hHandlerRoutine = DllCallbackRegister(HandlerRoutine, "BOOL", "DWORD")
$pHandlerRoutine = DllCallbackGetPtr($hHandlerRoutine)

SetConsoleCtrlHandler($pHandlerRoutine, 1)

Global $tCOORD = DllStructCreate($tagCOORD)

GLobal $oScreenBufferInfo = DllStructExCreate($tagCONSOLE_SCREEN_BUFFER_INFO)
Global Const $tScreenBufferInfo = DllStructExGetStruct($oScreenBufferInfo)
Global Const $tSrWindow = DllStructExGetStruct($oScreenBufferInfo.srWindow)

GetConsoleScreenBufferInfo($hStdOut, $tScreenBufferInfo)
Global $columns = $tSrWindow.Right - $tSrWindow.Left + 1
Global $rows = $tSrWindow.Bottom - $tSrWindow.Top + 1

;ConsoleWrite(StringFormat("columns: %d\n", $columns))
;ConsoleWrite(StringFormat("rows: %d\n", $rows))
;ConsoleWrite(StringFormat("%s x %s\n", $oScreenBufferInfo.dwSize.X, $oScreenBufferInfo.dwSize.y))
Global $tScreenBufferSize = DllStructCreate($tagCOORD)
$tScreenBufferSize.X = $columns
$tScreenBufferSize.Y = $rows
SetConsoleScreenBufferSize($hStdOut, $tScreenBufferSize)

_ConsoleWriteCenter($columns, $rows, $hStdOut, StringFormat("%s x %s", $columns, $rows))

$tCOORD.X = 0
$tCOORD.Y = 0

Global Const $FOCUS_EVENT = 0x0010
Global Const $KEY_EVENT = 0x0001
Global Const $MENU_EVENT = 0x0008
Global Const $MOUSE_EVENT = 0x0002
Global Const $WINDOW_BUFFER_SIZE_EVENT = 0x0004

$oINPUT_RECORD = DllStructExCreate($tagINPUT_RECORD)
$tINPUT_RECORD = DllStructExGetStruct($oINPUT_RECORD)
$tEvent = DllStructExGetStruct($oINPUT_RECORD.Event)
$tFocusEvent = DllStructExGetStruct($oINPUT_RECORD.Event.FocusEvent)
$tKeyEvent = DllStructExGetStruct($oINPUT_RECORD.Event.KeyEvent)

ConsoleWrite(StringFormat("$tFocusEvent: %s\n", DllStructGetPtr($tFocusEvent)))
ConsoleWrite(StringFormat("$tKeyEvent: %s\n", DllStructGetPtr($tKeyEvent)))

DllStructExDisplay($oINPUT_RECORD)

While $iRun
    Sleep(10)
    If DllCall("kernel32.dll", "BOOL", "GetNumberOfConsoleInputEvents", "handle", $hStdIn, "DWORD*", 0)[2] > 0 Then
        SetConsoleCursorPosition($hStdOut, $tCOORD)
        _WinAPI_WriteConsole($hStdOut, _StringRepeat(" ", $columns))
        $tCOORD.Y = 1
        SetConsoleCursorPosition($hStdOut, $tCOORD)
        _WinAPI_WriteConsole($hStdOut, _StringRepeat(" ", $columns))
        $tCOORD.Y = 0
        SetConsoleCursorPosition($hStdOut, $tCOORD)
        ReadConsoleInput($hStdIn, $tINPUT_RECORD, 1)
        ;_WinAPI_WriteConsole($hStdOut, DllStructGetData($tINPUT_RECORD, 1))
        Switch ($oINPUT_RECORD.EventType)
            Case $EventType_FOCUS_EVENT
                _WinAPI_WriteConsole($hStdOut, "FOCUS_EVENT")
                $tCOORD.Y = 1
                SetConsoleCursorPosition($hStdOut, $tCOORD)
                ;_WinAPI_WriteConsole($hStdOut, StringFormat("bSetFocus: %s", $oINPUT_RECORD.Event.FocusEvent.bSetFocus))
                _WinAPI_WriteConsole($hStdOut, StringFormat("bSetFocus: %s", $tFocusEvent.bSetFocus))
            Case $EventType_KEY_EVENT
                _WinAPI_WriteConsole($hStdOut, "KEY_EVENT")
                $tCOORD.Y = 1
                SetConsoleCursorPosition($hStdOut, $tCOORD)
                $oEvent = $oINPUT_RECORD.Event
                $oKeyEvent = $oEvent.KeyEvent
                $ouChar = $oKeyEvent.uChar
                _WinAPI_WriteConsole($hStdOut, StringFormat("uChar: %s", $oINPUT_RECORD.Event.KeyEvent.uChar.AsciiChar))
            Case $EventType_MENU_EVENT
                _WinAPI_WriteConsole($hStdOut, "MENU_EVENT")
                $tCOORD.Y = 1
                SetConsoleCursorPosition($hStdOut, $tCOORD)
                _WinAPI_WriteConsole($hStdOut, StringFormat("dwCommandId: %s", $oINPUT_RECORD.Event.MenuEvent.dwCommandId))
            Case $EventType_MOUSE_EVENT
                ;ConsoleWrite("MOUSE_EVENT"&@CRLF)
                _WinAPI_WriteConsole($hStdOut, "MOUSE_EVENT")
            Case $EventType_WINDOW_BUFFER_SIZE_EVENT
                _ConsoleHideCursor($hStdOut)

                $tCOORD.X = 0
                $tCOORD.Y = 0
                
                GetConsoleScreenBufferInfo($hStdOut, $tScreenBufferInfo)
                $columns = $tSrWindow.Right - $tSrWindow.Left + 1
                $rows = $tSrWindow.Bottom - $tSrWindow.Top + 1

                $tScreenBufferSize.X = $columns
                $tScreenBufferSize.Y = $rows
                SetConsoleScreenBufferSize($hStdOut, $tScreenBufferSize)

                FillConsoleOutputCharacter($hStdOut, " ", $columns * $rows, $tCOORD)

                SetConsoleCursorPosition($hStdOut, $tCOORD)
                _WinAPI_WriteConsole($hStdOut, "WINDOW_BUFFER_SIZE_EVENT")

                $tCOORD.Y = 1
                SetConsoleCursorPosition($hStdOut, $tCOORD)
                _WinAPI_WriteConsole($hStdOut, StringFormat("SIZE: %sx%s", $oINPUT_RECORD.Event.WindowBufferSizeEvent.dwSize.X, $oINPUT_RECORD.Event.WindowBufferSizeEvent.dwSize.Y))

                _ConsoleWriteCenter($columns, $rows, $hStdOut, StringFormat("%s x %s", $columns, $rows))
            Case Else
                _WinAPI_WriteConsole($hStdOut, "*UNKNOWN*")
        EndSwitch
    EndIf
WEnd
Exit
