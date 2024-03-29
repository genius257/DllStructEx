#include "DllStructEx.au3"

$tagINPUT_RECORD = _
"WORD  EventType;"& _
"union {"& _
"    KEY_EVENT_RECORD          KeyEvent;"& _
"    MOUSE_EVENT_RECORD        MouseEvent;"& _
"    WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeEvent;"& _
"    MENU_EVENT_RECORD         MenuEvent;"& _
"    FOCUS_EVENT_RECORD        FocusEvent;"& _
"} Event;"

$tagKEY_EVENT_RECORD = _
"BOOL  bKeyDown;"& _
"WORD  wRepeatCount;"& _
"WORD  wVirtualKeyCode;"& _
"WORD  wVirtualScanCode;"& _
"union {"& _
"  WCHAR UnicodeChar;"& _
"  CHAR  AsciiChar;"& _
"} uChar;"& _
"DWORD dwControlKeyState;"

$tagMOUSE_EVENT_RECORD = _
"COORD dwMousePosition;" & _
"DWORD dwButtonState;" & _
"DWORD dwControlKeyState;" & _
"DWORD dwEventFlags;"

$tagCOORD = _
"SHORT X;"& _
"SHORT Y;"

$tagWINDOW_BUFFER_SIZE_RECORD = _
"COORD dwSize;"

$tagMENU_EVENT_RECORD = _
"UINT dwCommandId;"

$tagFOCUS_EVENT_RECORD = _
"BOOL bSetFocus;"

$txINPUT_RECORD = DllStructExCreate($tagINPUT_RECORD)
If @error <> 0 Then
    ConsoleWrite(StringFormat("ERROR: %d : %d\n", @error, @extended))
    $txINPUT_RECORD = 0
    Exit 1
EndIf
$tINPUT_RECORD = DllStructExGetStruct($txINPUT_RECORD) ;Get the normal DllStruct from the DllStructEx
DllStructSetData($tINPUT_RECORD, "EventType", 123)
ConsoleWrite($txINPUT_RECORD.EventType&@CRLF)
$txINPUT_RECORD.EventType = 321
ConsoleWrite(DllStructGetData($tINPUT_RECORD, "EventType")&@CRLF)

$tKEY_EVENT_RECORD = DllStructExGetStruct($txINPUT_RECORD.Event.KeyEvent)
DllStructSetData($tKEY_EVENT_RECORD, "bKeyDown", 1)

ConsoleWrite($txINPUT_RECORD.Event.KeyEvent.bKeyDown&@CRLF)
$keyEvent = $txINPUT_RECORD.Event.KeyEvent
ConsoleWrite($KeyEvent(1)&@CRLF);getting value of index 1: bKeyDown
$keyEvent = 0
$txINPUT_RECORD = 0
