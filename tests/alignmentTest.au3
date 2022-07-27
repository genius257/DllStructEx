#include "../au3pm/au3unit/Unit/assert.au3"
#include "../DllStructEx.au3"
#include "../DllStructEx.debug.au3"

Global Const $tagINPUT_RECORD = _
"WORD  EventType;"& _
"union {"& _
"    KEY_EVENT_RECORD          KeyEvent;"& _
"    MOUSE_EVENT_RECORD        MouseEvent;"& _
"    WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeEvent;"& _
"    MENU_EVENT_RECORD         MenuEvent;"& _
"    FOCUS_EVENT_RECORD        FocusEvent;"& _
"} Event;"

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

$oINPUT_RECORD = DllStructExCreate($tagINPUT_RECORD)
If @error <> 0 Then
    ConsoleWrite("@error: "&@error&@CRLF)
    Exit
EndIf

$pINPUT_RECORD = DllStructExGetPtr($oINPUT_RECORD)
$pEvent = DllStructExGetPtr($oINPUT_RECORD.Event)

$tINPUT_RECORD = DllStructCreate("WORD EventType;STRUCT;BOOL bKeyDown;WORD wRepeatCount;WORD  wVirtualKeyCode;WORD  wVirtualScanCode;STRUCT;WCHAR UnicodeChar;ENDSTRUCT;DWORD dwControlKeyState;ENDSTRUCT;", DllStructExGetPtr($oINPUT_RECORD))
assertEquals( _
    Number(DllStructGetPtr($tINPUT_RECORD, "UnicodeChar") - DllStructGetPtr($tINPUT_RECORD, "EventType")), _
    Number(DllStructExGetPtr($oINPUT_RECORD.Event.KeyEvent.uChar, "UnicodeChar") - $pINPUT_RECORD) _
)
$tINPUT_RECORD = DllStructCreate("WORD EventType;BOOL bSetFocus;", DllStructExGetPtr($oINPUT_RECORD))
assertEquals( _
    Number(DllStructGetPtr($tINPUT_RECORD, "bSetFocus") - DllStructGetPtr($tINPUT_RECORD, "EventType")), _
    Number(DllStructExGetPtr($oINPUT_RECORD.Event.FocusEvent, "bSetFocus") - $pINPUT_RECORD) _
)
assertEquals(Number($pEvent), Number(DllStructExGetPtr($oINPUT_RECORD.Event.KeyEvent)))
assertEquals(Number($pEvent), Number(DllStructExGetPtr($oINPUT_RECORD.Event.MouseEvent)))
assertEquals(Number($pEvent), Number(DllStructExGetPtr($oINPUT_RECORD.Event.WindowBufferSizeEvent)))
assertEquals(Number($pEvent), Number(DllStructExGetPtr($oINPUT_RECORD.Event.MenuEvent)))
assertEquals(Number($pEvent), Number(DllStructExGetPtr($oINPUT_RECORD.Event.FocusEvent)))
