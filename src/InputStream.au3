#include-once

Func InputStream($sInput)
    Local $t = DllStructCreate("WCHAR input["&StringLen($sInput)&"]; LONG length; LONG pos; BOOL eof;")
    DllStructSetData($t, "input", $sInput)
    $t.length = StringLen($sInput)
    $t.pos = 1

    Return $t
EndFunc

Func InputStream_PeekMany($t, $iCount)
    If $t.pos + $iCount > $t.length Then $iCount = $t.length - $t.pos + 1
    ; Position times 2, because we are dealing with a WCHAR array
    Local $_t = DllStructCreate("WCHAR["&$iCount&"]", DllStructGetPtr($t, "input") + ($t.pos - 1) * 2)
    Return DllStructGetData($_t, 1)
EndFunc

Func InputStream_Peek($t)
    If $t.eof Then Return Null
    Return DllStructGetData($t, "input", $t.pos)
EndFunc

Func InputStream_Consume($t)
    If $t.eof Then Return Null
    Local $c = DllStructGetData($t, "input", $t.pos)
    InputStream_StepForward($t)
    InputStream_CheckEOF($t)
    Return $c
EndFunc

Func InputStream_StepBack($t, $iSteps = 1)
    InputStream_StepForward($t, -$iSteps)
EndFunc

Func InputStream_StepForward($t, $iSteps = 1)
    $t.pos += $iSteps
    InputStream_CheckEOF($t)
EndFunc

Func InputStream_StepTo($t, $iPos)
    $t.pos = $iPos
    InputStream_CheckEOF($t)
EndFunc

Func InputStream_GetPos($t)
    Return $t.pos
EndFunc

Func InputStream_CheckEOF($t)
    $t.eof = $t.pos > $t.length
EndFunc

Func InputStream_GetEOF($t)
    Return $t.eof
EndFunc

Func InputStream_GetSubstring($t, $iStartPos, $iEndPos)
    If $iEndPos > $t.length Then $iEndPos = $t.length
    If $iStartPos > $t.length Then $iStartPos = $t.length
    If $iStartPos > $iEndPos Then
        Local $tmp = $iStartPos
        $iStartPos = $iEndPos
        $iEndPos = $tmp
    EndIf
    If $iStartPos = $iEndPos Then
        Return ""
    EndIf

    Local $_t = DllStructCreate("WCHAR input["&($iEndPos - $iStartPos)&"];", DllStructGetPtr($t, "input") + ($iStartPos - 1) * 2)

    Return DllStructGetData($_t, 1)
EndFunc
