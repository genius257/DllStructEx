#include-once
#include "error.au3"
#include "InputStream.au3"

#cs
# ^(?&hn)*(?&struct_line_declaration)+(?&hn)*$
#ce
Func __DllStructEx_ParseStruct($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    #Region
        Local $aStructLineDeclarations[0]
        
        While 1
            Local $iPos = InputStream_GetPos($tInputStream)
            $mStructLineDeclaration = __DllStructEx_Parse_struct_line_declaration($tInputStream, UBound($aStructLineDeclarations) = 0)
            If Not (@error = 0) Then
                InputStream_StepTo($tInputStream, $iPos)
                ExitLoop
            EndIf
            Redim $aStructLineDeclarations[UBound($aStructLineDeclarations) + 1]
            $aStructLineDeclarations[UBound($aStructLineDeclarations) - 1] = $mStructLineDeclaration

            ; expect zero or more horizontal-spaces/newlines
            Do
                __DllStructEx_Parse_hn($tInputStream, False)
            Until Not (@error = 0)
        WEnd

        If UBound($aStructLineDeclarations) = 0 Then
            If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct. Expected one or more struct line declarations, but found none at offset: %s', InputStream_GetPos($tInputStream)))
            Return SetError(1, 0, Null)
        EndIf
    #EndRegion

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    If Not InputStream_GetEOF($tInputStream) Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct. Expected end-of-file, found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    Return $aStructLineDeclarations
EndFunc

#cs
# (?<hn>[\h\n])
#ce
Func __DllStructEx_Parse_hn($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    Local $c = InputStream_Peek($tInputStream)
    If Not (__DllStructEx_Parse_isHorizontalWhitespace($c) Or AscW($c) = 0x0A) Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse horizontal-whitespace or new-line, found "%s" at offset: %s', $c, InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    InputStream_StepForward($tInputStream)
    Return $c
EndFunc

#cs
# (?<struct_line_declaration>(?:(?&declaration)|(?&struct)|(?&union));)
#ce
Func __DllStructEx_Parse_struct_line_declaration($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    ; Save current position, will be used to restore the position between the function calls
    Local $iPos = InputStream_GetPos($tInputStream)
    Local $result = __DllStructEx_Parse_declaration($tInputStream, False)
    If Not (@error = 0) Then
        InputStream_StepTo($tInputStream, $iPos)
        $result = __DllStructEx_Parse_struct($tInputStream, False)
        If Not (@error = 0) Then
            InputStream_StepTo($tInputStream, $iPos)
            $result = __DllStructEx_Parse_union($tInputStream, False)
            If Not (@error = 0) Then
                If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct line, expected declaration, struct or union, found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
                Return SetError(1, 0, Null)
            EndIf
        EndIf
    EndIf

    If Not InputStream_Peek($tInputStream) = ";" Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct line expected ";", found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    InputStream_StepForward($tInputStream)

    Return $result
EndFunc

#cs
# (?<union>union(?&hn)*{(?&hn)*(?:(?&struct_line_declaration)(?&hn)*)+}(?:\h+(?&identifier))?(?&hn)*)
#ce
Func __DllStructEx_Parse_union($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    If Not (InputStream_PeekMany($tInputStream, 5) = "union") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse union expected "union", found "%s" at offset: %s', InputStream_PeekMany($tInputStream, 5), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    InputStream_StepForward($tInputStream, 5)

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    If Not (InputStream_Peek($tInputStream) = "{") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse union expected "{", found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf
    InputStream_StepForward($tInputStream)

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    Local $iContentStartPosition = InputStream_GetPos($tInputStream)
    Local $aStructLineDeclarations[0]

    While 1
        Local $iPos = InputStream_GetPos($tInputStream)
        $mStructLineDeclaration = __DllStructEx_Parse_struct_line_declaration($tInputStream, UBound($aStructLineDeclarations) = 0)
        If Not (@error = 0) Then
            InputStream_StepTo($tInputStream, $iPos)
            ExitLoop
        EndIf
        Redim $aStructLineDeclarations[UBound($aStructLineDeclarations) + 1]
        $aStructLineDeclarations[UBound($aStructLineDeclarations) - 1] = $mStructLineDeclaration

        ; expect zero or more horizontal-spaces/newlines
        Do
            __DllStructEx_Parse_hn($tInputStream, False)
        Until Not (@error = 0)
    WEnd

    If UBound($aStructLineDeclarations) = 0 Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse union expected one or more struct line declarations, but found none at offset: %s', InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    Local $content = InputStream_GetSubstring($tInputStream, $iContentStartPosition, InputStream_GetPos($tInputStream))

    If Not (InputStream_Peek($tInputStream) = "}") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse union expected "}", found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf
    InputStream_StepForward($tInputStream)

    ; Optional identifier case
    #Region
        Local $identifier = Null
        ; Save current position, will be used to restore the position if the identifier is not found
        Local $iPos = InputStream_GetPos($tInputStream)

        ; Expect one or more horizontal spaces
        If __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream)) Then
            InputStream_StepForward($tInputStream)

            While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
                InputStream_StepForward($tInputStream)
            WEnd

            #cs
            While __DllStructEx_Parse_isWordCharacter(InputStream_Peek($tInputStream))
                InputStream_StepForward($tInputStream)
            WEnd
            #ce

            Local $_identifier = __DllStructEx_Parse_identifier($tInputStream, False)
            If @error = 0 Then
                $identifier = $_identifier
            Else
                ; Restore position
                InputStream_StepTo($tInputStream, $iPos)
            EndIf
        Else
            ; Restore position
            InputStream_StepTo($tInputStream, $iPos)
        EndIf
    #EndRegion

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    Local $result[]
    $result['type'] = 'union'
    $result['content'] = $content
    $result['identifier'] = $identifier
    $result['structLineDeclarations'] = $aStructLineDeclarations

    Return $result
EndFunc

#cs
# (?<struct>struct(?&hn)*{(?&hn)*(?:(?&struct_line_declaration)(?&hn)*)+}(?:\h+(?&identifier))?(?&hn)*)
#ce
Func __DllStructEx_Parse_struct($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    If Not (InputStream_PeekMany($tInputStream, 6) = "struct") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct expected "struct", found "%s" at offset: %s', InputStream_PeekMany($tInputStream, 6), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    InputStream_StepForward($tInputStream, 6)

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    If Not (InputStream_Peek($tInputStream) = "{") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct expected "{", found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf
    InputStream_StepForward($tInputStream)

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    Local $iContentStartPosition = InputStream_GetPos($tInputStream)
    Local $aStructLineDeclarations[0]
    While 1
        Local $iPos = InputStream_GetPos($tInputStream)
        $mStructLineDeclaration = __DllStructEx_Parse_struct_line_declaration($tInputStream, UBound($aStructLineDeclarations) = 0)
        If Not (@error = 0) Then
            InputStream_StepTo($tInputStream, $iPos)
            ExitLoop
        EndIf
        Redim $aStructLineDeclarations[UBound($aStructLineDeclarations) + 1]
        $aStructLineDeclarations[UBound($aStructLineDeclarations) - 1] = $mStructLineDeclaration

        ; expect zero or more horizontal-spaces/newlines
        Do
            __DllStructEx_Parse_hn($tInputStream, False)
        Until Not (@error = 0)
    WEnd

    If UBound($aStructLineDeclarations) = 0 Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct expected one or more struct line declarations, but found none at offset: %s', InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    Local $content = InputStream_GetSubstring($tInputStream, $iContentStartPosition, InputStream_GetPos($tInputStream))

    If Not (InputStream_Peek($tInputStream) = "}") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct expected "}", found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf
    InputStream_StepForward($tInputStream)

    ; Optional identifier case
    #Region
        Local $identifier = Null
        ; Save current position, will be used to restore the position if the identifier is not found
        Local $iPos = InputStream_GetPos($tInputStream)

        ; Expect one or more horizontal spaces
        If __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream)) Then
            InputStream_StepForward($tInputStream)

            While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
                InputStream_StepForward($tInputStream)
            WEnd

            #cs
            While __DllStructEx_Parse_isWordCharacter(InputStream_Peek($tInputStream))
                InputStream_StepForward($tInputStream)
            WEnd
            #ce

            Local $_identifier = __DllStructEx_Parse_identifier($tInputStream, False)
            If @error = 0 Then
                $identifier = $_identifier
            Else
                ; Restore position
                InputStream_StepTo($tInputStream, $iPos)
            EndIf
        Else
            ; Restore position
            InputStream_StepTo($tInputStream, $iPos)
        EndIf
    #EndRegion

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    Local $result[]
    $result['type'] = 'struct'
    $result['content'] = $content
    $result['identifier'] = $identifier
    $result['structLineDeclarations'] = $aStructLineDeclarations

    Return $result
EndFunc

#cs
# (?<declaration>(?&hn)*(?&type)\h+(?&identifier)(?&hn)*)
#ce
Func __DllStructEx_Parse_declaration($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    Local $sDataType = __DllStructEx_Parse_type($tInputStream, $bErrorMessages)
    If Not (@error = 0) Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse declaration at offset: %s', InputStream_GetPos($tInputStream)))
    EndIf

    ; Expect one or more horizontal spaces
    #Region
        If Not __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream)) Then
            If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse declaration expected horizontal whitespace, found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
            Return SetError(1, 0, Null)
        EndIf

        InputStream_StepForward($tInputStream)

        While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
            InputStream_StepForward($tInputStream)
        WEnd
    #EndRegion

    $identifier = __DllStructEx_Parse_identifier($tInputStream, $bErrorMessages)
    If Not (@error = 0) Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse declaration at offset: %s', InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    Local $result[]
    $result['type'] = 'declaration'
    $result['dataType'] = $sDataType
    $result['identifier'] = $identifier
    Return $result
EndFunc

#cs
# Checks if character is a word character.
# @see https://www.pcre.org/current/doc/html/pcre2pattern.html source
#ce
Func __DllStructEx_Parse_isWordCharacter($c)
    Return StringIsAlNum($c) Or $c == "_"
EndFunc

#cs
# Checks if character is a whitespace.
#ce
Func __DllStructEx_Parse_isWhitespace($c)
    Return __DllStructEx_Parse_isHorizontalWhitespace($c) Or __DllStructEx_Parse_isVerticalWhitespace($c)
EndFunc

#cs
# Checks if character is a horizontal whitespace.
# @see https://www.pcre.org/current/doc/html/pcre2pattern.html source
#ce
Func __DllStructEx_Parse_isHorizontalWhitespace($c)
    Switch AscW($c)
        Case 0x0009, _ ; Horizontal tab (HT)
             0x0020, _ ; Space
             0x00A0, _ ; Non-break space
             0x1680, _ ; Ogham space mark
             0x180E, _ ; Mongolian vowel separator
             0x2000, _ ; En quad
             0x2001, _ ; Em quad
             0x2002, _ ; En space
             0x2003, _ ; Em space
             0x2004, _ ; Three-per-em space
             0x2005, _ ; Four-per-em space
             0x2006, _ ; Six-per-em space
             0x2007, _ ; Figure space
             0x2008, _ ; Punctuation space
             0x2009, _ ; Thin space
             0x200A, _ ; Hair space
             0x202F, _ ; Narrow no-break space
             0x205F, _ ; Medium mathematical space
             0x3000    ; Ideographic space
            Return True
    EndSwitch

    Return False
EndFunc

#cs
# Checks if character is a vertical whitespace.
# @see https://www.pcre.org/current/doc/html/pcre2pattern.html source
#ce
Func __DllStructEx_Parse_isVerticalWhitespace($c)
    Switch AscW($c)
        Case 0x000A, _ ; Linefeed (LF)
             0x000B, _ ; Vertical tab (VT)
             0x000C, _ ; Form feed (FF)
             0x000D, _ ; Carriage return (CR)
             0x0085, _ ; Next line (NEL)
             0x2028, _ ; Line separator
             0x2029    ; Paragraph separator
        Return True
    EndSwitch

    Return False
EndFunc

#cs
# (?<type>\w+)
#ce
Func __DllStructEx_Parse_type($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    $sDataType = ""

    While __DllStructEx_Parse_isWordCharacter(InputStream_Peek($tInputStream))
        $sDataType &= InputStream_Consume($tInputStream)
    WEnd

    If $sDataType = "" Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse type expected word character, found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    Return $sDataType
EndFunc

#cs
# (?<identifier>[*]*\w+)
#ce
Func __DllStructEx_Parse_identifier($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    $iIndirectionLevel = 0
    While InputStream_Peek($tInputStream) == "*"
        InputStream_StepForward($tInputStream)
        $iIndirectionLevel += 1
    WEnd

    $sName = InputStream_Peek($tInputStream)

    If Not __DllStructEx_Parse_isWordCharacter($sName) Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse identifier. Expected word character, got "%s" at offset: %s', $sName, InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    InputStream_StepForward($tInputStream)

    While __DllStructEx_Parse_isWordCharacter(InputStream_Peek($tInputStream))
        $sName &= InputStream_Consume($tInputStream)
    WEnd

    Local $result[]
    $result['indirectionLevel'] = $iIndirectionLevel
    $result['name'] = $sName

    Return $result
EndFunc

;----------------------------------------------------------------------------------------------------------------------

#include "C:/Users/Frank/Documents/GitHub/au3json/json.au3"

$tagVARIANT = _
"  union {"& _
"    struct {"& _
"      VARTYPE vt;"& _
"      WORD    wReserved1;"& _
"      WORD    wReserved2;"& _
"      WORD    wReserved3;"& _
"      union {"& _
"        LONGLONG     llVal;"& _
"        LONG         lVal;"& _
"        BYTE         bVal;"& _
"        SHORT        iVal;"& _
"        FLOAT        fltVal;"& _
"        DOUBLE       dblVal;"& _
"        VARIANT_BOOL boolVal;"& _
"        VARIANT_BOOL __OBSOLETE__VARIANT_BOOL;"& _
"        SCODE        scode;"& _
"        CY           cyVal;"& _
"        DATE         date;"& _
"        BSTR         bstrVal;"& _
"        IUnknown     *punkVal;"& _
"        IDispatch    *pdispVal;"& _
"        SAFEARRAY    *parray;"& _
"        BYTE         *pbVal;"& _
"        SHORT        *piVal;"& _
"        LONG         *plVal;"& _
"        LONGLONG     *pllVal;"& _
"        FLOAT        *pfltVal;"& _
"        DOUBLE       *pdblVal;"& _
"        VARIANT_BOOL *pboolVal;"& _
"        VARIANT_BOOL *__OBSOLETE__VARIANT_PBOOL;"& _
"        SCODE        *pscode;"& _
"        CY           *pcyVal;"& _
"        DATE         *pdate;"& _
"        BSTR         *pbstrVal;"& _
"        IUnknown     **ppunkVal;"& _
"        IDispatch    **ppdispVal;"& _
"        SAFEARRAY    **pparray;"& _
"        VARIANT      *pvarVal;"& _
"        PVOID        byref;"& _
"        CHAR         cVal;"& _
"        USHORT       uiVal;"& _
"        ULONG        ulVal;"& _
"        ULONGLONG    ullVal;"& _
"        INT          intVal;"& _
"        UINT         uintVal;"& _
"        DECIMAL      *pdecVal;"& _
"        CHAR         *pcVal;"& _
"        USHORT       *puiVal;"& _
"        ULONG        *pulVal;"& _
"        ULONGLONG    *pullVal;"& _
"        INT          *pintVal;"& _
"        UINT         *puintVal;"& _
"        struct {"& _
"          PVOID       pvRecord;"& _
"          IRecordInfo *pRecInfo;"& _
"        } __VARIANT_NAME_4;"& _
"      } __VARIANT_NAME_3;"& _
"    } __VARIANT_NAME_2;"& _
"    DECIMAL decVal;"& _
"  } __VARIANT_NAME_1;"

$tInputStream = InputStream($tagVARIANT)
$time = TimerInit()
Global $result = __DllStructEx_ParseStruct($tInputStream)
ConsoleWrite(TimerDiff($time)&@CRLF)
ConsoleWrite("@error: "&@error&@CRLF)
ConsoleWrite(_json_encode_pretty($result[0])&@CRLF)
