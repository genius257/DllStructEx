#include-once
#include "error.au3"
#include "InputStream.au3"

#cs
# ^(?&hn)*(?&struct_line_declaration)+(?&hn)*$
#ce
Func __DllStructEx_Parse_Root($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
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
    Local $result = __DllStructEx_Parse_align_directive($tInputStream, False)
    If Not (@error = 0) Then
        $result = __DllStructEx_Parse_struct_directive($tInputStream, False)
        If Not (@error = 0) Then
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
        EndIf
    EndIf

    If Not (InputStream_Peek($tInputStream) = ";") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct line expected ";", found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    InputStream_StepForward($tInputStream)

    Return $result
EndFunc

#cs
# (?<union>union(?&hn)*{(?&hn)*(?:(?&struct_line_declaration)(?&hn)*)+}(?:\h+(?&identifier))?(\h*\[\h*\d+\h*\])?(?&hn)*)
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

    ; Optional array case
    #Region
        Local $arraySize = Null
        ; Save current position, will be used to restore the position if the array is not found
        Local $iPos = InputStream_GetPos($tInputStream)

        ; Expect zero or more horizontal spaces
        While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
            InputStream_StepForward($tInputStream)
        WEnd

        If InputStream_Peek($tInputStream) = "[" Then
            InputStream_StepForward($tInputStream)

            ; Expect zero or more horizontal spaces
            While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
                InputStream_StepForward($tInputStream)
            WEnd

            ; Expect one or more digits
            If StringIsDigit(InputStream_Peek($tInputStream)) Then
                $arraySize = InputStream_Consume($tInputStream)

                While StringIsDigit(InputStream_Peek($tInputStream))
                    $arraySize &= InputStream_Consume($tInputStream)
                WEnd

                $arraySize = Int($arraySize)
                If @error <> 0 Then
                    If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to convert array size to integer at offset: %s', InputStream_GetPos($tInputStream)))
                    Return SetError(1, 0, Null)
                EndIf
                If $arraySize <= 0 Then
                    If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Array size must be greater than zero at offset: %s', InputStream_GetPos($tInputStream)))
                    Return SetError(1, 0, Null)
                EndIf

                ; Expect zero or more horizontal spaces
                While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
                    InputStream_StepForward($tInputStream)
                WEnd

                If InputStream_Peek($tInputStream) = "]" Then
                    InputStream_StepForward($tInputStream)
                Else
                    ; Restore position
                    InputStream_StepTo($tInputStream, $iPos)
                    ;TODO: We could add an error here
                EndIf
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
    $result['arraySize'] = $arraySize

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
# (?<declaration>(?&hn)*(?&type)\h+(?&identifier)(\h*\[\h*\d+\h*\])?(?&hn)*)
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

    ; Optional array case
    #Region
        Local $arraySize = Null
        ; Save current position, will be used to restore the position if the array is not found
        Local $iPos = InputStream_GetPos($tInputStream)

        ; Expect zero or more horizontal spaces
        While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
            InputStream_StepForward($tInputStream)
        WEnd

        If InputStream_Peek($tInputStream) = "[" Then
            InputStream_StepForward($tInputStream)

            ; Expect zero or more horizontal spaces
            While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
                InputStream_StepForward($tInputStream)
            WEnd

            ; Expect one or more digits
            If StringIsDigit(InputStream_Peek($tInputStream)) Then
                $arraySize = InputStream_Consume($tInputStream)

                While StringIsDigit(InputStream_Peek($tInputStream))
                    $arraySize &= InputStream_Consume($tInputStream)
                WEnd

                $arraySize = Int($arraySize)
                If @error <> 0 Then
                    If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to convert array size to integer at offset: %s', InputStream_GetPos($tInputStream)))
                    Return SetError(1, 0, Null)
                EndIf
                If $arraySize <= 0 Then
                    If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Array size must be greater than zero at offset: %s', InputStream_GetPos($tInputStream)))
                    Return SetError(1, 0, Null)
                EndIf

                ; Expect zero or more horizontal spaces
                While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
                    InputStream_StepForward($tInputStream)
                WEnd

                If InputStream_Peek($tInputStream) = "]" Then
                    InputStream_StepForward($tInputStream)
                Else
                    ; Restore position
                    InputStream_StepTo($tInputStream, $iPos)
                    ;TODO: We could add an error here
                EndIf
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
    $result['type'] = 'declaration'
    $result['dataType'] = $sDataType
    $result['identifier'] = $identifier
    $result['arraySize'] = $arraySize
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

#cs
# (?<struct_directive>(struct|endstruct)(?&hn)*(?=;))
#ce
Func __DllStructEx_Parse_struct_directive($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    Local $iPos = InputStream_GetPos($tInputStream)
    If InputStream_PeekMany($tInputStream, 6) = "struct" Then
        InputStream_StepForward($tInputStream, 6)
    ElseIf InputStream_PeekMany($tInputStream, 9) = "endstruct" Then
        InputStream_StepForward($tInputStream, 9)
    Else
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct directive, expected "struct" or "endstruct", found "%s" at offset: %s', InputStream_PeekMany($tInputStream, 6), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    Local $sContent = InputStream_GetSubstring($tInputStream, $iPos, InputStream_GetPos($tInputStream))

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    ; We check for a semicolon at the end of the struct, but we don't consume it.
    If Not (InputStream_Peek($tInputStream) = ";") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct directive, expected ";", found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    Local $result[]
    $result['type'] = 'directive'
    $result['dataType'] = $sContent
    $result['content'] = $sContent

    Return $result
EndFunc

#cs
# (?<align_directive>align\h+\d+(?&hn)*(?=;))
#ce
Func __DllStructEx_Parse_align_directive($tInputStream, $bErrorMessages = $__DllStructEx_bErrorMessages)
    Local $iPos = InputStream_GetPos($tInputStream)

    If Not (InputStream_PeekMany($tInputStream, 5) = "align") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse align directive expected "align" at offset: %s', InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    InputStream_StepForward($tInputStream, 5)

    Local $sContent = InputStream_GetSubstring($tInputStream, $iPos, InputStream_GetPos($tInputStream))

    ; Expect one or more horizontal spaces
    #Region
        If Not __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream)) Then
            If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse align directive, expected horizontal whitespace, found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
            Return SetError(1, 0, Null)
        EndIf

        InputStream_StepForward($tInputStream)

        While __DllStructEx_Parse_isHorizontalWhitespace(InputStream_Peek($tInputStream))
            InputStream_StepForward($tInputStream)
        WEnd
    #EndRegion

    ; Expect one or more digits
    If StringIsDigit(InputStream_Peek($tInputStream)) Then
        Local $boundary = InputStream_Consume($tInputStream)

        While StringIsDigit(InputStream_Peek($tInputStream))
            $boundary &= InputStream_Consume($tInputStream)
        WEnd

        $boundary = Int($boundary)
        If @error <> 0 Then
            If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to convert align boundary to integer at offset: %s', InputStream_GetPos($tInputStream)))
            Return SetError(1, 0, Null)
        EndIf
        #cs
        ;FIXME: we could check for valid boundary values
        If $boundary <= 0 Then
            If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Array size must be greater than zero at offset: %s', InputStream_GetPos($tInputStream)))
            Return SetError(1, 0, Null)
        EndIf
        #ce
    Else
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse align directive, expected integer at offset: %s', InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    ; expect zero or more horizontal-spaces/newlines
    Do
        __DllStructEx_Parse_hn($tInputStream, False)
    Until Not (@error = 0)

    ; We check for a semicolon at the end of the struct, but we don't consume it.
    If Not (InputStream_Peek($tInputStream) = ";") Then
        If $bErrorMessages Then __DllStructEx_ConsoleWriteErrorLine(StringFormat('ERROR: Failed to parse struct directive, expected ";", found "%s" at offset: %s', InputStream_Peek($tInputStream), InputStream_GetPos($tInputStream)))
        Return SetError(1, 0, Null)
    EndIf

    Local $result[]
    $result['type'] = 'directive'
    $result['dataType'] = $sContent
    $result['content'] = $sContent & " " & $boundary
    $result['boundary'] = $boundary

    Return $result
EndFunc
