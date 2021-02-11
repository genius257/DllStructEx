#include <Memory.au3>
#include <WinAPIMem.au3>

If Not IsDeclared("IID_IUnknown") Then Global Const $IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
If Not IsDeclared("IID_IDispatch") Then Global Const $IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"

If Not IsDeclared("S_OK") Then Global Const $S_OK = 0x00000000
If Not IsDeclared("E_NOTIMPL") Then Global Const $E_NOTIMPL = 0x80004001
If Not IsDeclared("E_NOINTERFACE") Then Global Const $E_NOINTERFACE = 0x80004002
If Not IsDeclared("E_POINTER") Then Global Const $E_POINTER = 0x80004003
If Not IsDeclared("E_ABORT") Then Global Const $E_ABORT = 0x80004004
If Not IsDeclared("E_FAIL") Then Global Const $E_FAIL = 0x80004005
If Not IsDeclared("E_ACCESSDENIED") Then Global Const $E_ACCESSDENIED = 0x80070005
If Not IsDeclared("E_HANDLE") Then Global Const $E_HANDLE = 0x80070006
If Not IsDeclared("E_OUTOFMEMORY") Then Global Const $E_OUTOFMEMORY = 0x8007000E
If Not IsDeclared("E_INVALIDARG") Then Global Const $E_INVALIDARG = 0x80070057
If Not IsDeclared("E_UNEXPECTED") Then Global Const $E_UNEXPECTED = 0x8000FFFF

If Not IsDeclared("DISP_E_UNKNOWNINTERFACE") Then Global Const $DISP_E_UNKNOWNINTERFACE = 0x80020001
If Not IsDeclared("DISP_E_MEMBERNOTFOUND") Then Global Const $DISP_E_MEMBERNOTFOUND = 0x80020003
If Not IsDeclared("DISP_E_PARAMNOTFOUND") Then Global Const $DISP_E_PARAMNOTFOUND = 0x80020004
If Not IsDeclared("DISP_E_TYPEMISMATCH") Then Global Const $DISP_E_TYPEMISMATCH = 0x80020005
If Not IsDeclared("DISP_E_UNKNOWNNAME") Then Global Const $DISP_E_UNKNOWNNAME = 0x80020006
If Not IsDeclared("DISP_E_NONAMEDARGS") Then Global Const $DISP_E_NONAMEDARGS = 0x80020007
If Not IsDeclared("DISP_E_BADVARTYPE") Then Global Const $DISP_E_BADVARTYPE = 0x80020008
If Not IsDeclared("DISP_E_EXCEPTION") Then Global Const $DISP_E_EXCEPTION = 0x80020009
If Not IsDeclared("DISP_E_OVERFLOW") Then Global Const $DISP_E_OVERFLOW = 0x8002000A
If Not IsDeclared("DISP_E_BADINDEX") Then Global Const $DISP_E_BADINDEX = 0x8002000B
If Not IsDeclared("DISP_E_UNKNOWNLCID") Then Global Const $DISP_E_UNKNOWNLCID = 0x8002000C
If Not IsDeclared("DISP_E_ARRAYISLOCKED") Then Global Const $DISP_E_ARRAYISLOCKED = 0x8002000D
If Not IsDeclared("DISP_E_BADPARAMCOUNT") Then Global Const $DISP_E_BADPARAMCOUNT = 0x8002000E
If Not IsDeclared("DISP_E_PARAMNOTOPTIONAL") Then Global Const $DISP_E_PARAMNOTOPTIONAL = 0x8002000F
If Not IsDeclared("DISP_E_BADCALLEE") Then Global Const $DISP_E_BADCALLEE = 0x80020010
If Not IsDeclared("DISP_E_NOTACOLLECTION") Then Global Const $DISP_E_NOTACOLLECTION = 0x80020011

Global Const $tagVARIANT = "ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2";FIXME: rename varaible

Global Enum $VT_EMPTY,$VT_NULL,$VT_I2,$VT_I4,$VT_R4,$VT_R8,$VT_CY,$VT_DATE,$VT_BSTR,$VT_DISPATCH, _
    $VT_ERROR,$VT_BOOL,$VT_VARIANT,$VT_UNKNOWN,$VT_DECIMAL,$VT_I1=16,$VT_UI1,$VT_UI2,$VT_UI4,$VT_I8, _
    $VT_UI8,$VT_INT,$VT_UINT,$VT_VOID,$VT_HRESULT,$VT_PTR,$VT_SAFEARRAY,$VT_CARRAY,$VT_USERDEFINED, _
    $VT_LPSTR,$VT_LPWSTR,$VT_RECORD=36,$VT_FILETIME=64,$VT_BLOB,$VT_STREAM,$VT_STORAGE,$VT_STREAMED_OBJECT, _
    $VT_STORED_OBJECT,$VT_BLOB_OBJECT,$VT_CF,$VT_CLSID,$VT_BSTR_BLOB=0xfff,$VT_VECTOR=0x1000, _
    $VT_ARRAY=0x2000,$VT_BYREF=0x4000,$VT_RESERVED=0x8000,$VT_ILLEGAL=0xffff,$VT_ILLEGALMASKED=0xfff, _
    $VT_TYPEMASK=0xfff

Global Const $g__DllStructEx_tagObject = "int RefCount;int Size;ptr Object;ptr Methods[7];ptr szStruct;ptr szTranslatedStruct;ptr pStruct;int cElements;ptr pElements;"
Global Const $g__DllStructEx_tagElement = "int iType;ptr szName;ptr szStruct;ptr szTranslatedStruct;int cElements;ptr pElements;"
;The size of $g__DllStructEx_tagElement in bytes 
Global Const $g__DllStructEx_iElement = DllStructGetSize(DllStructCreate($g__DllStructEx_tagElement))

Global Enum $g__DllStructEx_eElementType_NONE, $g__DllStructEx_eElementType_UNION, $g__DllStructEx_eElementType_Struct, $g__DllStructEx_eElementType_Element

; FIXME: support $pData parameter
Func DllStructExCreate($sStruct, $pData = 0)
    Local $aResult = __DllStructEx_ParseStruct($sStruct)
    If @error <> 0 Then Return SetError(1, @error, "")
    Local $sTranslatedStruct = $aResult[0]
    ConsoleWrite($sTranslatedStruct&@CRLF);FIXME: remove this

    Local $tObject = DllStructCreate($g__DllStructEx_tagObject)

    $tObject.cElements = $aResult[1].Index
    Local $sElements = StringFormat("BYTE[%d]", $g__DllStructEx_iElement * $tObject.cElements)
    $tObject.pElements = DllStructGetPtr(__DllStructEx_DllStructAlloc($sElements, DllStructCreate($sElements, DllStructGetPtr($aResult[1], "Elements"))))

    Local $pString = _WinAPI_CreateString($sStruct, 0, -1, True, False)
    DllStructSetData($tObject, "szStruct", $pString)
    $pString = _WinAPI_CreateString($sTranslatedStruct, 0, -1, True, False)
    DllStructSetData($tObject, "szTranslatedStruct", $pString)

    If $pData = 0 Then
        Local $tStruct = DllStructCreate($sTranslatedStruct)
        Local $iStruct = DllStructGetSize($tStruct)
        Local $hStruct = _MemGlobalAlloc($iStruct, $GMEM_MOVEABLE)
        Local $pStruct = _MemGlobalLock($hStruct)
        _MemMoveMemory(DllStructGetPtr($tStruct), $pStruct, $iStruct)
        $pData = $pStruct
    EndIf
    DllStructCreate($sTranslatedStruct, $pData)
    If @error <> 0 Then Return SetError(2, 0, 0)
    DllStructSetData($tObject, "pStruct", $pData)

    Local Static $QueryInterface = DllCallbackRegister(__DllStructEx_QueryInterface, "LONG", "ptr;ptr;ptr")
    DllStructSetData($tObject, "Methods", DllCallbackGetPtr($QueryInterface), 1)

    Local Static $AddRef = DllCallbackRegister(__DllStructEx_AddRef, "dword", "PTR")
    DllStructSetData($tObject, "Methods", DllCallbackGetPtr($AddRef), 2)

    Local Static $Release = DllCallbackRegister(__DllStructEx_Release, "dword", "PTR")
    DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Release), 3)

    Local Static $GetTypeInfoCount = DllCallbackRegister(__DllStructEx_GetTypeInfoCount, "long", "ptr;ptr")
    DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfoCount), 4)

    Local Static $GetTypeInfo = DllCallbackRegister(__DllStructEx_GetTypeInfo, "long", "ptr;uint;int;ptr")
    DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfo), 5)

    Local Static $GetIDsOfNames = DllCallbackRegister(__DllStructEx_GetIDsOfNames, "long", "ptr;ptr;ptr;uint;int;ptr")
    DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetIDsOfNames), 6)

    Local Static $Invoke = DllCallbackRegister(__DllStructEx_Invoke, "long", "ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr")
    DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Invoke), 7)

    DllStructSetData($tObject, "RefCount", 2) ; initial ref count is 1
    DllStructSetData($tObject, "Size", 7) ; number of interface methods

    $tObject = __DllStructEx_DllStructAlloc($g__DllStructEx_tagObject, $tObject)

    DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
    Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IDispatch, Default, True) ; pointer that's wrapped into object
EndFunc

Func __DllStructEx_QueryInterface($pSelf, $pRIID, $pObj)
    If $pObj=0 Then Return $E_POINTER
    Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
    If (Not ($sGUID=$IID_IDispatch)) And (Not ($sGUID=$IID_IUnknown)) Then Return $E_NOINTERFACE
    Local $tStruct = DllStructCreate("ptr", $pObj)
    DllStructSetData($tStruct, 1, $pSelf)
    Return $S_OK
EndFunc

Func __DllStructEx_AddRef($pSelf)
    Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
    $tStruct.Ref += 1
    Return $tStruct.Ref
EndFunc

Func __DllStructEx_Release($pSelf)
    Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
    $tStruct.Ref -= 1
    If $tStruct.Ref = 0 Then; initiate garbage collection

        Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf-8)
        ;releases all properties
        _WinAPI_FreeMemory(DllStructGetData($tObject, "szStruct"))
        _WinAPI_FreeMemory(DllStructGetData($tObject, "szTranslatedStruct"))

        Local $aRet = DllCall("Kernel32.dll", "ptr", "GlobalHandle", "ptr", $pSelf-8)
        If @error <> 0 Or $aRet[0] = 0 Then Return $tStruct.Ref ;An error occured
        _MemGlobalFree($aRet[1])
        Return 0
    EndIf
    Return $tStruct.Ref
EndFunc

Func __DllStructEx_GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
    Local $tDispId = DllStructCreate("long DispId;", $rgDispId)
    Local $pName = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)
    Local $sName = _WinAPI_GetString($pName, True)
    ConsoleWrite($sName&@CRLF)

    Local Static $CSTR_EQUAL = 2
    Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf - 8)
    Local $i
    For $i = 0 To $tObject.cElements - 1 Step +1
        Local $tElement = DllStructCreate($g__DllStructEx_tagElement, $tObject.pElements + $g__DllStructEx_iElement * $i)
        Local $aRet = DllCall("kernel32.dll", "int", "CompareStringOrdinal", "ptr", $pName, "int", -1, "ptr", $tElement.szName, "int", -1, "BOOL", 1)
        If $aRet[0] = $CSTR_EQUAL Then
            DllStructSetData($tDispId, "DispId", $i + 1)
            Return $S_OK
        EndIf
    Next

    ;CompareStringOrdinal
    ;_WinAPI_CompareString
    DllStructSetData($tDispId, "DispId", -1)
    Return $S_OK
EndFunc

Func __DllStructEx_GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
    If $iTInfo<>0 Then Return $DISP_E_BADINDEX
    If $ppTInfo=0 Then Return $E_INVALIDARG
    Return $S_OK
EndFunc

Func __DllStructEx_GetTypeInfoCount($pSelf, $pctinfo)
    DllStructSetData(DllStructCreate("UINT",$pctinfo), 1, 0)
    Return $S_OK
EndFunc

Func __DllStructEx_Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
    ;ConsoleWrite($dispIdMember&@CRLF)
    If $dispIdMember = -1 Then Return $DISP_E_MEMBERNOTFOUND

    #cs
    ;For debugging: getting vt of autoit types
    ;FIXME: remove this
    $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
    If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
    Local $tVARIANT=DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
    ConsoleWrite($tVARIANT.vt&@CRLF)
    Return $S_OK
    #ce

    If $dispIdMember = 0 Then
        ;FIXME: look for first parameter, to see if index is requested
        ;VariantInit($pVariant)
        Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf-8)
        Local $tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
        $tVARIANT.vt = $VT_BSTR
        $tVARIANT.data = DllStructGetData($tObject, "szStruct")
        Return $S_OK
    EndIf

    If $dispIdMember < -1 Then
        ;
    EndIf

    If $dispIdMember > 0 Then
        Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf - 8)
        If $tObject.cElements >= $dispIdMember Then
            Local $tElement = DllStructCreate($g__DllStructEx_tagElement, $tObject.pElements + $g__DllStructEx_iElement * ($dispIdMember - 1))
            Local $tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
            Local $tStruct = DllStructCreate(_WinAPI_GetString($tObject.szTranslatedStruct, True), $tObject.pStruct)
            Local $vData = DllStructGetData($tStruct, _WinAPI_GetString($tElement.szName, True));FIXME: add support for getting struct data slice by index
            __DllStructEx_DataToVariant($vData, $tVARIANT)
            Return $S_OK
        EndIf
    EndIf

    Return $DISP_E_MEMBERNOTFOUND
EndFunc

Func __DllStructEx_Create()
    ;FIXME: used from internal to create and return nested struct object ref, without parsing strings again.
    ;FIXME: move defning and setting object methods into here. The other function uses this.
EndFunc

Func __DllStructEx_DataToVariant($vData, $tVARIANT = Null)
    If Not IsDllStruct($tVARIANT) Then $tVARIANT = DllStructCreate($tagVARIANT)
    #cs
        $VT_I4 = 3 => int32
        $VT_I8 = 20 => int64
        $VT_R8 = 5 => Double
        $VT_BSTR = 8 => String
        $VT_ARRAY + $VT_BOOL = 8209 (0x2011) => Binary
        $VT_UI4 = 19 => Pointer
    #ce

    Switch VarGetType($vData)
        Case 'Int32'
            $tVARIANT.vt = $VT_I4
            Local $tData = DllStructCreate("INT data;", DllStructGetPtr($tVARIANT, 'data'))
            $tVARIANT.data = $vData
        Case 'Int64'
            $tVARIANT.vt = $VT_I8
            Local $tData = DllStructCreate("INT64 data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        Case 'Double'
            $tVARIANT.vt = $VT_R8
            Local $tData = DllStructCreate("DOUBLE data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        Case 'String'
            $tVARIANT.vt = $VT_BSTR
            Local $tData = DllStructCreate("PTR data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        #cs
        Case 'Binary'
            $tVARIANT.vt = BitOR($VT_ARRAY, $VT_BOOL)
            ;FIXME: support Binary
        #ce
        Case 'Pointer'
            $tVARIANT.vt = @AutoItX64 ? $VT_UI8 : $VT_UI4
            Local $tData = DllStructCreate("PTR data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        ;Case 'Object'
            ;FIXME: return IDispatch
        Case Else
            If Not @Compiled Then ConsoleWriteError(StringFormat('ERROR: Unsupported Datatype "%s" for VARIANT\n', VarGetType($vData)))
            Return SetError(1, 0, $tVARIANT)
    EndSwitch

    Return $tVARIANT
EndFunc

Func __DllStructEx_CalcStructSize()
    ;FIXME
EndFunc

Func __DllStructEx_ParseStruct($sStruct)
    Local $aUnions = StringRegExp($sStruct, "(?i)(^|;)\s*union\s*{[^}]+}\s*(?:\w+)?", 3)
    If @error = 2 Then Return SetError(3, @error, "")
    Local $tUnions = DllStructCreate(StringFormat("INT Index;INT Size;BYTE Elements[%d];", $g__DllStructEx_iElement * UBound($aUnions, 1)))
    $sStruct = StringRegExpReplaceCallbackEx($sStruct, "(?i)(^|;)\s*union\s*{([^}]+)}\s*(\w+)?", __DllStructEx_ParseUnion, 0, $tUnions)
    If @error <> 0 Then Return SetError(1, @error, "")
    Local $aElements = StringRegExp($sStruct, "(^|;)\s*\w+(?:\s+\w+)?", 3)
    If @error <> 0 Then Return SetError(4, @error, "")
    Local $tElements = DllStructCreate(StringFormat("INT Index;INT Size;BYTE Elements[%d];", $g__DllStructEx_iElement * UBound($aElements, 1)))
    $sStruct = StringRegExpReplaceCallbackEx($sStruct, "(^|;)\s*(\w+)(?:\s+(\w+))?", __DllStructEx_ParseStructTypeCallback, 0, $tElements)
    If @error <> 0 Then Return SetError(2, @error, "")

    ;FIXME: insert tUnions into  $tElements

    ;FIXME: validate final struct with regex.
    Local $aResult = [$sStruct, $tElements]
    Return $aResult
    Return $sStruct
EndFunc

Func __DllStructEx_ParseStructType($sType, $tElements = Null)
    Local $tElement = IsDllStruct($tElements) ? DllStructCreate($g__DllStructEx_tagElement, DllStructGetPtr($tElements, "Elements") + $g__DllStructEx_iElement * $tElements.Index) : DllStructCreate($g__DllStructEx_tagElement)
    ; The size of type in bytes
    Local $iSize = Null
    Local $sTranslatedType = $sType

    $tElement.iType = $g__DllStructEx_eElementType_Element

    Switch ($sType)
        Case "BYTE", "BOOLEAN", "CHAR"
            $iSize = 1
        Case "WCHAR", "SHORT", "USHORT", "WORD"
            $iSize = 2
        Case "INT", "LONG", "BOOL", "UINT", "ULONG", "DWORD", "FLOAT"
            $iSize = 4
        Case "INT64", "UINT64", "DOUBLE"
            $iSize = 8
        Case "PTR", "HWND", "HANDLE", "INT_PTR", "LONG_PTR", "LRESULT", "LPARAM", "UINT_PTR", "ULONG_PTR", "DWORD_PTR", "WPARAM"
            $iSize = @AutoItX64 ? 8 : 4
        Case "STRUCT", "ENDSTRUCT", "ALIGN"
            $tElement.iType = $g__DllStructEx_eElementType_NONE
            $iSize = 0
        Case Else
            ;FIXME: maybe we should assume type element, if the solved variable results in single native struct element type?
            $tElement.iType = $g__DllStructEx_eElementType_Struct
            If IsDeclared("tag" & $sType) Then
                Local $aResult = __DllStructEx_ParseStruct(Eval("tag" & $sType))
                If @error <> 0 Then Return SetError(2, @error, "")
                Local $sStruct = $aResult[0]
                $tStruct = DllStructCreate($sStruct)
                If @error <> 0 Then Return SetError(2, 0, "")
                $iSize = DllStructGetSize($tStruct)
                $sTranslatedType = StringFormat("BYTE[%d]", $iSize)
                ;Local $tmp = [$sType, $aResult[1]];FIXME: remmove so only string is returned
                ;$sType = $tmp
                ;$aResult[1]
                $tElement.cElements = 0;FIXME: set element count from $aResult[1]
                ;$tElement.pElements = ? ;FIXME: add ref to nested elements, if or union struct
            Elseif Not @Compiled Then
                ConsoleWrite(StringFormat('Error: Variable "tag%s" not defined!\n', $sType))
            EndIf
    EndSwitch

    If $iSize > 0 Then
        $tElement.szStruct = _WinAPI_CreateString($sType, 0, -1, True, False)
        $tElement.szTranslatedStruct = _WinAPI_CreateString($sTranslatedType, 0, -1, True, False)
        ;$tElement.pStruct = ? ;FIXME: add ref to struct start pointer
        $tElements.Index += 1
    EndIf

    If Null = $iSize Then Return SetError(1, 0, "")

    Return SetExtended($iSize, $sTranslatedType)
EndFunc

Func __DllStructEx_ParseStructTypeCallback($aMatches, $tElements)
    Local Static $anonymousElementCount
    Local $sPrepend = $aMatches[1]
    Local $sType = $aMatches[2]
    Local $sName = $aMatches[3]

    If "" = $sName Then
        $sName = StringFormat("_anonymousElement%d", $anonymousElementCount)
        $anonymousElementCount += 1
    EndIf

    Local $tElement = DllStructCreate($g__DllStructEx_tagElement, DllStructGetPtr($tElements, "Elements") + $g__DllStructEx_iElement * $tElements.Index)
    ;TODO: maybe we send $tElement instead of all of them, to reduce extra compute time?

    $sType = __DllStructEx_ParseStructType($sType, $tElements)
    Local $iSize = @extended
    If @error <> 0 Then Return SetError(1, 0, "")

    If $iSize > 0 Then
        Local $pName = _WinAPI_CreateString($sName, 0, -1, True, False)
        If @error <> 0 Then Return SetError(2, @error, "")
        $tElement.szName = $pName
    EndIf

    If IsArray($sType) Then
        ;FIXME: $sType[1] is a struct
        $sType = $sType[0]
    EndIf
    
    If Not (Null = $tElements) Then
        $tElements.Size = $tElements.Size < $iSize ? $iSize : $tElements.Size
    EndIf

    ;FIXME: generate a name if none exists
    If "" = $sName Then Return StringFormat("%s%s", $sPrepend, $sType)
    Local $aType = StringRegExp($sType, "(\w+)(\[\d+\])?", 1)
    Return StringFormat("%s%s %s%s", $sPrepend, $aType[0], $sName, UBound($aType, 1) > 1 ? $aType[1] : "")
EndFunc

Func __DllStructEx_ParseUnion($aUnion, $tUnions)
    Local Static $anonymousUnionCount = 0
    Local $sPrepend = $aUnion[1]
    Local $sName = UBound($aUnion) > 3 ? $aUnion[3] : ""
    If $sName = "" Then
        $sName = StringFormat("_anonymousUnion%s", $anonymousUnionCount)
        $anonymousUnionCount += 1
    EndIf
    Local $sUnionStruct = $aUnion[2]
    ;FIXME: if no name found, generate UID to match generated Element with parsed union

    $tUnions.Index += 1

    Local $aUnions = StringRegExp($sUnionStruct, "(?i)(^|;)\s*union\s*{[^}]+}\s*(?:\w+)?", 3)
    If @error = 2 Then Return SetError(3, @error, "")
    Local $_tUnions = DllStructCreate(StringFormat("INT Index;BYTE Elements[%d];", $g__DllStructEx_iElement * UBound($aUnions, 1)))
    ;FIXME: use $_tUnions

    Local $sStruct = StringRegExpReplaceCallbackEx($sUnionStruct, "(?i)(^|;)\s*union\s*{([^}]+)}\s*([^\s;]+)?", __DllStructEx_ParseUnion)
    If @error <> 0 Then Return SetError(1, @error, "")

    Local $aElements = StringRegExp($sStruct, "(^|;)\s*\w+(?:\s+\w+)?", 3)
    If @error <> 0 Then Return SetError(4, @error, "")
    Local $tElements = DllStructCreate(StringFormat("INT Index;INT Size;BYTE Elements[%d];", $g__DllStructEx_iElement * UBound($aElements, 1)))

    ;Local $tSize = DllStructCreate("INT Size;");FIXME: should be a $tElements instead
    $sStruct = StringRegExpReplaceCallbackEx($sStruct, "(^|;)\s*(\w+)(?:\s+(\w+))?", __DllStructEx_ParseStructTypeCallback, 0, $tElements)
    If @error <> 0 Then Return SetError(2, @error, "")
    Local $iOffset = 1
    Local $iBytes = $tElements.Size
    Return StringFormat("%sBYTE %s[%d]", $sPrepend, $sName, $iBytes)
EndFunc

#cs
# Get the DllStruct for the top level of the struct.
# @param DllStructEx $oDllStructEx A DllStructEx Object
# @return DllStruct
#ce
Func DllStructExGetStruct($oDllStructEx)
    Local $pDllStructEx = Ptr($oDllStructEx)
    Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pDllStructEx - 8)
    Local $sStruct = _WinAPI_GetString(DllStructGetData($tObject, "szTranslatedStruct"), True)
    Local $pStruct = DllStructGetData($tObject, "pStruct")
    Return DllStructCreate($sStruct, $pStruct)
EndFunc

#cs
# Get the string used when the DllStruct was created.
# @param DllStructEx $oDllStructEx A DllStructEx Object
# @return string
#ce
Func DllStructExGetStructString($oDllStructEx)
    Local $pDllStructEx = Ptr($oDllStructEx)
    Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pDllStructEx - 8)
    Return _WinAPI_GetString(DllStructGetData($tObject, "szStruct"), True)
EndFunc

Func DllStructExGetSize($oDllStructEx)
    $tStruct = DllStructExGetStruct($oDllStructEx)
    Return DllStructGetSize($tStruct)
EndFunc

;FIXME: either rename to fit namespace, or use a include.
Func StringRegExpReplaceCallbackEx($sString, $sPattern, $sFunc, $iLimit = 0, $vExtra = Null)
    Local $iOffset = 1, $iDone = 0, $iMatchOffset

    While True
        $aRes = StringRegExp($sString, $sPattern, 2, $iOffset)
        If @error Then ExitLoop

        $sRet = Call($sFunc, $aRes, $vExtra)
        If @error Then Return SetError(@error, $iDone, $sString)

        $iOffset = StringInStr($sString, $aRes[0], 1, 1, $iOffset)
        $sString = StringLeft($sString, $iOffset - 1) & $sRet & StringMid($sString, $iOffset + StringLen($aRes[0]))
        $iOffset += StringLen($sRet)

        $iDone += 1
        If $iDone = $iLimit Then ExitLoop
    WEnd

    Return SetExtended($iDone, $sString)
EndFunc   ;==>StringRegExpReplaceCallback

Func __DllStructEx_DllStructAlloc($sStruct, $tStruct = Null)
    If Not IsDllStruct($tStruct) Then $tStruct = DllStructCreate($sStruct)
    If @error <> 0 Then Return SetError(@error, @extended, 0)
    Local $iStruct = DllStructGetSize($tStruct)
    Local $hStruct = _MemGlobalAlloc($iStruct, $GMEM_MOVEABLE)
    Local $pStruct = _MemGlobalLock($hStruct)
    _MemMoveMemory(DllStructGetPtr($tStruct), $pStruct, $iStruct)
    Return DllStructCreate($sStruct, $pStruct)
EndFunc