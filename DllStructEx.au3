#include <Memory.au3>
#include <WinAPIMem.au3>

If Not IsDeclared("IID_IUnknown") Then Global Const $IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
If Not IsDeclared("IID_IDispatch") Then Global Const $IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"

If Not IsDeclared("DISPATCH_METHOD") Then Global Const $DISPATCH_METHOD = 1
If Not IsDeclared("DISPATCH_PROPERTYGET") Then Global Const $DISPATCH_PROPERTYGET = 2
If Not IsDeclared("DISPATCH_PROPERTYPUT") Then Global Const $DISPATCH_PROPERTYPUT = 4
If Not IsDeclared("DISPATCH_PROPERTYPUTREF") Then Global Const $DISPATCH_PROPERTYPUTREF = 8

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

Global Const $g__DllStructEx_tagVARIANT = "ushort vt;ushort r1;ushort r2;ushort r3;PTR data" & (@AutoItX64 ? "" : ";PTR data2")

If Not IsDeclared("VT_EMPTY") Then Global Enum $VT_EMPTY,$VT_NULL,$VT_I2,$VT_I4,$VT_R4,$VT_R8,$VT_CY,$VT_DATE,$VT_BSTR,$VT_DISPATCH, _
    $VT_ERROR,$VT_BOOL,$VT_VARIANT,$VT_UNKNOWN,$VT_DECIMAL,$VT_I1=16,$VT_UI1,$VT_UI2,$VT_UI4,$VT_I8, _
    $VT_UI8,$VT_INT,$VT_UINT,$VT_VOID,$VT_HRESULT,$VT_PTR,$VT_SAFEARRAY,$VT_CARRAY,$VT_USERDEFINED, _
    $VT_LPSTR,$VT_LPWSTR,$VT_RECORD=36,$VT_FILETIME=64,$VT_BLOB,$VT_STREAM,$VT_STORAGE,$VT_STREAMED_OBJECT, _
    $VT_STORED_OBJECT,$VT_BLOB_OBJECT,$VT_CF,$VT_CLSID,$VT_BSTR_BLOB=0xfff,$VT_VECTOR=0x1000, _
    $VT_ARRAY=0x2000,$VT_BYREF=0x4000,$VT_RESERVED=0x8000,$VT_ILLEGAL=0xffff,$VT_ILLEGALMASKED=0xfff, _
    $VT_TYPEMASK=0xfff

Global Const $g__DllStructEx_tagObject = "int RefCount;int Size;ptr Object;ptr Methods[7];ptr szStruct;ptr szTranslatedStruct;ptr pStruct;boolean ownPStruct;int cElements;ptr pElements;ptr pParent;"
Global Const $g__DllStructEx_tagElement = "int iType;ptr szName;ptr szStruct;ptr szTranslatedStruct;int cElements;ptr pElements;"
#cs
# The size of $g__DllStructEx_tagElement in bytes 
# @var int
#ce
Global Const $g__DllStructEx_iElement = DllStructGetSize(DllStructCreate($g__DllStructEx_tagElement))
Global Const $g__DllStructEx_tagElements = "INT Index;INT Size;BYTE Elements[%d];"

Global Enum $g__DllStructEx_eElementType_NONE, $g__DllStructEx_eElementType_UNION, $g__DllStructEx_eElementType_STRUCT, $g__DllStructEx_eElementType_Element, $g__DllStructEx_eElementType_PTR

Global Const $g__DllStructEx_sStructRegex = "(?ix)(?(DEFINE)(?<hn>[\h\n])(?<struct_line_declaration>(?:(?&declaration)|(?&struct)|(?&union));)(?<union>union(?&hn)*{(?&hn)*(?:(?&struct_line_declaration)(?&hn)*)+}(?:\h+(?&identifier))?(?&hn)*)(?<struct>struct(?&hn)*{(?&hn)*(?:(?&struct_line_declaration)(?&hn)*)+}(?:\h+(?&identifier))?(?&hn)*)(?<declaration>(?&type)\h+(?&identifier))(?<type>\w+)(?<identifier>[*]*\w+))"

#cs
# Creates a C/C++ style structure.
# @param string $sStruct A string representing the structure to create.
# @param ptr    $pData   If supplied the struct will not allocate memory but use the pointer supplied.
# @return DllStructEx
#ce
Func DllStructExCreate($sStruct, $pData = 0)
    Local $aResult = __DllStructEx_ParseStruct($sStruct)
    If @error <> 0 Then Return SetError(__DllStructEx_Error("Failed to create DllStruct from the final translated structure string", 1, @error), @error, "")
    Local $sTranslatedStruct = $aResult[0]
    Local $tElements = $aResult[1]
    ;ConsoleWrite($sTranslatedStruct&@CRLF);NOTE: for debugging

    ;Indicating if struct ptr is self created and should be released on object desctruction
    $ownPStruct = (0 = $pData)

    If $pData = 0 Then
        Local $tStruct = __DllStructEx_DllStructAlloc($sTranslatedStruct)
        If @error <> 0 Then Return SetError(__DllStructEx_Error("Failed to create DllStruct from the final translated structure string", 2, @error), @error, 0)
        $pData = DllStructGetPtr($tStruct)
    EndIf
    DllStructCreate($sTranslatedStruct, $pData)
    If @error <> 0 Then Return SetError(__DllStructEx_Error("Failed to create DllStruct from the final translated structure string, with data pointer", 2, @error), 0, 0)

    $oDllStructEx = __DllStructEx_Create($sStruct, $sTranslatedStruct, $tElements, $pData)
    If @error <> 0 Then Return SetError(__DllStructEx_Error("Failed creating DllStructEx(IDispatch) Object", 2, @extended), @error, 0)

    $tObject = DllStructCreate($g__DllStructEx_tagObject, Ptr($oDllStructEx) - 8)
    $tObject.ownPStruct = $ownPStruct

    Return $oDllStructEx
EndFunc

#cs
# Queries a COM object for a pointer to one of its interface; identifying the interface by a reference to its interface identifier (IID). If the COM object implements the interface, then it returns a pointer to that interface after calling __DllStructEx_AddRef on it.
# @internal
#ce
Func __DllStructEx_QueryInterface($pSelf, $pRIID, $pObj)
    If $pObj=0 Then Return $E_POINTER
    Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
    If (Not ($sGUID=$IID_IDispatch)) And (Not ($sGUID=$IID_IUnknown)) Then Return $E_NOINTERFACE
    Local $tStruct = DllStructCreate("ptr", $pObj)
    DllStructSetData($tStruct, 1, $pSelf)
    __DllStructEx_AddRef($pSelf)
    Return $S_OK
EndFunc

#cs
# Increments the reference count for an interface pointer to a COM object. You should call this method whenever you make a copy of an interface pointer.
# @internal
#ce
Func __DllStructEx_AddRef($pSelf)
    Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
    $tStruct.Ref += 1
    Return $tStruct.Ref
EndFunc

#cs
# Decrements the reference count for an interface on a COM object.
# @internal
#ce
Func __DllStructEx_Release($pSelf)
    Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
    $tStruct.Ref -= 1
    If $tStruct.Ref = 0 Then; initiate garbage collection
        Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf-8)
        If $tObject.pParent = 0 Then
            ;releases all properties
            _WinAPI_FreeMemory(DllStructGetData($tObject, "szStruct"))
            _WinAPI_FreeMemory(DllStructGetData($tObject, "szTranslatedStruct"))
            ;FIXME: release pElements data Tree!

            If $tObject.ownPStruct Then
                __DllStructEx_DllStructFree($tObject.pStruct)
            EndIf
        Else
            __DllStructEx_Release($tObject.pParent)
        EndIf
        $tObject = Null
        __DllStructEx_DllStructFree($pSelf-8)
        If @error <> 0 And Not @Compiled Then ConsoleWriteError(StringFormat("ERROR: could not release object (%s)\n", $pSelf - 8))
        Return 0
    EndIf
    Return $tStruct.Ref
EndFunc

#cs
# Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs, which can be used on subsequent calls to __DllStructEx_Invoke.
# @internal
#ce
Func __DllStructEx_GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
    Local $tDispId = DllStructCreate("long DispId;", $rgDispId)
    Local $pName = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)
    Local $sName = _WinAPI_GetString($pName, True) ;TODO: remove or outcomment. This only gets used for debugging
    ;ConsoleWrite($sName&@CRLF);NOTE: for debugging

    Local Static $CSTR_EQUAL = 2
    Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf - 8)
    Local $i
    For $i = 0 To $tObject.cElements - 1 Step +1
        Local $tElement = DllStructCreate($g__DllStructEx_tagElement, $tObject.pElements + $g__DllStructEx_iElement * $i)
        ;ConsoleWrite(StringFormat("\t[%02d]: %s\n", $i, _WinAPI_GetString($tElement.szName)));NOTE: for debugging
        If __DllStructEx_szStringsEqual($pName, $tElement.szName) Then
            DllStructSetData($tDispId, "DispId", $i + 1)
            Return $S_OK
        EndIf
    Next

    ;ConsoleWrite("Not found: "&$sName&@CRLF);NOTE: for debugging

    DllStructSetData($tDispId, "DispId", -1)
    Return $S_OK
EndFunc

#cs
# Retrieves the type information for an object, which can then be used to get the type information for an interface.
# @internal
#ce
Func __DllStructEx_GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
    If $iTInfo<>0 Then Return $DISP_E_BADINDEX
    If $ppTInfo=0 Then Return $E_INVALIDARG
    Return $S_OK
EndFunc

#cs
# Retrieves the number of type information interfaces that an object provides (either 0 or 1).
# @internal
#ce
Func __DllStructEx_GetTypeInfoCount($pSelf, $pctinfo)
    DllStructSetData(DllStructCreate("UINT",$pctinfo), 1, 0)
    Return $S_OK
EndFunc

#cs
# Provides access to properties and methods exposed by an object.
# @internal
#ce
Func __DllStructEx_Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
    If $dispIdMember = -1 Then Return $DISP_E_MEMBERNOTFOUND

    #cs
    ;For debugging: getting vt of autoit types
    ;TODO: remove this
    $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
    If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
    Local $tVARIANT=DllStructCreate($g__DllStructEx_tagVARIANT, $tDISPPARAMS.rgvargs)
    ConsoleWrite($tVARIANT.vt&@CRLF)
    Return $S_OK
    #ce

    If $dispIdMember = 0 Then
        If ((BitAND($wFlags, $DISPATCH_PROPERTYPUT)=$DISPATCH_PROPERTYPUT)) Then Return $DISP_E_EXCEPTION
        Local $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
        If $tDISPPARAMS.cArgs>1 Then Return $DISP_E_BADPARAMCOUNT
        If $tDISPPARAMS.cArgs = 1 Then
            Local $tVARIANT=DllStructCreate($g__DllStructEx_tagVARIANT, $tDISPPARAMS.rgvargs)
            If $tVARIANT.vt = $VT_BSTR Then
                Local $name = StringLower(_WinAPI_GetString($tVARIANT.data))
                Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf - 8)
                Local $i
                Local $tElement
                For $i = 0 To $tObject.cElements - 1 Step +1
                    $tElement = DllStructCreate($g__DllStructEx_tagElement, $tObject.pElements + $g__DllStructEx_iElement * $i)
                    If $name = StringLower(_WinAPI_GetString($tElement.szName)) Then
                        $dispIdMember = $i + 1
                        ExitLoop
                    EndIf
                Next
                If $dispIdMember = 0 Then Return $DISP_E_MEMBERNOTFOUND
            ElseIf $tVARIANT.vt = $VT_I4 Or $tVARIANT.vt = $VT_I8 Then
                $dispIdMember = DllStructGetData(DllStructCreate($tVARIANT.vt = $VT_I4 ? "INT" : "INT64", DllStructGetPtr($tVARIANT, 'data')), 1)
                If $dispIdMember < 1 Then Return $DISP_E_MEMBERNOTFOUND
            Else
                Return $DISP_E_BADVARTYPE
            EndIf

            If $dispIdMember = -1 Then Return $DISP_E_MEMBERNOTFOUND
        Else
            ;VariantInit($pVariant) ;TODO: it seems unclear if the return variant should go through VariantInit before usage?
            Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf-8)
            Local $tVARIANT = DllStructCreate($g__DllStructEx_tagVARIANT, $pVarResult)
            $tVARIANT.vt = $VT_BSTR
            $tVARIANT.data = DllStructGetData($tObject, "szTranslatedStruct")

            Return $S_OK
        EndIf
    EndIf

    If $dispIdMember < -1 Then
        ;
    EndIf

    If $dispIdMember > 0 Then
        Return __DllStructEx_Invoke_ProcessElement($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
    EndIf

    Return $DISP_E_MEMBERNOTFOUND
EndFunc

#cs
# Extracts member from object and assigns the converted value to the provided variant pointer.
# This needed to be it's own function, to allow recursive calls.
# @internal
# @param ptr    $pSelf        Pointer to the Object member of a DllStructEx_tagObject structure
# @param int    $dispIdMember The requested member identifier.
# @param ptr    $riid
# @param int    $lcid
# @param ushort $wFlags
# @param ptr    $pDispParams
# @param ptr    $pVarResult   Pointer to the variant where the result will be stored.
# @param ptr    $pExcepInfo
# @param ptr    $puArgErr
# @return long The HRESULT, indicating success state.
#ce
Func __DllStructEx_Invoke_ProcessElement($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
    Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pSelf - 8)
    If $tObject.cElements >= $dispIdMember Then
        Local $tElement = DllStructCreate($g__DllStructEx_tagElement, $tObject.pElements + $g__DllStructEx_iElement * ($dispIdMember - 1))
        If @error <> 0 Then Return $DISP_E_EXCEPTION
        Local $tVARIANT = DllStructCreate($g__DllStructEx_tagVARIANT, $pVarResult)
        If @error <> 0 Then Return $DISP_E_EXCEPTION
        Local $tStruct = DllStructCreate(_WinAPI_GetString($tObject.szTranslatedStruct, True), $tObject.pStruct)
        If @error <> 0 Then Return $DISP_E_EXCEPTION
        Switch $tElement.iType
            Case $g__DllStructEx_eElementType_Element
                If ((BitAND($wFlags, $DISPATCH_PROPERTYGET)=$DISPATCH_PROPERTYGET)) Then
                    Local $vData = DllStructGetData($tStruct, _WinAPI_GetString($tElement.szName, True));TODO: add support for getting struct data slice by index
                    __DllStructEx_DataToVariant($vData, $tVARIANT)
                    If @error <> 0 Then Return $DISP_E_EXCEPTION
                ElseIf ((BitAND($wFlags, $DISPATCH_PROPERTYPUT)=$DISPATCH_PROPERTYPUT)) Then
                    Local $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
                    If $tDISPPARAMS.cArgs<1 Then Return $DISP_E_BADPARAMCOUNT
                    Local $_tVARIANT=DllStructCreate($g__DllStructEx_tagVARIANT, $tDISPPARAMS.rgvargs)
                    Local $vData = __DllStructEx_VariantToData($_tVARIANT)
                    If @error <> 0 Then Return $DISP_E_EXCEPTION
                    DllStructSetData($tStruct, _WinAPI_GetString($tElement.szName, True), $vData)
                Else
                    Return $DISP_E_EXCEPTION
                EndIf
            Case $g__DllStructEx_eElementType_STRUCT
                If ((BitAND($wFlags, $DISPATCH_PROPERTYPUT)=$DISPATCH_PROPERTYPUT)) Then Return $DISP_E_EXCEPTION
                Local $tElements = DllStructCreate(StringFormat($g__DllStructEx_tagElements, 1)); WARNING: Elements property in this struct is "BYTE[1]", not intended for use!
                If @error <> 0 Then Return $DISP_E_EXCEPTION
                $tElements.Index = $tElement.cElements
                Local $vData = __DllStructEx_Create($tElement.szStruct, $tElement.szTranslatedStruct, $tElements, DllStructGetPtr($tStruct, _WinAPI_GetString($tElement.szName)), $tElement.pElements)
                Local $_tObject = DllStructCreate($g__DllStructEx_tagObject, Ptr($vData)-8)
                $_tObject.pParent = $pSelf
                __DllStructEx_AddRef($pSelf)
                __DllStructEx_AddRef(Ptr($vData)) ; add ref, as we pass it into a variant
                __DllStructEx_DataToVariant($vData, $tVARIANT)
                If @error <> 0 Then Return $DISP_E_EXCEPTION
            Case $g__DllStructEx_eElementType_UNION
                If ((BitAND($wFlags, $DISPATCH_PROPERTYPUT)=$DISPATCH_PROPERTYPUT)) Then Return $DISP_E_EXCEPTION
                Local $tElements = DllStructCreate(StringFormat($g__DllStructEx_tagElements, 1)); WARNING: Elements property in this struct is "BYTE[1]", not intended for use!
                If @error <> 0 Then Return $DISP_E_EXCEPTION
                $tElements.Index = $tElement.cElements
                Local $vData = __DllStructEx_Create($tElement.szStruct, $tElement.szTranslatedStruct, $tElements, DllStructGetPtr($tStruct, _WinAPI_GetString($tElement.szName)), $tElement.pElements)
                Local $_tObject = DllStructCreate($g__DllStructEx_tagObject, Ptr($vData)-8)
                $_tObject.pParent = $pSelf
                __DllStructEx_AddRef($pSelf)
                __DllStructEx_AddRef(Ptr($vData)) ; add ref, as we pass it into a variant
                __DllStructEx_DataToVariant($vData, $tVARIANT)
                If @error <> 0 Then Return $DISP_E_EXCEPTION
            Case $g__DllStructEx_eElementType_PTR
                Local $iPtrLevelCount = $tElement.cElements
                Local $sType = _WinAPI_GetString($tElement.szStruct)
                Local $pLevel = DllStructGetData($tStruct, _WinAPI_GetString($tElement.szName))
                Local $iLevel = 1
                For $iLevel = 1 To $iPtrLevelCount - 1 Step +1 ;This loop handles all but final ptr.
                    If 0 = $pLevel Then
                        __DllStructEx_Error("ptr ref empty")
                        Return $DISP_E_EXCEPTION;FIXME: should we return null instead of throwing an exception, or maybe do both?
                    EndIf
                    $pLevel = DllStructGetData(DllStructCreate("PTR", $pLevel), 1)
                Next

                If 0 = $pLevel Then
                    __DllStructEx_Error("ptr ref empty")
                    Return $DISP_E_EXCEPTION;FIXME: should we return null instead of throwing an exception, or maybe do both?
                EndIf

                Switch $sType
                    Case 'IUnknown'
                        If ((BitAND($wFlags, $DISPATCH_PROPERTYPUT)=$DISPATCH_PROPERTYPUT)) Then Return $DISP_E_EXCEPTION; FIXME: remove this, when propertyput is supported!
                        $tVariant.vt = $VT_UNKNOWN
                        $tVariant.data = $pLevel
                        Return $S_OK
                    Case 'IDispatch'
                        If ((BitAND($wFlags, $DISPATCH_PROPERTYPUT)=$DISPATCH_PROPERTYPUT)) Then Return $DISP_E_EXCEPTION; FIXME: remove this, when propertyput is supported!
                        $tVariant.vt = $VT_DISPATCH
                        $tVariant.data = $pLevel
                        Return $S_OK
                    Case Else
                        Local $tElements = DllStructCreate(StringFormat($g__DllStructEx_tagElements, $g__DllStructEx_iElement))
                        $tElements.iSize = 1; NOTE: maybe this line is not needed. Look into it, for optimization.

                        Local $sTranslatedType = __DllStructEx_ParseStructType($sType, $tElements)
                        
                        $_tObject = DllStructCreate($g__DllStructEx_tagObject)
                        $_tObject.cElements = 1
                        $_tObject.pElements = DllStructGetPtr($tElements, "Elements")
                        $_tObject.pStruct = $pLevel
                        $_tObject.szTranslatedStruct = __DllStructEx_CreateString($sTranslatedType)
                        $iResponse = __DllStructEx_Invoke_ProcessElement(DllStructGetPtr($_tObject, "Object"), 1, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
                        ;Cleanup manually managed memory, before returning
                        _WinAPI_FreeMemory($_tObject.szTranslatedStruct)
                        Return $iResponse
                EndSwitch
            Case Else
                __DllStructEx_Error(StringFormat('struct element type not supported "%d"', $tElement.iType))
                Return $DISP_E_EXCEPTION
        EndSwitch
        Return $S_OK
    EndIf
EndFunc

#cs
# Create DllStructEx object.
# @internal
# @param string $sStruct
# @param string $sTranslatedStruct
# @param DllStruct $tElements
# @param ptr $pData
# @param ptr|null $pElements if not null, specifies the ptr for the elements to use. Fallback is the Elements index from $tElements
# @return DllStructEx
#ce
Func __DllStructEx_Create($sStruct, $sTranslatedStruct, $tElements, $pData, $pElements = Null)
    Local $tObject = DllStructCreate($g__DllStructEx_tagObject)

    $tObject.cElements = $tElements.Index
    Local $sElements = StringFormat("BYTE[%d]", $g__DllStructEx_iElement * $tObject.cElements)
    $tObject.pElements = DllStructGetPtr(__DllStructEx_DllStructAlloc($sElements, DllStructCreate($sElements, $pElements = Null ? DllStructGetPtr($tElements, "Elements") : $pElements)))

    Local $pString = IsString($sStruct) ? __DllStructEx_CreateString($sStruct) : $sStruct
    DllStructSetData($tObject, "szStruct", $pString)
    $pString = IsString($sTranslatedStruct) ? __DllStructEx_CreateString($sTranslatedStruct) : $sTranslatedStruct
    DllStructSetData($tObject, "szTranslatedStruct", $pString)

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

    DllStructSetData($tObject, "RefCount", 1) ; initial ref count is 1
    DllStructSetData($tObject, "Size", 7) ; number of interface methods

    $tObject = __DllStructEx_DllStructAlloc($g__DllStructEx_tagObject, $tObject)

    DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
    Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IDispatch, Default, True) ; pointer that's wrapped into object
EndFunc

#cs
# Convert AutoIt data to variant.
# @internal
# @param mixed          $vData
# @param DllStruct|null $tVariant variant
# @return DllStruct variant
#ce
Func __DllStructEx_DataToVariant($vData, $tVARIANT = Null)
    If Not IsDllStruct($tVARIANT) Then $tVARIANT = DllStructCreate($g__DllStructEx_tagVARIANT)
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
            $tData.data = __DllStructEx_CreateString($vData) ;TODO: should use SysAllocString instead.
        #cs
        Case 'Binary'
            $tVARIANT.vt = BitOR($VT_ARRAY, $VT_BOOL)
            ;FIXME: support Binary
        #ce
        Case 'Pointer'
            $tVARIANT.vt = @AutoItX64 ? $VT_UI8 : $VT_UI4
            Local $tData = DllStructCreate("PTR data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        Case 'Object'
            Local $tGUID = DllStructCreate("byte[16]")
            Local $pRIID = DllCall("ole32.dll", "LONG", "CLSIDFromString", "wstr", $IID_IDispatch, "struct*", $tGUID)
            Local $pQueryInterface = DllStructGetData(DllStructCreate("PTR", DllStructGetData(DllStructCreate("PTR", Ptr($vData)), 1)), 1)
            Local $pObj = DllStructCreate("ptr")
            Local $response = DllCallAddress("LONG", $pQueryInterface, "ptr", Ptr($vData), "struct*", $tGUID, "ptr", DllStructGetPtr($pObj))
            If @error <> 0 Then Return SetError(2, 0, $tVARIANT)
            If $response[0] = $S_OK Then
                Local $pRelease = DllStructGetData(DllStructCreate("PTR", DllStructGetData(DllStructCreate("PTR", DllStructGetData($pObj, 1)), 1) + (@AutoItX64 ? 8 : 4) * 2), 1)
                DllCallAddress("dword", $pRelease, "PTR", DllStructGetData($pObj, 1)); release the idispatch object we got back
            EndIf
            $tVARIANT.vt = $response[0] = $S_OK ? $VT_DISPATCH : $VT_UNKNOWN
            Local $tData = DllStructCreate("PTR data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        Case Else
            If Not @Compiled Then ConsoleWriteError(StringFormat('ERROR: Unsupported Datatype "%s" for VARIANT\n', VarGetType($vData)))
            Return SetError(1, 0, $tVARIANT)
    EndSwitch

    Return $tVARIANT
EndFunc

Func __DllStructEx_VariantToData($tVARIANT)
    Local $sType = Null
    Switch $tVARIANT.vt
        Case $VT_BSTR, $VT_LPSTR, $VT_LPWSTR
            Return _WinAPI_GetString($tVARIANT.data, Not ($tVARIANT.vt = $VT_LPSTR))
        Case $VT_I2
            $sType = "SHORT"
        Case $VT_I4, $VT_INT
            $sType = "INT"
        Case $VT_I8
            $sType = "INT64"
        Case $VT_NULL
            Return Null
        Case $VT_PTR
            Return $tVARIANT.data
        Case $VT_R4
            $sType = "FLOAT"
        Case $VT_R8
            $sType = "DOUBLE"
        Case $VT_UI4, $VT_UI8
            Return Ptr($tVARIANT.data)
        Case $VT_UINT
            $sType = "UINT"
        ;Case $VT_DISPATCH
        ;Case $VT_UNKNOWN
            ;TODO: support dispatch and unknown...
        Case Else
            Return SetError(1, 0, Null)
    EndSwitch
    Return DllStructGetData(DllStructCreate(StringFormat("%s data", $sType), DllStructGetPtr($tVARIANT, "data")), "data")
EndFunc

#cs
# Parse DllStructEx string.
# @internal
# @param string $sStruct
# @return [string, DllStruct]
#ce
Func __DllStructEx_ParseStruct($sStruct)
    If Not StringRegExp($sStruct, $g__DllStructEx_sStructRegex&"^(?&struct_line_declaration)+$", 0) Then Return SetError(__DllStructEx_Error("Regex validation for the entire struct provided failed!", 3), @error, "")
    Local $aStructLineDeclarations = StringRegExp($sStruct, $g__DllStructEx_sStructRegex&"(?&struct_line_declaration)", 3)
    If @error <> 0 Then Return SetError(1, @error, "")
    Local $tElements = DllStructCreate(StringFormat($g__DllStructEx_tagElements, $g__DllStructEx_iElement * UBound($aStructLineDeclarations, 1)))

    $sStruct = ""
    Local $i
    Local $sStructLineDeclaration
    For $i = 0 To UBound($aStructLineDeclarations) - 1 Step +1
        $sStructLineDeclaration = $aStructLineDeclarations[$i]
        Select
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&union);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $aUnion = StringRegExp($sStructLineDeclaration, "(?i)^\s*union\s*{(.*)}\s*(\w+)?;?$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for union-declaration failed", 2), @error, "")
                $sStruct &= __DllStructEx_ParseUnion($aUnion, $tElements)
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&struct);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $_aStruct = StringRegExp($sStructLineDeclaration, "(?i)^\s*struct\s*{(.*)}\s*(\w+)?;?$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                $sStruct &= __DllStructEx_ParseNestedStruct($_aStruct, $tElements)
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&declaration);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $aMatches = StringRegExp($sStructLineDeclaration, "^\s*(\w+)(?:\s+([*]*)(\w+))?;$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for struct-line-declaration failed", 2), @error, "")
                $sStruct &= __DllStructEx_ParseStructTypeCallback($aMatches, $tElements)
            Case Else
                ConsoleWriteError("a struct-line-declaration could not be recognized!")
        EndSelect
    Next

    Local $aResult = [$sStruct, $tElements]
    Return $aResult
EndFunc

#cs
# Parse DllStruct type.
# The @extended value will be equal to the size of type in bytes.
# @internal
# @param string         $sType
# @param DllStruct|null $tElements
# @return string
#ce
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
            $tElement.iType = $g__DllStructEx_eElementType_STRUCT
            If IsDeclared("tag" & $sType) Then
                Local $aResult = __DllStructEx_ParseStruct(Eval("tag" & $sType))
                If @error <> 0 Then Return SetError(__DllStructEx_Error(StringFormat('Parsing of struct in variable "tag%s" failed.', $sType), 2), @error, "")
                Local $sStruct = $aResult[0]
                Local $_tElements = $aResult[1]
                $tStruct = DllStructCreate($sStruct)
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Failed creating DllStruct", 2), @error, "")
                $iSize = DllStructGetSize($tStruct)
                $sTranslatedType = StringFormat("BYTE[%d]", $iSize)
                $tElement.cElements = $_tElements.Index
                ;If $tElement.cElements > 0 Then ConsoleWrite("> "&_WinAPI_GetString(DllStructCreate($g__DllStructEx_tagElement, DllStructGetPtr($_tElements, "Elements")).szName)&@CRLF)
                $tElement.pElements = DllStructGetPtr(__DllStructEx_DllStructAlloc($sTranslatedType, DllStructCreate($sTranslatedType, DllStructGetPtr($_tElements, "Elements"))))
                $tElement.szStruct = __DllStructEx_CreateString($sType)
                $tElement.szTranslatedStruct = __DllStructEx_CreateString($sStruct)
                $tElements.Index += 1
                Return SetExtended($iSize, $sTranslatedType)
            Elseif Not @Compiled Then
                ConsoleWriteError(StringFormat('Error: Variable "tag%s" not defined!\n', $sType))
            EndIf
    EndSwitch

    If $iSize > 0 Then
        $tElement.szStruct = __DllStructEx_CreateString($sType)
        $tElement.szTranslatedStruct = __DllStructEx_CreateString($sTranslatedType)
        ;$tElement.pStruct = ? ;NOTE: not needed here, only here for clarity
        $tElements.Index += 1
    EndIf

    If Null = $iSize Then Return SetError(1, 0, "")

    Return SetExtended($iSize, $sTranslatedType)
EndFunc

#cs
# Callback for parsing a union string.
# @internal
# @param string[]  $aUnion
# @param DllStruct $tUnions
# @return string
#ce
Func __DllStructEx_ParseStructTypeCallback($aMatches, $tElements)
    Local Static $anonymousElementCount
    Local $sType = $aMatches[0]
    Local $sPtr = UBound($aMatches) > 1 ? $aMatches[1] : ""
    Local $sName = UBound($aMatches) > 2 ? $aMatches[2] : ""

    If "" = $sName Then
        $sName = StringFormat("_anonymousElement%d", $anonymousElementCount)
        $anonymousElementCount += 1
    EndIf

    Local $tElement = DllStructCreate($g__DllStructEx_tagElement, DllStructGetPtr($tElements, "Elements") + $g__DllStructEx_iElement * $tElements.Index)
    
    If Not ("" = $sPtr) Then
        $tElement.iType = $g__DllStructEx_eElementType_PTR
        $tElement.szName = __DllStructEx_CreateString($sName)
        $tElement.szStruct = __DllStructEx_CreateString($sType)
        $tElement.szTranslatedStruct = 0
        $tElement.cElements = StringLen($sPtr)
        $tElement.pElements = 0
        $tElements.Index += 1
        Return StringFormat("PTR %s;", $sName)
    EndIf

    ;TODO: maybe we send $tElement instead of all of them, to reduce extra compute time?
    $sType = __DllStructEx_ParseStructType($sType, $tElements)
    Local $iSize = @extended
    If @error <> 0 Then Return SetError(1, 0, "")

    If $iSize > 0 Then
        Local $pName = __DllStructEx_CreateString($sName)
        If @error <> 0 Then Return SetError(2, @error, "")
        $tElement.szName = $pName
    EndIf

    If Not (Null = $tElements) Then
        $tElements.Size = $tElements.Size < $iSize ? $iSize : $tElements.Size
    EndIf

    Local $aType = StringRegExp($sType, "(\w+)(\[\d+\])?", 1)
    Return StringFormat("%s %s%s;", $aType[0], $sName, UBound($aType, 1) > 1 ? $aType[1] : "")
EndFunc

#cs
# Callback for parsing a union string.
# @internal
# @param string[]  $aUnion
# @param DllStruct $tUnions
# @return string
#ce
Func __DllStructEx_ParseUnion($aUnion, $tUnions)
    Local Static $anonymousUnionCount = 0
    ;$sStruct = __DllStructEx_StringRegExpReplaceCallbackEx($sStruct, "(?i)(^|;)\s*union\s*{([^}]+)}\s*(\w+)?", __DllStructEx_ParseUnion, 0, $tUnions)
    Local $sName = UBound($aUnion) > 1 ? $aUnion[1] : ""
    If $sName = "" Then
        $sName = StringFormat("_anonymousUnion%s", $anonymousUnionCount)
        $anonymousUnionCount += 1
    EndIf
    Local $sUnionStruct = $aUnion[0]

    $tUnion = DllStructCreate($g__DllStructEx_tagElement, DllStructGetPtr($tUnions, "Elements") + $g__DllStructEx_iElement * $tUnions.Index)
    $tUnion.iType = $g__DllStructEx_eElementType_UNION
    $tUnion.szName = __DllStructEx_CreateString($sName)
    $tUnion.szStruct = __DllStructEx_CreateString($sUnionStruct)
    If @error <> 0 Then Return SetError(@error, @extended, "")
    $tUnions.Index += 1

    Local $aStructLineDeclarations = StringRegExp($sUnionStruct, $g__DllStructEx_sStructRegex&"(?&struct_line_declaration)", 3)
    If @error <> 0 Then Return SetError(1, @error, "")
    Local $tElements = DllStructCreate(StringFormat($g__DllStructEx_tagElements, $g__DllStructEx_iElement * UBound($aStructLineDeclarations, 1)))

    $sUnionStruct = ""
    Local $i
    Local $sStructLineDeclaration
    For $i = 0 To UBound($aStructLineDeclarations) - 1 Step +1
        $sStructLineDeclaration = $aStructLineDeclarations[$i]
        Select
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&union);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $_aUnion = StringRegExp($sStructLineDeclaration, "(?i)^\s*union\s*{(.*)}\s*(\w+)?;?$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for union-declaration failed", 2), @error, "")
                $sUnionStruct &= __DllStructEx_ParseUnion($_aUnion, $tElements)
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&struct);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $aStruct = StringRegExp($sStructLineDeclaration, "(?i)^\s*struct\s*{(.*)}\s*(\w+)?;?$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for struct-declaration failed", 2), @error, "")
                $sUnionStruct &= __DllStructEx_ParseNestedStruct($aStruct, $tElements)
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&declaration);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $aMatches = StringRegExp($sStructLineDeclaration, "^\s*(\w+)(?:\s+([*]*)(\w+))?;$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for struct-line-declaration failed", 2), @error, "")
                $sUnionStruct &= __DllStructEx_ParseStructTypeCallback($aMatches, $tElements)
            Case Else
                ConsoleWriteError("a struct-line-declaration could not be recognized!")
        EndSelect
    Next

    $tUnion.szTranslatedStruct = __DllStructEx_CreateString($sUnionStruct)
    $tUnion.cElements = $tElements.Index
    Local $sElements = StringFormat("BYTE[%d]", $g__DllStructEx_iElement * $tUnion.cElements)
    $tUnion.pElements = DllStructGetPtr(__DllStructEx_DllStructAlloc($sElements, DllStructCreate($sElements, DllStructGetPtr($tElements, "Elements"))))

    Local $iBytes = $tElements.Size
    Return StringFormat("BYTE %s[%d];", $sName, $iBytes)
EndFunc

#cs
# Callback for parsing a struct string (struct {...})
# @internal
# @param string[]  $aStruct
# @param DllStruct $tStructs
# @return string
#ce
Func __DllStructEx_ParseNestedStruct($aStruct, $tStructs)
    Local Static $anonymousStructCount = 0
    Local $sName = UBound($aStruct) > 1 ? $aStruct[1] : ""
    If $sName = "" Then
        $sName = StringFormat("_anonymousStruct%s", $anonymousStructCount)
        $anonymousStructCount += 1
    EndIf
    Local $sStruct = $aStruct[0]

    $tStruct = DllStructCreate($g__DllStructEx_tagElement, DllStructGetPtr($tStructs, "Elements") + $g__DllStructEx_iElement * $tStructs.Index)
    $tStruct.iType = $g__DllStructEx_eElementType_STRUCT
    $tStruct.szName = __DllStructEx_CreateString($sName)
    $tStruct.szStruct = __DllStructEx_CreateString($sStruct)
    If @error <> 0 Then Return SetError(@error, @extended, "")
    $tStructs.Index += 1

    Local $aStructLineDeclarations = StringRegExp($sStruct, $g__DllStructEx_sStructRegex&"(?&struct_line_declaration)", 3)
    If @error <> 0 Then Return SetError(1, @error, "")
    Local $tElements = DllStructCreate(StringFormat($g__DllStructEx_tagElements, $g__DllStructEx_iElement * UBound($aStructLineDeclarations, 1)))

    $sStruct = ""
    Local $i
    Local $sStructLineDeclaration
    For $i = 0 To UBound($aStructLineDeclarations) - 1 Step +1
        $sStructLineDeclaration = $aStructLineDeclarations[$i]
        Select
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&union);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $aUnion = StringRegExp($sStructLineDeclaration, "(?i)^\s*union\s*{(.*)}\s*(\w+)?;?$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                $sStruct &= __DllStructEx_ParseUnion($aUnion, $tElements)
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&struct);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $_aStruct = StringRegExp($sStructLineDeclaration, "(?i)^\s*struct\s*{(.*)}\s*(\w+)?;?$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                $sStruct &= __DllStructEx_ParseNestedStruct($_aStruct, $tElements)
            Case StringRegExp($sStructLineDeclaration, $g__DllStructEx_sStructRegex&"^(?&declaration);$", $STR_REGEXPMATCH);TODO: should the semicolon be optional in the regex?
                Local $aMatches = StringRegExp($sStructLineDeclaration, "^\s*(\w+)(?:\s+([*]*)(\w+))?;$", $STR_REGEXPARRAYMATCH);FIXME: update regex to be more clear
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for struct-line-declaration failed", 2), @error, "")
                $sStruct &= __DllStructEx_ParseStructTypeCallback($aMatches, $tElements)
            Case Else
                ConsoleWriteError("a struct-line-declaration could not be recognized!")
        EndSelect
    Next

    $tStruct.szTranslatedStruct = __DllStructEx_CreateString($sStruct)
    $tStruct.cElements = $tElements.Index
    Local $sElements = StringFormat("BYTE[%d]", $g__DllStructEx_iElement * $tStruct.cElements)
    $tStruct.pElements = DllStructGetPtr(__DllStructEx_DllStructAlloc($sElements, DllStructCreate($sElements, DllStructGetPtr($tElements, "Elements"))))

    Local $iBytes = $tElements.Size
    Return StringFormat("STRUCT;BYTE %s[%d];ENDSTRUCT;", $sName, $iBytes)
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

#cs
# Get the AutoIt dllstruct compatible string transpiled from the string used when the DllStruct was created.
# @param DllStructEx $oDllStructEx A DllStructEx Object
# @return string
#ce
Func DllStructExGetTranspiledStructString($oDllStructEx)
    Local $pDllStructEx = Ptr($oDllStructEx)
    Local $tObject = DllStructCreate($g__DllStructEx_tagObject, $pDllStructEx - 8)
    Return _WinAPI_GetString(DllStructGetData($tObject, "szTranslatedStruct"), True)
EndFunc

#cs
# Get the stuct size in bytes.
# @param DllStructEx $oDllStructEx A DllStructEx Object
# @return int
#ce
Func DllStructExGetSize($oDllStructEx)
    $tStruct = DllStructExGetStruct($oDllStructEx)
    Return DllStructGetSize($tStruct)
EndFunc

#cs
# Perform a regular expression search and replace using a callback.
# @author Mat
# @modifier genius257
# @param string          $sString
# @param string          $sPattern
# @param function|string $sFunc
# @param int             $iLimit
# @param mixed           $vExtra
# @return string
#ce
Func __DllStructEx_StringRegExpReplaceCallbackEx($sString, $sPattern, $sFunc, $iLimit = 0, $vExtra = Null)
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

#cs
# Allocate DllStruct memory
# @internal
# @param string         $sStruct
# @param DllStruct|null $tStruct
# @return DllStruct
#ce
Func __DllStructEx_DllStructAlloc($sStruct, $tStruct = Null)
    If Not IsDllStruct($tStruct) Then $tStruct = DllStructCreate($sStruct)
    If @error <> 0 Then Return SetError(@error, @extended, 0)
    Local $iStruct = DllStructGetSize($tStruct)
    Local $hStruct = _MemGlobalAlloc($iStruct, $GMEM_MOVEABLE)
    Local $pStruct = _MemGlobalLock($hStruct)
    _MemMoveMemory(DllStructGetPtr($tStruct), $pStruct, $iStruct)
    Return DllStructCreate($sStruct, $pStruct)
EndFunc

#cs
# Free DllStruct memory
# @internal
# @param ptr $pStruct
# @return bool
#ce
Func __DllStructEx_DllStructFree($pStruct)
    If IsDllStruct($pStruct) Then $pStruct = DllStructGetPtr($pStruct)
    Local $aRet = DllCall("Kernel32.dll", "ptr", "GlobalHandle", "ptr", $pStruct)
    If @error <> 0 Or $aRet[0] = 0 Then Return SetError(@error, @extended, False)
    Local $hMemory = _MemGlobalFree($aRet[1])
    if $hMemory <> 0 Then SetError(-1, 0, False)
    Return True
EndFunc

#cs
# Conpare strings
# @internal
# @param ptr $szString1
# @param ptr $szString2
# @return int
#ce
Func __DllStructEx_CompareSzStrings($szString1, $szString2)
    ; FIXME: add support for comparing strings on win xp
    ;CompareStringOrdinal
    ;_WinAPI_CompareString
    Local $aRet = DllCall("kernel32.dll", "int", "CompareStringOrdinal", "ptr", $szString1, "int", -1, "ptr", $szString2, "int", -1, "BOOL", 1)
    If @error <> 0 Then Return SetError(@error, @extended, 0)
    Return $aRet[0]
EndFunc

#cs
# Check if strings are equal.
# @internal
# @param ptr $szString1
# @param ptr $szString2
# @return bool
#ce
Func __DllStructEx_szStringsEqual($szString1, $szString2)
    Local Static $CSTR_EQUAL = 2
    $iResult = __DllStructEx_CompareSzStrings($szString1, $szString2)
    If @error <> 0 Then Return SetError(@error, @extended, False)
    Return $iResult = $CSTR_EQUAL
EndFunc

#cs
# Copies a specified string to the newly allocated memory block and returns its pointer
# @internal
# @param string $sString
# @return ptr
#ce
Func __DllStructEx_CreateString($sString)
    Local $pString = _WinAPI_CreateString($sString, 0, -1, True, False)
    If @error <> 0 Then Return SetError(@error, @extended, 0)
    Return $pString
EndFunc

#cs 
# Helper function for debugging.
# Allows to mimic VERY BASIC exceptions
# @internal
# @param string $message  The exception message to "throw"
# @param int    $code     The Exception code
# @param mixed  $extended Extended error code
# @param int    $line     Scipt line number in which the exceotuib was created
# @param string $file     name of the file in chich the exception was created
# @returns int The exception code
#ce
Func __DllStructEx_Error($message = "", $code = 0, $extended = 0, $line = @ScriptLineNumber, $file = "DllStructEx.au3")
    ;NOTE: would be nice to have $file default to @ScriptFullPath, but silly AutoIt3 only returns the main executing script...
    If Not @Compiled And @NumParams > 3 Then ConsoleWriteError(StringFormat("WARNING: following error had scriptline manually overridden\n"))
    If Not @Compiled And @NumParams > 4 Then ConsoleWriteError(StringFormat("WARNING: following error had file manually overridden\n"))
    If Not @Compiled Then ConsoleWriteError(StringFormat("Error: %s\n  in %s on line %s\n", $message, $file, $line))
    Return $code
EndFunc
