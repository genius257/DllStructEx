#include-once

#include <Memory.au3>
#include <WinAPIMem.au3>

;FIXME: verify if ptr objects made, have proper cleanup, as they make their memory JIT style!

Global Const $__g_DllStructEx_IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Global Const $__g_DllStructEx_IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"

Global Const $__g_DllStructEx_DISPATCH_METHOD = 1
Global Const $__g_DllStructEx_DISPATCH_PROPERTYGET = 2
Global Const $__g_DllStructEx_DISPATCH_PROPERTYPUT = 4
Global Const $__g_DllStructEx_DISPATCH_PROPERTYPUTREF = 8

Global Const $__g_DllStructEx_S_OK = 0x00000000
Global Const $__g_DllStructEx_E_NOTIMPL = 0x80004001
Global Const $__g_DllStructEx_E_NOINTERFACE = 0x80004002
Global Const $__g_DllStructEx_E_POINTER = 0x80004003
Global Const $__g_DllStructEx_E_ABORT = 0x80004004
Global Const $__g_DllStructEx_E_FAIL = 0x80004005
Global Const $__g_DllStructEx_E_ACCESSDENIED = 0x80070005
Global Const $__g_DllStructEx_E_HANDLE = 0x80070006
Global Const $__g_DllStructEx_E_OUTOFMEMORY = 0x8007000E
Global Const $__g_DllStructEx_E_INVALIDARG = 0x80070057
Global Const $__g_DllStructEx_E_UNEXPECTED = 0x8000FFFF

Global Const $__g_DllStructEx_DISP_E_UNKNOWNINTERFACE = 0x80020001
Global Const $__g_DllStructEx_DISP_E_MEMBERNOTFOUND = 0x80020003
Global Const $__g_DllStructEx_DISP_E_PARAMNOTFOUND = 0x80020004
Global Const $__g_DllStructEx_DISP_E_TYPEMISMATCH = 0x80020005
Global Const $__g_DllStructEx_DISP_E_UNKNOWNNAME = 0x80020006
Global Const $__g_DllStructEx_DISP_E_NONAMEDARGS = 0x80020007
Global Const $__g_DllStructEx_DISP_E_BADVARTYPE = 0x80020008
Global Const $__g_DllStructEx_DISP_E_EXCEPTION = 0x80020009
Global Const $__g_DllStructEx_DISP_E_OVERFLOW = 0x8002000A
Global Const $__g_DllStructEx_DISP_E_BADINDEX = 0x8002000B
Global Const $__g_DllStructEx_DISP_E_UNKNOWNLCID = 0x8002000C
Global Const $__g_DllStructEx_DISP_E_ARRAYISLOCKED = 0x8002000D
Global Const $__g_DllStructEx_DISP_E_BADPARAMCOUNT = 0x8002000E
Global Const $__g_DllStructEx_DISP_E_PARAMNOTOPTIONAL = 0x8002000F
Global Const $__g_DllStructEx_DISP_E_BADCALLEE = 0x80020010
Global Const $__g_DllStructEx_DISP_E_NOTACOLLECTION = 0x80020011

Global Const $__g_DllStructEx_tagVARIANT = "ushort vt;ushort r1;ushort r2;ushort r3;PTR data" & (@AutoItX64 ? "" : ";PTR data2")

Global Enum $__g_DllStructEx_VT_EMPTY,$__g_DllStructEx_VT_NULL,$__g_DllStructEx_VT_I2,$__g_DllStructEx_VT_I4,$__g_DllStructEx_VT_R4,$__g_DllStructEx_VT_R8,$__g_DllStructEx_VT_CY,$__g_DllStructEx_VT_DATE,$__g_DllStructEx_VT_BSTR,$__g_DllStructEx_VT_DISPATCH, _
    $__g_DllStructEx_VT_ERROR,$__g_DllStructEx_VT_BOOL,$__g_DllStructEx_VT_VARIANT,$__g_DllStructEx_VT_UNKNOWN,$__g_DllStructEx_VT_DECIMAL,$__g_DllStructEx_VT_I1=16,$__g_DllStructEx_VT_UI1,$__g_DllStructEx_VT_UI2,$__g_DllStructEx_VT_UI4,$__g_DllStructEx_VT_I8, _
    $__g_DllStructEx_VT_UI8,$__g_DllStructEx_VT_INT,$__g_DllStructEx_VT_UINT,$__g_DllStructEx_VT_VOID,$__g_DllStructEx_VT_HRESULT,$__g_DllStructEx_VT_PTR,$__g_DllStructEx_VT_SAFEARRAY,$__g_DllStructEx_VT_CARRAY,$__g_DllStructEx_VT_USERDEFINED, _
    $__g_DllStructEx_VT_LPSTR,$__g_DllStructEx_VT_LPWSTR,$__g_DllStructEx_VT_RECORD=36,$__g_DllStructEx_VT_FILETIME=64,$__g_DllStructEx_VT_BLOB,$__g_DllStructEx_VT_STREAM,$__g_DllStructEx_VT_STORAGE,$__g_DllStructEx_VT_STREAMED_OBJECT, _
    $__g_DllStructEx_VT_STORED_OBJECT,$__g_DllStructEx_VT_BLOB_OBJECT,$__g_DllStructEx_VT_CF,$__g_DllStructEx_VT_CLSID,$__g_DllStructEx_VT_BSTR_BLOB=0xfff,$__g_DllStructEx_VT_VECTOR=0x1000, _
    $__g_DllStructEx_VT_ARRAY=0x2000,$__g_DllStructEx_VT_BYREF=0x4000,$__g_DllStructEx_VT_RESERVED=0x8000,$__g_DllStructEx_VT_ILLEGAL=0xffff,$__g_DllStructEx_VT_ILLEGALMASKED=0xfff, _
    $__g_DllStructEx_VT_TYPEMASK=0xfff

Global Const $__g_DllStructEx_tagObject = "int RefCount;int Size;ptr Object;ptr Methods[7];ptr szStruct;ptr szTranslatedStruct;ptr pStruct;boolean ownPStruct;int cElements;ptr pElements;ptr pParent;BYTE bUnion;"
Global Const $__g_DllStructEx_tagElement = "int iType;ptr szName;ptr szStruct;ptr szTranslatedStruct;int cElements;ptr pElements;"
#cs
# The size of $__g_DllStructEx_tagElement in bytes 
# @var int
#ce
Global Const $__g_DllStructEx_iElement = DllStructGetSize(DllStructCreate($__g_DllStructEx_tagElement))
Global Const $__g_DllStructEx_tagElements = "INT Index;INT Size;BYTE Elements[%d];"

Global Enum $__g_DllStructEx_eElementType_NONE, $__g_DllStructEx_eElementType_UNION, $__g_DllStructEx_eElementType_STRUCT, $__g_DllStructEx_eElementType_Element, $__g_DllStructEx_eElementType_PTR

Global Const $__g_DllStructEx_sStructRegex = "(?ix)(?(DEFINE)(?<hn>[\h\n])(?<struct_line_declaration>(?:(?&declaration)|(?&struct)|(?&union));)(?<union>union(?&hn)*{(?&hn)*(?:(?&struct_line_declaration)(?&hn)*)+}(?:\h+(?&identifier))?(?&hn)*)(?<struct>struct(?&hn)*{(?&hn)*(?:(?&struct_line_declaration)(?&hn)*)+}(?:\h+(?&identifier))?(?&hn)*)(?<declaration>(?&type)\h+(?&identifier))(?<type>\w+)(?<identifier>[*]*\w+))"
Global Const $__g_DllStructEx_sStructRegex_union = $__g_DllStructEx_sStructRegex&"^(?&union);$";TODO: should the semicolon be optional in the regex?
Global Const $__g_DllStructEx_sStructRegex_struct = $__g_DllStructEx_sStructRegex&"^(?&struct);$";TODO: should the semicolon be optional in the regex?
Global Const $__g_DllStructEx_sStructRegex_declaration = $__g_DllStructEx_sStructRegex&"^(?&declaration);$";TODO: should the semicolon be optional in the regex?

Global Const $__g_DllStructEx_sUnionRegex = "(?i)^\s*union\s*{(.*)}\s*(\w+)?;?$"
Global Const $__g_DllStructEx_sSubStructRegex = "(?i)^\s*struct\s*{(.*)}\s*(\w+)?;?$"
Global Const $__g_DllStructEx_sStructLineDeclaration = "^\s*(\w+)(?:\s+([*]*)(\w+))?;$"

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

    $tObject = DllStructCreate($__g_DllStructEx_tagObject, Ptr($oDllStructEx) - 8)
    $tObject.ownPStruct = $ownPStruct

    Return $oDllStructEx
EndFunc

#cs
# Queries a COM object for a pointer to one of its interface; identifying the interface by a reference to its interface identifier (IID). If the COM object implements the interface, then it returns a pointer to that interface after calling __DllStructEx_AddRef on it.
# @internal
#ce
Func __DllStructEx_QueryInterface($pSelf, $pRIID, $pObj)
    If $pObj=0 Then Return $__g_DllStructEx_E_POINTER
    Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
    If (Not ($sGUID=$__g_DllStructEx_IID_IDispatch)) And (Not ($sGUID=$__g_DllStructEx_IID_IUnknown)) Then Return $__g_DllStructEx_E_NOINTERFACE
    Local $tStruct = DllStructCreate("ptr", $pObj)
    DllStructSetData($tStruct, 1, $pSelf)
    __DllStructEx_AddRef($pSelf)
    Return $__g_DllStructEx_S_OK
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
        Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, $pSelf-8)
        If $tObject.pParent = 0 Then
            ;releases all properties
            _WinAPI_FreeMemory(DllStructGetData($tObject, "szStruct"))
            _WinAPI_FreeMemory(DllStructGetData($tObject, "szTranslatedStruct"))
            __DllStructEx_FreeElements($tObject)

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
    Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, $pSelf - 8)
    Local $i
    For $i = 0 To $tObject.cElements - 1 Step +1
        Local $tElement = DllStructCreate($__g_DllStructEx_tagElement, $tObject.pElements + $__g_DllStructEx_iElement * $i)
        ;ConsoleWrite(StringFormat("\t[%02d]: %s\n", $i, _WinAPI_GetString($tElement.szName)));NOTE: for debugging
        If __DllStructEx_szStringsEqual($pName, $tElement.szName) Then
            DllStructSetData($tDispId, "DispId", $i + 1)
            Return $__g_DllStructEx_S_OK
        EndIf
    Next

    ;ConsoleWrite("Not found: "&$sName&@CRLF);NOTE: for debugging

    DllStructSetData($tDispId, "DispId", -1)
    Return $__g_DllStructEx_S_OK
EndFunc

#cs
# Retrieves the type information for an object, which can then be used to get the type information for an interface.
# @internal
#ce
Func __DllStructEx_GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
    If $iTInfo<>0 Then Return $__g_DllStructEx_DISP_E_BADINDEX
    If $ppTInfo=0 Then Return $__g_DllStructEx_E_INVALIDARG
    Return $__g_DllStructEx_S_OK
EndFunc

#cs
# Retrieves the number of type information interfaces that an object provides (either 0 or 1).
# @internal
#ce
Func __DllStructEx_GetTypeInfoCount($pSelf, $pctinfo)
    DllStructSetData(DllStructCreate("UINT",$pctinfo), 1, 0)
    Return $__g_DllStructEx_S_OK
EndFunc

#cs
# Provides access to properties and methods exposed by an object.
# @internal
#ce
Func __DllStructEx_Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
    If $dispIdMember = -1 Then Return $__g_DllStructEx_DISP_E_MEMBERNOTFOUND

    #cs
    ;For debugging: getting vt of autoit types
    ;TODO: remove this
    $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
    If $tDISPPARAMS.cArgs<>1 Then Return $__g_DllStructEx_DISP_E_BADPARAMCOUNT
    Local $tVARIANT=DllStructCreate($__g_DllStructEx_tagVARIANT, $tDISPPARAMS.rgvargs)
    ConsoleWrite($tVARIANT.vt&@CRLF)
    Return $__g_DllStructEx_S_OK
    #ce

    If $dispIdMember = 0 Then
        If ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT)) Then Return $__g_DllStructEx_DISP_E_EXCEPTION
        Local $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
        If $tDISPPARAMS.cArgs>1 Then Return $__g_DllStructEx_DISP_E_BADPARAMCOUNT
        If $tDISPPARAMS.cArgs = 1 Then
            Local $tVARIANT=DllStructCreate($__g_DllStructEx_tagVARIANT, $tDISPPARAMS.rgvargs)
            If $tVARIANT.vt = $__g_DllStructEx_VT_BSTR Then
                Local $name = StringLower(_WinAPI_GetString($tVARIANT.data))
                Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, $pSelf - 8)
                Local $i
                Local $tElement
                For $i = 0 To $tObject.cElements - 1 Step +1
                    $tElement = DllStructCreate($__g_DllStructEx_tagElement, $tObject.pElements + $__g_DllStructEx_iElement * $i)
                    If $name = StringLower(_WinAPI_GetString($tElement.szName)) Then
                        $dispIdMember = $i + 1
                        ExitLoop
                    EndIf
                Next
                If $dispIdMember = 0 Then Return $__g_DllStructEx_DISP_E_MEMBERNOTFOUND
            ElseIf $tVARIANT.vt = $__g_DllStructEx_VT_I4 Or $tVARIANT.vt = $__g_DllStructEx_VT_I8 Then
                $dispIdMember = DllStructGetData(DllStructCreate($tVARIANT.vt = $__g_DllStructEx_VT_I4 ? "INT" : "INT64", DllStructGetPtr($tVARIANT, 'data')), 1)
                If $dispIdMember < 1 Then Return $__g_DllStructEx_DISP_E_MEMBERNOTFOUND
            Else
                Return $__g_DllStructEx_DISP_E_BADVARTYPE
            EndIf

            If $dispIdMember = -1 Then Return $__g_DllStructEx_DISP_E_MEMBERNOTFOUND
        Else
            ;VariantInit($pVariant) ;TODO: it seems unclear if the return variant should go through VariantInit before usage?
            Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, $pSelf-8)
            Local $tVARIANT = DllStructCreate($__g_DllStructEx_tagVARIANT, $pVarResult)
            $tVARIANT.vt = $__g_DllStructEx_VT_BSTR
            $tVARIANT.data = DllStructGetData($tObject, "szTranslatedStruct")

            Return $__g_DllStructEx_S_OK
        EndIf
    EndIf

    If $dispIdMember < -1 Then
        ;
    EndIf

    If $dispIdMember > 0 Then
        Return __DllStructEx_Invoke_ProcessElement($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
    EndIf

    Return $__g_DllStructEx_DISP_E_MEMBERNOTFOUND
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
    Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, $pSelf - 8)
    If $tObject.cElements >= $dispIdMember Then
        Local $tElement = DllStructCreate($__g_DllStructEx_tagElement, $tObject.pElements + $__g_DllStructEx_iElement * ($dispIdMember - 1))
        If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
        Local $tVARIANT = DllStructCreate($__g_DllStructEx_tagVARIANT, $pVarResult)
        If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
        If $tObject.bUnion = 1 And ($tElement.iType = $__g_DllStructEx_eElementType_UNION Or $tElement.iType = $__g_DllStructEx_eElementType_Element) Then;FIXME: test with ptr
            Local $tStruct = DllStructCreate(_WinAPI_GetString($tElement.szTranslatedStruct, True)&" "&_WinAPI_GetString($tElement.szName, True), $tObject.pStruct);TODO: currently the name is missing from the element type struct. as a result is hacky way solves it for now.
        Else
            Local $tStruct = DllStructCreate(_WinAPI_GetString($tObject.szTranslatedStruct, True), $tObject.pStruct)
        EndIf
        If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
        Switch $tElement.iType
            Case $__g_DllStructEx_eElementType_Element
                If ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYGET)=$__g_DllStructEx_DISPATCH_PROPERTYGET)) Then
                    Local $vData = DllStructGetData($tStruct, _WinAPI_GetString($tElement.szName, True));TODO: add support for getting struct data slice by index
                    __DllStructEx_DataToVariant($vData, $tVARIANT)
                    If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
                ElseIf ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT)) Then
                    Local $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
                    If $tDISPPARAMS.cArgs<1 Then Return $__g_DllStructEx_DISP_E_BADPARAMCOUNT
                    Local $_tVARIANT=DllStructCreate($__g_DllStructEx_tagVARIANT, $tDISPPARAMS.rgvargs)
                    Local $vData = __DllStructEx_VariantToData($_tVARIANT)
                    If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
                    DllStructSetData($tStruct, _WinAPI_GetString($tElement.szName, True), $vData)
                Else
                    Return $__g_DllStructEx_DISP_E_EXCEPTION
                EndIf
            Case $__g_DllStructEx_eElementType_STRUCT
                If ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT)) Then Return $__g_DllStructEx_DISP_E_EXCEPTION
                Local $tElements = DllStructCreate(StringFormat($__g_DllStructEx_tagElements, 1)); WARNING: Elements property in this struct is "BYTE[1]", not intended for use!
                If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
                $tElements.Index = $tElement.cElements
                Local $vData = __DllStructEx_Create($tElement.szStruct, $tElement.szTranslatedStruct, $tElements, DllStructGetPtr($tStruct, _WinAPI_GetString($tElement.szName)), $tElement.pElements)
                Local $_tObject = DllStructCreate($__g_DllStructEx_tagObject, Ptr($vData)-8)
                $_tObject.pParent = $pSelf
                __DllStructEx_AddRef($pSelf)
                __DllStructEx_AddRef(Ptr($vData)) ; add ref, as we pass it into a variant
                __DllStructEx_DataToVariant($vData, $tVARIANT)
                If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
            Case $__g_DllStructEx_eElementType_UNION
                If ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT)) Then Return $__g_DllStructEx_DISP_E_EXCEPTION
                Local $tElements = DllStructCreate(StringFormat($__g_DllStructEx_tagElements, 1)); WARNING: Elements property in this struct is "BYTE[1]", not intended for use!
                If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
                $tElements.Index = $tElement.cElements
                Local $vData = __DllStructEx_Create($tElement.szStruct, $tElement.szTranslatedStruct, $tElements, DllStructGetPtr($tStruct, _WinAPI_GetString($tElement.szName)), $tElement.pElements)
                Local $_tObject = DllStructCreate($__g_DllStructEx_tagObject, Ptr($vData)-8)
                $_tObject.pParent = $pSelf
                $_tObject.bUnion = 1
                __DllStructEx_AddRef($pSelf)
                __DllStructEx_AddRef(Ptr($vData)) ; add ref, as we pass it into a variant
                __DllStructEx_DataToVariant($vData, $tVARIANT)
                If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
            Case $__g_DllStructEx_eElementType_PTR
                Local $iPtrLevelCount = $tElement.cElements
                Local $sType = _WinAPI_GetString($tElement.szStruct)
                Local $pLevel = DllStructGetData($tStruct, _WinAPI_GetString($tElement.szName))
                Local $ppLevel = DllStructGetPtr($tStruct, _WinAPI_GetString($tElement.szName))
                Local $iLevel = 1
                For $iLevel = 1 To $iPtrLevelCount - 1 Step +1 ;This loop handles all but final ptr.
                    If 0 = $pLevel Then
                        __DllStructEx_Error("ptr ref empty")
                        Return $__g_DllStructEx_DISP_E_EXCEPTION;TODO: should we return null instead of throwing an exception, or maybe do both?
                    EndIf
                    $ppLevel = $pLevel
                    $pLevel = DllStructGetData(DllStructCreate("PTR", $pLevel), 1)
                Next

                If 0 = $pLevel And Not ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT) Or (BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUTREF)=$__g_DllStructEx_DISPATCH_PROPERTYPUTREF)) Then
                    __DllStructEx_Error("ptr ref empty")
                    Return $__g_DllStructEx_DISP_E_EXCEPTION;TODO: should we return null instead of throwing an exception, or maybe do both?
                EndIf

                Switch $sType
                    Case 'IUnknown'
                        If ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT) Or (BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUTREF)=$__g_DllStructEx_DISPATCH_PROPERTYPUTREF)) Then
                            If Not ($pLevel = 0) Then
                                ;we will release previous value ref
                                Local $pRelease = DllStructGetData(DllStructCreate("PTR", DllStructGetData(DllStructCreate("PTR", $pLevel), 1) + (@AutoItX64 ? 8 : 4) * 2), 1)
                                DllCallAddress("dword", $pRelease, "PTR", $pLevel); release the iunknown object
                            EndIf

                            Local $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
                            If $tDISPPARAMS.cArgs<1 Then Return $__g_DllStructEx_DISP_E_BADPARAMCOUNT
                            Local $_tVARIANT=DllStructCreate($__g_DllStructEx_tagVARIANT, $tDISPPARAMS.rgvargs)
                            ;Local $vData = __DllStructEx_VariantToData($_tVARIANT)
                            ;If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
                            ;DllStructSetData($tStruct, _WinAPI_GetString($tElement.szName, True), $vData)
                            $pLevel = $_tVARIANT.data
                            DllStructSetData(DllStructCreate("PTR", $ppLevel), 1, $pLevel)
                        EndIf
                        ;If ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT)) Then Return $__g_DllStructEx_DISP_E_EXCEPTION; TODO: remove this, when propertyput is supported!
                        $tVariant.vt = $__g_DllStructEx_VT_UNKNOWN
                        $tVariant.data = $pLevel
                        Local $pAddRef = DllStructGetData(DllStructCreate("PTR", DllStructGetData(DllStructCreate("PTR", $pLevel), 1) + (@AutoItX64 ? 8 : 4) * 1), 1)
                        DllCallAddress("dword", $pAddRef, "PTR", $pLevel); AddRef the iunknown object we got back
                        Return $__g_DllStructEx_S_OK
                    Case 'IDispatch'
                        If ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT) Or (BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUTREF)=$__g_DllStructEx_DISPATCH_PROPERTYPUTREF)) Then
                            If Not ($pLevel = 0) Then
                                ;we will release previous value ref
                                Local $pRelease = DllStructGetData(DllStructCreate("PTR", DllStructGetData(DllStructCreate("PTR", $pLevel), 1) + (@AutoItX64 ? 8 : 4) * 2), 1)
                                DllCallAddress("dword", $pRelease, "PTR", $pLevel); release the idispatch object
                            EndIf

                            Local $tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
                            If $tDISPPARAMS.cArgs<1 Then Return $__g_DllStructEx_DISP_E_BADPARAMCOUNT
                            Local $_tVARIANT=DllStructCreate($__g_DllStructEx_tagVARIANT, $tDISPPARAMS.rgvargs)
                            ;Local $vData = __DllStructEx_VariantToData($_tVARIANT)
                            ;If @error <> 0 Then Return $__g_DllStructEx_DISP_E_EXCEPTION
                            ;DllStructSetData($tStruct, _WinAPI_GetString($tElement.szName, True), $vData)
                            $pLevel = $_tVARIANT.data
                            DllStructSetData(DllStructCreate("PTR", $ppLevel), 1, $pLevel)
                        Else
                            $tVariant.vt = $__g_DllStructEx_VT_DISPATCH
                            $tVariant.data = $pLevel
                        EndIf

                        ;If ((BitAND($wFlags, $__g_DllStructEx_DISPATCH_PROPERTYPUT)=$__g_DllStructEx_DISPATCH_PROPERTYPUT)) Then Return $__g_DllStructEx_DISP_E_EXCEPTION; TODO: remove this, when propertyput is supported!
                        If Not ($pLevel = 0) Then
                            Local $pAddRef = DllStructGetData(DllStructCreate("PTR", DllStructGetData(DllStructCreate("PTR", $pLevel), 1) + (@AutoItX64 ? 8 : 4) * 1), 1)
                            DllCallAddress("dword", $pAddRef, "PTR", $pLevel); AddRef the idispatch object we got back
                        EndIf
                        Return $__g_DllStructEx_S_OK
                    Case Else
                        Local $tElements = DllStructCreate(StringFormat($__g_DllStructEx_tagElements, $__g_DllStructEx_iElement))
                        $tElements.iSize = 1; NOTE: maybe this line is not needed. Look into it, for optimization.

                        Local $sTranslatedType = __DllStructEx_ParseStructType($sType, $tElements)
                        
                        $_tObject = DllStructCreate($__g_DllStructEx_tagObject)
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
                Return $__g_DllStructEx_DISP_E_EXCEPTION
        EndSwitch
        Return $__g_DllStructEx_S_OK
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
    Local $tObject = DllStructCreate($__g_DllStructEx_tagObject)

    $tObject.cElements = $tElements.Index
    Local $sElements = StringFormat("BYTE[%d]", $__g_DllStructEx_iElement * $tObject.cElements)
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

    $tObject = __DllStructEx_DllStructAlloc($__g_DllStructEx_tagObject, $tObject)

    DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
    Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $__g_DllStructEx_IID_IDispatch, Default, True) ; pointer that's wrapped into object
EndFunc

#cs
# Convert AutoIt data to variant.
# @internal
# @param mixed          $vData
# @param DllStruct|null $tVariant variant
# @return DllStruct variant
#ce
Func __DllStructEx_DataToVariant($vData, $tVARIANT = Null)
    If Not IsDllStruct($tVARIANT) Then $tVARIANT = DllStructCreate($__g_DllStructEx_tagVARIANT)
    #cs
        $__g_DllStructEx_VT_I4 = 3 => int32
        $__g_DllStructEx_VT_I8 = 20 => int64
        $__g_DllStructEx_VT_R8 = 5 => Double
        $__g_DllStructEx_VT_BSTR = 8 => String
        $__g_DllStructEx_VT_ARRAY + $__g_DllStructEx_VT_BOOL = 8209 (0x2011) => Binary
        $__g_DllStructEx_VT_UI4 = 19 => Pointer
    #ce

    Switch VarGetType($vData)
        Case 'Int32'
            $tVARIANT.vt = $__g_DllStructEx_VT_I4
            Local $tData = DllStructCreate("INT data;", DllStructGetPtr($tVARIANT, 'data'))
            $tVARIANT.data = $vData
        Case 'Int64'
            $tVARIANT.vt = $__g_DllStructEx_VT_I8
            Local $tData = DllStructCreate("INT64 data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        Case 'Double'
            $tVARIANT.vt = $__g_DllStructEx_VT_R8
            Local $tData = DllStructCreate("DOUBLE data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        Case 'String'
            $tVARIANT.vt = $__g_DllStructEx_VT_BSTR
            Local $tData = DllStructCreate("PTR data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = __DllStructEx_CreateString($vData) ;TODO: should use SysAllocString instead.
        #cs
        Case 'Binary'
            $tVARIANT.vt = BitOR($__g_DllStructEx_VT_ARRAY, $__g_DllStructEx_VT_BOOL)
            ;FIXME: support Binary
        #ce
        Case 'Ptr'
            $tVARIANT.vt = @AutoItX64 ? $__g_DllStructEx_VT_UI8 : $__g_DllStructEx_VT_UI4
            Local $tData = DllStructCreate("PTR data;", DllStructGetPtr($tVARIANT, 'data'))
            $tData.data = $vData
        Case 'Object'
            Local $tGUID = DllStructCreate("byte[16]")
            Local $pRIID = DllCall("ole32.dll", "LONG", "CLSIDFromString", "wstr", $__g_DllStructEx_IID_IDispatch, "struct*", $tGUID)
            Local $pQueryInterface = DllStructGetData(DllStructCreate("PTR", DllStructGetData(DllStructCreate("PTR", Ptr($vData)), 1)), 1)
            Local $pObj = DllStructCreate("ptr")
            Local $response = DllCallAddress("LONG", $pQueryInterface, "ptr", Ptr($vData), "struct*", $tGUID, "ptr", DllStructGetPtr($pObj))
            If @error <> 0 Then Return SetError(2, 0, $tVARIANT)
            If $response[0] = $__g_DllStructEx_S_OK Then
                Local $pRelease = DllStructGetData(DllStructCreate("PTR", DllStructGetData(DllStructCreate("PTR", DllStructGetData($pObj, 1)), 1) + (@AutoItX64 ? 8 : 4) * 2), 1)
                DllCallAddress("dword", $pRelease, "PTR", DllStructGetData($pObj, 1)); release the idispatch object we got back
            EndIf
            $tVARIANT.vt = $response[0] = $__g_DllStructEx_S_OK ? $__g_DllStructEx_VT_DISPATCH : $__g_DllStructEx_VT_UNKNOWN
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
        Case $__g_DllStructEx_VT_BSTR, $__g_DllStructEx_VT_LPSTR, $__g_DllStructEx_VT_LPWSTR
            Return _WinAPI_GetString($tVARIANT.data, Not ($tVARIANT.vt = $__g_DllStructEx_VT_LPSTR))
        Case $__g_DllStructEx_VT_I2
            $sType = "SHORT"
        Case $__g_DllStructEx_VT_I4, $__g_DllStructEx_VT_INT
            $sType = "INT"
        Case $__g_DllStructEx_VT_I8
            $sType = "INT64"
        Case $__g_DllStructEx_VT_NULL
            Return Null
        Case $__g_DllStructEx_VT_PTR
            Return $tVARIANT.data
        Case $__g_DllStructEx_VT_R4
            $sType = "FLOAT"
        Case $__g_DllStructEx_VT_R8
            $sType = "DOUBLE"
        Case $__g_DllStructEx_VT_UI4, $__g_DllStructEx_VT_UI8
            Return Ptr($tVARIANT.data)
        Case $__g_DllStructEx_VT_UINT
            $sType = "UINT"
        ;Case $__g_DllStructEx_VT_DISPATCH
        ;Case $__g_DllStructEx_VT_UNKNOWN
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
    If Not StringRegExp($sStruct, $__g_DllStructEx_sStructRegex&"^(?&struct_line_declaration)+$", 0) Then Return SetError(__DllStructEx_Error("Regex validation for the entire struct provided failed!", 3), @error, "")
    Local $aStructLineDeclarations = StringRegExp($sStruct, $__g_DllStructEx_sStructRegex&"(?&struct_line_declaration)", 3)
    If @error <> 0 Then Return SetError(1, @error, "")
    Local $tElements = DllStructCreate(StringFormat($__g_DllStructEx_tagElements, $__g_DllStructEx_iElement * UBound($aStructLineDeclarations, 1)))

    $sStruct = ""
    Local $i
    Local $sStructLineDeclaration
    For $i = 0 To UBound($aStructLineDeclarations) - 1 Step +1
        $sStructLineDeclaration = $aStructLineDeclarations[$i]
        Select
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_union, $STR_REGEXPMATCH)
                Local $aUnion = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sUnionRegex, $STR_REGEXPARRAYMATCH)
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for union-declaration failed", 2), @error, "")
                $sStruct &= __DllStructEx_ParseUnion($aUnion, $tElements)
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_struct, $STR_REGEXPMATCH)
                Local $_aStruct = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sSubStructRegex, $STR_REGEXPARRAYMATCH)
                $sStruct &= __DllStructEx_ParseNestedStruct($_aStruct, $tElements)
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_declaration, $STR_REGEXPMATCH)
                Local $aMatches = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructLineDeclaration, $STR_REGEXPARRAYMATCH)
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
    Local $tElement = IsDllStruct($tElements) ? DllStructCreate($__g_DllStructEx_tagElement, DllStructGetPtr($tElements, "Elements") + $__g_DllStructEx_iElement * $tElements.Index) : DllStructCreate($__g_DllStructEx_tagElement)
    ; The size of type in bytes
    Local $iSize = Null
    Local $sTranslatedType = $sType

    $tElement.iType = $__g_DllStructEx_eElementType_Element

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
            $tElement.iType = $__g_DllStructEx_eElementType_NONE
            $iSize = 0
        Case Else
            ;TODO: maybe we should assume type element, if the solved variable results in single native struct element type?
            $tElement.iType = $__g_DllStructEx_eElementType_STRUCT
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
                Local $sElements = StringFormat("BYTE[%d]", $__g_DllStructEx_iElement * $tElement.cElements)
                $tElement.pElements = DllStructGetPtr(__DllStructEx_DllStructAlloc($sElements, DllStructCreate($sElements, DllStructGetPtr($_tElements, "Elements"))))
                ;If $tElement.cElements > 0 Then ConsoleWrite("> "&_WinAPI_GetString(DllStructCreate($__g_DllStructEx_tagElement, DllStructGetPtr($_tElements, "Elements")).szName)&@CRLF)
                ;$tElement.pElements = DllStructGetPtr(__DllStructEx_DllStructAlloc($sTranslatedType, DllStructCreate($sTranslatedType, DllStructGetPtr($_tElements, "Elements"))))
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

    Local $tElement = DllStructCreate($__g_DllStructEx_tagElement, DllStructGetPtr($tElements, "Elements") + $__g_DllStructEx_iElement * $tElements.Index)
    
    If Not ("" = $sPtr) Then
        $tElement.iType = $__g_DllStructEx_eElementType_PTR
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

    $tUnion = DllStructCreate($__g_DllStructEx_tagElement, DllStructGetPtr($tUnions, "Elements") + $__g_DllStructEx_iElement * $tUnions.Index)
    $tUnion.iType = $__g_DllStructEx_eElementType_UNION
    $tUnion.szName = __DllStructEx_CreateString($sName)
    $tUnion.szStruct = __DllStructEx_CreateString($sUnionStruct)
    If @error <> 0 Then Return SetError(@error, @extended, "")
    $tUnions.Index += 1

    Local $aStructLineDeclarations = StringRegExp($sUnionStruct, $__g_DllStructEx_sStructRegex&"(?&struct_line_declaration)", 3)
    If @error <> 0 Then Return SetError(1, @error, "")
    Local $tElements = DllStructCreate(StringFormat($__g_DllStructEx_tagElements, $__g_DllStructEx_iElement * UBound($aStructLineDeclarations, 1)))

    $sUnionStruct = ""
    Local $i
    Local $sStructLineDeclaration
    For $i = 0 To UBound($aStructLineDeclarations) - 1 Step +1
        $sStructLineDeclaration = $aStructLineDeclarations[$i]
        Select
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_union, $STR_REGEXPMATCH)
                Local $_aUnion = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sUnionRegex, $STR_REGEXPARRAYMATCH)
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for union-declaration failed", 2), @error, "")
                $sUnionStruct &= __DllStructEx_ParseUnion($_aUnion, $tElements)
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_struct, $STR_REGEXPMATCH)
                Local $aStruct = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sSubStructRegex, $STR_REGEXPARRAYMATCH)
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for struct-declaration failed", 2), @error, "")
                $sUnionStruct &= __DllStructEx_ParseNestedStruct($aStruct, $tElements)
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_declaration, $STR_REGEXPMATCH)
                Local $aMatches = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructLineDeclaration, $STR_REGEXPARRAYMATCH)
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for struct-line-declaration failed", 2), @error, "")
                $sUnionStruct &= __DllStructEx_ParseStructTypeCallback($aMatches, $tElements)
            Case Else
                ConsoleWriteError("a struct-line-declaration could not be recognized!")
        EndSelect
    Next

    $tUnion.szTranslatedStruct = __DllStructEx_CreateString($sUnionStruct)
    $tUnion.cElements = $tElements.Index
    Local $sElements = StringFormat("BYTE[%d]", $__g_DllStructEx_iElement * $tUnion.cElements)
    $tUnion.pElements = DllStructGetPtr(__DllStructEx_DllStructAlloc($sElements, DllStructCreate($sElements, DllStructGetPtr($tElements, "Elements"))))

    Local $iBytes = $tElements.Size
    $tUnions.Size = $tUnions.Size < $iBytes ? $iBytes : $tUnions.Size
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

    $tStruct = DllStructCreate($__g_DllStructEx_tagElement, DllStructGetPtr($tStructs, "Elements") + $__g_DllStructEx_iElement * $tStructs.Index)
    $tStruct.iType = $__g_DllStructEx_eElementType_STRUCT
    $tStruct.szName = __DllStructEx_CreateString($sName)
    $tStruct.szStruct = __DllStructEx_CreateString($sStruct)
    If @error <> 0 Then Return SetError(@error, @extended, "")
    $tStructs.Index += 1

    Local $aStructLineDeclarations = StringRegExp($sStruct, $__g_DllStructEx_sStructRegex&"(?&struct_line_declaration)", 3)
    If @error <> 0 Then Return SetError(1, @error, "")
    Local $tElements = DllStructCreate(StringFormat($__g_DllStructEx_tagElements, $__g_DllStructEx_iElement * UBound($aStructLineDeclarations, 1)))

    $sStruct = ""
    Local $i
    Local $sStructLineDeclaration
    For $i = 0 To UBound($aStructLineDeclarations) - 1 Step +1
        $sStructLineDeclaration = $aStructLineDeclarations[$i]
        Select
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_union, $STR_REGEXPMATCH)
                Local $aUnion = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sUnionRegex, $STR_REGEXPARRAYMATCH)
                $sStruct &= __DllStructEx_ParseUnion($aUnion, $tElements)
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_struct, $STR_REGEXPMATCH)
                Local $_aStruct = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sSubStructRegex, $STR_REGEXPARRAYMATCH)
                $sStruct &= __DllStructEx_ParseNestedStruct($_aStruct, $tElements)
            Case StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructRegex_declaration, $STR_REGEXPMATCH)
                Local $aMatches = StringRegExp($sStructLineDeclaration, $__g_DllStructEx_sStructLineDeclaration, $STR_REGEXPARRAYMATCH)
                If @error <> 0 Then Return SetError(__DllStructEx_Error("Regex for struct-line-declaration failed", 2), @error, "")
                $sStruct &= __DllStructEx_ParseStructTypeCallback($aMatches, $tElements)
            Case Else
                ConsoleWriteError("a struct-line-declaration could not be recognized!")
        EndSelect
    Next

    $tStruct.szTranslatedStruct = __DllStructEx_CreateString($sStruct)
    $tStruct.cElements = $tElements.Index
    Local $sElements = StringFormat("BYTE[%d]", $__g_DllStructEx_iElement * $tStruct.cElements)
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
    Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, $pDllStructEx - 8)
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
    Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, $pDllStructEx - 8)
    Return _WinAPI_GetString(DllStructGetData($tObject, "szStruct"), True)
EndFunc

#cs
# Get the AutoIt dllstruct compatible string transpiled from the string used when the DllStruct was created.
# @param DllStructEx $oDllStructEx A DllStructEx Object
# @return string
#ce
Func DllStructExGetTranspiledStructString($oDllStructEx)
    Local $pDllStructEx = Ptr($oDllStructEx)
    Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, $pDllStructEx - 8)
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

#cs
# Frees the elements tree from an object
# @param DllStruct $tObject
#ce
Func __DllStructEx_FreeElements($tObject, $iLevel = 1)
    ;If $iLevel = 1 Then ConsoleWrite("__DllStructEx_FreeElements"&@CRLF&_WinAPI_GetString($tObject.szTranslatedStruct, True)&@CRLF)
    Local $pElements = $tObject.pElements
    Local $cElements = $tObject.cElements
    If $cElements = 0 Then Return
    If $pElements = 0 Then Return SetError(__DllStructEx_Error("pElements was null, but cElements is non zero!", 1))
    Local $tElement
    Local $i
    For $i = 0 To $cElements - 1 Step +1
        $tElement = DllStructCreate($__g_DllStructEx_tagElement, $pElements + $__g_DllStructEx_iElement * $i)
        Switch ($tElement.iType)
            Case $__g_DllStructEx_eElementType_UNION
                ;ConsoleWrite(StringFormat("[UNION]: %s\n", _WinAPI_GetString(DllStructGetData($tElement, "szName"))))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szName"))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szStruct"))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szTranslatedStruct"))
                Local $tUnionObject = DllStructCreate($__g_DllStructEx_tagObject)
                    $tUnionObject.cElements = $tElement.cElements
                    $tUnionObject.pElements = $tElement.pElements
                __DllStructEx_FreeElements($tUnionObject, $iLevel + 1)
                If @error <> 0 Then Return SetError(@error, @extended)
            Case $__g_DllStructEx_eElementType_STRUCT
                ;ConsoleWrite(StringFormat("[STRUCT]: %s\n", _WinAPI_GetString(DllStructGetData($tElement, "szName"))))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szName"))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szStruct"))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szTranslatedStruct"))
                Local $tStructObject = DllStructCreate($__g_DllStructEx_tagObject)
                    $tStructObject.cElements = $tElement.cElements
                    $tStructObject.pElements = $tElement.pElements
                __DllStructEx_FreeElements($tStructObject, $iLevel + 1)
                If @error <> 0 Then Return SetError(@error, @extended)
            Case $__g_DllStructEx_eElementType_PTR
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szName"))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szStruct"))
            Case $__g_DllStructEx_eElementType_Element
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szName"))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szStruct"))
                _WinAPI_FreeMemory(DllStructGetData($tElement, "szTranslatedStruct"))
            Case $__g_DllStructEx_eElementType_NONE
                ;
            Case Else
                ;ConsoleWrite(StringFormat("Level: %s\n", $iLevel))
                ;ConsoleWrite(StringFormat("%s\n", _WinAPI_GetString($tElement.szStruct)))
                ;ConsoleWrite(StringFormat("cElements: %s\n", $cElements))
                ;ConsoleWrite(StringFormat("element index: %s\n", $i))
                Return SetError(__DllStructEx_Error(StringFormat('Unknown Element type: "%s"', $tElement.iType), 1))
        EndSwitch
        __DllStructEx_DllStructFree($tObject.pElements)
    Next
EndFunc

#cs
# Get the pointer to the struct or element in the struct.
# @param DllStructEx $oDllStructEx A DllStructEx Object
# @param int|string  $vElement     The element of the struct whose pointer you need, starting at 1 or the element name, as defined when the struct was created.
# @return ptr
#ce
Func DllStructExGetPtr($oDllStructEx, $vElement = Null)
    Local $tDllStructEx = DllStructExGetStruct($oDllStructEx)
    If @error <> 0 Then Return SetError(3, 0, 0)
    Local $ptr = (Not (Null = $vElement)) ? DllStructGetPtr($tDllStructEx, $vElement) : DllStructGetPtr($tDllStructEx)
    Return SetError(@error, @extended, $ptr)
EndFunc

;FIXME: add new helper function for overriding the pStruct value, if pOwnPStruct is false. (this will make it possible to re-use an DllStructEx object instead of discarding and re-creating for one small change.)
