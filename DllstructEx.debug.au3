#include-once

#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include "DllStructEx.au3"

Global $__g_DllStructEx_DllStructExDisplay_treeviewID
Global $__g_DllStructEx_DllStructExDisplay_hTreeView
Global $__g_DllStructEx_DllStructExDisplay_iName
Global $__g_DllStructEx_DllStructExDisplay_iType
Global $__g_DllStructEx_DllStructExDisplay_iSize
Global $__g_DllStructEx_DllStructExDisplay_iStructure
Global $__g_DllStructEx_DllStructExDisplay_iDllStructure

#cs
# Displays a DllStructEx information as a TreeView
# @param DllStructEx $oDllStructEx A DllStructEx Object
#ce
Func DllStructExDisplay($oDllStructEx)
    Local $tObject = DllStructCreate($__g_DllStructEx_tagObject, Ptr($oDllStructEx)-8)

    Local $hGUI = GUICreate("DllStructExDisplay", 700, 500)

    GUICtrlCreateLabel("Name:", 410, 13)
    $__g_DllStructEx_DllStructExDisplay_iName = GUICtrlCreateInput("", 450, 10, 240, -1, $ES_READONLY)
    GUICtrlCreateLabel("Type:", 410, 38)
    $__g_DllStructEx_DllStructExDisplay_iType = GUICtrlCreateInput("", 450, 35, 240, -1, $ES_READONLY)
    GUICtrlCreateLabel("Size:", 410, 63)
    $__g_DllStructEx_DllStructExDisplay_iSize = GUICtrlCreateInput("", 450, 60, 240, -1, $ES_READONLY)
    GUICtrlCreateLabel("Structure:", 410, 95)
    $__g_DllStructEx_DllStructExDisplay_iStructure = GUICtrlCreateEdit("", 410, 115, 280, 170, $ES_READONLY)
    GUICtrlCreateLabel("DllStructure:", 410, 305)
    $__g_DllStructEx_DllStructExDisplay_iDllStructure = GUICtrlCreateEdit("", 410, 325, 280, 170, $ES_READONLY)

    $__g_DllStructEx_DllStructExDisplay_treeviewID = GUICtrlCreateTreeView(0, 0, 400, 500)
    $__g_DllStructEx_DllStructExDisplay_hTreeView = GUICtrlGetHandle($__g_DllStructEx_DllStructExDisplay_treeviewID)
    
    _GUICtrlTreeView_BeginUpdate($__g_DllStructEx_DllStructExDisplay_treeviewID)
    Local $treeviewItemID = GUICtrlCreateTreeViewItem("[STRUCT]", $__g_DllStructEx_DllStructExDisplay_treeviewID)
    Local $tElement = DllStructCreate($__g_DllStructEx_tagElement)
        DllStructSetData($tElement, "iType", $__g_DllStructEx_eElementType_STRUCT)
        DllStructSetData($tElement, "szName", 0)
        DllStructSetData($tElement, "szStruct", $tObject.szStruct)
        DllStructSetData($tElement, "szTranslatedStruct", $tObject.szTranslatedStruct)
        DllStructSetData($tElement, "cElements", $tObject.cElements)
        DllStructSetData($tElement, "pElements", $tObject.pElements)
    Local $_tObject = DllStructCreate($__g_DllStructEx_tagObject)
        DllStructSetData($_tObject, "cElements", 1)
        DllStructSetData($_tObject, "pElements", DllStructGetPtr($tElement))
    _GUICtrlTreeView_SetItemParam($__g_DllStructEx_DllStructExDisplay_hTreeView, GUICtrlGetHandle($treeviewItemID), DllStructGetPtr($tElement))
    __DllStructEx_DisplayPopulate($treeviewItemID, $tObject)
    _GUICtrlTreeView_EndUpdate($__g_DllStructEx_DllStructExDisplay_treeviewID)

    ; Switch to GetMessage mode
    Local $iOnEventMode = Opt("GUIOnEventMode", 0), $iMsg

    GUIRegisterMsg($WM_NOTIFY, "__DllStructEx_DllStructExDisplay_WM_NOTIFY")

    GUISetState(@SW_SHOW, $hGUI)

    While 1
        $iMsg = GUIGetMsg() ; Variable needed to check which "Copy" button was pressed
        Switch $iMsg
            case $GUI_EVENT_CLOSE
                ExitLoop
        EndSwitch
    WEnd

    GUIDelete($hGUI)
    Opt("GUIOnEventMode", $iOnEventMode)
EndFunc

#cs
# Populate a treeview identifier with fields information from the provided DllstructEx object Structure
# @internal
# @param int $treeviewID treeview identifier
# @param DllStruct $tObject DllStructEx Object structure ($__g_DllStructEx_tagObject)
#ce
Func __DllStructEx_DisplayPopulate($treeviewID, $tObject)
    Local $pElements = $tObject.pElements
    Local $cElements = $tObject.cElements

    If $cElements = 0 Then Return

    Local $treeviewItemID
    Local $tElement
    Local $i
    For $i = 0 To $cElements - 1 Step +1
        $tElement = DllStructCreate($__g_DllStructEx_tagElement, $pElements + $__g_DllStructEx_iElement * $i)
        Switch ($tElement.iType)
            Case $__g_DllStructEx_eElementType_UNION
                $treeviewItemID = GUICtrlCreateTreeViewItem(_WinAPI_GetString(DllStructGetData($tElement, "szName")), $treeviewID)
                Local $tUnionObject = DllStructCreate($__g_DllStructEx_tagObject)
                    $tUnionObject.cElements = $tElement.cElements
                    $tUnionObject.pElements = $tElement.pElements
                __DllStructEx_DisplayPopulate($treeviewItemID, $tUnionObject)
            Case $__g_DllStructEx_eElementType_STRUCT
                $treeviewItemID = GUICtrlCreateTreeViewItem(_WinAPI_GetString(DllStructGetData($tElement, "szName")), $treeviewID)
                Local $tStructObject = DllStructCreate($__g_DllStructEx_tagObject)
                    $tStructObject.cElements = $tElement.cElements
                    $tStructObject.pElements = $tElement.pElements
                __DllStructEx_DisplayPopulate($treeviewItemID, $tStructObject)
            Case $__g_DllStructEx_eElementType_PTR
                $treeviewItemID = GUICtrlCreateTreeViewItem(_WinAPI_GetString(DllStructGetData($tElement, "szName")), $treeviewID)
            Case $__g_DllStructEx_eElementType_Element
                $treeviewItemID = GUICtrlCreateTreeViewItem(_WinAPI_GetString(DllStructGetData($tElement, "szName")), $treeviewID)
            Case $__g_DllStructEx_eElementType_NONE
                $treeviewItemID = GUICtrlCreateTreeViewItem("[NONE]", $treeviewID)
            Case Else
                $treeviewItemID = GUICtrlCreateTreeViewItem("[!UNKNOWN]", $treeviewID)
        EndSwitch
        ;_GUICtrlTreeView_GetItemParam
        ;_GUICtrlTreeView_SetItemParam($__g_DllStructEx_DllStructExDisplay_hTreeView, $treeviewItemID, DllStructGetPtr($tElement))
        _GUICtrlTreeView_SetItemParam($__g_DllStructEx_DllStructExDisplay_hTreeView, GUICtrlGetHandle($treeviewItemID), DllStructGetPtr($tElement))
    Next
EndFunc

Func __DllStructEx_DllStructExDisplay_WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
    ; Most of this code is from a post by Melba23: https://www.autoitscript.com/forum/topic/186975-action-on-selection-in-treeview/?do=findComment&comment=1342750
    ; Create NMTREEVIEW structure
    Local $tStruct = DllStructCreate("struct;hwnd hWndFrom;uint_ptr IDFrom;INT Code;endstruct;" & _
            "uint Action;struct;uint OldMask;handle OldhItem;uint OldState;uint OldStateMask;" & _
            "ptr OldText;int OldTextMax;int OldImage;int OldSelectedImage;int OldChildren;lparam OldParam;endstruct;" & _
            "struct;uint NewMask;handle NewhItem;uint NewState;uint NewStateMask;" & _
            "ptr NewText;int NewTextMax;int NewImage;int NewSelectedImage;int NewChildren;lparam NewParam;endstruct;" & _
            "struct;long PointX;long PointY;endstruct", $lParam)
    ;ConsoleWrite(DllStructGetData($tStruct, "hWndFrom")&@CRLF)
    ;ConsoleWrite(VarGetType(DllStructGetData($tStruct, "hWndFrom"))&" "&DllStructGetData($tStruct, "hWndFrom")&" "&VarGetType($__g_DllStructEx_DllStructExDisplay_hTreeView)&" "&$__g_DllStructEx_DllStructExDisplay_hTreeView&@CRLF)
    If DllStructGetData($tStruct, "hWndFrom") = $__g_DllStructEx_DllStructExDisplay_hTreeView Then
        Switch DllStructGetData($tStruct, "Code")
            ; If item selection changed
            Case $TVN_SELCHANGEDA, $TVN_SELCHANGEDW
                Local $hItem = DllStructGetData($tStruct, "NewhItem")
                ; Set flag to selected item handle
                ;If $hItem Then $hItemSelected = $hItem
                If $hItem Then
                    Local $pElement = _GUICtrlTreeView_GetItemParam($__g_DllStructEx_DllStructExDisplay_hTreeView, $hItem)
                    Local $tElement = DllStructCreate($__g_DllStructEx_tagElement, $pElement)
                    ;ConsoleWrite($pElement&@CRLF);
                    GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iName, _WinAPI_GetString($tElement.szName))

                    Local $sType = "Unknown"
                    Switch ($tElement.iType)
                        Case $__g_DllStructEx_eElementType_Element
                            $sType = "Element"
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iStructure, _WinAPI_GetString($tElement.szStruct))
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iDllStructure, _WinAPI_GetString($tElement.szTranslatedStruct))
                        Case $__g_DllStructEx_eElementType_NONE
                            $sType = "NONE"
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iStructure, "")
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iDllStructure, "")
                        Case $__g_DllStructEx_eElementType_PTR
                            $sType = "PTR"
                            ;FIXME: get correct information from ptr (dll struct)
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iStructure, "")
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iDllStructure, "")
                        Case $__g_DllStructEx_eElementType_STRUCT
                            $sType = "Struct"
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iStructure, _WinAPI_GetString($tElement.szStruct))
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iDllStructure, _WinAPI_GetString($tElement.szTranslatedStruct))
                        Case $__g_DllStructEx_eElementType_UNION
                            $sType = "Union"
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iStructure, _WinAPI_GetString($tElement.szStruct))
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iDllStructure, _WinAPI_GetString($tElement.szTranslatedStruct))
                        Case Else
                            $sType = "Unknown"
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iStructure, "")
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iDllStructure, "")
                    EndSwitch
                    GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iType, $sType)
                    Switch ($tElement.iType)
                        Case $__g_DllStructEx_eElementType_UNION
                            Local $iSize = 0
                            Local $iElementSize = 0
                            Local $pElements = $tElement.pElements
                            Local $_tElement
                            For $i = 0 To $tElement.cElements - 1
                                $_tElement = DllStructCreate($__g_DllStructEx_tagElement, $pElements + $__g_DllStructEx_iElement * $i)
                                $iElementSize = DllStructGetSize(DllStructCreate(_WinAPI_GetString($_tElement.szTranslatedStruct)))
                                $iSize = $iElementSize > $iSize ? $iElementSize : $iSize
                            Next
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iSize, $iSize)
                        Case Else
                            GUICtrlSetData($__g_DllStructEx_DllStructExDisplay_iSize, DllStructGetSize(DllStructCreate(_WinAPI_GetString($tElement.szTranslatedStruct))))
                    EndSwitch

                    ;ConsoleWrite("change"&@CRLF)
                    ;ConsoleWrite($hItem)
                    ;ConsoleWrite(_GUICtrlTreeView_GetText($__g_DllStructEx_DllStructExDisplay_hTreeView, $hItem)&@CRLF)
                    ;_GUICtrlTreeView_SetText($hItem)
                EndIf
        EndSwitch
    EndIf
EndFunc
