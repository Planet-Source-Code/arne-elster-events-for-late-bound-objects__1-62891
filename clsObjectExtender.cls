VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsObjectExtender
'
' event support for Late-Bound objects
' low level COM Projekt - by [rm_code] 2005

'
' HOW IT WORKS
'
' An object, wich supports events,
' should implement IConnectionPointContainer.
' This interface let's you enumerate all
' event interfaces.
' Once you have the event interface you want,
' you can get IConnectionPoint for this interface,
' advise your event sink to it, and start
' recieving events on it.
' If the event interface has IDispatch implemented,
' all raised events will go through
' IDispatch::Invoke.
'
' And that's what we use:
' We just take the first best event interface we
' get from IEnumConnectionPoints, create a
' event sink which implements IDispatch,
' and capture all events through Invoke().
'

Private Type IUnknown
    QueryInterface              As Long
    AddRef                      As Long
    Release                     As Long
End Type

Private Type IDispatch
    IUnk                        As IUnknown
    GetTypeInfoCount            As Long
    GetTypeInfo                 As Long
    GetIDsOfNames               As Long
    Invoke                      As Long
End Type

Private Type IConnectionPointContainer
    IUnk                        As IUnknown
    EnumConnectionPoints        As Long
    FindConnectionPoint         As Long
End Type

Private Type IConnectionPoint
    IUnk                        As IUnknown
    GetConnectionInterface      As Long
    GetConnectionPointContainer As Long
    Advise                      As Long
    Unadvise                    As Long
    EnumConnections             As Long
End Type

Private Type IEnumConnectionPoints
    IUnk                        As IUnknown
    Next                        As Long
    Skip                        As Long
    Reset                       As Long
    Clone                       As Long
End Type

Private Type ITypeInfo
    IUnk                        As IUnknown
    GetTypeAttr                 As Long
    GetTypeComp                 As Long
    GetFuncDesc                 As Long
    GetVarDesc                  As Long
    GetNames                    As Long
    GetRefTypeOfImplType        As Long
    GetImplTypeFlags            As Long
    GetIDsOfNames               As Long
    Invoke                      As Long
    GetDocumentation            As Long
    GetDllEntry                 As Long
    GetRefTypeInfo              As Long
    AddressOfMember             As Long
    CreateInstance              As Long
    GetMops                     As Long
    GetContainingTypeLib        As Long
    ReleaseTypeAttr             As Long
    ReleaseFuncDesc             As Long
    ReleaseVarDesc              As Long
End Type

Private Type ITypeLib
    IUnk                        As IUnknown
    GetTypeInfoCount            As Long
    GetTypeInfo                 As Long
    GetTypeInfoType             As Long
    GetTypeInfoOfGuid           As Long
    GetLibAttr                  As Long
    GetTypeComp                 As Long
    GetDocumentation            As Long
    IsName                      As Long
    FindName                    As Long
    ReleaseTLibAttr             As Long
End Type

Private Type DISPPARAMS
    rgPointerToVariantArray     As Long
    rgPointerToLONGNamedArgs    As Long
    cArgs                       As Long
    cNamedArgs                  As Long
End Type

Private Type SAFEARRAYBOUND
    cElements                   As Long
    lLBound                     As Long
End Type

Private Type SAFEARRAY_1D
   cDims                        As Integer
   fFeatures                    As Integer
   cbElements                   As Long
   cLocks                       As Long
   pvData                       As Long
   Bounds(0 To 0)               As SAFEARRAYBOUND
End Type

Private Enum VARENUM
    VT_EMPTY = 0
    VT_NULL = 1
    VT_I2 = 2
    VT_I4 = 3
    VT_R4 = 4
    VT_R8 = 5
    VT_CY = 6
    VT_DATE = 7
    VT_BSTR = 8
    VT_DISPATCH = 9
    VT_ERROR = 10
    VT_BOOL = 11
    VT_VARIANT = 12
    VT_UNKNOWN = 13
    VT_DECIMAL = 14
    VT_I1 = 16
    VT_UI1 = 17
    VT_UI2 = 18
    VT_UI4 = 19
    VT_I8 = 20
    VT_UI8 = 21
    VT_INT = 22
    VT_UINT = 23
    VT_VOID = 24
    VT_HRESULT = 25
    VT_PTR = 26
    VT_SAFEARRAY = 27
    VT_CARRAY = 28
    VT_USERDEFINED = 29
    VT_LPSTR = 30
    VT_LPWSTR = 31
    VT_FILETIME = 64
    VT_BLOB = 65
    VT_STREAM = 66
    VT_STORAGE = 67
    VT_STREAMED_OBJECT = 68
    VT_STORED_OBJECT = 69
    VT_BLOB_OBJECT = 70
    VT_CF = 71
    VT_CLSID = 72
    VT_STREAMED_PROPSET = 73
    VT_STORED_PROPSET = 74
    VT_BLOB_PROPSET = 75
    VT_VERBOSE_ENUM = 76
    VT_BSTR_BLOB = &HFFF
    VT_VECTOR = &H1000
    VT_ARRAY = &H2000
    VT_BYREF = &H4000
    VT_RESERVED = &H8000
    VT_ILLEGAL = &HFFFF
    VT_ILLEGALMASKED = &HFFF
    VT_TYPEMASK = &HFFF
End Enum

Public Event EventRaised(ByVal strName As String, params() As Variant)

Private oCPC        As IConnectionPointContainer
Private pCPC        As Long
Private pVTblCPC    As Long

Private oECP        As IEnumConnectionPoints
Private pECP        As Long
Private pVTblECP    As Long

Private oCP         As IConnectionPoint
Private pCP         As Long
Private pVTblCP     As Long

Private oUnk        As IUnknown
Private pUnk        As Long
Private pVTblUnk    As Long

Private oTarget     As Object
Private pTarget     As Long
Private pVTblTarget As Long

Private oSink       As Object
Private dwCookie    As Long

Private iid_event   As UUID

Private blnAttached As Boolean

' by Edanmo
Private Sub pvSetParamArray(ByVal Ptr As Long, aParams() As Variant)
    Dim tDPAR       As DISPPARAMS
    Dim SafeArray   As SAFEARRAY_1D
    Dim iVarType    As Integer
    Dim lIdx        As Long
    Dim lPtr        As Long
    Dim lVarPtr     As Long

    ' DISPPARAMS structure
    CpyMem tDPAR, ByVal Ptr, Len(tDPAR)
    If tDPAR.cArgs = 0 Then Exit Sub

    ' pointer to the first variant
    lVarPtr = tDPAR.rgPointerToVariantArray

    For lIdx = 0 To tDPAR.cArgs - 1

        ' get the variant's type
        CpyMem iVarType, ByVal lVarPtr + (lIdx * 16&), 2&

        If (iVarType And VT_BYREF) = VT_BYREF Then

            ' ByRef Parameter

            ' get the pointer
            CpyMem lPtr, ByVal lVarPtr + (lIdx * 16&) + 8&, 4&

            Select Case iVarType And VT_TYPEMASK
                Case vbString

                    Dim lStrPtr As Long
                    ' String Pointer
                    CpyMem lStrPtr, ByVal lPtr, 4&

                    If lStrPtr = 0 Then
                        ' new string
                        lStrPtr = SysAllocStringPtr(StrPtr(aParams(tDPAR.cArgs - lIdx)))
                    Else
                        ' reallocate the string
                        lStrPtr = SysReAllocString(lStrPtr, StrPtr(aParams(tDPAR.cArgs - lIdx)))
                    End If

                    ' copy the string pointer
                    CpyMem ByVal lPtr, lStrPtr, 4&

                Case vbInteger
                    CpyMem ByVal lPtr, CInt(aParams(tDPAR.cArgs - lIdx)), 2&

                Case vbBoolean
                    CpyMem ByVal lPtr, CBool(aParams(tDPAR.cArgs - lIdx)), 2&

                Case vbInteger
                    CpyMem ByVal lPtr, CInt(aParams(tDPAR.cArgs - lIdx)), 2&

                Case vbLong
                    CpyMem ByVal lPtr, CLng(aParams(tDPAR.cArgs - lIdx)), 4&

                Case vbSingle
                    CpyMem ByVal lPtr, CSng(aParams(tDPAR.cArgs - lIdx)), 4&

                Case vbDouble
                    CpyMem ByVal lPtr, CDbl(aParams(tDPAR.cArgs - lIdx)), 8&

                Case vbDate
                    CpyMem ByVal lPtr, CDate(aParams(tDPAR.cArgs - lIdx)), 8&

                Case vbByte
                    CpyMem ByVal lPtr, CByte(aParams(tDPAR.cArgs - lIdx)), 1&

                Case vbCurrency
                    CpyMem ByVal lPtr, CCur(aParams(tDPAR.cArgs - lIdx)), 8&

                Case vbVariant, vbDecimal
                    VariantCopyIndPtr lPtr, VarPtr(aParams(tDPAR.cArgs - lIdx))

            End Select

        End If

    Next

End Sub

' by Edanmo
Private Function pvGetParamArray(ByVal Ptr As Long) As Variant()
    Dim tDPAR           As DISPPARAMS
    Dim SafeArray       As SAFEARRAY_1D
    Dim aTmpParams()    As Variant
    Dim aParams()       As Variant
    Dim lIdx            As Long

    ' DISPPARAMS structure
    CpyMem tDPAR, ByVal Ptr, Len(tDPAR)
    If tDPAR.cArgs = 0 Then Exit Function

    ' array pointing to the param array
    With SafeArray
        .Bounds(0).cElements = tDPAR.cArgs
        .Bounds(0).lLBound = 0
        .cDims = 1
        .cbElements = 16
        .pvData = tDPAR.rgPointerToVariantArray
    End With
    CpyMem ByVal VarPtrArray(aTmpParams), VarPtr(SafeArray), 4&

    ' copy elements
    ReDim aParams(1 To tDPAR.cArgs)
    For lIdx = 1 To tDPAR.cArgs
        aParams(lIdx) = aTmpParams(tDPAR.cArgs - lIdx)
    Next

    ' return the parameters
    pvGetParamArray = aParams

    ' destroy the array
    CpyMem ByVal VarPtrArray(aTmpParams), 0&, 4&
End Function

Private Function GetMemberName(obj As Object, ByVal dispid As Long, iid As UUID) As String
    Dim oTypeLib    As ITypeLib
    Dim pTypeLib    As Long
    Dim pVTblTpLib  As Long

    Dim oTypeInfo   As ITypeInfo
    Dim pTypeInfo   As Long
    Dim pVTblTpInfo As Long

    Dim oDispatch   As IDispatch
    Dim hRet        As Long
    Dim dwIndex     As Long
    Dim pcNames     As Long
    Dim pVTbl       As Long

    Dim strName     As String

    ' get IDispatch from the object
    pVTbl = ObjPtr(obj)
    CpyMem pVTbl, ByVal pVTbl, 4
    CpyMem oDispatch, ByVal pVTbl, Len(oDispatch)

    ' get ITypeInfo
    hRet = CallPointer(oDispatch.GetTypeInfo, ObjPtr(obj), 0, LCID, VarPtr(pTypeInfo))
    If hRet Then Exit Function

    ' ITypeInfo VTable
    CpyMem pVTblTpInfo, ByVal pTypeInfo, 4
    CpyMem oTypeInfo, ByVal pVTblTpInfo, Len(oTypeInfo)

    ' let's first ty to get the name
    ' of the member by using the current TypeInfo
    hRet = CallPointer(oTypeInfo.GetNames, pTypeInfo, dispid, VarPtr(strName), 1, VarPtr(pcNames))
    If Len(strName) > 0 Then
        GetMemberName = strName
        Exit Function
    End If

    ' no, that didn't work.
    ' go for the whole type library

    ' GetContainingTypeLib
    hRet = CallPointer(oTypeInfo.GetContainingTypeLib, pTypeInfo, VarPtr(pTypeLib), VarPtr(dwIndex))
    If hRet Then
        GetMemberName = dispid
        Exit Function
    End If

    ' ITypeLib VTable
    CpyMem pVTblTpLib, ByVal pTypeLib, 4
    CpyMem oTypeLib, ByVal pVTblTpLib, Len(oTypeLib)

    ' GetTypeInfoOfGUID
    hRet = CallPointer(oTypeLib.GetTypeInfoOfGuid, pTypeLib, VarPtr(iid_event), VarPtr(pTypeInfo))
    If hRet Then
        GetMemberName = dispid
        Exit Function
    End If

    ' ITypeInfo VTable
    CpyMem pVTblTpInfo, ByVal pTypeInfo, 4
    CpyMem oTypeInfo, ByVal pVTblTpInfo, Len(oTypeInfo)

    ' GetNames
    hRet = CallPointer(oTypeInfo.GetNames, pTypeInfo, dispid, VarPtr(strName), 1, VarPtr(pcNames))

    If Len(strName) = 0 Then
        ' no string... :(
        ' instead return the dispip
        GetMemberName = dispid
    Else
        GetMemberName = strName
    End If
End Function

Public Sub FireEvent(ByVal dispid As Long, ByVal params As Long)
    Dim strEvent    As String
    Dim vParams()   As Variant

    ' get the name of the event
    strEvent = GetMemberName(oTarget, dispid, iid_event)

    ' param array
    vParams = pvGetParamArray(params)

    ' forward the event
    RaiseEvent EventRaised(strEvent, vParams)

    ' sync byref params
    pvSetParamArray params, vParams
End Sub

Public Sub Detach()
    Dim hRet    As Long

    If Not blnAttached Then Exit Sub

    ' clean up
    hRet = CallPointer(oCP.Unadvise, pCP, dwCookie)
    hRet = CallPointer(oCP.IUnk.Release, pCP)
    hRet = CallPointer(oECP.IUnk.Release, pECP)
    hRet = CallPointer(oCPC.IUnk.Release, pCPC)
    Set oTarget = Nothing

    blnAttached = False
End Sub

Public Function Attach(obj As Object) As Boolean
    Dim cReturned   As Long
    Dim hRet        As Long
    Dim iid_ICP     As UUID

    ' already connected to an object?
    If blnAttached Then Detach

    ' IUnknown VTable
    Set oTarget = obj
    pTarget = ObjPtr(obj)
    CpyMem pVTblTarget, ByVal pTarget, 4
    CpyMem oUnk, ByVal pVTblTarget, Len(oUnk)

    ' IID string -> GUID struct
    hRet = CLSIDFromString(StrPtr(IIDSTR_IConnectionPointContainer), iid_ICP)
    If hRet Then Exit Function

    ' get IConnectionPointContainer
    hRet = CallPointer(oUnk.QueryInterface, pTarget, VarPtr(iid_ICP), VarPtr(pCPC))
    If hRet Then Exit Function

    ' IConnectionPointContainer VTable
    CpyMem pVTblCPC, ByVal pCPC, 4
    CpyMem oCPC, ByVal pVTblCPC, Len(oCPC)

    ' get IEnumConnectionPoints
    hRet = CallPointer(oCPC.EnumConnectionPoints, pCPC, VarPtr(pECP))
    If hRet Then Exit Function

    ' IEnumConnectionPoints VTable
    CpyMem pVTblECP, ByVal pECP, 4
    CpyMem oECP, ByVal pVTblECP, Len(oECP)

    ' take the first best Connection Point
    hRet = CallPointer(oECP.Next, pECP, 1, VarPtr(pCP), VarPtr(cReturned))
    If hRet Then Exit Function

    ' IConnectionPoint VTable
    CpyMem pVTblCP, ByVal pCP, 4
    CpyMem oCP, ByVal pVTblCP, Len(oCP)

    ' IID of the event interface
    hRet = CallPointer(oCP.GetConnectionInterface, pCP, VarPtr(iid_event))
    If hRet Then Exit Function

    ' create a new event sink
    Set oSink = CreateEventSink(iid_event, Me)

    ' advise the event sink
    hRet = CallPointer(oCP.Advise, pCP, ObjPtr(oSink), VarPtr(dwCookie))
    If hRet Then Exit Function

    ' wohoo, done!
    blnAttached = True
    Attach = True
End Function

Private Sub Class_Initialize()
    InitObjExtender
End Sub
