Attribute VB_Name = "modObjectExtender"
Option Explicit

' modObjectExtender
'
' event support for Late-Bound objects
' low level COM Projekt - by [rm_code] 2005

Public Type UUID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(7)    As Byte
End Type

Public Type EventSink
    pVTable     As Long     ' VTable pointer
    pClass      As Long     ' clsObjectExtender pointer
    cRef        As Long     ' reference counter
    iid         As UUID     ' interface IID
    hMem        As Long     ' memory address
End Type

Public Declare Function CallWindowProcA Lib "user32" ( _
    ByVal adr As Long, ByVal p1 As Long, ByVal p2 As Long, _
    ByVal p3 As Long, ByVal p4 As Long) As Long

Public Declare Function VariantCopyIndPtr Lib "oleaut32" Alias "VariantCopyInd" ( _
    ByVal pvargDest As Long, ByVal pvargSrc As Long) As Long

Public Declare Function SysAllocStringPtr Lib "oleaut32" ( _
    ByVal pStr As Long) As Long

Public Declare Function SysReAllocString Lib "oleaut32" ( _
    ByVal StrSrc As Long, ByVal StrNew As Long) As Long

Public Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" ( _
    PtrDest() As Any) As Long

Public Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal dwLen As Long)

Public Declare Sub FillMem Lib "kernel32" Alias "RtlFillMemory" ( _
    pDst As Any, ByVal dlen As Long, ByVal Fill As Byte)

Public Declare Function IsEqualGUID Lib "ole32" ( _
    rguid1 As UUID, rguid2 As UUID) As Long

Public Declare Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, UUID As Any) As Long

Public Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal uFlags As Long, ByVal dwBytes As Long) As Long

Public Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Public Declare Function LCID Lib "kernel32" Alias "GetSystemDefaultLCID" ( _
    ) As Long

Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_NOTIMPL As Long = &H80004001

'             GMEM | GMEM_ZEROINIT
Private Const GPTR As Long = &H40&

Public Const IIDSTR_IUnknown As String = _
    "{00000000-0000-0000-C000-000000000046}"

Public Const IIDSTR_IDispatch As String = _
    "{00020400-0000-0000-C000-000000000046}"

Public Const IIDSTR_IConnectionPoint As String = _
    "{B196B286-BAB4-101A-B69C-00AA00341D07}"

Public Const IIDSTR_IEnumConnectionPoints As String = _
    "{B196B285-BAB4-101A-B69C-00AA00341D07}"

Public Const IIDSTR_IConnectionPointContainer As String = _
    "{B196B284-BAB4-101A-B69C-00AA00341D07}"

Public IID_IUnknown     As UUID
Public IID_IDispatch    As UUID

Private Const MAXCODE   As Long = &HEC00&

Private ObjExt_vtbl(6) As Long

Public Sub InitObjExtender()
    Static blnInit  As Boolean

    If blnInit Then Exit Sub

    CLSIDFromString StrPtr(IIDSTR_IUnknown), IID_IUnknown
    CLSIDFromString StrPtr(IIDSTR_IDispatch), IID_IDispatch

    ObjExt_vtbl(0) = addr(AddressOf ObjExt_QueryInterface)
    ObjExt_vtbl(1) = addr(AddressOf ObjExt_AddRef)
    ObjExt_vtbl(2) = addr(AddressOf ObjExt_Release)
    ObjExt_vtbl(3) = addr(AddressOf ObjExt_GetTypeInfoCount)
    ObjExt_vtbl(4) = addr(AddressOf ObjExt_GetTypeInfo)
    ObjExt_vtbl(5) = addr(AddressOf ObjExt_GetIDsOfNames)
    ObjExt_vtbl(6) = addr(AddressOf ObjExt_Invoke)

    blnInit = True
End Sub

' IUnknown::QueryInterface
Private Function ObjExt_QueryInterface(This As EventSink, riid As UUID, pObj As Long) As Long

    ' IUnknown
    If IsEqualGUID(riid, IID_IUnknown) Then
        pObj = VarPtr(This)
        ObjExt_AddRef This

    ' IDispatch
    ElseIf IsEqualGUID(riid, IID_IDispatch) Then
        pObj = VarPtr(This)
        ObjExt_AddRef This

    ' event interface
    ElseIf IsEqualGUID(riid, This.iid) Then
        pObj = VarPtr(This)
        ObjExt_AddRef This

    ' not an implemented interface
    Else
        pObj = 0
        ObjExt_QueryInterface = E_NOINTERFACE

    End If
End Function

' IUnknown::AddRef
Private Function ObjExt_AddRef(This As EventSink) As Long
    This.cRef = This.cRef + 1
    ObjExt_AddRef = This.cRef
End Function

' IUnknown::Release
Private Function ObjExt_Release(This As EventSink) As Long
    This.cRef = This.cRef - 1
    ObjExt_Release = This.cRef

    ' if reference count is 0, free the object
    If This.cRef = 0 Then GlobalFree This.hMem
End Function

' IDispatch::GetTypeInfoCount
Private Function ObjExt_GetTypeInfoCount(This As EventSink, pctinfo As Long) As Long
    pctinfo = 0
    ObjExt_GetTypeInfoCount = E_NOTIMPL
End Function

' IDispatch::GetTypeInfo
Private Function ObjExt_GetTypeInfo(This As EventSink, ByVal iTInfo As Long, ByVal LCID As Long, ppTInfo As Long) As Long
    ppTInfo = 0
    ObjExt_GetTypeInfo = E_NOTIMPL
End Function

' IDispatch::GetIDsOfNames
Private Function ObjExt_GetIDsOfNames(This As EventSink, riid As UUID, rgszNames As Long, ByVal cNames As Long, ByVal LCID As Long, rgDispId As Long) As Long
    ObjExt_GetIDsOfNames = E_NOTIMPL
End Function

' IDispatch::Invoke
Private Function ObjExt_Invoke(This As EventSink, _
         ByVal dispIdMember As Long, _
         riid As UUID, _
         ByVal LCID As Long, _
         ByVal wFlags As Integer, _
         ByVal pDispParams As Long, _
         ByVal pVarResult As Long, _
         ByVal pExcepInfo As Long, _
         puArgErr As Long) As Long

    ' get the object extender class
    ' which owns this event sink
    Dim objext  As clsObjectExtender
    Set objext = ResolveObjPtr(This.pClass)

    ' forward the event
    objext.FireEvent dispIdMember, pDispParams
End Function

Public Function CreateEventSink(iid As UUID, objext As clsObjectExtender) As Object
    Dim sink    As EventSink

    ' our event sink object :)
    With sink
        .cRef = 1
        .iid = iid
        .pClass = ObjPtr(objext)
        .pVTable = VarPtr(ObjExt_vtbl(0))
    End With

    ' allocate some memory for our object
    sink.hMem = GlobalAlloc(GPTR, Len(sink))
    If sink.hMem = 0 Then Exit Function
    CpyMem ByVal sink.hMem, sink, Len(sink)

    ' return the object
    CpyMem CreateEventSink, sink.hMem, 4&
End Function

Private Function addr(p As Long) As Long
    addr = p
End Function

' Pointer->Object
Private Function ResolveObjPtr(ByVal Ptr As Long) As IUnknown
    Dim oUnk As IUnknown

    CpyMem oUnk, Ptr, 4&
    Set ResolveObjPtr = oUnk
    CpyMem oUnk, 0&, 4&
End Function

Public Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long
    Dim btASM(MAXCODE - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    FillMem ByVal pASM, MAXCODE, &HCC

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(params) = 0 Then
        If IsArray(params(0)) Then
            For i = UBound(params(0)) To 0 Step -1
                AddPush pASM, CLng(params(0)(i))    ' PUSH dword
            Next
        Else
            For i = UBound(params) To 0 Step -1
                AddPush pASM, CLng(params(i))       ' PUSH dword
            Next
        End If
    Else
        For i = UBound(params) To 0 Step -1
            AddPush pASM, CLng(params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallPointer = CallWindowProcA(VarPtr(btASM(0)), _
                                  0, 0, 0, 0)
End Function

Private Sub AddPush(pASM As Long, lng As Long)
    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall(pASM As Long, addr As Long)
    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong(pASM As Long, lng As Long)
    CpyMem ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte(pASM As Long, bt As Byte)
    CpyMem ByVal pASM, bt, 1
    pASM = pASM + 1
End Sub
