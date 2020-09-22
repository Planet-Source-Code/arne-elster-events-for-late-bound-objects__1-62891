VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m    As clsObjectExtender
Attribute m.VB_VarHelpID = -1
Private c               As Object

Private Sub Form_Load()
    ' new object extender instance
    Set m = New clsObjectExtender

    ' testobj is an internet explorer instance
    Set c = CreateObject("InternetExplorer.Application")

    ' connect to c
    If Not m.Attach(c) Then
        MsgBox "couldn't connect to c", vbExclamation
        Exit Sub
    End If

    ' fire some events
    c.Navigate2 "http://www.google.de/"

    ' unadvise the event sink
    m.Detach
End Sub

Private Sub m_EventRaised(ByVal strName As String, params() As Variant)
    On Error Resume Next
    Dim i    As Long

    MsgBox "Event " & strName

    ' event name
    Debug.Print "m_Event: " & strName, ;

    ' test the bounds
    i = UBound(params)
    If Err Then
        Debug.Print ""
        Exit Sub
    End If

    ' parameters values
    For i = 1 To UBound(params)
        Debug.Print "Param " & i & ": " & params(i), ;
    Next

    Debug.Print ""
End Sub
