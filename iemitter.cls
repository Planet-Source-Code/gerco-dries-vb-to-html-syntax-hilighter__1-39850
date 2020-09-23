VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Let OutputFile(TheFile As String)
End Property
Public Property Get OutputFile() As String
End Property

Public Property Set Statuslabel(TheLabel As Label)
End Property

Public Sub EmitStart()
End Sub

Public Sub EmitEnd()
End Sub

Public Sub EmitKeyword(Token As String)
End Sub

Public Sub EmitComment(Token As String)
End Sub

Public Sub EmitString(Token As String)
End Sub

Public Sub EmitNumber(Token As String)
End Sub

Public Sub EmitText(Token As String)
End Sub

Public Sub EmitDelimiter(Token As String)
End Sub
