VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDebugEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEmitter

Private Sub IEmitter_EmitDelimiter(Token As String)
    Debug.Print Token
End Sub

Public Sub IEmitter_EmitStart()
    Debug.Print "Output started"
End Sub

Public Sub IEmitter_EmitEnd()
    Debug.Print "Output stopped"
End Sub

Public Sub IEmitter_EmitKeyword(Token As String)
    Debug.Print Token;
End Sub

Public Sub IEmitter_EmitComment(Token As String)
    Debug.Print Token;
End Sub

Public Sub IEmitter_EmitString(Token As String)
    Debug.Print Token;
End Sub

Public Sub IEmitter_EmitNumber(Token As String)
    Debug.Print Token;
End Sub

Private Sub IEmitter_EmitText(Token As String)
    Debug.Print Token
End Sub

Private Property Let IEmitter_OutputFile(RHS As String)

End Property

Private Property Get IEmitter_OutputFile() As String

End Property

Private Property Set IEmitter_Statuslabel(RHS As Label)

End Property
