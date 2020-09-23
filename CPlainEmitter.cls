VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlainEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEmitter

Private m_Filename   As String
Private m_Filenumber As Long
Private m_Fileopen   As Boolean
Private m_Label      As Label
Private m_lTokens    As Long

Private Property Let Filename(TheFilename As String)
    If Not m_Fileopen Then m_Filename = TheFilename
End Property
Private Property Get Filename() As String
    Filename = m_Filename
End Property
Private Property Get Filenumber() As Long
    Filenumber = m_Filenumber
End Property
Private Property Get FileOpen() As Boolean
    FileOpen = m_Fileopen
End Property

Private Sub Class_Terminate()
    If FileOpen Then
        IEmitter_EmitEnd
    End If
End Sub

Private Sub IEmitter_EmitComment(Token As String)
    EmitToken Token
End Sub

Private Sub IEmitter_EmitDelimiter(Token As String)
    EmitToken Token
End Sub

Private Sub IEmitter_EmitString(Token As String)
    EmitToken Token
End Sub

Private Sub IEmitter_EmitKeyword(Token As String)
    EmitToken Token
End Sub

Private Sub IEmitter_EmitNumber(Token As String)
    EmitToken Token
End Sub

Private Sub EmitToken(ByVal Token As String)
    If m_lTokens Mod 1000 = 0 Then
        m_Label.Caption = "Status: " & m_lTokens & " tokens written..."
        DoEvents
    End If

    Token = Replace(Token, vbCrLf, vbLf)
    Token = Replace(Token, vbCr, vbLf)
    Token = Replace(Token, vbLf, vbCrLf)
    Print #m_Filenumber, Token;
    
    m_lTokens = m_lTokens + 1
End Sub

Private Sub IEmitter_EmitStart()
    If FileOpen Then
        IEmitter.EmitEnd
    End If
        
    m_Fileopen = False
    m_Filenumber = FreeFile
    Open m_Filename For Output As m_Filenumber
End Sub

Private Sub IEmitter_EmitEnd()
    m_Label.Caption = "Status: Done, " & m_lTokens & " tokens written."
    Close m_Filenumber
End Sub

Private Sub IEmitter_EmitText(Token As String)
    EmitToken Token
End Sub

Private Property Get IEmitter_OutputFile() As String
    IEmitter_OutputFile = Filename
End Property

Private Property Let IEmitter_OutputFile(RHS As String)
    Filename = RHS
End Property

Private Property Set IEmitter_Statuslabel(RHS As Label)
    Set m_Label = RHS
End Property
