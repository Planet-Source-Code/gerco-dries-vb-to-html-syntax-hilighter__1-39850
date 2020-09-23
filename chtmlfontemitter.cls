VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTMLFontEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEmitter

Private m_Filename   As String
Private m_Filenumber As Long
Private m_Fileopen   As Boolean

Private HTMLSTART As String
Private HTMLEND   As String

Private FONTSTART   As String
Private FONTEND     As String

Private CurColor    As String

Public KeywordColor As String
Public CommentColor As String
Public StringColor  As String
Public TextColor    As String
Public NumberColor  As String
 
Public Property Let Filename(TheFilename As String)
    If Not m_Fileopen Then m_Filename = TheFilename
End Property
Public Property Get Filename() As String
    Filename = m_Filename
End Property
Public Property Get Filenumber() As Long
    Filenumber = m_Filenumber
End Property
Public Property Get FileOpen() As Boolean
    FileOpen = m_Fileopen
End Property

Private Sub Class_Initialize()
    HTMLSTART = _
        "<html>" & vbCrLf & _
        "  <head>" & vbCrLf & _
        "  </head>" & vbCrLf & _
        "  <body>" & vbCrLf & _
        "    <tt><pre>" & vbCrLf
        
    HTMLEND = _
        "    </pre></tt>" & vbCrLf & _
        "  </body>" & vbCrLf & _
        "</html>" & vbCrLf
        
    FONTSTART = _
        "<font color=""$FONTCOLOR"">"
        
    FONTEND = _
        "</font>"
End Sub

Private Sub Class_Terminate()
    If FileOpen Then
        IEmitter_EmitEnd
    End If
End Sub

Private Sub IEmitter_EmitComment(Token As String)
    EmitToken Token, CommentColor
End Sub

Private Sub IEmitter_EmitDelimiter(Token As String)
    If Trim$(Token) = "" Then
        EmitToken Token, CurColor
    ElseIf Token = vbLf Then
        EmitToken Token
    ElseIf InStr("()", Token) > 0 Then
        EmitToken Token, KeywordColor
    Else
        EmitToken Token, TextColor
    End If
End Sub

Private Sub IEmitter_EmitString(Token As String)
    EmitToken Token, StringColor
End Sub

Private Sub IEmitter_EmitKeyword(Token As String)
    EmitToken Token, KeywordColor
End Sub

Private Sub IEmitter_EmitNumber(Token As String)
    EmitToken Token, NumberColor
End Sub

Private Sub EmitToken(ByVal Token As String, Optional FontColor As String = "")
    Token = HTMLEntities(Token)
    
    If FontColor <> CurColor Then
        If CurColor <> "" Then Print #m_Filenumber, FONTEND;
        If FontColor <> "" Then Print #m_Filenumber, Replace(FONTSTART, "$FONTCOLOR", FontColor);
        CurColor = FontColor
    End If
    
    Print #m_Filenumber, Token;
End Sub

Private Sub IEmitter_EmitStart()
    If FileOpen Then
        IEmitter.EmitEnd
    End If
        
    m_Fileopen = False
    m_Filenumber = FreeFile
    Open m_Filename For Output As m_Filenumber
        Print #m_Filenumber, HTMLSTART;
End Sub

Private Sub IEmitter_EmitEnd()
    Print #m_Filenumber, HTMLEND;
    Close m_Filenumber
End Sub

Private Sub IEmitter_EmitText(Token As String)
    EmitToken Token, TextColor
End Sub

Private Function HTMLEntities(ByVal Token As String) As String
    Token = Replace(Token, vbLf, vbCrLf)
    Token = Replace(Token, "&", "&amp;")
    Token = Replace(Token, """", "&quot;")
    
    HTMLEntities = Token
End Function


