VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTMLEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEmitter

Private m_Filename   As String
Private m_Filenumber As Long
Private m_Fileopen   As Boolean

Private CSS         As String

Private HTMLSTART   As String
Private HTMLEND     As String

Private SPANSTART   As String
Private SPANEND     As String

Private CurClass    As String
Private m_Label     As Label
Private m_lTokens   As Long


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

Private Sub Class_Initialize()
    CSS = _
    "SPAN.keyword { color: #000088; }" & vbCrLf & _
    "SPAN.comment { color: #008800; }" & vbCrLf & _
    "SPAN.String  { color: #888888; }" & vbCrLf & _
    "SPAN.Number  { color: #880000; }" & vbCrLf & _
    "SPAN.Text    { color: #000000; }" & vbCrLf
    
    HTMLSTART = _
        "<html>" & vbCrLf & _
        "  <head>" & vbCrLf & _
        "    <link rel=""stylesheet"" type=""text/css"" href=""$CSSFILE"">" & vbCrLf & _
        "  </head>" & vbCrLf & _
        "  <body>" & vbCrLf & _
        "    <tt><pre>" & vbCrLf
        
    HTMLEND = _
        "    </pre></tt>" & vbCrLf & _
        "  </body>" & vbCrLf & _
        "</html>" & vbCrLf
        
    SPANSTART = _
        "<span class=""$SPANCLASS"">"
        
    SPANEND = _
        "</span>"
End Sub

Private Sub Class_Terminate()
    If FileOpen Then
        IEmitter_EmitEnd
    End If
End Sub

Private Sub IEmitter_EmitComment(Token As String)
    EmitToken Token, "comment"
End Sub

Private Sub IEmitter_EmitDelimiter(Token As String)
    If Trim$(Token) = "" Then
        EmitToken Token, CurClass
    ElseIf Token = vbLf Then
        EmitToken Token
    ElseIf InStr("()", Token) > 0 Then
        EmitToken Token, "keyword"
    Else
        EmitToken Token, "text"
    End If
End Sub

Private Sub IEmitter_EmitString(Token As String)
    EmitToken Token, "string"
End Sub

Private Sub IEmitter_EmitKeyword(Token As String)
    EmitToken Token, "keyword"
End Sub

Private Sub IEmitter_EmitNumber(Token As String)
    EmitToken Token, "number"
End Sub

Private Sub EmitToken(ByVal Token As String, Optional SpanClass As String = "")
    If m_lTokens Mod 1000 = 0 Then
        m_Label.Caption = "Status: " & m_lTokens & " tokens written..."
        DoEvents
    End If
    
    Token = HTMLEntities(Token)
    
    If SpanClass <> CurClass Then
        If CurClass <> "" Then Print #m_Filenumber, SPANEND;
        If SpanClass <> "" Then Print #m_Filenumber, Replace(SPANSTART, "$SPANCLASS", SpanClass);
        CurClass = SpanClass
    End If
    
    Print #m_Filenumber, Token;
    m_lTokens = m_lTokens + 1
End Sub

Private Sub IEmitter_EmitStart()
    If FileOpen Then
        IEmitter.EmitEnd
    End If
        
    m_Fileopen = True
    m_Filenumber = FreeFile
    
    ' Write CSS
    Open m_Filename & ".css" For Output As m_Filenumber
        Print #m_Filenumber, CSS;
    Close m_Filenumber
    
    ' Write HTML header
    Open m_Filename For Output As m_Filenumber
        Print #m_Filenumber, Replace(HTMLSTART, "$CSSFILE", m_Filename & ".css");
End Sub

Private Sub IEmitter_EmitEnd()
    Print #m_Filenumber, HTMLEND;
    Close m_Filenumber
    m_Fileopen = False
    
    m_Label.Caption = "Status: Done, " & m_lTokens & " tokens written..."
    DoEvents
End Sub

Private Sub IEmitter_EmitText(Token As String)
    EmitToken Token, "text"
End Sub

Private Function HTMLEntities(ByVal Token As String) As String
    Token = Replace(Token, vbLf, vbCrLf)
    Token = Replace(Token, "&", "&amp;")
    Token = Replace(Token, """", "&quot;")
    Token = Replace(Token, "<", "&lt;")
    Token = Replace(Token, ">", "&gt;")
    
    HTMLEntities = Token
End Function

Private Property Get IEmitter_OutputFile() As String
    IEmitter_OutputFile = Filename
End Property

Private Property Let IEmitter_OutputFile(RHS As String)
    Filename = RHS
End Property

Private Property Set IEmitter_Statuslabel(RHS As Label)
    Set m_Label = RHS
End Property
