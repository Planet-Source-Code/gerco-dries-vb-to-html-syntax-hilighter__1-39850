VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C4GLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IParser

Event Status(Text As String)

Private Enum eState
    stInCode
    stInComment
    stInString
    stInNumber
End Enum

Private m_State        As eState
Private m_PrevState    As eState
Private m_Code         As String
Private m_st           As CStringTokenizer
Private m_RW()         As String
Private m_Delim        As String
Private m_PrevToken    As String
Private m_Token        As String
Private m_Emitter      As IEmitter
Private m_CommentLevel As Long
Private m_StringDelim  As String

Public lExpandKwd As Boolean

Public Property Let IParser_Code(TheCode As String)
    m_Code = TheCode
End Property

Public Property Set IParser_Emitter(TheEmitter As IEmitter)
    Set m_Emitter = TheEmitter
End Property

Private Property Let State(TheState As eState)
    m_PrevState = m_State
    m_State = TheState
End Property


Public Sub IParser_Start()
    Dim sTmp     As String
    Dim sPrevTmp As String

    m_Code = Replace(m_Code, vbCrLf, vbLf)

    Set m_st = New CStringTokenizer
    m_st.Init m_Code, m_Delim, True
    
    State = stInCode
    
    'Open "c:\temp\cf.txt" For Output As #99
    
    Do While m_st.hasMoreTokens
        m_PrevToken = m_Token
        sPrevTmp = sTmp
        sTmp = m_st.nextToken
        
        'Print #99, "m_State: " & m_State; ", m_CommentLevel: " & m_CommentLevel & ", sTmp: |" & sTmp & "|"
        
        Select Case m_State
            Case stInCode:
                Select Case sTmp
                    Case """", "'":
                        m_Token = sTmp
                        m_StringDelim = sTmp
                        State = stInString
                                                
                    Case CStr(Val(sTmp)):
                        m_Token = sTmp
                        State = stInNumber
                        
                    Case Else
                        If InStr(m_Delim, sTmp) > 0 Then
                            m_Emitter.EmitDelimiter sTmp
                            m_Token = sTmp
                        Else
                            If isReservedWord(sTmp) Then
                                m_Emitter.EmitKeyword sTmp
                            Else
                                m_Emitter.EmitText sTmp
                            End If
                            m_Token = sTmp
                        End If
                End Select
                
            Case stInComment:
                m_Token = m_Token & sTmp
                
            Case stInString:
                m_Token = m_Token & sTmp
                If sTmp = m_StringDelim Then
                    m_Emitter.EmitString m_Token
                    State = stInCode
                End If
                
            Case stInNumber:
                If InStr(m_Delim, sTmp) = 0 Or sTmp = "." Then
                    m_Token = m_Token & sTmp
                Else
                    m_Emitter.EmitNumber m_Token
                    m_Emitter.EmitDelimiter sTmp
                    State = stInCode
                End If
        End Select
        
        If sPrevTmp = "*" And sTmp = "/" Then
            m_CommentLevel = m_CommentLevel - 1
            
            If m_CommentLevel = 0 Then
                m_Emitter.EmitComment m_Token
                m_State = stInCode
            End If
            
            sTmp = ""
        ElseIf sPrevTmp = "/" And sTmp = "*" Then
            m_State = stInComment
            If m_CommentLevel = 0 Then m_Token = ""
            m_CommentLevel = m_CommentLevel + 1
            
            sTmp = ""
        End If

    Loop
    
    'Close #99
End Sub

Private Function isReservedWord(ByRef Token As String) As Boolean
    Dim i            As Integer
    Dim sTmp         As String
    Dim iMatchLength As Integer
    Dim sMatchToken  As String
    
    sTmp = UCase$(Token)
    If Len(sTmp) < 2 Then Exit Function
    
    ' Find best possible match for Token
    For i = LBound(m_RW) To UBound(m_RW)
        If InStr(m_RW(i), "(") > 0 Then
            If Len(sTmp) >= InStr(m_RW(i), "(") - 1 Then
                If sTmp = Mid$(Replace(m_RW(i), "(", ""), 1, Len(sTmp)) Then
                    If InStr(m_RW(i), "(") > iMatchLength Then
                        iMatchLength = Len(sTmp)
                        If lExpandKwd Then
                            sMatchToken = Replace(m_RW(i), "(", "")
                        Else
                            sMatchToken = UCase$(Token)
                        End If
                    End If
                End If
            End If
        Else
            ' An exact match overrides all other matches
            If sTmp = m_RW(i) Then
                iMatchLength = Len(sTmp)
                sMatchToken = UCase$(Token)
                Exit For
            End If
        End If
    Next i
    
    If iMatchLength > 0 And sMatchToken <> "" Then
        Token = sMatchToken
        isReservedWord = True
    End If
End Function

Private Sub Class_Initialize()
    Dim i     As Integer
    Dim hFile As Long
    Dim st    As CStringTokenizer
    
    m_Delim = " ()+,.:;<>=?@[]^{}~""'/*" & vbLf & vbTab

    hFile = FreeFile
    Open App.Path & "\4GL.ini" For Input As hFile
        Set st = New CStringTokenizer
        st.Init Input(LOF(hFile), hFile), "," & vbCr & vbLf, False
    Close hFile

    While st.hasMoreTokens
        i = i + 1
        ReDim Preserve m_RW(1 To i) As String
        m_RW(i) = UCase$(st.nextToken)
    Wend
    
    Set st = Nothing
End Sub
