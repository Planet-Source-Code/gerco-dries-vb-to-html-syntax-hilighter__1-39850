VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBParser"
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

Private m_State As eState
Private m_PrevState As eState
Private m_Code As String
Private m_st As CStringTokenizer
Private m_RW As Variant
Private m_Delim As String
Private m_Token As String
Private m_Emitter As IEmitter

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
    Dim sTmp As String

    m_Code = Replace(m_Code, vbCrLf, vbLf)

    Set m_st = New CStringTokenizer
    m_st.Init m_Code, m_Delim, True
    
    State = stInCode
    
    While m_st.hasMoreTokens
        sTmp = m_st.nextToken
        
        Select Case m_State
            Case stInCode:
                Select Case sTmp
                    Case """":
                        m_Token = sTmp
                        State = stInString
                        
                    Case "'":
                        m_Token = sTmp
                        State = stInComment
                        
                    Case CStr(Val(sTmp)):
                        m_Token = sTmp
                        State = stInNumber
                        
                    Case Else
                        If InStr(m_Delim, sTmp) > 0 Then
                            m_Emitter.EmitDelimiter (sTmp)
                        Else
                            m_Token = sTmp
                            If isReservedWord(m_Token) Then
                                m_Emitter.EmitKeyword m_Token
                            Else
                                m_Emitter.EmitText m_Token
                            End If
                        End If
                        m_Token = ""
                End Select
                
            Case stInString:
                m_Token = m_Token & sTmp
                If sTmp = """" Then
                    m_Emitter.EmitString m_Token
                    m_Token = ""
                    State = stInCode
                End If
                
            Case stInComment:
                If sTmp <> vbLf Then
                    m_Token = m_Token & sTmp
                Else
                    m_Emitter.EmitComment m_Token
                    m_Emitter.EmitDelimiter sTmp
                    m_Token = ""
                    State = stInCode
                End If
                
            Case stInNumber:
                If InStr(m_Delim, sTmp) = 0 Or sTmp = "." Then
                    m_Token = m_Token & sTmp
                Else
                    m_Emitter.EmitNumber m_Token
                    m_Emitter.EmitDelimiter sTmp
                    m_Token = ""
                    State = stInCode
                End If
        End Select
        
    Wend
End Sub

Private Function isReservedWord(ByVal Token As String) As Boolean
    Dim i As Integer
    
    Token = UCase$(Token)
    
    For i = LBound(m_RW) To UBound(m_RW)
        If Token = m_RW(i) Then
            isReservedWord = True
            Exit Function
        End If
    Next i
End Function

Private Sub Class_Initialize()
    Dim i As Integer
    
    m_Delim = " .,()"":'-+/*^&" & vbLf & vbTab
    m_RW = Array("Const", "Else", "ElseIf", "If", "Alias", "And", "As", "Base", "Binary", "Boolean", "Byte", "ByVal", "Call", "Case", "CBool", _
               "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng", "Close", "Compare", "Const", "CSng", "CStr", "Currency", "CVar", "CVErr", _
               "Decimal", "Declare", "DefBool", "DefByte", "DefCur", "DefDate", "DefDbl", "DefDec", "DefInt", "DefLng", "DefObj", "DefSng", "DefStr", _
               "DefVar", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "End", "Enum", "Eqv", "Erase", "Error", "Exit", "Explicit", "False", "For", _
               "Function", "Get", "Global", "GoSub", "GoTo", "If", "Imp", "In", "Input", "Input", "Integer", "Is", "LBound", "Let", "Lib", "Like", "Line", _
               "Lock", "Long", "Loop", "LSet", "Name", "New", "Next", "Not", "Object", "Open", "Option", "On", "Or", "Output", "Preserve", "Print", "Private", _
               "Property", "Public", "Put", "Random", "Read", "ReDim", "Resume", "Return", "RSet", "Seek", "Select", "Set", "Single", "Spc", "Static", "String", _
               "Stop", "Sub", "Tab", "Then", "True", "UBound", "Variant", "While", "Wend", "With", "Event", "Implements", "Debug", "Attribute", "Begin", "InStr", _
               "MsgBox", "Lof")

    For i = LBound(m_RW) To UBound(m_RW)
        m_RW(i) = UCase$(m_RW(i))
    Next i
End Sub
