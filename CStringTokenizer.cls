VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_currentPosition As Long
Private m_maxPosition As Long
Private m_str As String
Private m_delimiters As String
Private m_retTokens As Boolean

Public Sub Init(str As String, delim As String, returnTokens As Boolean)
    m_currentPosition = 1
    m_str = str
    m_maxPosition = Len(str) + 1
    m_delimiters = delim
    m_retTokens = returnTokens
End Sub

Private Sub skipDelimiters()
    While (Not m_retTokens And _
           (m_currentPosition < m_maxPosition) And _
           (InStr(m_delimiters, Mid$(m_str, m_currentPosition, 1)) > 0))
           
        m_currentPosition = m_currentPosition + 1
    Wend
End Sub

Public Function hasMoreTokens() As Boolean
    skipDelimiters
    hasMoreTokens = (m_currentPosition < m_maxPosition)
End Function

Public Function nextToken() As String
    skipDelimiters
    
    If m_currentPosition >= m_maxPosition Then
        Exit Function
    End If
    
    Dim Start As Long
    Start = m_currentPosition
    While ((m_currentPosition < m_maxPosition) And _
          (InStr(m_delimiters, Mid$(m_str, m_currentPosition, 1)) = 0))
        m_currentPosition = m_currentPosition + 1
    Wend
    
    If m_retTokens And (Start = m_currentPosition) And _
       (InStr(m_delimiters, Mid$(m_str, m_currentPosition, 1)) > 0) Then
       m_currentPosition = m_currentPosition + 1
    End If
       
    nextToken = Mid$(m_str, Start, m_currentPosition - Start)
End Function

