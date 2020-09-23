VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Formatter"
   ClientHeight    =   2535
   ClientLeft      =   4395
   ClientTop       =   4740
   ClientWidth     =   6750
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6750
   Begin VB.CheckBox chkExpandKwd 
      Caption         =   "Expand keywords"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cmbOutputFormat 
      Height          =   315
      ItemData        =   "FMain.frx":0442
      Left            =   960
      List            =   "FMain.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox cmbInputFormat 
      Height          =   315
      ItemData        =   "FMain.frx":045D
      Left            =   960
      List            =   "FMain.frx":0467
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowseOutput 
      Caption         =   "..."
      Height          =   285
      Left            =   6240
      Picture         =   "FMain.frx":0487
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   5175
   End
   Begin VB.CommandButton cmdBrowseInput 
      Caption         =   "..."
      Height          =   285
      Left            =   6240
      Picture         =   "FMain.frx":08C9
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtInputFile 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   6735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Format :"
      Height          =   255
      Left            =   -120
      TabIndex        =   10
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Format :"
      Height          =   255
      Left            =   -120
      TabIndex        =   9
      Top             =   520
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Output file :"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   980
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Input file :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Start  As Boolean
Private m_Force  As Boolean
Private m_Hide   As Boolean
Private m_Expand As Boolean
Private m_Touch  As Boolean

Private Sub cmbInputFormat_Click()
    chkExpandKwd.Visible = (cmbInputFormat.ListIndex = 1)
End Sub

Private Sub cmdBrowseInput_Click()
    Dim s As String
    s = GetOpenFileName("Open file to convert", txtInputFile.Text, Me.hWnd)
    If s <> "" Then txtInputFile.Text = s
End Sub

Private Sub cmdBrowseOutput_Click()
    Dim s As String
    s = GetSaveFileName("Select file to save to", txtOutputFile.Text, Me.hWnd)
    If s <> "" Then txtOutputFile.Text = s
End Sub

Private Sub cmdConvert_Click()
    Dim sCode    As String
    Dim sNewCode As String
    Dim sExtra   As String
    Dim glp      As C4GLParser
    Dim p        As IParser
    Dim e        As IEmitter
    Dim dt       As cDateTime
    Dim dtInput  As Date
            
    Set dt = New cDateTime
            
    ' Parse file
    Select Case cmbInputFormat.ListIndex
        Case 0: Set p = New CVBParser
        Case 1:
            Set glp = New C4GLParser
            glp.lExpandKwd = (chkExpandKwd.Value = vbChecked)
            Set p = glp
        Case Else:
            MsgBox "ERROR"
            Exit Sub
    End Select
    
    Select Case cmbOutputFormat.ListIndex
        Case 0: Set e = New CPlainEmitter
        Case 1: Set e = New CHTMLEmitter
        Case Else:
            MsgBox "ERROR"
            Exit Sub
    End Select
    
    If Not m_Force Then
        If Dir(txtOutputFile.Text) <> "" Then
            If MsgBox("The output file already exists, overwrite?", vbExclamation + vbYesNo + vbDefaultButton2, "CodeFormatter") = vbNo Then Exit Sub
        End If
    End If
    
    ' Read input file
    On Error GoTo cmdConvert_Click_Error
    dtInput = dt.GetFileDate(txtInputFile.Text)
    Open txtInputFile.Text For Input As 1
        sCode = Input(LOF(1), 1)
    Close 1
    On Error GoTo 0
    
    p.Code = sCode
    e.OutputFile = txtOutputFile.Text
    Set e.Statuslabel = lblStatus
    Set p.Emitter = e
    
    On Error GoTo cmdConvert_Click_Error:
    Me.MousePointer = vbHourglass
    e.EmitStart
    p.Start
    e.EmitEnd
    On Error GoTo 0
    
    ' Sloopcheck
    If chkExpandKwd.Value <> vbChecked Then
        On Error GoTo cmdConvert_Click_Error
        Open txtOutputFile.Text For Input As 1
            sNewCode = Input(LOF(1), 1)
        Close 1
        
        If UCase$(sNewCode) <> UCase(sCode) Then
            If LCase$(txtInputFile.Text) = LCase$(txtOutputFile.Text) Then
                sExtra = vbCrLf & "Your code has been restored"
                Open txtInputFile.Text For Output As 1
                    Print #1, sCode;
                Close 1
            End If
                
            If Not m_Force Then _
                MsgBox "After formatting " & txtInputFile.Text & ", " & Replace(CStr(100 - Round((Len(sNewCode) / Len(sCode)) * 100, 2)), ",", ".") & "% of the code was destroyed" & sExtra, _
                    vbExclamation + vbOKOnly, "CodeFormatter"
                    
            chkExpandKwd.Enabled = False
            chkExpandKwd.Value = vbUnchecked
        Else
            chkExpandKwd.Enabled = True
            
            If Not m_Force Then
                If MsgBox("Would you like to expand keywords?", vbQuestion + vbYesNo + vbDefaultButton1, "CodeFormatter") = vbYes Then
                    chkExpandKwd.Value = vbChecked
                    m_Force = True
                    cmdConvert.Value = True
                    m_Force = False
                End If
            ElseIf m_Expand Then
                chkExpandKwd.Value = vbChecked
                cmdConvert.Value = True
            End If
        End If
        On Error GoTo 0
    End If
    
    If Not m_Touch Then dt.UpdateFileTime txtOutputFile, dtInput
    
    If Not (m_Start Or m_Hide Or m_Force) Then MsgBox "Done"

cmdConvert_Click_Exit:
    Me.MousePointer = vbDefault
    Exit Sub
    
cmdConvert_Click_Error:
    If Not m_Force Then _
        MsgBox "Error while parsing " & txtInputFile.Text & vbCrLf & _
               Err.Description & " (" & Err.Number & ")", vbCritical + vbOKOnly, "CodeFormatter"
    Resume cmdConvert_Click_Exit
End Sub

Private Sub Form_Load()
    cmbInputFormat.ListIndex = 1
    cmbOutputFormat.ListIndex = 0
    chkExpandKwd.Visible = True
    
    parseOptions
    
    If m_Start Then
        If Not m_Hide Then Me.Show
        cmdConvert.Value = True
        End
    End If
End Sub

Private Sub parseOptions()
    Dim sCmd       As String
    Dim stTok      As CStringTokenizer
    Dim sToken     As String
    Dim iState     As Integer
    Dim sCurOpt    As String
    Dim sParam     As String
    Dim sPrevToken As String
    
    sCmd = Command$ & " -"
    Set stTok = New CStringTokenizer
    stTok.Init sCmd, " -", True
    
    Do While stTok.hasMoreTokens
        sPrevToken = sToken
        sToken = stTok.nextToken
        
        Select Case iState
            Case 0: ' Whitespace
                If sToken = "-" Then iState = 1 ' Parsing option
                    
            Case 1: ' Parsing option
                If sToken = "-" Then
                    Select Case sCurOpt
                        Case "i": txtInputFile.Text = Trim(sParam)
                        Case "o": txtOutputFile.Text = Trim(sParam)
                        Case "p": cmbInputFormat.ListIndex = Int(Trim(sParam))
                        Case "e": cmbOutputFormat.ListIndex = Int(Trim(sParam))
                        Case "s": m_Start = True
                        Case "f": m_Force = True
                        Case "h": m_Hide = True
                        Case "x": m_Expand = True
                        Case "t": m_Touch = True
                    End Select
                    
                ElseIf sPrevToken = "-" And InStr("iopesfhxt", sToken) > 0 Then
                    sCurOpt = sToken
                    sParam = ""
                Else
                    sParam = sParam & sToken
                End If
        End Select
    Loop
End Sub

Private Sub txtInputFile_Change()
    chkExpandKwd.Enabled = False
    chkExpandKwd.Value = vbUnchecked
End Sub
