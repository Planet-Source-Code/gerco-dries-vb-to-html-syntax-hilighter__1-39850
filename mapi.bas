Attribute VB_Name = "MAPI"
Option Explicit

Private Declare Function API_GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function API_GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000

Public Function GetOpenFileName(Prompt As String, Filename As String, hWnd As Long) As String
        Dim of As OPENFILENAME
        Dim file As String * 1024
        Dim retval As String
        
        file = Filename & Chr$(0)
        of.hInstance = App.hInstance
        of.hwndOwner = hWnd
        of.lpstrFilter = "Alle Bestanden (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        of.flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        of.lpstrFile = file
        of.lpstrFileTitle = 0&
        of.nMaxFile = 1024
        of.lpstrTitle = Prompt & Chr$(0)
        of.lStructSize = Len(of)
        
        If API_GetOpenFileName(of) > 0 Then
            retval = Left$(of.lpstrFile, InStr(of.lpstrFile, Chr$(0)) - 1)
        End If
        
        GetOpenFileName = retval
End Function

Public Function GetSaveFileName(Prompt As String, Filename As String, hWnd As Long) As String
        Dim of As OPENFILENAME
        Dim file As String * 1024
        Dim retval As String
        
        file = Filename & Chr$(0)
        of.hInstance = App.hInstance
        of.hwndOwner = hWnd
        of.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        of.lpstrFilter = "Alle Bestanden (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        of.lpstrFile = file
        of.lpstrFileTitle = 0&
        of.nMaxFile = 1024
        of.lpstrTitle = Prompt & Chr$(0)
        of.lStructSize = Len(of)
        
        If API_GetSaveFileName(of) > 0 Then
            retval = Left$(of.lpstrFile, InStr(of.lpstrFile, Chr$(0)) - 1)
        End If
        
        GetSaveFileName = retval
End Function


