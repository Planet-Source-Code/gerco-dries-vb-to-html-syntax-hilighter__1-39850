VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function APIGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Public Sub UpdateFileTime(Filename As String, NewDateTime As Date)
 
  Dim lngHandle As Long
  Dim udtFileTime As FILETIME
  Dim udtLocalTime As FILETIME
  Dim udtSystemTime As SYSTEMTIME

  udtSystemTime.wYear = Year(NewDateTime)
  udtSystemTime.wMonth = Month(NewDateTime)
  udtSystemTime.wDay = Day(NewDateTime)
  udtSystemTime.wDayOfWeek = Weekday(NewDateTime) - 1
  udtSystemTime.wHour = Hour(NewDateTime)
  udtSystemTime.wMinute = Minute(NewDateTime) ' this was missing
  udtSystemTime.wSecond = Second(NewDateTime)
  udtSystemTime.wMilliseconds = 0

  ' convert system time to local time
  SystemTimeToFileTime udtSystemTime, udtLocalTime
  
  ' convert local time to GMT
  LocalFileTimeToFileTime udtLocalTime, udtFileTime
  
  ' open the file to get the filehandle
  lngHandle = CreateFile(Filename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
  
  ' change date/time property of the file
  SetFileTime lngHandle, udtFileTime, udtFileTime, udtFileTime
  
  ' close the handle
  CloseHandle lngHandle
  
End Sub

Public Function GetFileDate(Filename As String) As Date
    Dim ftCRTime As FILETIME
    Dim ftLATime As FILETIME
    Dim ftLWTime As FILETIME
    Dim ftLocalFileTime As FILETIME
    Dim stSystemTime As SYSTEMTIME
    Dim lngHandle As Long
    Dim dtDate As Date
    
    If Dir(Filename) <> "" Then
        lngHandle = CreateFile(Filename, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
        GetFileTime lngHandle, ftCRTime, ftLATime, ftLWTime
        FileTimeToLocalFileTime ftLWTime, ftLocalFileTime
        FileTimeToSystemTime ftLocalFileTime, stSystemTime
        CloseHandle lngHandle
        
        dtDate = DateSerial(stSystemTime.wYear, stSystemTime.wMonth, stSystemTime.wDay) + _
                 TimeSerial(stSystemTime.wHour, stSystemTime.wMinute, stSystemTime.wSecond)
    End If
    
    GetFileDate = dtDate
End Function

Public Function GetFileSize(Filename As String) As Long
    Dim lngHandle As Long
    
    If Dir(Filename) <> "" Then
        lngHandle = CreateFile(Filename, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
        GetFileSize = APIGetFileSize(lngHandle, 0&)
        CloseHandle lngHandle
    End If
End Function


