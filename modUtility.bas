Attribute VB_Name = "modUtility"
Option Explicit

'*** Monitoring a DOS Shell
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const STILL_ACTIVE = &H103
Const PROCESS_QUERY_INFORMATION = &H400


Function Shell32Bit(ByVal JobToDo As String)
    
    Dim hProcess As Long
    Dim RetVal As Long
    'The next line launches JobToDo as icon,
    
    'captures process ID
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, vbNormalFocus))
    
    Do
        
        'Get the status of the process
        GetExitCodeProcess hProcess, RetVal
        
        'Sleep command recommended as well as DoEvents
        DoEvents: Sleep 1000
        
    Loop While RetVal = STILL_ACTIVE

    CloseHandle hProcess

    Shell32Bit = RetVal

End Function
