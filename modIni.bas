Attribute VB_Name = "modIni"
' Taken from Visual Basic Source Code 2002 CD-ROM
' File Manipulation > File Manipulation 2
' Author: Stanley Campbell
' Modified by: Thraka

Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Const KEY_NOT_FOUND_VALUE As String = "!!ERR NOT FOUND ERR!!"

Function GetINI(FileName As String, section As String, key As String, Optional defaultValue As String) As String

    GetINI = ReadWriteINI(FileName, "READ", section, key)
    
    If GetINI = KEY_NOT_FOUND_VALUE Then
        GetINI = defaultValue
    End If

End Function

Function ReadWriteINI(FileName As String, Mode As String, tmpSecname As String, tmpKeyname As String, Optional tmpKeyValue) As String
    
    Dim secname As String
    Dim keyname As String
    Dim keyvalue As String
    Dim anInt
    
    On Error GoTo ReadWriteINIError
    '
    ' *** set the return value to OK
    'ReadWriteINI = "OK"
    ' *** test for good data to work with
    If IsNull(Mode) Or Len(Mode) = 0 Then
      ReadWriteINI = "ERROR MODE"    ' Set the return value
      Exit Function
    End If
    If IsNull(tmpSecname) Or Len(tmpSecname) = 0 Then
      ReadWriteINI = "ERROR Secname" ' Set the return value
      Exit Function
    End If
    If IsNull(tmpKeyname) Or Len(tmpKeyname) = 0 Then
      ReadWriteINI = "ERROR Keyname" ' Set the return value
      Exit Function
    End If
    
    ' ******* WRITE MODE *************************************
    If UCase(Mode) = "WRITE" Then
        If IsNull(tmpKeyValue) Or Len(tmpKeyValue) = 0 Then
          ReadWriteINI = "ERROR KeyValue"
          Exit Function
        Else
        
        secname = tmpSecname
        keyname = tmpKeyname
        keyvalue = tmpKeyValue
        anInt = WritePrivateProfileString(secname, keyname, keyvalue, FileName)
        End If
    End If
    ' *******************************************************
    '
    ' *******  READ MODE *************************************
    If UCase(Mode) = "GET" Or UCase(Mode) = "READ" Then
    
        secname = tmpSecname
        keyname = tmpKeyname
        keyvalue = String$(255, 32)
        anInt = GetPrivateProfileString(secname, keyname, KEY_NOT_FOUND_VALUE, keyvalue, Len(keyvalue), FileName)
        keyvalue = RTrim(keyvalue)
        keyvalue = Trim(Left(keyvalue, Len(keyvalue) - 1))
              
        ReadWriteINI = keyvalue
    End If
    Exit Function
     
    ' *******
ReadWriteINIError:
     MsgBox Error
     Stop
End Function

