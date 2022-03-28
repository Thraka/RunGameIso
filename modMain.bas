Attribute VB_Name = "modMain"
Option Explicit

Public Enum EmuSetting

    TurnOn
    TurnOff
    Ignore
    
End Enum

Const PROG_FILES_DTOOLS As String = "c:\program files\d-tools\daemon.exe"

Public Sub Main()

    Dim dtoolsPath As String
    Dim deviceNumber As Byte: deviceNumber = 0
    Dim imagePath As String
    Dim sleepProgram As Long
    Dim program As String
    Dim flagSafe As EmuSetting
    Dim flagSecu As EmuSetting
    Dim flagLaser As EmuSetting
    Dim flagRMPS As EmuSetting
    Dim unmount As Boolean
    Dim returnValue As Long
    Dim gameName As String
    Dim alertForm As Alert
    
    Dim parameters() As String
    
    ' Make sure command line parameters are correct
    If IsNull(Command) Or Command = "" Then
        ErrorCommandline
        Exit Sub
    End If
    
    If FileExist(Command) = False Then
        MsgBox "File doesn't exist:" + vbNewLine + Command
        Exit Sub
    End If
    
    ' Read all of the settings
    dtoolsPath = GetINI(Command, "DEFAULT", "daemon_executable", PROG_FILES_DTOOLS)
    deviceNumber = CInt(GetINI(Command, "DEFAULT", "device_number"))
    sleepProgram = CLng(GetINI(Command, "DEFAULT", "wait_seconds_before_program", "2"))
    imagePath = GetINI(Command, "DEFAULT", "image")
    program = GetINI(Command, "DEFAULT", "program", "")
    flagSafe = ConvertEmuEnum(GetINI(Command, "DEFAULT", "safedisc", "ignore"))
    flagSecu = ConvertEmuEnum(GetINI(Command, "DEFAULT", "securom", "ignore"))
    flagLaser = ConvertEmuEnum(GetINI(Command, "DEFAULT", "laserlock", "ignore"))
    flagRMPS = ConvertEmuEnum(GetINI(Command, "DEFAULT", "rmps", "ignore"))
    unmount = CBool(GetINI(Command, "DEFAULT", "unmount", "false"))
    gameName = GetINI(Command, "DEFAULT", "game_name", "CD-ROM Game")
    
    If dtoolsPath = "" Then dtoolsPath = PROG_FILES_DTOOLS
    
    ' Check if dtools and image paths exist
    If Not FileExist(dtoolsPath) Then
        MsgBox "Can't find D-Tools path:" + vbNewLine + "  " + dtoolsPath
        Exit Sub
    End If
    If Not FileExist(imagePath) Then
        MsgBox "Can't find image:" + vbNewLine + "  " + imagePath
        Exit Sub
    End If
    
    ' Build the command string
    Dim runCommand As String
    
    runCommand = Qoute(dtoolsPath) _
                 + " -mount " + CStr(deviceNumber) + "," + Qoute(imagePath)
    
    If flagSafe <> Ignore Then
        runCommand = runCommand + " -safedisc " + IIf(flagSafe = TurnOn, "on", "off")
    End If
    If flagSecu <> Ignore Then
        runCommand = runCommand + " -securom " + IIf(flagSecu = TurnOn, "on", "off")
    End If
    If flagLaser <> Ignore Then
        runCommand = runCommand + " -laserlock " + IIf(flagLaser = TurnOn, "on", "off")
    End If
    If flagRMPS <> Ignore Then
        runCommand = runCommand + " -rmps " + IIf(flagRMPS = TurnOn, "on", "off")
    End If
    
    Set alertForm = New Alert
    alertForm.GameTitle = gameName
    alertForm.Show False
    DoEvents
    
    ' Return value not working yet...
    returnValue = Shell(runCommand)
    
    ' If program is defined, run it
    If program <> "" Then
        If Not FileExist(program) Then
            MsgBox "Can't find program to run:" + vbNewLine + "  " + program
            Exit Sub
        End If
        
        Do Until sleepProgram = 0
        
            Sleep 1000
            DoEvents
            sleepProgram = sleepProgram - 1
            
        Loop
        
        alertForm.Hide
        DoEvents
        
        Rem Shell Qoute(program), vbNormalFocus
        Shell32Bit Qoute(program)
        
        If unmount Then
            runCommand = Qoute(dtoolsPath) _
                + " -unmount " + CStr(deviceNumber)
            
            Shell runCommand
        End If
    Else
        alertForm.Hide
    End If
    
    ' ShellWait(program
    GoTo ExitProgram
    
IniError:
    MsgBox "Unable to parse INI file: " + vbNewLine + Command

ExitProgram:
    Unload alertForm
    Set alertForm = Nothing
    
End Sub

Private Sub ErrorCommandline()

    MsgBox _
        "Command line parameters are incorrect" + vbNewLine + _
        vbNewLine + _
        "Format: " + vbNewLine + _
        "    VDL.EXE ""path to INI file""", vbOKOnly, "Invalid parameters"

End Sub

Private Function Qoute(value As String)
    Qoute = """" + value + """"
End Function

Private Function ConvertEmuEnum(value As String) As EmuSetting

    ConvertEmuEnum = Ignore

    value = Trim(LCase(value))
    
    If value = "on" Then
        ConvertEmuEnum = TurnOn
    ElseIf value = "off" Then
        ConvertEmuEnum = TurnOff
    End If
    
End Function

Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
    If IsNull(Fname) Or Fname = "" Then
        FileExist = False
    Else
        FileExist = (Dir(Fname) <> "")
    End If
End Function
