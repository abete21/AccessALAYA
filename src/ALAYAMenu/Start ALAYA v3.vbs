Option Explicit

Dim MenuFolder, WorkFolder, ModuleName, ReleaseDef
Dim RelModule, DevModule, LocModule
Dim WS, DriveName, DrivePath
Dim DBGFLG
'=====================================================================
' ALAYA Menu Startup Script. 2024/01/01
'
'<< Standard Module Settings. >>
' [MenuFolder]                             Shared Folder where Alaya Menu module exists. 
' [MenuFolder] & [ReleaseDef]              Release module definition file. 
' [MenuFolder] & [ModuleName]ver.accdb     AlayaMenu module.
' [WorkFolder]                             Client work folder where destributed client module be located.
'-------------------------------------
' Below DBGFLG is the code trace flag. If you need to trace this script, Turn on the flag, i.e. DBGFLG = True.
DBGFLG = False
' ------------------------------------
'MenuFolder = {AbsolutePath}
MenuFolder = Replace(WScript.ScriptFullName, WScript.ScriptName,"")
ModuleName = "ALAYA Menu"
ReleaseDef = "ALAYA Menu.rel"
WorkFolder = "C:\Temp\ALAYA\"

'-------------------------------------
'=====================================================================

'//////////
' Start up
'//////////
    'Reboot this script in CScript mode.
    RunCScript

    dbgEcho "Script Name = " & Wscript.ScriptName
    DevModule = GetDevModule(MenuFolder, ModuleName)
    RelModule = GetRelModule(MenuFolder, ReleaseDef)
    LocModule = GetLocModule(WorkFolder, ReleaseDef)

    Echo "****************************************"
    Echo ModuleName & " LancheMenu"
    Echo "****************************************"
    Echo "Developing Module: " & DevModule
    Echo "Destribute Module: " & RelModule
    Echo "Local Module     : " & LocModule

Dim Cmmd
    Cmmd = InputBox("Start up [" & ModuleName & "]. " & vbCrLf & "Are you sure? (Y/N)","ALAYA Startup Script","Y")
    Select Case Ucase(Cmmd)
    Case "Y"
        dbgEcho "Normal Mode"
        Call RunNormalMenu   'Nornal Menu
    Case "AD"
        dbgEcho "Administration Mode."
        Call RunAdminMenu     'Administrator's Menu
    Case Else
        Echo "Cancel." 
    End Select

    WScript.Sleep 3000
    WScript.Quit


'//////////////////////
' Normal Menu
'//////////////////////
Private Sub RunNormalMenu()

    If ChkWorkFolder = False Then Exit Sub
    dbgEcho "Checking distributed version."

    ' Compare the local module version and the distributing version.
    ' If New version exists, copy the new module to local.
    Echo "RunNormalMenu :Loc = " & LocModule & ", Rel =" & RelModule
    If LocModule < RelModule Then
        Echo "Copying the new version. " & LocModule & " => " & RelModule
        With CreateObject("Scripting.FileSystemObject")
            .CopyFile MenuFolder & ReleaseDef, WorkFolder 
            .CopyFile MenuFolder & RelModule , WorkFolder & ModuleName & ".accdb"
        End With
    Else
        Echo "Local module is the latest."
    End if

    ' RunLocalModule.
    Echo "Startingup " & ModuleName & "."
    CreateObject("WScript.Shell").Exec AccExe & Chr(34) & WorkFolder & ModuleName & ".accdb" & Chr(34) 

End Sub

'--------------------------------------
' Check Work folder for the client module.
'--------------------------------------
Private Function ChkWorkFolder()
Dim FS, P, PP, DLM
    Set FS = CreateObject("Scripting.FileSystemObject")
    If FS.FolderExists(WorkFolder) = False Then
        dbgEcho "Client workfolder [ " & WorkFolder & " ] does not exist."
        If MsgBox("Work Folder [" & WorkFolder & "] doen not exist." & vbCrLf & "Do you want to create work folder?",vbyesno) = vbYes Then
            If mkFolder(WorkFolder) Then
                 MsgBox WorkFolder & " is created."
                 ChkWorkFolder = True
            Else
                 MsgBox WorkFolder & " creation failed."
                 ChkWorkFolder = False
            End If
        Else
            Echo "Cancel."
            chkWorkFolder = False
        End If
    Else
        ChkWorkFolder = True
    End If
    dbgecho "chkWorkFolder ret = " & chkworkfolder
End Function

'--------------------------------------
' Create Work Folder.
' [Return]  True: success creation.
'           False: failed.
'--------------------------------------
Private Function mkFolder(FolderPath)
Dim FS, P, PP, DLM
    With CreateObject("Scripting.FileSystemObject")
        For Each P In Split(FolderPath,"\")
            If "" < P Then
                 PP = PP & DLM & P
                 DLM = "\"
                 If .FolderExists(PP) Then
                     Echo PP & " exists."
                 Else
                     .CreateFolder PP
                     Echo PP & " created. "
                 End If
            End IF
        Next
        mkFolder = .FolderExists(FolderPath)
    End With
End Function

'///////////////////
' Admin Menu
'///////////////////
Private Sub RunAdminMenu
Dim WS, Msg, Mode
    Echo "<Administration Menu>"

    Msg = Msg & "*********************************************" & vbCrLf
    Msg = Msg &  ModuleName & "<Admin mode>"                    & vbCrLf
    Msg = Msg & "Dev. :" & DevModule                            & vbCrLf
    Msg = Msg & "Rel. :" & RelModule                            & vbCrLf
    Msg = Msg & "Loc. :" & LocModule                            & vbCrLf
    Msg = Msg & "*********************************************" & vbCrLf
    Msg = Msg                                                   & vbCrLf
    Msg = Msg & " Select command"                               & vbCrLf
    Msg = Msg                                                   & vbCrLf
    Msg = Msg & "-- Run Menu module in the mode of --"          & vbCrLf
    Msg = Msg & "M: Maintain :Maintenance Mode /excl"           & vbCrLf
    Msg = Msg & "N: Normal   :Normal Menu"                      & vbCrLf
    Msg = Msg & "F: Folder   :Open Module Folder"               & vbCrLf
    Msg = Msg                                                   & vbCrLf
    Msg = Msg & "-- Module operation --"                        & vbCrLf
    Msg = Msg & "D: Develop  :Create new dev. module."          & vbCrLf
    Msg = Msg & "R: Release  :Release latest Dev. module."      & vbCrLf
    Msg = Msg & "O: Optimise :Optimise module."                 & vbCrLf
    Msg = Msg                                                   & vbCrLf
    Msg = Msg & "-: cancel." & vbCrLf

    Mode = InputBox(Msg, ModuleName,"M")
    Select Case Mode
        Case "M","m"
            Echo "Startup the Dev. module in maintenance mode."
            dbgEcho "WS.Exec " & AccExe & Chr(34) & MenuFolder & DevModule & Chr(34) & " /excl /cmd 1"
            CreateObject("WScript.Shell").Exec AccExe & Chr(34) & MenuFolder & DevModule & Chr(34) & " /excl /cmd 1"
        Case "N","n"
            Echo "Startup the normal menu."
            Call RunNormalMenu
        Case "F","f"
            Echo "Open the ALAYA Menu module folder."
            CreateObject("Shell.Application").Explore MenuFolder
        Case "R","r"
            Echo "Release the latest dev.module as the distrib. version."
            Call ReleaseDevModule
        Case "D","d"
            Echo "Create New Dev. module."
            Call CreateDevModule
        Case "O","o"
            Echo "Optimise module."
            Call OptimiseModule
        Case Else
            Echo "Cancel."
    End Select
End Sub

'--------------------------------------------------------------
' Get the development module name.
'--------------------------------------------------------------
Private Function GetDevModule(MenuFolder, ModuleName)
Dim FS, Module
   dbgEcho "GetDevModule: Enter MenuFolder=" & MenuFolder
   Set FS = CreateObject("Scripting.FileSystemObject")
   If FS.FolderExists(MenuFolder) Then 
        dbgEcho "GetDevModule: MenuFolder exixtx."
        For Each Module in FS.GetFolder(MenuFolder).Files
            dbgEcho "GetDevModule Module = " & Module.name
            If 0 < Instr(Ucase(Module.Name), UCase(ModuleName)) AND Right(Module.Name,6) = ".accdb" Then
                If GetDevModule < Module.Name Then GetDevModule = Module.Name
            End If
        Next
    Else
        MsgBox "Folder [" & MenuFolder & "] is not found."
    End If
   dbgEcho "GetDevModule: ret = " & GetDevModule
End Function

'--------------------------------------------------------------
' Ger the distribution module name.
'--------------------------------------------------------------
Private Function GetRelModule(MenuFolder, ReleaseDef)
Dim FS, FO, ReleaseDefFile
   dbgEcho "GetDevModule: Enter"
   Set FS = CreateObject("Scripting.FileSystemObject")
   ReleaseDefFile = MenuFolder & ReleaseDef

   If FS.FileExists(ReleaseDefFile) then
       Set FO = FS.OpenTextFile(ReleaseDefFile, 1, False)
       GetRelModule=FO.ReadLine
       Set FO = Nothing
   Else
       Echo "<<<< Caution !!!>>>>"
       Echo "Release definition file does not exist I"
       Echo "    File= "& ReleaseDefFile
   End If
   Set FS = Nothing
   dbgEcho "GetRelModule: ret = " & GetRelModule
End Function

'--------------------------------------------------------------
' Get the local module name.
'--------------------------------------------------------------
Private Function GetLocModule(LocalFolder, ReleaseDef)
Dim FS, FO, ReleaseDefFile, CurrentAccFile
   dbgEcho "GetLocModule: Enter"
   ReleaseDefFile = WorkFolder & ReleaseDef
   CurrentAccFile = WorkFolder & ModuleName & ".accdb"
   Set FS = CreateObject("Scripting.FileSystemObject")
   If FS.FileExists(ReleaseDefFile) And FS.FileExists(CurrentAccFile) Then
       Set FO = FS.OpenTextFile(ReleaseDefFile, 1, False)
       GetLocModule= FO.ReadLine
       Set FO = Nothing
   End If
   Set FS=Nothing
   dbgEcho "GetLocModule: ret = " & GetLocModule
End Function

'-------------------------------
' Full path of MSACCESS.exe
'-------------------------------
Private Function AccExe()
Dim AcApp 
    On Error Resume Next
    Set AcApp = CreateObject("Access.Application")
    If Err.Number = 0 then
        AccExe = AcApp.SysCmd(9) & "MSACCESS.exe "
    Else
        MsgBox "Access is not installed!"
        WScript.Quit
    End If
    Set AcApp = Nothing
End Function

'----------------------------
' Console Echo
'----------------------------
Private Sub Echo(Msg)
    WScript.Echo Msg
End Sub

Private Sub dbgEcho(Msg)
    If DBGFLG Then MsgBox Msg
End Sub


'********************************
' Create new version of the development module.
'********************************
Private Sub CreateDevModule()
Dim Msg, Ver, NewModule
    Echo "Development Module :" & DevModule
    Echo "Distribution Module:" & RelModule

    NewModule = ModuleName & "[Ver].accdb"
    Msg = "Create new version of the development module copying the latest one." & vbCrLf & _
           DevModule & " -->" & vbCrLf & _
           NewModule & vbCrLf & _
          "Enter the new version number[Ver]."
    Ver = InputBox(Msg)
    If "" < Ver Then
        NewModule = Replace(NewModule,"[Ver]",Ver)

        Msg = "Are you sure to copy as below?" & vbCrLf & _
               DevModule & " ->" & vbCrLf & _
               NewModule
        If MsgBox(Msg,vbYesNo) = vbYes Then
           CreateObject("Scripting.FileSystemObject").CopyFile MenuFolder & DevModule, MenuFolder & NewModule
            MsgBox "New module is created. Open the latest module and update the version record."
        Else
             Echo "Cancel."
        End If
    Else
        Echo "Cancel."
    End If
End Sub

'************************************************************
'  Release the latest dev.module as the distribution module.
'************************************************************
Private Sub ReleaseDevModule()
Dim Msg, FSys, FName, FO
    Msg = "Release the latest development module as the distribution module." & vbCrLf & _
          "Are you sure?"
    If MsgBox(Msg, vbYesNo) = vbYes Then
        '--Optimise the development module.
        Call DBOptimize(MenuFolder & DevModule)

        '-- Overwrite the release definision file.
        FName = MenuFolder & ReleaseDef

        Set FSys = CreateObject("Scripting.FileSystemObject")
        Set FO = FSys.OpenTextFile(FName, 2, True)
        FO.WriteLine DevModule
        Echo "'" & DevModule & "' write into " & FName
        MsgBox "Module is released. Run the normal menu and check whether new version is released."
     Else
        Echo "Cancel."
     End If
End Sub

'**********************************
' Optimise Menu the module.
'**********************************
Private Sub OptimiseModule()
Dim Msg, SelMode
    Msg = "Optimise the menu module."                    & vbCrLf _
        & "Select which one to optimise."                & vbCrLf _
        &                                                  vbCrLf _
        & " 1:Development Module : [" & DevModule & "]"  & vbCrLf  _
        & " 2:Distribition Module: [" & RelModule & "]"
    SelMode = InputBox(Msg, "DB Compact")
    '
    Select Case SelMode
    Case "1"
        Call DBOptimize(MenuFolder & DevModule)
        'MsgBox " J     W   [    œK     ܂    B" & vbCrLf & DevModule
        MsgBox "Optimise the development module." & vbCrLf & DevModule
    Case "2"
        Call DBOptimize(MenuFolder & RelModule)
        MsgBox "Optimise the distribution module." & vbCrLf & RelModule
    Case Else
        Echo "Cancel."
    End Select
End Sub

'******************
' Optimize module.
'******************
Private Sub DBOptimize(FullName)
Dim Ret, FSO, tmpDB
    Echo "Optimize module " & FullName
    tmpDB = WorkFolder & "CompactDB.accdb"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile FullName, tmpDB
    Set Ret = CreateObject("WScript.Shell").Exec(AccExe & Chr(34) & tmpDB & Chr(34) & " /compact")
    Do While Ret.Status = 0
         WScript.Sleep 500
    Loop
    If Err.Number = 0 then
        FSO.CopyFile tmpDB, FullName
        Echo "Module is optimized successful."
        FSO.DeleteFile tmpDB
    Else
        Echo "Module optimize error."
    End If
    Set FSO = Nothing
End Sub

'**************************************************************
' Reboot this script in CScript mode for prompting to console.
'**************************************************************
Private Sub RunCScript()
    If Instr(UCase(WScript.FullName),"WSCRIPT.EXE") Then
        CreateObject("WScript.Shell").Run """CScript.exe"" """ & WScript.ScriptFullName & """"
        WScript.Quit
    End If
End Sub
