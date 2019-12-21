Imports Microsoft.Win32
Module ModRegistry

    Public Sub DoRegistryMagic()
        On Error GoTo Err

        Dim DefaultFolder As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Sparkle", "DefaultFolder", "")

        If DefaultFolder = "" Then
            Dim Key As RegistryKey = Registry.CurrentUser.CreateSubKey("SOFTWARE\Sparkle")
            Key.SetValue("DefaultFolder", My.Application.Info.DirectoryPath)
            Key.Close()
            DefaultFolder = "<*>"   'Invalid file name characters to make sure we will not find them in SYSTEM PATH
        End If

        If DefaultFolder <> My.Application.Info.DirectoryPath Then

            'Get SYSTEM PATH string for user
            Dim P As String = Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.User)

            'Find Default Folder string in PATH string
            Dim I As Integer = Strings.InStr(P, DefaultFolder)

            If I > 0 Then
                Dim L, R As String
                L = Strings.Left(P, I - 1)

                R = Strings.Right(P, P.Length - (I - 1) - DefaultFolder.Length)

                'if PATH contains default Folder then replace Default Folder with Current Folder
                P = L + My.Application.Info.DirectoryPath + R
            Else
                'Otherwise, add Current Folder to PATH
                If P.Length > 0 Then
                    P = My.Application.Info.DirectoryPath + ";" + P
                Else
                    P = My.Application.Info.DirectoryPath
                End If
            End If

            'Update DefaultFolder Value in Registry with Current Folder
            Dim NewKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\Sparkle")
            NewKey.SetValue("DefaultFolder", My.Application.Info.DirectoryPath)
            NewKey.Close()

            'Save updated SYSTEM PATH for current user
            Environment.SetEnvironmentVariable("Path", P, EnvironmentVariableTarget.User)

        End If

        'Check if Sparkle is run as Administrator
        If My.User.IsInRole(ApplicationServices.BuiltInRole.Administrator) Then
            'Yes, so check file associations
            FrmMain.TsbAdmin.Visible = True
        Else
            FrmMain.TsbAdmin.Visible = False
        End If

        If Debugger.IsAttached = False Then         'Check if prg is run from IDE
            FrmMain.TsmTestDisk.Visible = False     'No, hide test disk option
            FrmMain.TssSep.Visible = False
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub AssociateSLS()
        On Error GoTo Err

        If My.Computer.Registry.ClassesRoot.OpenSubKey("Sparkle Loader Script\shell\open\command", True) Is Nothing Then
            If MsgBox("Do you want to associate the .sls file extension with Sparkle?", vbYesNo + vbQuestion, "Sparkle Admin Mode") = vbYes Then
                My.Computer.Registry.ClassesRoot.CreateSubKey(".sls").SetValue("", "Sparkle Loader Script", Microsoft.Win32.RegistryValueKind.String)
                My.Computer.Registry.ClassesRoot.CreateSubKey("Sparkle Loader Script\shell\open\command").SetValue("", Application.ExecutablePath & " ""%l"" ", Microsoft.Win32.RegistryValueKind.String)

                Using SK1 As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32).OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.sls")
                    If SK1 IsNot Nothing Then
                        My.Computer.Registry.CurrentUser.DeleteSubKeyTree("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.sls")
                        SK1.Close()
                    End If
                End Using
            Else
                Exit Sub
            End If
        Else
            If MsgBox("The .sls file extension is already associated with Sparkle." + vbNewLine + vbNewLine + "Do you want to update it?", vbYesNo + vbQuestion, "Sparkle Admin Mode") = vbYes Then
                My.Computer.Registry.ClassesRoot.CreateSubKey(".sls").SetValue("", "Sparkle Loader Script", Microsoft.Win32.RegistryValueKind.String)
                My.Computer.Registry.ClassesRoot.CreateSubKey("Sparkle Loader Script\shell\open\command").SetValue("", Application.ExecutablePath & " ""%l"" ", Microsoft.Win32.RegistryValueKind.String)

                Using SK2 As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32).OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.sls")
                    If SK2 IsNot Nothing Then
                        My.Computer.Registry.CurrentUser.DeleteSubKeyTree("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.sls")
                        SK2.Close()
                    End If
                End Using
            Else
                Exit Sub
            End If
        End If

        MsgBox("The .sls file extension has been successfully associated with Sparkle!", vbOKOnly + vbInformation, "Sparkle Admin Mode")

        Exit Sub
Err:

        MsgBox("Error during creating .sls file association with Sparkle" + vbNewLine + vbNewLine + ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub DeleteAssociation()
        On Error GoTo Err

        If MsgBox("Do you want to delete the .sls file association with Sparkle?", vbYesNo + vbQuestion, "Sparkle Admin Mode") = vbYes Then

            Using SK1 As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32).OpenSubKey("Sparkle Loader Script\shell\open\command")
                If SK1 IsNot Nothing Then
                    My.Computer.Registry.ClassesRoot.DeleteSubKeyTree("Sparkle Loader Script")
                    SK1.Close()
                End If
            End Using

            Using SK2 As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32).OpenSubKey(".sls")
                If SK2 IsNot Nothing Then
                    My.Computer.Registry.ClassesRoot.DeleteSubKeyTree(".sls")
                    SK2.Close()
                End If
            End Using

            Using SK3 As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32).OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.sls")
                If SK3 IsNot Nothing Then
                    My.Computer.Registry.CurrentUser.DeleteSubKeyTree("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.sls")
                    SK3.Close()
                End If
            End Using

            MsgBox("The .sls file extension is no longer associated with Sparkle.", vbOKOnly + vbInformation, "Sparkle Admin Mode")
        End If

        Exit Sub
Err:

        MsgBox("Error during deleting .sls file association with Sparkle" + vbNewLine + vbNewLine + ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Function DotNetVersion() As Boolean
        On Error GoTo Err

        Const SubKey As String = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"

        Using NDPKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(SubKey)
            If NDPKey IsNot Nothing AndAlso NDPKey.GetValue("Release") IsNot Nothing Then
                DotNetVersion = CheckFor45PlusVersion(NDPKey.GetValue("Release"))
            Else
                DotNetVersion = False
            End If
        End Using

		Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    'Checking the version using >= will enable forward compatibility.
    Private Function CheckFor45PlusVersion(releaseKey As Integer) As Boolean
        On Error GoTo Err

        If releaseKey >= 378389 Then
            CheckFor45PlusVersion = True
        Else
            CheckFor45PlusVersion = False
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
    End Function

End Module
