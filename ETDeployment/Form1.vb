Option Explicit On

Public Class Form1

    Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String,
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

    Private Declare Function GetDesktopWindow Lib "user32" () As Long

    Const SW_SHOWNORMAL = 1

    Const SE_ERR_FNF = 2&
    Const SE_ERR_PNF = 3&
    Const SE_ERR_ACCESSDENIED = 5&
    Const SE_ERR_OOM = 8&
    Const SE_ERR_DLLNOTFOUND = 32&
    Const SE_ERR_SHARE = 26&
    Const SE_ERR_ASSOCINCOMPLETE = 27&
    Const SE_ERR_DDETIMEOUT = 28&
    Const SE_ERR_DDEFAIL = 29&
    Const SE_ERR_DDEBUSY = 30&
    Const SE_ERR_NOASSOC = 31&
    Const ERROR_BAD_FORMAT = 11&

    Function StartDoc(DocName As String) As Long
        Dim Scr_hDC As Long
        Scr_hDC = GetDesktopWindow()
        StartDoc = ShellExecute(Scr_hDC, "Open", DocName,
          "", "C:\", SW_SHOWNORMAL)
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "EnableTransparency", Nothing) = 1 Then CheckedListBox1.SetItemChecked(0, 1)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "LaunchTo", Nothing) = 0 Then CheckedListBox1.SetItemChecked(1, 1)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = ""

        If CheckedListBox1.GetItemCheckState(0).ToString = "Checked" Then

            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "EnableTransparency", Nothing) = 0 Then
                TextBox1.AppendText("Transparency is disabled. Enabling." + vbNewLine)
                Try
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "EnableTransparency", 1)
                    TextBox1.AppendText("Transparency is now enabled." + vbNewLine)
                    TextBox1.AppendText("Restarting Explorer to apply changes." + vbNewLine)
                    For Each p As Process In Process.GetProcesses()
                        Try
                            If p.MainModule.FileName.ToLower().EndsWith(":\windows\explorer.exe") Then
                                p.Kill()
                                Exit Try
                            End If
                        Catch ex As Exception
                            TextBox1.AppendText("CRITICAL ERROR:" + ex.ToString + vbNewLine)
                        End Try
                    Next
                    TextBox1.AppendText("Explorer restarted." + vbNewLine)
                Catch ex As Exception
                    TextBox1.AppendText("CRITICAL ERROR:" + ex.ToString + vbNewLine)
                    TextBox1.AppendText("Fatal error. Aborting." + vbNewLine)
                End Try
            Else
                TextBox1.AppendText("Transparency is already enabled. Continuing." + vbNewLine)
            End If

        Else

            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "EnableTransparency", Nothing) = 1 Then
                TextBox1.AppendText("Transparency is enabled. Disabling." + vbNewLine)
                Try
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "EnableTransparency", 0)
                    TextBox1.AppendText("Transparency is now disabled." + vbNewLine)
                    TextBox1.AppendText("Restarting Explorer to apply changes." + vbNewLine)
                    For Each p As Process In Process.GetProcesses()
                        Try
                            If p.MainModule.FileName.ToLower().EndsWith(":\windows\explorer.exe") Then
                                p.Kill()
                                Exit Try
                            End If
                        Catch ex As Exception
                            TextBox1.AppendText("CRITICAL ERROR:" + ex.ToString + vbNewLine)
                        End Try
                    Next
                    TextBox1.AppendText("Explorer restarted." + vbNewLine)
                Catch ex As Exception
                    TextBox1.AppendText("CRITICAL ERROR:" + ex.ToString + vbNewLine)
                    TextBox1.AppendText("Fatal error. Aborting." + vbNewLine)
                End Try
            Else
                TextBox1.AppendText("Transparency is already disabled. Continuing." + vbNewLine)
            End If

        End If

        If CheckedListBox1.GetItemCheckState(1) = 1 Then
            TextBox1.AppendText("Setting File Explorer to open This PC." + vbNewLine)
            Try
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "LaunchTo", 0)
                TextBox1.AppendText("File Explorer is now set to open This PC." + vbNewLine)
            Catch ex As Exception
                TextBox1.AppendText("CRITICAL ERROR:" + ex.ToString + vbNewLine)
                TextBox1.AppendText("Fatal error. Aborting." + vbNewLine)
            End Try
        Else
            TextBox1.AppendText("Setting File Explorer to open Quick Access." + vbNewLine)
            Try
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "LaunchTo", 1)
                TextBox1.AppendText("File Explorer is now set to open Quick Access. You monster." + vbNewLine)
            Catch ex As Exception
                TextBox1.AppendText("CRITICAL ERROR:" + ex.ToString + vbNewLine)
                TextBox1.AppendText("Fatal error. Aborting." + vbNewLine)
            End Try
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox1.AppendText("Copying theme to Resources." + vbNewLine)
        Try
            Dim r As Long, msg As String
            r = StartDoc(My.Application.Info.DirectoryPath + "\ETTheme.deskthemepack")
            If r <= 32 Then
                Select Case r
                    Case SE_ERR_FNF
                        msg = "File not found: " + My.Application.Info.DirectoryPath + "\ETTheme.deskthemepack"
                    Case SE_ERR_PNF
                        msg = "Path not found."
                    Case SE_ERR_ACCESSDENIED
                        msg = "Access denied."
                    Case SE_ERR_OOM
                        msg = "Out of memory."
                    Case SE_ERR_DLLNOTFOUND
                        msg = "DLL not found."
                    Case SE_ERR_SHARE
                        msg = "A sharing violation occurred."
                    Case SE_ERR_ASSOCINCOMPLETE
                        msg = "Incomplete or invalid file association."
                    Case SE_ERR_DDETIMEOUT
                        msg = "DDE Time out."
                    Case SE_ERR_DDEFAIL
                        msg = "DDE transaction failed."
                    Case SE_ERR_DDEBUSY
                        msg = "DDE busy."
                    Case SE_ERR_NOASSOC
                        msg = "No association for file extension."
                    Case ERROR_BAD_FORMAT
                        msg = "Invalid EXE file or error in EXE image."
                    Case Else
                        msg = "Unknown error."
                End Select
                TextBox1.AppendText("CRITICAL ERROR: " + msg + vbNewLine)
                TextBox1.AppendText("Fatal error. Aborting." + vbNewLine)
                Exit Try
            End If
            TextBox1.AppendText("Theme copied successfully." + vbNewLine)
            TextBox1.AppendText("Opening Personalization window to apply theme." + vbNewLine)
            TextBox1.AppendText("Theme applied successfully." + vbNewLine)
        Catch ex As Exception
            TextBox1.AppendText("CRITICAL ERROR:" + ex.ToString + vbNewLine)
            TextBox1.AppendText("Fatal error. Aborting." + vbNewLine)
        End Try

    End Sub

End Class
