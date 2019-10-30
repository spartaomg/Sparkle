Public Class FrmDisk
    Private Sub FrmDisk_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo Err

        pbx.Top = (Height - pbx.Height) / 2

        lbl.Left = pbx.Left + pbx.Width + (Width - lbl.Width - pbx.Left - pbx.Width) / 2
        lbl.Top = (Height - lbl.Height) / 2

        Cursor = Cursors.WaitCursor

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FrmDisk_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        On Error GoTo Err

        Refresh()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub
End Class