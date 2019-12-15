Public Class FrmAbout
	Private Sub FrmAbout_Click(sender As Object, e As EventArgs) Handles Me.Click
		On Error GoTo Err

		Close()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FrmAbout_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
		On Error GoTo Err

		Close()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FrmAbout_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
		On Error GoTo Err

		Close()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FrmAbout_Load(sender As Object, e As EventArgs) Handles MyBase.Load

		lblDescription.Left = (Width - lblDescription.Width) / 2
		lblMe.Left = (Width - lblMe.Width) / 2
		With LblVersion
			.Text = "Version: " + My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString '+ "." + My.Application.Info.Version.Build.ToString
			'.Text = "Version: " + My.Application.Info.Version.ToString
			.Refresh()
			.Left = (Width - .Width) / 2
		End With
		'lblRights.Left = (Width - lblRights.Width) / 2

	End Sub

	Private Sub LblName_Click(sender As Object, e As EventArgs) Handles LblName.Click
		On Error GoTo Err

		Close()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub lblDescription_Click(sender As Object, e As EventArgs) Handles lblDescription.Click
		On Error GoTo Err

		Close()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub lblMe_Click(sender As Object, e As EventArgs) Handles lblMe.Click
		On Error GoTo Err

		Close()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub LblVersion_Click(sender As Object, e As EventArgs) Handles LblVersion.Click
		On Error GoTo Err

		Close()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub PbxLogo_Click(sender As Object, e As EventArgs) Handles PbxLogo.Click
		On Error GoTo Err

		Close()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

End Class