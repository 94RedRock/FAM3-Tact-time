Public Class Lock
    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        If (txtPassword.Text = ProgramConfig.ReadIniUserSetting("Password")) Then
            Me.DialogResult = DialogResult.OK
        Else
            MessageBox.Show("비밀번호를 확인하여 주십시오.", "warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

End Class