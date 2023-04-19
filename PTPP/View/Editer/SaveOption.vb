Imports System.Windows.Forms.CommonDialog

Public Class SaveOption

    Sub New()
        InitializeComponent()
        ReadIniFile()
    End Sub

    ''' <summary>
    ''' Apply Click Event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        APPIYSetting()
    End Sub

    ''' <summary>
    ''' Cancel Click Event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' Setting Apply
    ''' </summary>
    Private Sub APPIYSetting()
        If txtInputFilepath.Text = String.Empty Or txtSaveFilePath.Text = String.Empty Then
            MessageBox.Show("파일경로를 선택해 주십시오.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            WriteIniFile()
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' Read iniFile
    ''' </summary>
    Private Sub ReadIniFile()
        Me.txtInputFilepath.Text = ProgramConfig.ReadIniUserSetting("InputFilePath")
        Me.txtSaveFilePath.Text = ProgramConfig.ReadIniUserSetting("SaveFilePath")
    End Sub

    ''' <summary>
    ''' Write iniFile
    ''' </summary>
    Private Sub WriteIniFile()
        ProgramConfig.WriteIniSetting("InputFilePath", Me.txtInputFilepath.Text)
        ProgramConfig.WriteIniSetting("SaveFilePath", Me.txtSaveFilePath.Text)
    End Sub

    ''' <summary>
    ''' InputFile Select
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnInputFileSelect_Click(sender As Object, e As EventArgs) Handles btnInputFileSelect.Click
        Dim saveFileDailog As FolderBrowserDialog = New FolderBrowserDialog()

        If saveFileDailog.ShowDialog() = DialogResult.OK Then
            Me.txtInputFilepath.Text = saveFileDailog.SelectedPath.ToString()
        End If
    End Sub

    ''' <summary>
    ''' OutputFile Select
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnOutputFileSelect_Click(sender As Object, e As EventArgs) Handles btnOutputFileSelect.Click
        Dim saveFileDailog As FolderBrowserDialog = New FolderBrowserDialog()

        If saveFileDailog.ShowDialog() = DialogResult.OK Then
            Me.txtSaveFilePath.Text = saveFileDailog.SelectedPath.ToString()
        End If
    End Sub
End Class