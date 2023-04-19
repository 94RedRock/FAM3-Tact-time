Imports System.Data.SqlClient
Imports System.Data
Imports System.Threading
Imports log4net

Public Class MainView

    Sub New()
        InitializeComponent()

        log4net.Config.XmlConfigurator.Configure()
        SystemLogger.Instance.InfoLog(ProgramEnum.LogType.File, "Main()", "===== Execute PTPP =====")
    End Sub

    ''' <summary>
    ''' Close Program
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MainView_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SystemLogger.Instance.InfoLog(ProgramEnum.LogType.File, "MainView_FormClosing()", "===== Exit PTPP =====")
        SystemLogger.Shutdown()
    End Sub

    Private Sub UserControl1_Load(sender As Object, e As EventArgs) Handles UserControl1.Load

    End Sub
End Class
