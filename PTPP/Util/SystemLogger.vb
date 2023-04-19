Imports log4net
Imports System.Windows.Forms
Public NotInheritable Class SystemLogger
    Public Shared ReadOnly Property Instance As SystemLogger
        Get

            If _instance Is Nothing Then

                SyncLock _syncRoot
                    If _instance Is Nothing Then _instance = New SystemLogger()
                End SyncLock
            End If

            Return _instance
        End Get
    End Property

    Private Shared _instance As SystemLogger = Nothing
    Private Shared _syncRoot As Object = New Object()
    Private _fileLogger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub ErrorLog(ByVal logType As ProgramEnum.LogType, ByVal sender As String, ByVal message As String)
        Select Case logType
            Case ProgramEnum.LogType.Display
                DisplayErrorLog(message)
            Case ProgramEnum.LogType.File
                FileErrorLog(sender, message)
            Case ProgramEnum.LogType.Both
                DisplayErrorLog(message)
                FileErrorLog(sender, message)
        End Select
    End Sub

    Private Sub DisplayErrorLog(ByVal message As String)
        MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
    End Sub

    Private Sub FileErrorLog(ByVal sender As String, ByVal message As String)
        If _fileLogger.IsErrorEnabled Then
            Dim msg As String = String.Format("{0} : {1}", sender, message)
            _fileLogger.[Error](msg)
        End If
    End Sub

    Public Sub InfoLog(ByVal logType As ProgramEnum.LogType, ByVal sender As String, ByVal message As String)
        Select Case logType
            Case ProgramEnum.LogType.Display
                DisplayInfoLog(message)
            Case ProgramEnum.LogType.File
                FileInfoLog(sender, message)
            Case ProgramEnum.LogType.Both
                DisplayInfoLog(message)
                FileInfoLog(sender, message)
        End Select
    End Sub

    Private Sub DisplayInfoLog(ByVal message As String)
        MessageBox.Show(message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub FileInfoLog(ByVal sender As String, ByVal message As String)
        If _fileLogger.IsInfoEnabled Then
            Dim msg As String = String.Format("{0} : {1}", sender, message)
            _fileLogger.Info(msg)
        End If
    End Sub

    Public Sub WarningLog(ByVal logType As ProgramEnum.LogType, ByVal sender As String, ByVal message As String)
        Select Case logType
            Case ProgramEnum.LogType.Display
                DisplayWarningLog(message)
            Case ProgramEnum.LogType.File
                FileWarningLog(sender, message)
            Case ProgramEnum.LogType.Both
                DisplayWarningLog(message)
                FileWarningLog(sender, message)
        End Select
    End Sub

    Private Sub DisplayWarningLog(ByVal message As String)
        MessageBox.Show(message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub FileWarningLog(ByVal sender As String, ByVal message As String)
        If _fileLogger.IsWarnEnabled Then
            Dim msg As String = String.Format("{0} : {1}", sender, message)
            _fileLogger.Warn(msg)
        End If
    End Sub

    Public Shared Sub Shutdown()
        LogManager.Shutdown()
    End Sub

End Class
