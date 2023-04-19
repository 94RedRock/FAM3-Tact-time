Imports System.Data.OleDb
Public Class ExcelConnection
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Create a connection where first row contains column names
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <param name="IMEX"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()>
    Public Function HeaderConnectionString(ByVal FileName As String, Optional ByVal IMEX As Integer = 1) As String
        Dim Builder As New OleDbConnectionStringBuilder
        If IO.Path.GetExtension(FileName).ToUpper = ".XLS" Then
            Builder.Provider = "Microsoft.Jet.OLEDB.4.0"
            Builder.Add("Extended Properties", String.Format("Excel 8.0;IMEX={0};HDR=Yes;", IMEX))
        Else
            Builder.Provider = "Microsoft.ACE.OLEDB.12.0"
            Builder.Add("Extended Properties", String.Format("Excel 12.0;IMEX={0};HDR=Yes;", IMEX))
        End If

        Builder.DataSource = FileName

        Return Builder.ToString

    End Function
    ''' <summary>
    ''' Create a connection where first row contains data
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <param name="IMEX"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()>
    Public Function NoHeaderConnectionString(ByVal FileName As String, Optional ByVal IMEX As Integer = 1) As String
        Dim Builder As New OleDbConnectionStringBuilder
        If IO.Path.GetExtension(FileName).ToUpper = ".XLS" Then
            Builder.Provider = "Microsoft.Jet.OLEDB.4.0"
            Builder.Add("Extended Properties", String.Format("Excel 8.0;IMEX={0};HDR=No;", IMEX))
        Else
            Builder.Provider = "Microsoft.ACE.OLEDB.12.0"
            Builder.Add("Extended Properties", String.Format("Excel 12.0;IMEX={0};HDR=No;", IMEX))
        End If

        Builder.DataSource = FileName

        Return Builder.ToString

    End Function

    ''' <summary>
    ''' Create a connection where first row contains data
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()>
    Public Function WriteConnectionString(ByVal FileName As String) As String
        Dim Builder As New OleDbConnectionStringBuilder
        If IO.Path.GetExtension(FileName).ToUpper = ".XLS" Then
            Builder.Provider = "Microsoft.Jet.OLEDB.4.0"
            Builder.Add("Extended Properties", String.Format("Excel 8.0;HDR=YES;"))
        Else
            Builder.Provider = "Microsoft.ACE.OLEDB.12.0"
            Builder.Add("Extended Properties", String.Format("Excel 12.0 Xml;HDR=YES;Text;CharacterSet=UTF-8"))
        End If

        Builder.DataSource = FileName

        Return Builder.ToString

    End Function
End Class
