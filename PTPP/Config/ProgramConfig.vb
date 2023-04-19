Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.IO
Public Class ProgramConfig

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''MEMBER FUNCTIONS
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "MEMBER_FUNCTIONS"

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''' Read
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    <DllImport("kernel32")>
    Private Shared Function GetPrivateProfileString(ByVal section As String, ByVal key As String, ByVal defVal As String, ByVal retVal As StringBuilder, ByVal size As Integer, ByVal filePath As String) As Integer
    End Function

    ''' <summary>
    ''' Read Setting iniFile
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    Public Shared Function ReadIniDBSetting(ByVal key As String) As String
        Dim value As StringBuilder = New StringBuilder(128)
        GetPrivateProfileString("DBSetting", key, String.Empty, value, 128, ProgramDefine.INI_FILE_PATH)
        Return value.ToString()
    End Function

    ''' <summary>
    ''' Read Setting iniFile
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    Public Shared Function ReadIniUserSetting(ByVal key As String) As String
        Dim value As StringBuilder = New StringBuilder(128)
        GetPrivateProfileString("UserSetting", key, String.Empty, value, 128, ProgramDefine.INI_FILE_PATH)
        Return value.ToString()
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''Write
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    <DllImport("kernel32")>
    Private Shared Function WritePrivateProfileString(ByVal section As String, ByVal key As String, ByVal val As String, ByVal filePath As String) As Long
    End Function

    ''' <summary>
    ''' Write Setting iniFile
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="value"></param>
    Public Shared Sub WriteIniSetting(ByVal key As String, ByVal value As String)
        WritePrivateProfileString("UserSetting", key, value, ProgramDefine.INI_FILE_PATH)
    End Sub

#End Region
End Class
