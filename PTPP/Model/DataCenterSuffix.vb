Public Class DataCenterSuffix

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''MEMBER VARIABLES
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "MEMBER_VARIABLES"

    ''' <summary>
    ''' suffix
    ''' </summary>
    Private _suffix As String

    ''' <summary>
    ''' ADDITIONAL_MAOUNTING
    ''' </summary>
    Private _additional_mount As String

    ''' <summary>
    ''' ADDITIONAL_ASSEMBLY
    ''' </summary>
    Private _additional_assembly As String


#End Region


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''PROPERTY
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "PROPERTY"

    ''' <summary>
    ''' Suffix
    ''' </summary>
    ''' <returns></returns>
    Public Property Suffix As String
        Get
            Return _suffix
        End Get
        Set(value As String)
            _suffix = value
        End Set
    End Property

    ''' <summary>
    ''' ADDITIONAL_MAOUNTING
    ''' </summary>
    ''' <returns></returns>
    Public Property AdditionalMount As String
        Get
            Return _additional_mount
        End Get
        Set(value As String)
            _additional_mount = value
        End Set
    End Property

    ''' <summary>
    ''' 前付け
    ''' </summary>
    ''' <returns></returns>
    Public Property AdditionalAssembly As String
        Get
            Return _additional_assembly
        End Get
        Set(value As String)
            _additional_assembly = value
        End Set
    End Property


#End Region

    Public Sub New(suffix As String, additional_mount As String, additional_assembly As String)
        _suffix = suffix
        _additional_mount = additional_mount
        _additional_assembly = additional_assembly
    End Sub

End Class
