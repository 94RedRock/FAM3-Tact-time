Public Class DataCenterLimit

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''MEMBER VARIABLES
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "MEMBER_VARIABLES"

    ''' <summary>
    ''' MODEL
    ''' </summary>
    Private _model_limit As String

    ''' <summary>
    ''' CARRIER
    ''' </summary>
    Private _carrier_limit As String




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
    Public Property ModelLimit As String
        Get
            Return _model_limit
        End Get
        Set(value As String)
            _model_limit = value
        End Set
    End Property

    ''' <summary>
    ''' ADDITIONAL_MAOUNTING
    ''' </summary>
    ''' <returns></returns>
    Public Property CarrierLimit As String
        Get
            Return _carrier_limit
        End Get
        Set(value As String)
            _carrier_limit = value
        End Set
    End Property



#End Region

    Public Sub New(model_limit As String, carrier_limit As String)
        _model_limit = model_limit
        _carrier_limit = carrier_limit
    End Sub

End Class
