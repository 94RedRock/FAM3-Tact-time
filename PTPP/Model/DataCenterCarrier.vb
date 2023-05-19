Public Class DataCenterCarrier

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''MEMBER VARIABLES
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "MEMBER_VARIABLES"

    ''' <summary>
    ''' 모델
    ''' </summary>
    Private _carrier_model As String

    ''' <summary>
    ''' 캐리어명
    ''' </summary>
    Private _carrier_carrier As String


#End Region


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''PROPERTY
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "PROPERTY"

    ''' <summary>
    ''' CARRIER
    ''' </summary>
    ''' <returns></returns>
    Public Property CarrierModel As String
        Get
            Return _carrier_model
        End Get
        Set(value As String)
            _carrier_model = value
        End Set
    End Property

    ''' <summary>
    ''' LIMIT
    ''' </summary>
    ''' <returns></returns>
    Public Property CarrierCarrier As String
        Get
            Return _carrier_carrier
        End Get
        Set(value As String)
            _carrier_carrier = value
        End Set
    End Property

#End Region

    Public Sub New(carrierModel As String, carrierCarrier As String)
        _carrier_model = carrierModel
        _carrier_carrier = carrierCarrier
    End Sub

End Class
