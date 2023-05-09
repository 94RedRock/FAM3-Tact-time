Public Class DataCenterCarrier

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''MEMBER VARIABLES
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "MEMBER_VARIABLES"

    ''' <summary>
    ''' carrier name
    ''' </summary>
    Private _carrier As String

    ''' <summary>
    ''' 제한대수
    ''' </summary>
    Private _limit As String

    ''' <summary>
    ''' 캐리어 수량
    ''' </summary>
    Private _quantity As String


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
    Public Property Carrier As String
        Get
            Return _carrier
        End Get
        Set(value As String)
            _carrier = value
        End Set
    End Property

    ''' <summary>
    ''' LIMIT
    ''' </summary>
    ''' <returns></returns>
    Public Property Limit As String
        Get
            Return _limit
        End Get
        Set(value As String)
            _limit = value
        End Set
    End Property

    ''' <summary>
    ''' QUANTITY
    ''' </summary>
    ''' <returns></returns>
    Public Property Quantity As String
        Get
            Return _quantity
        End Get
        Set(value As String)
            _quantity = value
        End Set
    End Property


#End Region

    Public Sub New(carrier As String, limit As String, quantity As String)
        _carrier = carrier
        _limit = limit
        _quantity = quantity
    End Sub

End Class
