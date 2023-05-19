Public Class DataCenterLimit

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''MEMBER VARIABLES
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "MEMBER_VARIABLES"

    ''' <summary>
    ''' CARRIER
    ''' </summary>
    Private _carrier As String

    ''' <summary>
    ''' limit
    ''' </summary>
    Private _limit As String
    ''' <summary>
    ''' quantity 
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
    ''' 캐리어명
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
    ''' 제한대수
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
    ''' 수량
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
