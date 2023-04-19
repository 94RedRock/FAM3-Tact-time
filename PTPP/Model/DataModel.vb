Public Class DataModel

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''MEMBER VARIABLES
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "MEMBER_VARIABLES"

    ''' <summary>
    ''' Model Name
    ''' </summary>
    Private _model As List(Of String) = New List(Of String)

    ''' <summary>
    ''' 部品 SET
    ''' </summary>
    Private _component_set As List(Of String) = New List(Of String)

    ''' <summary>
    ''' 前付け
    ''' </summary>
    Private _maedzuke As List(Of String) = New List(Of String)

    ''' <summary>
    ''' MT
    ''' </summary>
    Private _mount As List(Of String) = New List(Of String)

    ''' <summary>
    ''' L/C
    ''' </summary>
    Private _lead_cutting As List(Of String) = New List(Of String)

    ''' <summary>
    ''' 目視
    ''' </summary>
    Private _visual_examination As List(Of String) = New List(Of String)

    ''' <summary>
    ''' pickup
    ''' </summary>
    Private _pickup As List(Of String) = New List(Of String)

    ''' <summary>
    ''' 組立
    ''' </summary>
    Private _assambly As List(Of String) = New List(Of String)

    ''' <summary>
    ''' 機能検査
    ''' </summary>
    Private _function_check As List(Of String) = New List(Of String)

    ''' <summary>
    ''' 2者検査
    ''' </summary>
    Private _person_examine As List(Of String) = New List(Of String)

    ''' <summary>
    ''' examine_time(검사 시간)
    ''' </summary>
    Private _examine_time As List(Of String) = New List(Of String)

#End Region

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''PROPERTY
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "PROPERTY"

    ''' <summary>
    ''' Model Name
    ''' </summary>
    ''' <returns></returns>
    Public Property Model As List(Of String)
        Get
            Return _model
        End Get
        Set(ByVal value As List(Of String))
            _model = value
        End Set
    End Property

    ''' <summary>
    ''' 部品 SET
    ''' </summary>
    ''' <returns></returns>
    Public Property ComponentSet As List(Of String)
        Get
            Return _component_set
        End Get
        Set(value As List(Of String))
            _component_set = value
        End Set
    End Property

    ''' <summary>
    ''' 前付け
    ''' </summary>
    ''' <returns></returns>
    Public Property Maedzuke As List(Of String)
        Get
            Return _maedzuke
        End Get
        Set(value As List(Of String))
            _maedzuke = value
        End Set
    End Property

    ''' <summary>
    ''' MT
    ''' </summary>
    ''' <returns></returns>
    Public Property Maunt As List(Of String)
        Get
            Return _mount
        End Get
        Set(value As List(Of String))
            _mount = value
        End Set
    End Property

    ''' <summary>
    ''' L/C
    ''' </summary>
    ''' <returns></returns>
    Public Property LeadCutting As List(Of String)
        Get
            Return _lead_cutting
        End Get
        Set(value As List(Of String))
            _lead_cutting = value
        End Set
    End Property

    ''' <summary>
    ''' 目視
    ''' </summary>
    ''' <returns></returns>
    Public Property VisualExamination As List(Of String)
        Get
            Return _visual_examination
        End Get
        Set(value As List(Of String))
            _visual_examination = value
        End Set
    End Property

    'Public Property 
#End Region

End Class
