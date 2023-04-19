Public Class DataCenter

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''
    ''MEMBER VARIABLES
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#Region "MEMBER_VARIABLES"

    ''' <summary>
    ''' Model Name
    ''' </summary>
    Private _model As String

    ''' <summary>
    ''' 部品 SET
    ''' </summary>
    Private _componentSet As String

    ''' <summary>
    ''' 前付け
    ''' </summary>
    Private _maedzuke As String

    ''' <summary>
    ''' MT
    ''' </summary>
    Private _mount As String

    ''' <summary>
    ''' L/C
    ''' </summary>
    Private _leadCutting As String

    ''' <summary>
    ''' 目視
    ''' </summary>
    Private _visualExamination As String

    ''' <summary>
    ''' pickup
    ''' </summary>
    Private _pickup As String

    ''' <summary>
    ''' 組立
    ''' </summary>
    Private _assambly As String

    ''' <summary>
    ''' 機能検査(수동)
    ''' </summary>
    Private _mFunctionCheck As String

    ''' <summary>
    ''' 機能検査(자동)
    ''' </summary>
    Private _aFunctionCheck As String

    ''' <summary>
    ''' 2者検査
    ''' </summary>
    Private _personExamine As String

    ''' <summary>
    ''' examineEquipment(검사 장치)
    ''' </summary>
    Private _examineEquipment As String

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
    Public Property Model As String
        Get
            Return _model
        End Get
        Set(value As String)
            _model = value
        End Set
    End Property

    ''' <summary>
    ''' 部品 SET
    ''' </summary>
    ''' <returns></returns>
    Public Property ComponentSet As String
        Get
            Return _componentSet
        End Get
        Set(value As String)
            _componentSet = value
        End Set
    End Property

    ''' <summary>
    ''' 前付け
    ''' </summary>
    ''' <returns></returns>
    Public Property Maedzuke As String
        Get
            Return _maedzuke
        End Get
        Set(value As String)
            _maedzuke = value
        End Set
    End Property

    ''' <summary>
    ''' MT
    ''' </summary>
    ''' <returns></returns>
    Public Property Mount As String
        Get
            Return _mount
        End Get
        Set(value As String)
            _mount = value
        End Set
    End Property

    ''' <summary>
    ''' L/C
    ''' </summary>
    ''' <returns></returns>
    Public Property LeadCutting As String
        Get
            Return _leadCutting
        End Get
        Set(value As String)
            _leadCutting = value
        End Set
    End Property

    ''' <summary>
    ''' 目視
    ''' </summary>
    ''' <returns></returns>
    Public Property VisualExamination As String
        Get
            Return _visualExamination
        End Get
        Set(value As String)
            _visualExamination = value
        End Set
    End Property

    ''' <summary>
    ''' Pickup
    ''' </summary>
    ''' <returns></returns>
    Public Property PickUp As String
        Get
            Return _pickup
        End Get
        Set(value As String)
            _pickup = value
        End Set
    End Property

    ''' <summary>
    ''' 조립
    ''' </summary>
    ''' <returns></returns>
    Public Property Assambly As String
        Get
            Return _assambly
        End Get
        Set(value As String)
            _assambly = value
        End Set
    End Property

    ''' <summary>
    ''' 기능검사(수동)
    ''' </summary>
    ''' <returns></returns>
    Public Property MFunctionCheck As String
        Get
            Return _mFunctionCheck
        End Get
        Set(value As String)
            _mFunctionCheck = value
        End Set
    End Property

    ''' <summary>
    ''' 기능검사(자동)
    ''' </summary>
    ''' <returns></returns>
    Public Property AFunctionCheck As String
        Get
            Return _aFunctionCheck
        End Get
        Set(value As String)
            _aFunctionCheck = value
        End Set
    End Property

    ''' <summary>
    ''' 이자검사
    ''' </summary>
    ''' <returns></returns>
    Public Property PersonalExamine As String
        Get
            Return _personExamine
        End Get
        Set(value As String)
            _personExamine = value
        End Set
    End Property

    ''' <summary>
    ''' 검사설비
    ''' </summary>
    ''' <returns></returns>
    Public Property ExamineEquipment As String
        Get
            Return _examineEquipment
        End Get
        Set(value As String)
            _examineEquipment = value
        End Set
    End Property
#End Region

    Public Sub New(model As String, componentSet As String, maedzuke As String, mount As String, leadCutting As String, visualExamination As String, pickup As String, assambly As String, mFunctionCheck As String, aFunctionCheck As String, personExamine As String, examineEquipment As String)
        _model = model
        _componentSet = componentSet
        _maedzuke = maedzuke
        _mount = mount
        _leadCutting = leadCutting
        _visualExamination = visualExamination
        _pickup = pickup
        _assambly = assambly
        _mFunctionCheck = mFunctionCheck
        _aFunctionCheck = aFunctionCheck
        _personExamine = personExamine
        _examineEquipment = examineEquipment
    End Sub

End Class
