﻿Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data.OleDb


Public Class userControl

    Private _userControlPresenter As userControlPresenter = Nothing
    Private Connection As ExcelConnection = New ExcelConnection
    Private detailTable As New System.Data.DataTable
    Private simpleTable As New System.Data.DataTable
    Private inspectionTable As New System.Data.DataTable
    Private DbTable As New System.Data.DataTable
    Private newDBTable As New System.Data.DataTable
    Public masterDatalist As New List(Of DataCenter)
    Public masterDatalistSuffix As New List(Of DataCenterSuffix)
    Public masterDatalistCarrier As New List(Of DataCenterCarrier)
    Public masterDatalistLimit As New List(Of DataCenterLimit)

    Public Sub New()
        'ByVal DbType As Object

        InitializeComponent()

        'End If
        _userControlPresenter = New userControlPresenter(Me)
        detailTable.Columns.Add("No", GetType(Int32))
        detailTable.Columns.Add("Model", GetType(String))
        detailTable.Columns.Add("부속품", GetType(String))     'hsj test
        detailTable.Columns.Add("部品SET", GetType(String))
        detailTable.Columns.Add("前付け", GetType(String))
        detailTable.Columns.Add("MT", GetType(String))
        detailTable.Columns.Add("LC", GetType(String))
        detailTable.Columns.Add("目視", GetType(String))
        detailTable.Columns.Add("추가마운팅", GetType(String))  'hsj test
        detailTable.Columns.Add("Pickup", GetType(String))
        detailTable.Columns.Add("組立", GetType(String))
        detailTable.Columns.Add("機能検査수동", GetType(String))
        detailTable.Columns.Add("機能検査자동", GetType(String))
        detailTable.Columns.Add("2者検査", GetType(String))
        detailTable.Columns.Add("추가조립", GetType(String))  'hsj test
        detailTable.Columns.Add("검사설비", GetType(String))
        detailTable.Columns.Add("총공정시간", GetType(String))
        detailTable.Columns.Add("인공수", GetType(Double))

        simpleTable.Columns.Add("No", GetType(Int32))
        simpleTable.Columns.Add("Model", GetType(String))
        simpleTable.Columns.Add("수량", GetType(String))
        simpleTable.Columns.Add("마운팅", GetType(String))
        simpleTable.Columns.Add("조립", GetType(String))
        simpleTable.Columns.Add("총공정시간", GetType(String))
        simpleTable.Columns.Add("인공수", GetType(String))

        inspectionTable.Columns.Add("No", GetType(Int32))
        inspectionTable.Columns.Add("검사설비", GetType(String))
        inspectionTable.Columns.Add("수량", GetType(String))
        inspectionTable.Columns.Add("총시간", GetType(String))
        inspectionTable.Columns.Add("부하율", GetType(String))

        DbTable.Columns.Add("No", GetType(Int32))
        DbTable.Columns.Add("모델명", GetType(String))
        DbTable.Columns.Add("部品SET", GetType(String))
        DbTable.Columns.Add("前付け", GetType(String))
        DbTable.Columns.Add("MT", GetType(String))
        DbTable.Columns.Add("L/C", GetType(String))
        DbTable.Columns.Add("目視", GetType(String))
        DbTable.Columns.Add("Pickup", GetType(String))
        DbTable.Columns.Add("組立", GetType(String))
        DbTable.Columns.Add(" 機能検査_수동", GetType(String))
        DbTable.Columns.Add(" 機能検査_자동", GetType(String))
        DbTable.Columns.Add("2者検査", GetType(String))
        DbTable.Columns.Add("검사설비", GetType(String))

        'DbTable.Columns.Add("No", GetType(String))
        'DbTable.Columns.Add("SUFFIX", GetType(String))
        'DbTable.Columns.Add("추가 마운팅", GetType(String))
        'DbTable.Columns.Add("추가 조립", GetType(String))
        '프로시저로 만들어서 배열을 

    End Sub

    ''' <summary>
    ''' 필요 공수시간 계산
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnStart_Click_1(sender As Object, e As EventArgs) Handles btnStart.Click
        Try
            detailTable.Rows.Clear()
            Cursor.Current = Cursors.WaitCursor

            '상세모델 출력
            Dim time As List(Of String()) = _userControlPresenter.DataCalculate()
            Dim mountTime As New List(Of Double)
            Dim assamblyTime As New List(Of Double)
            Dim timespanList As List(Of TimeSpan) = New List(Of TimeSpan)

            Dim simpleList As List(Of String()) = _userControlPresenter.simpleModelList
            Dim inspectionList As List(Of String()) = _userControlPresenter.inspectionList

            For i As Integer = 0 To time.Count() - 1
                'For j As Integer = 1 To 10
                For j As Integer = 1 To 11

                    If time(i)(j) = "-" Then
                        time(i)(j) = "00:00:00"
                    End If

                    If time(i)(j) = "" Then
                        timespanList.Add(TimeSpan.Parse("0")) 'TimeSpan.Parse() : 시간 간격을 나타내는 문자열 표현을 TimeSpan 개체로 변환하는 데 사용
                    Else
                        timespanList.Add(TimeSpan.Parse(time(i)(j)))
                    End If
                Next

                Dim mount = timespanList(0).TotalSeconds + timespanList(1).TotalSeconds + timespanList(2).TotalSeconds + timespanList(3).TotalSeconds + timespanList(4).TotalSeconds
                Dim assambly = timespanList(5).TotalSeconds + timespanList(6).TotalSeconds + timespanList(7).TotalSeconds + timespanList(8).TotalSeconds + timespanList(9).TotalSeconds
                Dim totalts = timespanList(0).TotalSeconds + timespanList(1).TotalSeconds + timespanList(2).TotalSeconds + timespanList(3).TotalSeconds + timespanList(4).TotalSeconds + timespanList(5).TotalSeconds + timespanList(6).TotalSeconds + timespanList(7).TotalSeconds + timespanList(8).TotalSeconds + timespanList(9).TotalSeconds
                Dim totalTimeSecond = TimeSpan.FromSeconds(totalts)

                mountTime.Add(mount)
                assamblyTime.Add(assambly)

                Dim worker = (totalts / 60) / 460 'why divide 60 and 460

                detailTable.Rows.Add(i + 1, Replace(time(i)(0), CChar(vbLf), ""), time(i)(1), time(i)(2), time(i)(3), time(i)(4), time(i)(5), time(i)(6), time(i)(7), time(i)(8), time(i)(9), time(i)(10), Replace(time(i)(11), "\c", ","), totalTimeSecond.ToString(), worker.ToString("f2"))

                timespanList.Clear()
            Next

            Dim mountWorker As Double = (mountTime.Sum / 60) / 460
            Dim AssamblyWorker As Double = (assamblyTime.Sum / 60) / 460
            Dim totalWorker As Double = ((assamblyTime.Sum + mountTime.Sum) / 60) / 460

            labMauntingTime.Text = TimeSpan.FromSeconds(mountTime.Sum).ToString()
            labMauntingWorker.Text = mountWorker.ToString("f2") + "人"

            labAssamblyTime.Text = TimeSpan.FromSeconds(assamblyTime.Sum).ToString()
            labAssamblyWorker.Text = AssamblyWorker.ToString("f2") + "人"

            labTotalTime.Text = TimeSpan.FromSeconds(mountTime.Sum + assamblyTime.Sum).ToString()
            labTotalWorker.Text = totalWorker.ToString("f2") + "人"

            detailTable.Rows.Add(Nothing, Nothing, Nothing, Nothing, Nothing, "마운팅 시간 총합", TimeSpan.FromSeconds(mountTime.Sum).ToString(), Nothing, Nothing, Nothing, "조립 시간 총합", TimeSpan.FromSeconds(assamblyTime.Sum).ToString(), Nothing, TimeSpan.FromSeconds(mountTime.Sum + assamblyTime.Sum).ToString(), totalWorker.ToString("f2"))
            detailTable.Rows.Add(Nothing, Nothing, Nothing, Nothing, Nothing, "마운팅 총 인공수", mountWorker.ToString("f2"), Nothing, Nothing, Nothing, "조립 총 인공수", AssamblyWorker.ToString("f2"))

            GridSetting()

            mountTime.Clear()
            assamblyTime.Clear()

            '대표모델 출력
            simpleTable.Rows.Clear()
            Dim simpleData As List(Of String) = New List(Of String)

            For i As Integer = 0 To simpleList.Count() - 1
                If (simpleData.Contains(simpleList(i)(0)) = False) Then
                    simpleData.Add(simpleList(i)(0))
                End If
            Next

            For i As Integer = 0 To simpleData.Count() - 1
                Dim simpleTime As List(Of String()) = New List(Of String())
                For j As Integer = 0 To simpleList.Count() - 1
                    If simpleList(j)(0) = simpleData(i) Then

                        simpleTime.Add(simpleList(j))

                    End If
                Next

                For f As Integer = 0 To simpleTime.Count() - 1
                    For j As Integer = 1 To 10

                        If simpleTime(f)(j) = "-" Then
                            simpleTime(f)(j) = "00:00:00"
                        End If

                        timespanList.Add(TimeSpan.Parse(simpleTime(f)(j)))
                    Next

                    Dim mount = timespanList(0).TotalSeconds + timespanList(1).TotalSeconds + timespanList(2).TotalSeconds + timespanList(3).TotalSeconds + timespanList(4).TotalSeconds
                    Dim assambly = timespanList(5).TotalSeconds + timespanList(6).TotalSeconds + timespanList(7).TotalSeconds + timespanList(8).TotalSeconds + timespanList(9).TotalSeconds
                    Dim totalts = mount + assambly

                    mountTime.Add(mount)
                    assamblyTime.Add(assambly)

                    timespanList.Clear()
                Next

                totalWorker = ((assamblyTime.Sum + mountTime.Sum) / 60) / 460
                simpleTable.Rows.Add(i + 1, Replace(simpleTime(0)(0), CChar(vbLf), ""), simpleTime.Count(), TimeSpan.FromSeconds(mountTime.Sum).ToString(), TimeSpan.FromSeconds(assamblyTime.Sum).ToString(), TimeSpan.FromSeconds(mountTime.Sum + assamblyTime.Sum).ToString(), totalWorker.ToString("f2"))

                mountTime.Clear()
                assamblyTime.Clear()
            Next

            grdSimpleModel.DataSource = simpleTable

            grdSimpleModel.EnableHeadersVisualStyles = False
            grdSimpleModel.Columns(3).HeaderCell.Style.BackColor = Color.Moccasin
            grdSimpleModel.Columns(4).HeaderCell.Style.BackColor = Color.LightGreen
            grdSimpleModel.Columns(5).HeaderCell.Style.BackColor = Color.SkyBlue

            For i As Integer = 0 To 6
                grdSimpleModel.Columns(i).HeaderCell.Style.Font = New System.Drawing.Font("Tahoma", 10)
            Next

            grdSimpleModel.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            grdSimpleModel.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            '검사설비 출력
            inspectionTable.Rows.Clear()
            Dim inspectionData As List(Of String) = New List(Of String)

            For i As Integer = 0 To inspectionList.Count() - 1
                If (inspectionData.Contains(inspectionList(i)(3)) = False) Then
                    inspectionData.Add(inspectionList(i)(3))
                End If
            Next

            For i As Integer = 0 To inspectionData.Count() - 1
                Dim inspectionTime As List(Of String()) = New List(Of String())

                For j As Integer = 0 To inspectionList.Count() - 1
                    If inspectionList(j)(3) = inspectionData(i) Then

                        inspectionTime.Add(inspectionList(j))

                    End If
                Next

                For f As Integer = 0 To inspectionTime.Count() - 1
                    For j As Integer = 1 To 2
                        If inspectionTime(f)(j) = "-" Then
                            inspectionTime(f)(j) = "00:00:00"
                        End If
                        timespanList.Add(TimeSpan.Parse(inspectionTime(f)(j)))
                    Next

                    Dim totalts = timespanList(0).TotalSeconds + timespanList(1).TotalSeconds

                    mountTime.Add(totalts)

                    timespanList.Clear()
                Next

                Dim add As String() = Replace(inspectionTime(0)(3), "\c", ",").Split(CChar(","))
                totalWorker = ((mountTime.Sum / 60) / (460 * add.Count())) * 100

                inspectionTable.Rows.Add(i + 1, Replace(inspectionTime(0)(3), "\c", ","), inspectionTime.Count(), TimeSpan.FromSeconds(mountTime.Sum).ToString(), totalWorker.ToString("f2") + " %")

                mountTime.Clear()
                assamblyTime.Clear()
            Next

            grdInspectionEquipment.DataSource = inspectionTable

            For i As Integer = 0 To 4
                grdInspectionEquipment.Columns(i).HeaderCell.Style.Font = New System.Drawing.Font("Tahoma", 10)
            Next

            grdInspectionEquipment.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            grdInspectionEquipment.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            simpleList.Clear()
            inspectionList.Clear()

        Catch ex As Exception
            SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "btnStart_Click()", ex.Message)
        End Try
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub GridSetting()

        grdDetailModel.DataSource = detailTable
        grdDetailModel.EnableHeadersVisualStyles = False

        grdDetailModel.Columns(2).HeaderCell.Style.BackColor = Color.Moccasin
        grdDetailModel.Columns(3).HeaderCell.Style.BackColor = Color.Moccasin
        grdDetailModel.Columns(4).HeaderCell.Style.BackColor = Color.Moccasin
        grdDetailModel.Columns(5).HeaderCell.Style.BackColor = Color.Moccasin
        grdDetailModel.Columns(6).HeaderCell.Style.BackColor = Color.Moccasin
        grdDetailModel.Columns(7).HeaderCell.Style.BackColor = Color.LightGreen
        grdDetailModel.Columns(8).HeaderCell.Style.BackColor = Color.LightGreen
        grdDetailModel.Columns(9).HeaderCell.Style.BackColor = Color.LightGreen
        grdDetailModel.Columns(10).HeaderCell.Style.BackColor = Color.LightGreen
        grdDetailModel.Columns(11).HeaderCell.Style.BackColor = Color.LightGreen
        grdDetailModel.Columns(13).HeaderCell.Style.BackColor = Color.SkyBlue
        grdDetailModel.Columns(14).HeaderCell.Style.BackColor = Color.SkyBlue

        grdDetailModel.Columns(0).Width = 40
        grdDetailModel.Columns(12).Width = 90
        grdDetailModel.Columns(13).Width = 100
        grdDetailModel.Columns(14).Width = 80

        For i As Integer = 0 To 14
            grdDetailModel.Columns(i).HeaderCell.Style.Font = New System.Drawing.Font("Tahoma", 10)
        Next

        grdDetailModel.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdDetailModel.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdDetailModel.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

    End Sub

    ''' <summary>
    ''' Save Button Click Event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnSave_Click_1(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim sheetList As List(Of String) = New List(Of String)
        Dim tableList As List(Of System.Data.DataTable) = New List(Of System.Data.DataTable)
        sheetList.Add("전체모델")
        sheetList.Add("대표모델")
        sheetList.Add("검사설비")

        tableList.Add(detailTable)
        tableList.Add(simpleTable)
        tableList.Add(inspectionTable)

        _userControlPresenter.Save_Excel(sheetList, tableList)
    End Sub

    ''' <summary>
    ''' DB Data 전체조회
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    'Private Sub btnAllRead_Click_1(sender As Object, e As EventArgs)
    '    txtPath.Text = ""
    '    AllRead()
    'End Sub

    Private Sub AllRead()
        Try
            Cursor.Current = Cursors.WaitCursor

            DbTable.Rows.Clear()
            Dim ResultData As String = Nothing
            Dim SqlCMD As String = " Select RECNO, MODEL, COMPONENT_SET, MAEDZUKE, MAUNT, LEAD_CUTTING, VISUAL_EXAMINATION, PICKUP, ASSAMBLY, M_FUNCTION_CHECK, A_FUNCTION_CHECK, PERSON_EXAMINE, INSPECTION_EQUIPMENT " & "from FAM3_PRODUCT_TIME_TB"
            EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)

            Dim rowArray As String() = ResultData.Split(CChar(vbCrLf))
            For i = 1 To rowArray.Length - 2
                Dim colArray As String() = rowArray(i).Split(CChar(","))

                DbTable.Rows.Add(Replace(colArray(0), CChar(vbLf), ""), colArray(1), colArray(2), colArray(3), colArray(4), colArray(5), colArray(6), colArray(7), colArray(8), colArray(9), colArray(10), colArray(11), Replace(colArray(12), "\c", ","))
            Next
            DbTable.DefaultView.Sort = "No"
            'grdRead.DataSource = DbTable  'hsj test할려고 제거
            grd_master.DataSource = DbTable

        Catch ex As Exception
            SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "btnAllRead_Click()", ex.Message)
        End Try
        Cursor.Current = Cursors.Default

    End Sub

    ''' <summary>
    ''' Master File Select
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    'Private Sub txtPath_Click(sender As Object, e As EventArgs)
    '    Dim ofd As OpenFileDialog = New OpenFileDialog With {
    '        .Filter = "모든 파일 (*.*) | *.*"
    '    }
    '    ofd.ShowDialog()
    '    Dim path As String = ofd.FileName
    '    txtPath.Text = ofd.FileName
    '    'Dim variableType As Type = txtPath.GetType() 'hsj test
    '    Console.WriteLine(txtPath.GetType.Name) 'hsj test
    '    If txtPath.Text IsNot "" Then
    '        NewfileLead()
    '    End If
    'End Sub

    'txtbox test용
    'Private Sub TextBox4_Click(sender As Object, e As EventArgs) Handles txtbox_master_path.Click 'hsj test 마스터DB 클릭 시 파일 선택, DB 확인, txt 박스 클릭
    '    Dim ofd As OpenFileDialog = New OpenFileDialog With {
    '        .Filter = "모든 파일 (*.*) | *.*"
    '    }
    '    ofd.ShowDialog()
    '    Dim path As String = ofd.FileName
    '    'txtPath.Text = ofd.FileName 'hsj del
    '    txtbox_master_path.Text = ofd.FileName
    '    If txtbox_master_path.Text IsNot "" Then
    '        NewfileLead()
    '    End If
    'End Sub
    '  업로드 db 경로 불러오기 기능 - 공통 함수로 test 중, 고의로 master db 찾는 부분은 제외할 거임
    Private Sub txtbox_path_click(sender As Object, e As EventArgs) Handles txtbox_suffix_path.Click, txtbox_carrier_path.Click, txtbox_limit_path.Click, txtbox_master_path.Click '
        Dim ofd As OpenFileDialog = New OpenFileDialog With {
            .Filter = "모든 파일 (*.*) | *.*"
        }
        ofd.ShowDialog()
        Dim path As String = ofd.FileName
        'txtPath.Text = ofd.FileName 'hsj del
        '공통 기능 수행을 위해 sender 매개 변수를 사용하여, 클릭한 텍스트 박스에 따라  별도 기능을 사용하도록 하게 한다.
        If sender Is txtbox_suffix_path Then
            txtbox_suffix_path.Text = ofd.FileName
        ElseIf sender Is txtbox_carrier_path Then
            txtbox_carrier_path.Text = ofd.FileName
        ElseIf sender Is txtbox_limit_path Then
            txtbox_limit_path.Text = ofd.FileName
        ElseIf sender Is txtbox_master_path Then
            txtbox_master_path.Text = ofd.FileName
        End If
        'TextBox4.Text = ofd.FileName

        'If TextBox4.Text IsNot "" Then
        '    NewfileLead()
        'End If
        '파일 선택 클릭시 db table clear
        DbTable.Columns.Clear()

        'hsj 배열로 테이블 만들기
        Dim list_dbtable_master() As String = {"No", "부속품", "모델 명", "部品SET", "前付け", "MT", "L/C", "目視", "Pick up", "組立", "機能検査(수동)", "機能検査(자동)", "2者検査", "검사 설비"}
        Dim list_dbtable_suffix() As String = {"No", "SUFFIX", "추가 마운팅", "추가 조립"}
        Dim list_dbtable_carrier() As String = {"No", "모델명", "사용 캐리어"}
        Dim list_dbtable_limit() As String = {"No", "캐리어 명", "제한 대수", "수량"}
        'If txtbox_suffix_path.Text.IndexOf("SUFFIX") >= 0 Then

        If sender.Text.IndexOf("SUFFIX") >= 0 Then '클릭시 파일 선택 ,SUFFIX
            For i As Integer = 0 To list_dbtable_suffix.Count() - 1
                If i = 0 Then
                    DbTable.Columns.Add(list_dbtable_suffix(i), GetType(Int32))
                Else
                    DbTable.Columns.Add(list_dbtable_suffix(i), GetType(String))
                End If
            Next
        ElseIf sender.Text.IndexOf("MODEL") >= 0 Then
            For i As Integer = 0 To list_dbtable_carrier.Count() - 1
                If i = 0 Then
                    DbTable.Columns.Add(list_dbtable_carrier(i), GetType(Int32))
                Else
                    DbTable.Columns.Add(list_dbtable_carrier(i), GetType(String))
                End If
            Next
        ElseIf sender.Text.IndexOf("LIMIT") >= 0 Then
            For i As Integer = 0 To list_dbtable_limit.Count() - 1
                If i = 0 Then
                    DbTable.Columns.Add(list_dbtable_limit(i), GetType(Int32))
                Else
                    DbTable.Columns.Add(list_dbtable_limit(i), GetType(String))
                End If
            Next
        ElseIf sender.Text.IndexOf("Master Data") >= 0 Then 'sender로 변수 통합해도 되는거 아닌가
            For i As Integer = 0 To list_dbtable_master.Count() - 1
                If i = 0 Then
                    DbTable.Columns.Add(list_dbtable_master(i), GetType(Int32))
                Else
                    DbTable.Columns.Add(list_dbtable_master(i), GetType(String))
                End If
            Next
        End If

        If sender.Text IsNot "" Then
            NewfileLoad(sender)
        End If
    End Sub

    'Private Sub NewfileLead()
    '    AllRead()
    '    newDBTable.Clear()
    '    Dim SelectStatement As String = "SELECT [No], [모델 명], FORMAT([部品SET], 'HH:mm:ss') as [部品SET], FORMAT([前付け], 'HH:mm:ss') as [前付け], FORMAT([MT], 'HH:mm:ss') as [MT], FORMAT([L/C], 'HH:mm:ss') as [L/C], FORMAT([目視], 'HH:mm:ss') as [目視], FORMAT([Pick up], 'HH:mm:ss') as [Pick up],
    '                                       FORMAT([組立], 'HH:mm:ss') as [組立], FORMAT([機能検査(수동)], 'HH:mm:ss') as [機能検査_수동], FORMAT([機能検査(자동)], 'HH:mm:ss') as [機能検査_자동], FORMAT([2者検査], 'HH:mm:ss') as [2者検査], [검사 설비]  FROM [Sheet1$]"

    '    Using cn As New OleDb.OleDbConnection With {.ConnectionString = Connection.HeaderConnectionString(txtbox_master_path.Text)}
    '        'Using cn As New OleDb.OleDbConnection With {.ConnectionString = Connection.HeaderConnectionString(txtPath.Text)}
    '        Using cmd As New OleDbCommand With {.Connection = cn, .CommandText = SelectStatement}

    '            cn.Open()
    '            Try
    '                newDBTable.Load(cmd.ExecuteReader)
    '                'grdRead.DataSource = newDBTable
    '                grd_master.DataSource = newDBTable

    '                CompareData()

    '            Catch ex As Exception
    '                Console.WriteLine(ex.Message)
    '                MessageBox.Show("잘못된 Master 파일형식입니다.")
    '            End Try

    '        End Using
    '    End Using
    'End Sub
    ' hsj 공통으로 업로드할 엑셀 파일 불러오는 함수 제작 진행중
    Private Sub NewfileLoad(ByVal sender As Object)
        AllReadNew(sender)
        newDBTable.Clear()

        Dim SelectStatement As String

        If sender Is txtbox_suffix_path Then
            SelectStatement = "SELECT [No], [SUFFIX], FORMAT([추가 마운팅], 'HH:mm:ss') as [추가 마운팅], FORMAT([추가 조립], 'HH:mm:ss') as [추가 조립]  FROM [Sheet1$]"
        ElseIf sender Is txtbox_carrier_path Then
            SelectStatement = "SELECT [No], [모델명],[사용 캐리어] FROM [Sheet1$]"
        ElseIf sender Is txtbox_limit_path Then
            SelectStatement = "SELECT [No], [캐리어 명], [제한 대수], [수량]  FROM [Sheet1$]"
        ElseIf sender Is txtbox_master_path Then
            SelectStatement = "SELECT [No], [모델 명], FORMAT([부속품], 'HH:mm:ss') as [부속품], FORMAT([部品SET], 'HH:mm:ss') as [部品SET], FORMAT([前付け], 'HH:mm:ss') as [前付け], FORMAT([MT], 'HH:mm:ss') as [MT], FORMAT([L/C], 'HH:mm:ss') as [L/C], FORMAT([目視], 'HH:mm:ss') as [目視], FORMAT([Pick up], 'HH:mm:ss') as [Pick up],
                                           FORMAT([組立], 'HH:mm:ss') as [組立], FORMAT([機能検査(수동)], 'HH:mm:ss') as [機能検査_수동], FORMAT([機能検査(자동)], 'HH:mm:ss') as [機能検査_자동], FORMAT([2者検査], 'HH:mm:ss') as [2者検査], [검사 설비]  FROM [Sheet1$]"
        End If


        'Dim SelectStatement As String = "SELECT [No], [모델 명], FORMAT([部品SET], 'HH:mm:ss') as [部品SET], FORMAT([前付け], 'HH:mm:ss') as [前付け], FORMAT([MT], 'HH:mm:ss') as [MT], FORMAT([L/C], 'HH:mm:ss') as [L/C], FORMAT([目視], 'HH:mm:ss') as [目視], FORMAT([Pick up], 'HH:mm:ss') as [Pick up],
        'Format([組立], 'HH:mm:ss') as [組立], FORMAT([機能検査(수동)], 'HH:mm:ss') as [機能検査_수동], FORMAT([機能検査(자동)], 'HH:mm:ss') as [機能検査_자동], FORMAT([2者検査], 'HH:mm:ss') as [2者検査], [검사 설비]  FROM [Sheet1$]"

        Using cn As New OleDb.OleDbConnection With {.ConnectionString = Connection.HeaderConnectionString(sender.Text)}
            'Using cn As New OleDb.OleDbConnection With {.ConnectionString = Connection.HeaderConnectionString(txtbox_master_path.Text)}
            'Using cn As New OleDb.OleDbConnection With {.ConnectionString = Connection.HeaderConnectionString(txtPath.Text)}
            Using cmd As New OleDbCommand With {.Connection = cn, .CommandText = SelectStatement}

                cn.Open()
                Try
                    newDBTable.Load(cmd.ExecuteReader)
                    'grdRead.DataSource = newDBTable
                    'grd_master.DataSource = newDBTable
                    If sender Is txtbox_suffix_path Then
                        grd_suffix.DataSource = DbTable
                    ElseIf sender Is txtbox_carrier_path Then
                        grd_carrier.DataSource = DbTable
                    ElseIf sender Is txtbox_limit_path Then
                        grd_limit.DataSource = DbTable
                    ElseIf sender Is txtbox_master_path Then
                        grd_master.DataSource = DbTable
                    End If

                    CompareDataNew(sender)

                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    MessageBox.Show("잘못된 Master 파일형식입니다.")
                End Try

            End Using
        End Using
    End Sub

    ' hsj - DB 종류에 따라서 SqlCMD를 다르게 하기 위해서 새로 제작하는 함수
    Private Sub AllReadNew(ByVal sender As Object)
        Try
            Cursor.Current = Cursors.WaitCursor

            DbTable.Rows.Clear()
            Dim ResultData As String = Nothing
            'Console.WriteLine(sender.GetType.Name)
            If sender Is txtbox_suffix_path Then
                Dim SqlCMD As String = " Select RECNO, SUFFIX, ADDITIONAL_MAOUNTING, ADDITIONAL_ASSEMBLY " & "from FAM3_SUFFIX_TIME_TB" 'suffix db 데이터 불러오는 커맨드
                EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)   ' 중복 제거 필요? - 조건문 밖으로 빼면, SqlCMD 정의를 다시 해야함
            ElseIf sender Is txtbox_carrier_path Then
                Dim SqlCMD As String = " Select RECNO, MODEL, CARRIER " & "from FAM3_CARRIER_TB" ' 캐리어 종류 db 불러오기
                EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)
            ElseIf sender Is txtbox_limit_path Then
                Dim SqlCMD As String = " Select RECNO, CARRIER, LIMIT, QUANTITY " & "from FAM3_LIMIT_TB" ' 캐리어 제한대수 db 불러오기
                EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)
            ElseIf sender Is txtbox_master_path Then
                Dim SqlCMD As String = " Select RECNO, MODEL, ACCESSORY, COMPONENT_SET, MAEDZUKE, MAUNT, LEAD_CUTTING, VISUAL_EXAMINATION, PICKUP, ASSEMBLY, M_FUNCTION_CHECK, A_FUNCTION_CHECK, PERSON_EXAMINE, INSPECTION_EQUIPMENT " & "from FAM3_MODEL_TIME_TB" ' 마스터 db 불러오기
                EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)
            End If

            '=======================EtherUty.EtherSendSQL 빼려고 했는데 SqlCMD가 정의 되지 않아서 err

            '조건문에 삽입할 내용 start
            'Dim SqlCMD As String = " Select RECNO, MODEL, COMPONENT_SET, MAEDZUKE, MAUNT, LEAD_CUTTING, VISUAL_EXAMINATION, PICKUP, ASSAMBLY, M_FUNCTION_CHECK, A_FUNCTION_CHECK, PERSON_EXAMINE, INSPECTION_EQUIPMENT " & "from FAM3_PRODUCT_TIME_TB"
            'EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)

            ' hsjdb테이블에 별도로 저장하기
            Dim rowArray As String() = ResultData.Split(CChar(vbCrLf))
            For i = 1 To rowArray.Length - 2
                Dim colArray As String() = rowArray(i).Split(CChar(","))
                If sender Is txtbox_suffix_path Then
                    'DbTable.Rows.Add(Replace(colArray(0), CChar(vbLf), ""), colArray(1), colArray(2), colArray(3), colArray(4))
                    DbTable.Rows.Add(Replace(colArray(0), CChar(vbLf), ""), colArray(1), colArray(2), colArray(3))
                ElseIf sender Is txtbox_carrier_path Then
                    DbTable.Rows.Add(Replace(colArray(0), CChar(vbLf), ""), colArray(1), colArray(2))
                ElseIf sender Is txtbox_limit_path Then
                    DbTable.Rows.Add(Replace(colArray(0), CChar(vbLf), ""), colArray(1), colArray(2), colArray(3))
                ElseIf sender Is txtbox_master_path Then
                    DbTable.Rows.Add(Replace(colArray(0), CChar(vbLf), ""), colArray(1), colArray(2), colArray(3), colArray(4), colArray(5), colArray(6), colArray(7), colArray(8), colArray(9), colArray(10), colArray(11), colArray(12), Replace(colArray(13), "\c", ","))
                End If
            Next
            'DbTable.DefaultView.Sort = "No"
            DbTable.DefaultView.Sort = DbTable.Columns(0).ColumnName 'DbTable 첫 번째 컬럼 "No"로 sort
            'grdRead.DataSource = DbTable  'hsj test할려고 제거
            If sender Is txtbox_suffix_path Then
                grd_suffix.DataSource = DbTable
            ElseIf sender Is txtbox_carrier_path Then
                grd_carrier.DataSource = DbTable
            ElseIf sender Is txtbox_limit_path Then
                grd_limit.DataSource = DbTable
            ElseIf sender Is txtbox_master_path Then
                grd_master.DataSource = DbTable
            End If

        Catch ex As Exception
            SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "btnAllRead_Click()", ex.Message)
        End Try
        Cursor.Current = Cursors.Default

    End Sub
    Private Sub CompareDataNew(ByVal sender As Object)
        'Dim arrOb As Object() = DbTable.[Select]().[Select](Function(x) x("모델명")).ToArray() 'x("모델명")을 x(1)로 변경 검토중
        Dim arrOb As Object() = DbTable.[Select]().[Select](Function(x) x(1)).ToArray()
        Dim dbModel As String() = arrOb.Cast(Of String)().ToArray()
        Dim set_grd As Object 'DB별 data grid 설정 변수
        'Dim k As Integer '인덱스 길이 설정 변수
        If sender Is txtbox_suffix_path Then
            set_grd = grd_suffix
            'k = 3
        ElseIf sender Is txtbox_carrier_path Then
            set_grd = grd_carrier
        ElseIf sender Is txtbox_limit_path Then
            set_grd = grd_limit
        ElseIf sender Is txtbox_master_path Then
            set_grd = grd_master
        End If
        'k = DbTable.Columns.Count - 1

        For i As Integer = 0 To newDBTable.Rows.Count - 1 ' Range.Rows.Count
            Try ' 에러 확인 용 try
                If dbModel.Contains(Replace(newDBTable.Rows(i).ItemArray(1).ToString(), " ", "")) Then
                    For j As Integer = 0 To DbTable.Rows.Count - 1
                        If Replace(newDBTable.Rows(i).ItemArray(1).ToString(), " ", "") = DbTable.Rows(j).ItemArray(1).ToString() Then
                            'For t As Integer = 2 To 12 'DB 종류별로 컬럼 길이 설정을 따로 해줘야 할 것 같습니다. 그냥 그대로 사용하면 err가 발생할까?
                            For t As Integer = 2 To DbTable.Columns.Count - 1 'ｋ를　코드로　변환
                                If Replace(newDBTable.Rows(i).ItemArray(t).ToString(), " ", "") = DbTable.Rows(j).ItemArray(t).ToString() Then
                                Else
                                    set_grd.Rows(i).Cells(t).Style.BackColor = Color.Yellow
                                    set_grd.Rows(i).Cells(0).Style.BackColor = Color.Yellow
                                    set_grd.Rows(i).Cells(1).Style.BackColor = Color.Yellow
                                    'grdRead.Rows(i).Cells(t).Style.BackColor = Color.Yellow  ' grdRead 변수 변경 필요
                                    'grdRead.Rows(i).Cells(0).Style.BackColor = Color.Yellow  ' Cells(0) = No, (1) = 모델명
                                    'grdRead.Rows(i).Cells(1).Style.BackColor = Color.Yellow
                                End If
                            Next
                        End If
                    Next
                Else
                    If newDBTable.Rows.Count = DbTable.Rows.Count Then 'TPROD 서버 데이터와 엑셀데이터 행의 개수가 같을때?
                        For f As Integer = 0 To DbTable.Columns.Count - 1
                            set_grd.Rows(i).Cells(f).Style.BackColor = Color.DarkOrange
                            'grdRead.Rows(i).Cells(f).Style.BackColor = Color.DarkOrange
                        Next
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                MessageBox.Show("Data Compare NG")

            End Try


        Next
    End Sub

    ''' <summary>
    ''' Data Compare
    ''' </summary>
    'Private Sub CompareData()
    '    Dim arrOb As Object() = DbTable.[Select]().[Select](Function(x) x("모델명")).ToArray()
    '    Dim dbModel As String() = arrOb.Cast(Of String)().ToArray()

    '    For i As Integer = 0 To newDBTable.Rows.Count - 1 ' Range.Rows.Count
    '        Try ' 에러 확인 용 try
    '            If dbModel.Contains(Replace(newDBTable.Rows(i).ItemArray(1).ToString(), " ", "")) Then
    '                For j As Integer = 0 To DbTable.Rows.Count - 1
    '                    If Replace(newDBTable.Rows(i).ItemArray(1).ToString(), " ", "") = DbTable.Rows(j).ItemArray(1).ToString() Then
    '                        For t As Integer = 2 To 12
    '                            If Replace(newDBTable.Rows(i).ItemArray(t).ToString(), " ", "") = DbTable.Rows(j).ItemArray(t).ToString() Then
    '                            Else
    '                                grd_master.Rows(i).Cells(t).Style.BackColor = Color.Yellow
    '                                grd_master.Rows(i).Cells(0).Style.BackColor = Color.Yellow
    '                                grd_master.Rows(i).Cells(1).Style.BackColor = Color.Yellow
    '                                'grdRead.Rows(i).Cells(t).Style.BackColor = Color.Yellow  ' grdRead 변수 변경 필요
    '                                'grdRead.Rows(i).Cells(0).Style.BackColor = Color.Yellow  ' Cells(0), (1)가 뭐냐
    '                                'grdRead.Rows(i).Cells(1).Style.BackColor = Color.Yellow
    '                            End If
    '                        Next
    '                    End If
    '                Next
    '            Else
    '                For f As Integer = 0 To 12
    '                    grdRead.Rows(i).Cells(f).Style.BackColor = Color.DarkOrange
    '                Next
    '            End If
    '        Catch ex As Exception
    '            Console.WriteLine(ex.Message)
    '            MessageBox.Show("Data Compare NG.")

    '        End Try


    '    Next
    'End Sub

    ''' <summary>
    ''' Model 검색
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    'Private Sub btnSearch_Click_1(sender As Object, e As EventArgs)
    '    Cursor.Current = Cursors.WaitCursor
    '    Try
    '        DbTable.Rows.Clear()
    '        If txtSearch.Text = Nothing Then
    '            MsgBox("찾으실 모델명을 입력하여 주십시오")
    '        Else

    '            Dim list As List(Of String()) = New List(Of String())
    '            Dim ResultData As String = Nothing
    '            Dim SqlCMD As String = "Select * from FAM3_PRODUCT_TIME_TB WHERE MODEL Like" & "'" & (txtSearch.Text).ToUpper & "'"

    '            If (txtSearch.Text).Contains("*") Then
    '                SqlCMD = "select * from FAM3_PRODUCT_TIME_TB WHERE MODEL LIKE" & "'%" & Replace(txtSearch.Text, CChar("*"), "").ToUpper & "%'"
    '            Else
    '                SqlCMD = "select * from FAM3_PRODUCT_TIME_TB WHERE MODEL LIKE" & "'" & (txtSearch.Text).ToUpper & "'"
    '            End If

    '            EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)

    '            If ResultData = "ERROR:A0 Nothing" Then
    '                MessageBox.Show("검색 결과가 존재하지 않습니다.", "warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '            Else
    '                Dim rowArray As String() = ResultData.Split(CChar(vbCrLf))
    '                For i = 1 To rowArray.Length - 2
    '                    Dim colArray As String() = rowArray(i).Split(CChar(","))
    '                    list.Add(colArray)
    '                    DbTable.Rows.Add(i, colArray(1), colArray(2), colArray(3), colArray(4), colArray(5), colArray(6), colArray(7), colArray(8), colArray(9), colArray(10), colArray(11), Replace(colArray(12), "\c", ","))
    '                Next
    '            End If
    '            grdRead.DataSource = DbTable
    '        End If

    '    Catch ex As Exception
    '        SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "btnSearch_Click()", ex.Message)
    '    End Try
    '    Cursor.Current = Cursors.Default
    'End Sub

    ''' <summary>
    ''' Master 파일 수정
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    'db 업로드 설정
    Private Sub btnMasterSet_Click_1(sender As Object, e As EventArgs) Handles btn_master_upload.Click, btn_suffix_upload.Click, btn_carrier_upload.Click, btn_limit_upload.Click
        Dim login As Lock = New Lock()

        'If txtPath.Text = "" Or txtPath.Text = "클릭시 파일 선택" Then
        '    MessageBox.Show("변경할 마스터 데이터 파일을 선택하여 주십시오.", "warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'Else
        If sender.Text = "" Or sender.Text = "클릭시 파일 선택" Then
            MessageBox.Show("변경할 마스터 데이터 파일을 선택하여 주십시오.", "warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            login.ShowDialog()

            If login.DialogResult = DialogResult.OK Then
                Cursor.Current = Cursors.WaitCursor
                MasterReadNew() 'MasterReadNew로 대체 
                _userControlPresenter.MasterDataInputNew(sender)
                newDBTable.Rows.Clear()
                'masterDatalist 종류별로 초기화
                If sender.Text.IndexOf("Suffix") >= 0 Then
                    masterDatalistSuffix.Clear()
                    txtbox_suffix_path.Text = ""
                ElseIf sender.Text.IndexOf("Carrier") >= 0 Then
                    masterDatalistCarrier.Clear() '리스트 변경 필요
                    txtbox_carrier_path.Text = ""
                ElseIf sender.Text.IndexOf("Limit") >= 0 Then
                    masterDatalistLimit.Clear() '리스트 변경 필요
                    txtbox_limit_path.Text = ""
                ElseIf sender.Text.IndexOf("Master") >= 0 Then
                    masterDatalist.Clear() '리스트 변경 필요
                    txtbox_master_path.Text = ""
                End If

                'txtPath.Text = ""
                Cursor.Current = Cursors.Default
                MsgBox("Database 수정이 완료되었습니다.")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Master(NEW DB Table Read_Upload)
    ''' </summary>
    Private Sub MasterRead()

        Try
            For i As Integer = 0 To newDBTable.Rows.Count - 1 ' Range.Rows.Count

                Dim aa As String() = {newDBTable.Rows(i).ItemArray(0).ToString(), newDBTable.Rows(i).ItemArray(1).ToString(), newDBTable.Rows(i).ItemArray(2).ToString(), newDBTable.Rows(i).ItemArray(3).ToString(), newDBTable.Rows(i).ItemArray(4).ToString(),
                                      newDBTable.Rows(i).ItemArray(5).ToString(), newDBTable.Rows(i).ItemArray(6).ToString(), newDBTable.Rows(i).ItemArray(7).ToString(), newDBTable.Rows(i).ItemArray(8).ToString(), newDBTable.Rows(i).ItemArray(9).ToString(),
                                      newDBTable.Rows(i).ItemArray(10).ToString(), newDBTable.Rows(i).ItemArray(11).ToString(), newDBTable.Rows(i).ItemArray(12).ToString()}

                Dim list As DataCenter = New DataCenter(Replace(aa(1), " ", ""), aa(2), aa(3), aa(4), aa(5), aa(6), aa(7), aa(8), aa(9), aa(10), aa(11), aa(12), aa(13)) 'datacenter 부속품 추가

                masterDatalist.Add(list)
            Next
        Catch ex As Exception
            SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "MasterRead()", ex.Message)
        End Try
    End Sub

    Private Sub MasterReadNew() 'hsj db 종류별로 데이터 매칭을 다르게 설정한다

        Try
            For i As Integer = 0 To newDBTable.Rows.Count - 1 ' Range.Rows.Count
                If txtbox_suffix_path.Text.IndexOf("SUFFIX") >= 0 Then
                    Dim aa As String() = {newDBTable.Rows(i).ItemArray(0).ToString(), newDBTable.Rows(i).ItemArray(1).ToString(), newDBTable.Rows(i).ItemArray(2).ToString(), newDBTable.Rows(i).ItemArray(3).ToString()}

                    Dim list As DataCenterSuffix = New DataCenterSuffix(Replace(aa(1), " ", ""), aa(2), aa(3))
                    masterDatalistSuffix.Add(list)
                ElseIf txtbox_carrier_path.Text.IndexOf("MODEL") >= 0 Then '캐리어 0-2 변경 필요
                    Dim aa As String() = {newDBTable.Rows(i).ItemArray(0).ToString(), newDBTable.Rows(i).ItemArray(1).ToString(), newDBTable.Rows(i).ItemArray(2).ToString()}

                    Dim list As DataCenterCarrier = New DataCenterCarrier(Replace(aa(1), " ", ""), aa(2))
                    masterDatalistCarrier.Add(list)
                ElseIf txtbox_limit_path.Text.IndexOf("LIMIT") >= 0 Then
                    Dim aa As String() = {newDBTable.Rows(i).ItemArray(0).ToString(), newDBTable.Rows(i).ItemArray(1).ToString(), newDBTable.Rows(i).ItemArray(2).ToString(), newDBTable.Rows(i).ItemArray(3).ToString()}

                    Dim list As DataCenterLimit = New DataCenterLimit(Replace(aa(1), " ", ""), aa(2), aa(3))
                    masterDatalistLimit.Add(list)
                ElseIf txtbox_master_path.Text.IndexOf("Master Data") >= 0 Then '마스터 0-13
                    Dim aa As String() = {newDBTable.Rows(i).ItemArray(0).ToString(), newDBTable.Rows(i).ItemArray(1).ToString(), newDBTable.Rows(i).ItemArray(2).ToString(), newDBTable.Rows(i).ItemArray(3).ToString(), newDBTable.Rows(i).ItemArray(4).ToString(),
                                      newDBTable.Rows(i).ItemArray(5).ToString(), newDBTable.Rows(i).ItemArray(6).ToString(), newDBTable.Rows(i).ItemArray(7).ToString(), newDBTable.Rows(i).ItemArray(8).ToString(), newDBTable.Rows(i).ItemArray(9).ToString(),
                                      newDBTable.Rows(i).ItemArray(10).ToString(), newDBTable.Rows(i).ItemArray(11).ToString(), newDBTable.Rows(i).ItemArray(12).ToString(), newDBTable.Rows(i).ItemArray(13).ToString()}

                    Dim list As DataCenter = New DataCenter(Replace(aa(1), " ", ""), aa(2), aa(3), aa(4), aa(5), aa(6), aa(7), aa(8), aa(9), aa(10), aa(11), aa(12), aa(13))
                    masterDatalist.Add(list)
                End If

            Next
        Catch ex As Exception
            SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "MasterRead()", ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' background text
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    'Private Sub txtSearch_Leave(sender As Object, e As EventArgs)
    '    If txtSearch.Text = "" Then
    '        txtSearch.Text = "모델명 입력"
    '        txtSearch.ForeColor = Color.Gray
    '    End If
    'End Sub

    '''' <summary>
    ''''  background text
    '''' </summary>
    '''' <param name="sender"></param>
    '''' <param name="e"></param>    
    'Private Sub txtSearch_Enter(sender As Object, e As EventArgs)  'hsj 텍스트 상자에서 포커스가 벗어났을때 발생하는 이벤트 핸들러
    '    If txtSearch.Text = "모델명 입력" Then
    '        txtSearch.Text = ""
    '        txtSearch.ForeColor = Color.Black
    '    End If
    'End Sub

    '''' <summary>
    '''' SaveOption Dialog Show
    '''' </summary>
    '''' <param name="sender"></param>
    '''' <param name="e"></param>
    'Private Sub 파일경로설정ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 메뉴얼ToolStripMenuItem.Click
    '    _userControlPresenter.ShowSaveOption()
    'End Sub

    'Private Sub txtPath_Leave(sender As Object, e As EventArgs)
    '    If txtPath.Text = "" Then
    '        txtPath.Text = "클릭시 파일 선택"
    '        txtPath.ForeColor = Color.Gray
    '    End If
    'End Sub

    'Private Sub txtPath_Enter(sender As Object, e As EventArgs)
    '    If txtPath.Text = "클릭시 파일 선택" Then
    '        txtPath.Text = ""
    '        txtPath.ForeColor = Color.Black
    '    End If
    'End Sub

    Private Sub 설정파일관리ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 시스템정보ToolStripMenuItem.Click
        System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + "\Config\Program.ini")
    End Sub

    Private Sub 시스템정보ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 설정파일관리ToolStripMenuItem.Click
        Dim aboutDialog As ProgramInfo = New ProgramInfo()
        aboutDialog.ShowDialog()
    End Sub

    Private Sub 계획공수계산ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 계획공수계산ToolStripMenuItem.Click
        Dim Manu As UserManual = New UserManual()
        Manu.계획공수계산()
        Manu.ShowDialog()
    End Sub

    Private Sub 마스터데이터관리ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 마스터데이터관리ToolStripMenuItem.Click
        Dim Manu As UserManual = New UserManual()
        Manu.마스터데이터관리()
        Manu.ShowDialog()
    End Sub

    Private Sub grdDetailModel_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdDetailModel.CellContentClick

    End Sub

    Private Sub TableLayoutPanel3_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub grdRead_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    'Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
    'Private Sub txtPath1_TextChanged(sender As Object, e As EventArgs) Handles txtPath1.TextChanged 'txtPath1 = Master-"클릭시 파일 선택"

    'End Sub

    Private Sub txtPath_TextChanged(sender As Object, e As EventArgs)

    End Sub

    'Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles txtbox_limit_path.TextChanged 'txtPath2 = Limit-"클릭시 파일 선택"

    End Sub

    'Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles txtbox_suffix_path.TextChanged 'txtPath2 = Suffix-"클릭시 파일 선택"

    End Sub

    'Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles txtbox_carrier_path.TextChanged 'txtPath3 = Carrier-"클릭시 파일 선택"

    End Sub

    Private Sub txtPath1_TextChanged(sender As Object, e As EventArgs) Handles txtbox_master_path.TextChanged 'hsj - 마스터 db ;클릭시 파일 선택; 눌렀을때 이벤트 처리?

    End Sub

    Private Sub txtbox_search_suffix_TextChanged(sender As Object, e As EventArgs) Handles txtbox_search_suffix.TextChanged

    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub

    Private Sub TableLayoutPanel7_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel7.Paint

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grid_special.CellContentClick

    End Sub
End Class