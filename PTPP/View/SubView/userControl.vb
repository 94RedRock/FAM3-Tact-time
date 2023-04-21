Imports Microsoft.Office.Interop.Excel
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

    Public Sub New()

        InitializeComponent()

        _userControlPresenter = New userControlPresenter(Me)
        detailTable.Columns.Add("No", GetType(Int32))
        detailTable.Columns.Add("Model", GetType(String))
        detailTable.Columns.Add("部品SET", GetType(String))
        detailTable.Columns.Add("前付け", GetType(String))
        detailTable.Columns.Add("MT", GetType(String))
        detailTable.Columns.Add("LC", GetType(String))
        detailTable.Columns.Add("目視", GetType(String))
        detailTable.Columns.Add("Pickup", GetType(String))
        detailTable.Columns.Add("組立", GetType(String))
        detailTable.Columns.Add("機能検査수동", GetType(String))
        detailTable.Columns.Add("機能検査자동", GetType(String))
        detailTable.Columns.Add("2者検査", GetType(String))
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
                For j As Integer = 1 To 10

                    If time(i)(j) = "-" Then
                        time(i)(j) = "00:00:00"
                    End If

                    If time(i)(j) = "" Then
                        timespanList.Add(TimeSpan.Parse("0"))
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

                Dim worker = (totalts / 60) / 460

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
    Private Sub btnAllRead_Click_1(sender As Object, e As EventArgs) Handles btnAllRead.Click
        txtPath.Text = ""
        AllRead()
    End Sub

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
    Private Sub txtPath_Click(sender As Object, e As EventArgs) Handles txtPath.Click
        Dim ofd As OpenFileDialog = New OpenFileDialog With {
            .Filter = "모든 파일 (*.*) | *.*"
        }
        ofd.ShowDialog()
        Dim path As String = ofd.FileName
        txtPath.Text = ofd.FileName
        'Dim variableType As Type = txtPath.GetType() 'hsj test
        Console.WriteLine(txtPath.GetType.Name) 'hsj test
        If txtPath.Text IsNot "" Then
            NewfileLead()
        End If
    End Sub

    Private Sub TextBox4_Click(sender As Object, e As EventArgs) Handles TextBox4.Click 'hsj test 마스터DB 클릭 시 파일 선택, DB 확인, txt 박스 클릭
        Dim ofd As OpenFileDialog = New OpenFileDialog With {
            .Filter = "모든 파일 (*.*) | *.*"
        }
        ofd.ShowDialog()
        Dim path As String = ofd.FileName
        'txtPath.Text = ofd.FileName 'hsj del
        TextBox4.Text = ofd.FileName
        If TextBox4.Text IsNot "" Then
            NewfileLead()
        End If
    End Sub
    '  업로드 db 경로 불러오기 기능 - 공통 함수로 test 중, 고의로 master db 찾는 부분은 제외할 거임
    Private Sub txtbox_path_click(sender As Object, e As EventArgs) Handles txtbox_suffix_path.Click, txtbox_carrier_path.Click, txtbox_limit_path.Click, TextBox4.Click '
        Dim ofd As OpenFileDialog = New OpenFileDialog With {
            .Filter = "모든 파일 (*.*) | *.*"
        }
        ofd.ShowDialog()
        Dim path As String = ofd.FileName
        'txtPath.Text = ofd.FileName 'hsj del
        TextBox4.Text = ofd.FileName
        If TextBox4.Text IsNot "" Then
            NewfileLead()
        End If
    End Sub

    Private Sub NewfileLead()
        AllRead()
        newDBTable.Clear()
        Dim SelectStatement As String = "SELECT [No], [모델 명], FORMAT([部品SET], 'HH:mm:ss') as [部品SET], FORMAT([前付け], 'HH:mm:ss') as [前付け], FORMAT([MT], 'HH:mm:ss') as [MT], FORMAT([L/C], 'HH:mm:ss') as [L/C], FORMAT([目視], 'HH:mm:ss') as [目視], FORMAT([Pick up], 'HH:mm:ss') as [Pick up],
                                           FORMAT([組立], 'HH:mm:ss') as [組立], FORMAT([機能検査(수동)], 'HH:mm:ss') as [機能検査_수동], FORMAT([機能検査(자동)], 'HH:mm:ss') as [機能検査_자동], FORMAT([2者検査], 'HH:mm:ss') as [2者検査], [검사 설비]  FROM [Sheet1$]"

        Using cn As New OleDb.OleDbConnection With {.ConnectionString = Connection.HeaderConnectionString(TextBox4.Text)}
            'Using cn As New OleDb.OleDbConnection With {.ConnectionString = Connection.HeaderConnectionString(txtPath.Text)}
            Using cmd As New OleDbCommand With {.Connection = cn, .CommandText = SelectStatement}

                cn.Open()
                Try
                    newDBTable.Load(cmd.ExecuteReader)
                    'grdRead.DataSource = newDBTable
                    grd_master.DataSource = newDBTable

                    CompareData()

                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    MessageBox.Show("잘못된 Master 파일형식입니다.")
                End Try

            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Data Compare
    ''' </summary>
    Private Sub CompareData()
        Dim arrOb As Object() = DbTable.[Select]().[Select](Function(x) x("모델명")).ToArray()
        Dim dbModel As String() = arrOb.Cast(Of String)().ToArray()

        For i As Integer = 0 To newDBTable.Rows.Count - 1 ' Range.Rows.Count
            Try ' 에러 확인 용 try
                If dbModel.Contains(Replace(newDBTable.Rows(i).ItemArray(1).ToString(), " ", "")) Then
                    For j As Integer = 0 To DbTable.Rows.Count - 1
                        If Replace(newDBTable.Rows(i).ItemArray(1).ToString(), " ", "") = DbTable.Rows(j).ItemArray(1).ToString() Then
                            For t As Integer = 2 To 12
                                If Replace(newDBTable.Rows(i).ItemArray(t).ToString(), " ", "") = DbTable.Rows(j).ItemArray(t).ToString() Then
                                Else
                                    grd_master.Rows(i).Cells(t).Style.BackColor = Color.Yellow
                                    grd_master.Rows(i).Cells(0).Style.BackColor = Color.Yellow
                                    grd_master.Rows(i).Cells(1).Style.BackColor = Color.Yellow
                                    'grdRead.Rows(i).Cells(t).Style.BackColor = Color.Yellow  ' grdRead 변수 변경 필요
                                    'grdRead.Rows(i).Cells(0).Style.BackColor = Color.Yellow  ' Cells(0), (1)가 뭐냐
                                    'grdRead.Rows(i).Cells(1).Style.BackColor = Color.Yellow
                                End If
                            Next
                        End If
                    Next
                Else
                    For f As Integer = 0 To 12
                        grdRead.Rows(i).Cells(f).Style.BackColor = Color.DarkOrange
                    Next
                End If
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                MessageBox.Show("Data Compare NG.")

            End Try


        Next
    End Sub

    ''' <summary>
    ''' Model 검색
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnSearch_Click_1(sender As Object, e As EventArgs) Handles btnSearch.Click
        Cursor.Current = Cursors.WaitCursor
        Try
            DbTable.Rows.Clear()
            If txtSearch.Text = Nothing Then
                MsgBox("찾으실 모델명을 입력하여 주십시오")
            Else

                Dim list As List(Of String()) = New List(Of String())
                Dim ResultData As String = Nothing
                Dim SqlCMD As String = "Select * from FAM3_PRODUCT_TIME_TB WHERE MODEL Like" & "'" & (txtSearch.Text).ToUpper & "'"

                If (txtSearch.Text).Contains("*") Then
                    SqlCMD = "select * from FAM3_PRODUCT_TIME_TB WHERE MODEL LIKE" & "'%" & Replace(txtSearch.Text, CChar("*"), "").ToUpper & "%'"
                Else
                    SqlCMD = "select * from FAM3_PRODUCT_TIME_TB WHERE MODEL LIKE" & "'" & (txtSearch.Text).ToUpper & "'"
                End If

                EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)

                If ResultData = "ERROR:A0 Nothing" Then
                    MessageBox.Show("검색 결과가 존재하지 않습니다.", "warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Dim rowArray As String() = ResultData.Split(CChar(vbCrLf))
                    For i = 1 To rowArray.Length - 2
                        Dim colArray As String() = rowArray(i).Split(CChar(","))
                        list.Add(colArray)
                        DbTable.Rows.Add(i, colArray(1), colArray(2), colArray(3), colArray(4), colArray(5), colArray(6), colArray(7), colArray(8), colArray(9), colArray(10), colArray(11), Replace(colArray(12), "\c", ","))
                    Next
                End If
                grdRead.DataSource = DbTable
            End If

        Catch ex As Exception
            SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "btnSearch_Click()", ex.Message)
        End Try
        Cursor.Current = Cursors.Default
    End Sub

    ''' <summary>
    ''' Master 파일 수정
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnMasterSet_Click_1(sender As Object, e As EventArgs) Handles btnMasterSet.Click
        Dim login As Lock = New Lock()

        If txtPath.Text = "" Or txtPath.Text = "클릭시 파일 선택" Then
            MessageBox.Show("변경할 마스터 데이터 파일을 선택하여 주십시오.", "warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            login.ShowDialog()

            If login.DialogResult = DialogResult.OK Then
                Cursor.Current = Cursors.WaitCursor
                MasterRead()
                _userControlPresenter.MasterDataInput()
                masterDatalist.Clear()
                newDBTable.Rows.Clear()
                txtPath.Text = ""
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

                Dim list As DataCenter = New DataCenter(Replace(aa(1), " ", ""), aa(2), aa(3), aa(4), aa(5), aa(6), aa(7), aa(8), aa(9), aa(10), aa(11), aa(12))

                masterDatalist.Add(list)
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
    Private Sub txtSearch_Leave(sender As Object, e As EventArgs) Handles txtSearch.Leave
        If txtSearch.Text = "" Then
            txtSearch.Text = "모델명 입력"
            txtSearch.ForeColor = Color.Gray
        End If
    End Sub

    ''' <summary>
    '''  background text
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>    
    Private Sub txtSearch_Enter(sender As Object, e As EventArgs) Handles txtSearch.Enter 'hsj 텍스트 상자에서 포커스가 벗어났을때 발생하는 이벤트 핸들러
        If txtSearch.Text = "모델명 입력" Then
            txtSearch.Text = ""
            txtSearch.ForeColor = Color.Black
        End If
    End Sub

    ''' <summary>
    ''' SaveOption Dialog Show
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub 파일경로설정ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 메뉴얼ToolStripMenuItem.Click
        _userControlPresenter.ShowSaveOption()
    End Sub

    Private Sub txtPath_Leave(sender As Object, e As EventArgs) Handles txtPath.Leave
        If txtPath.Text = "" Then
            txtPath.Text = "클릭시 파일 선택"
            txtPath.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub txtPath_Enter(sender As Object, e As EventArgs) Handles txtPath.Enter
        If txtPath.Text = "클릭시 파일 선택" Then
            txtPath.Text = ""
            txtPath.ForeColor = Color.Black
        End If
    End Sub

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

    Private Sub TableLayoutPanel3_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel3.Paint

    End Sub

    Private Sub grdRead_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdRead.CellContentClick

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    'Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
    'Private Sub txtPath1_TextChanged(sender As Object, e As EventArgs) Handles txtPath1.TextChanged 'txtPath1 = Master-"클릭시 파일 선택"

    'End Sub

    Private Sub txtPath_TextChanged(sender As Object, e As EventArgs) Handles txtPath.TextChanged

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

    Private Sub txtPath1_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged 'hsj - 마스터 db ;클릭시 파일 선택; 눌렀을때 이벤트 처리?

    End Sub

    Private Sub txtbox_search_suffix_TextChanged(sender As Object, e As EventArgs) Handles txtbox_search_suffix.TextChanged

    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged

    End Sub
End Class