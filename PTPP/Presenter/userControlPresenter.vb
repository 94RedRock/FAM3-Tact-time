Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.IO
Imports System.Data
Imports Microsoft.Office.Interop.Excel

Public Class userControlPresenter

    Public _hostIP As String = ProgramConfig.ReadIniDBSetting("HostIP")
    Public _hostPort As String = ProgramConfig.ReadIniDBSetting("HostPort")
    Public _table As String = ProgramConfig.ReadIniDBSetting("MasterTable")

    Public _tableSuffix As String = ProgramConfig.ReadIniDBSetting("SuffixTable") ' 테이블 설정, 
    Public _tableCarrier As String = ProgramConfig.ReadIniDBSetting("CarrierTable")
    Public _tableLimit As String = ProgramConfig.ReadIniDBSetting("LimitTable")

    Private Connection As ExcelConnection = New ExcelConnection

    Shared excelApp As New Excel.Application
    Private book As Excel.Workbook
    Private sheet1 As Excel.Worksheet
    Private sheet2 As Excel.Worksheet

    Private _userControl As userControl = Nothing

    Public simpleModelList As New List(Of String())
    Public inspectionList As New List(Of String())

    Public Sub New(ByVal ucMonitoringView As userControl)
        _userControl = ucMonitoringView
    End Sub

    ''' <summary>
    ''' SaveOption Dialog Show
    ''' </summary>
    Public Sub ShowSaveOption()

        Dim saveOption As SaveOption = New SaveOption()
        saveOption.ShowDialog()

    End Sub

    ''' <summary>
    ''' ProductTime Data Calculator
    ''' </summary>
    ''' <returns></returns>
    Public Function DataCalculate() As List(Of String())

        Dim ErrMsg As String = Nothing
        Dim field As String() = {"MODEL", "ACCESSORY", "COMPONENT_SET", "MAEDZUKE", "MAUNT", "LEAD_CUTTING", "VISUAL_EXAMINATION", "PICKUP", "ASSAMBLY", "M_FUNCTION_CHECK", "A_FUNCTION_CHECK", "PERSON_EXAMINE", "INSPECTION_EQUIPMENT"}
        Dim list As New List(Of String())
        Dim modelList As List(Of ReadModel)
        Dim suffixList As List(Of ReadSuffix)

        (modelList, suffixList) = ModelNameRead()
        'Dim rdData(12) As String
        Dim rdData(13) As String
        Dim SqlCMD As String = ""
        Dim ResultData As String = Nothing

        Try
            For i As Integer = 0 To modelList.Count() - 1
                'SqlCMD &= " select" & " MODEL, ACCESSORY, COMPONENT_SET, MAEDZUKE, MAUNT, LEAD_CUTTING, VISUAL_EXAMINATION, PICKUP, ASSAMBLY, M_FUNCTION_CHECK, A_FUNCTION_CHECK, PERSON_EXAMINE, INSPECTION_EQUIPMENT " & "from FAM3_PRODUCT_TIME_TB WHERE MODEL = " & "'" & modelList(i).ModelName & "' union all select" & "'" & modelList(i).ModelName & " ','','','','','','','','','','','' FROM DUAL WHERE NOT EXISTS(Select * " & "from FAM3_PRODUCT_TIME_TB WHERE MODEL = " & "'" & modelList(i).ModelName & "'" & ")"
                SqlCMD &= " select" & " MODEL, ACCESSORY, COMPONENT_SET, MAEDZUKE, MAUNT, LEAD_CUTTING, VISUAL_EXAMINATION, PICKUP, ASSAMBLY, M_FUNCTION_CHECK, A_FUNCTION_CHECK, PERSON_EXAMINE, INSPECTION_EQUIPMENT " & "from FAM3_PRODUCT_TIME_TB WHERE MODEL = " & "'" & modelList(i).ModelName.Substring(0, 9) & "' union all select" & "'" & modelList(i).ModelName.Substring(0, 9) & " ','','','','','','','','','','','' FROM DUAL WHERE NOT EXISTS(Select * " & "from FAM3_PRODUCT_TIME_TB WHERE MODEL = " & "'" & modelList(i).ModelName.Substring(0, 9) & "'" & ")"
                If i < modelList.Count() - 1 Then
                    SqlCMD += " UNION ALL"
                End If

            Next

            EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)

            Dim rowArray As String() = ResultData.Split(CChar(vbCrLf))
            For i = 1 To rowArray.Length - 2
                Dim colArray As String() = rowArray(i).Split(CChar(","))
                If colArray(1) = "" Then '*****왜 두 번째 배열이 공란일때를 조건으로 설정 했을까?
                    list.Add(colArray)
                Else
                    list.Add(colArray)
                    'Replace(colArray(11), "\c", ",")
                    Replace(colArray(12), "\c", ",")
                    simpleModelList.Add({colArray(0).Substring(0, 7), colArray(1), colArray(2), colArray(3), colArray(4), colArray(5), colArray(6), colArray(7), colArray(8), colArray(9), colArray(10), colArray(11), colArray(12)})
                    inspectionList.Add({colArray(0), colArray(9), colArray(10), colArray(12)})
                    'simpleModelList.Add({colArray(0).Substring(0, 7), colArray(1), colArray(2), colArray(3), colArray(4), colArray(5), colArray(6), colArray(7), colArray(8), colArray(9), colArray(10), colArray(11)})
                    'inspectionList.Add({colArray(0), colArray(8), colArray(9), colArray(11)})
                End If

            Next
        Catch ex As Exception
            SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "DataCalculate()", ex.Message)
        End Try

        Return list

    End Function
    ''' <summary>
    ''' SuffixTime Data Calculator, hsj 
    ''' </summary>
    ''' <returns></returns>
    Public Function DataCalculateSuffix() As List(Of String()) 'data calculate test용)

        Dim ErrMsg As String = Nothing
        Dim field As String() = {"MODEL", "ACCESSORY", "COMPONENT_SET", "MAEDZUKE", "MAUNT", "LEAD_CUTTING", "VISUAL_EXAMINATION", "PICKUP", "ASSAMBLY", "M_FUNCTION_CHECK", "A_FUNCTION_CHECK", "PERSON_EXAMINE", "INSPECTION_EQUIPMENT"}
        Dim fieldSuffix As String() = {"SUFFIX", "ADDITIONAL_MAOUNTING", "ADDITIONAL_ASSEMBLY"}
        Dim list As New List(Of String())
        Dim modelList As List(Of ReadModel) = ModelNameReadSuffix()
        Dim rdData(3) As String
        Dim SqlCMD As String = ""
        Dim ResultData As String = Nothing

        Try
            For i As Integer = 0 To modelList.Count() - 1
                'SqlCMD &= " select" & " MODEL, ACCESSORY, COMPONENT_SET, MAEDZUKE, MAUNT, LEAD_CUTTING, VISUAL_EXAMINATION, PICKUP, ASSAMBLY, M_FUNCTION_CHECK, A_FUNCTION_CHECK, PERSON_EXAMINE, INSPECTION_EQUIPMENT " & "from FAM3_PRODUCT_TIME_TB WHERE MODEL = " & "'" & modelList(i).ModelName & "' union all select" & "'" & modelList(i).ModelName & " ','','','','','','','','','','','' FROM DUAL WHERE NOT EXISTS(Select * " & "from FAM3_PRODUCT_TIME_TB WHERE MODEL = " & "'" & modelList(i).ModelName & "'" & ")"
                SqlCMD &= " select" & " SUFFIX, ADDITIONAL_MAOUNTING, ADDITIONAL_ASSEMBLY " & "from FAM3_SUFFIX_TIME_TB WHERE MODEL = " & "'" & modelList(i).ModelName & "' union all select" & "'" & modelList(i).ModelName & " ','','','','','','','','','','','' FROM DUAL WHERE NOT EXISTS(Select * " & "from FAM3_SUFFIX_TIME_TB WHERE MODEL = " & "'" & modelList(i).ModelName & "'" & ")"
                If i < modelList.Count() - 1 Then
                    SqlCMD += " UNION ALL"
                End If

            Next

            EtherUty.EtherSendSQL(ProgramConfig.ReadIniDBSetting("HostIP"), 2005, SqlCMD, ResultData)

            Dim rowArray As String() = ResultData.Split(CChar(vbCrLf))
            For i = 1 To rowArray.Length - 2
                Dim colArray As String() = rowArray(i).Split(CChar(","))
                If colArray(1) = "" Then '*****왜 두 번째 배열이 공란일때를 조건으로 설정 했을까?
                    list.Add(colArray)
                Else
                    list.Add(colArray)
                    'Replace(colArray(11), "\c", ",")
                    Replace(colArray(12), "\c", ",")
                    simpleModelList.Add({colArray(0).Substring(0, 7), colArray(1), colArray(2), colArray(3), colArray(4), colArray(5), colArray(6), colArray(7), colArray(8), colArray(9), colArray(10), colArray(11), colArray(12)})
                    inspectionList.Add({colArray(0), colArray(9), colArray(10), colArray(12)})
                    'simpleModelList.Add({colArray(0).Substring(0, 7), colArray(1), colArray(2), colArray(3), colArray(4), colArray(5), colArray(6), colArray(7), colArray(8), colArray(9), colArray(10), colArray(11)})
                    'inspectionList.Add({colArray(0), colArray(8), colArray(9), colArray(11)})
                End If

            Next
        Catch ex As Exception
            SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "DataCalculate()", ex.Message)
        End Try

        Return list

    End Function

    ''' <summary>
    ''' ExcelFile ModelName Read list
    ''' </summary>
    Public Function ModelNameRead() As (List(Of ReadModel), List(Of ReadSuffix))
        Dim ModelNamelist As New List(Of ReadModel)
        Dim SuffixNamelist As New List(Of ReadSuffix) 'hsj add

        'Dim strFile As String = ProgramConfig.ReadIniUserSetting("InputfilePath") + "\" + ProgramConfig.ReadIniUserSetting("InputFileName") + "20210322"
        Dim strFile As String = ProgramConfig.ReadIniUserSetting("InputfilePath") + "\" + ProgramConfig.ReadIniUserSetting("InputFileName") + DateTime.Now.ToString("yyyyMMdd")

        Dim fileInfo As FileInfo = New FileInfo(strFile + ".xls")
        Dim fileInfo2 As FileInfo = New FileInfo(strFile + ".xlsx")

        If fileInfo.Exists Or fileInfo2.Exists Then
            'Dim oBook As Object = excelApp.Workbooks.Open(ProgramConfig.ReadIniUserSetting("InputfilePath") + "\" + ProgramConfig.ReadIniUserSetting("InputFileName") + "20210322")
            Dim oBook As Object = excelApp.Workbooks.Open(ProgramConfig.ReadIniUserSetting("InputfilePath") + "\" + ProgramConfig.ReadIniUserSetting("InputFileName") + DateTime.Now.ToString("yyyyMMdd"))
            Dim oSheet As Object = excelApp.Worksheets(1)

            Dim range As Range = oSheet.UsedRange
            Dim data As String = Nothing

            Try
                For i As Integer = 3 To range.Rows.Count

                    '워크시트 변경에 따라 모델명 읽어오는 행 수정 (14 → 16) - Ver 1.01 KJ
                    '모델 명 외에도 읽어오는 에러 수정

                    'data = range.Cells(i, 14).Value
                    data = range.Cells(i, 16).Value

                    If data IsNot Nothing Then
                        If Len(data) > 0 And Not data.Equals("ORDER ENTRY CODE") Then
                            Dim data1 As ReadModel = New ReadModel(data)
                            ModelNamelist.Add(data1)
                        End If
                    End If
                Next

            Catch ex As Exception
                SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "ModelNameRead()", ex.Message)
            End Try
            Try
                excelApp.DisplayAlerts = False
            Catch __unusedException1__ As Exception
            Finally
                excelApp.Workbooks.Close()
                excelApp.Quit()
            End Try
        Else
            MsgBox(DateTime.Now.ToString("yyyy/MM/dd") + " 작업지시서파일이 존재하지않습니다")
        End If

        Return (ModelNamelist, SuffixNamelist)  '맞음?


        '====================== hsj add start======================
        'Try
        '    For i As Integer = 3 To Range.Rows.Count

        '        '워크시트 변경에 따라 모델명 읽어오는 행 수정 (14 → 16) - Ver 1.01 KJ
        '        '모델 명 외에도 읽어오는 에러 수정

        '        'data = range.Cells(i, 14).Value
        '        Data = Range.Cells(i, 16).Value

        '        If Data IsNot Nothing Then
        '            If Len(Data) > 0 And Not Data.Equals("ORDER ENTRY CODE") Then
        '                Dim data1 As ReadModel = New ReadModel(Data)
        '                ModelNamelist.Add(data1)
        '            End If
        '        End If
        '    Next

        'Catch ex As Exception
        '    SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "ModelNameRead()", ex.Message)
        'End Try
        'Try
        '    excelApp.DisplayAlerts = False
        'Catch __unusedException1__ As Exception
        'Finally
        '    excelApp.Workbooks.Close()
        '    excelApp.Quit()
        'End Try
        'Else
        'MsgBox(DateTime.Now.ToString("yyyy/MM/dd") + " 작업지시서파일이 존재하지않습니다")
        'End If

    End Function
    Public Function ModelNameReadSuffix() As List(Of ReadModel)
        Dim ModelNamelist As New List(Of ReadModel)

        'Dim strFile As String = ProgramConfig.ReadIniUserSetting("InputfilePath") + "\" + ProgramConfig.ReadIniUserSetting("InputFileName") + "20210322"
        Dim strFile As String = ProgramConfig.ReadIniUserSetting("InputfilePath") + "\" + ProgramConfig.ReadIniUserSetting("InputFileName") + DateTime.Now.ToString("yyyyMMdd")

        Dim fileInfo As FileInfo = New FileInfo(strFile + ".xls")
        Dim fileInfo2 As FileInfo = New FileInfo(strFile + ".xlsx")

        If fileInfo.Exists Or fileInfo2.Exists Then
            'Dim oBook As Object = excelApp.Workbooks.Open(ProgramConfig.ReadIniUserSetting("InputfilePath") + "\" + ProgramConfig.ReadIniUserSetting("InputFileName") + "20210322")
            Dim oBook As Object = excelApp.Workbooks.Open(ProgramConfig.ReadIniUserSetting("InputfilePath") + "\" + ProgramConfig.ReadIniUserSetting("InputFileName") + DateTime.Now.ToString("yyyyMMdd"))
            Dim oSheet As Object = excelApp.Worksheets(1)

            Dim range As Range = oSheet.UsedRange
            Dim data As String = Nothing

            Try
                For i As Integer = 3 To range.Rows.Count

                    '워크시트 변경에 따라 모델명 읽어오는 행 수정 (14 → 16) - Ver 1.01 KJ
                    '모델 명 외에도 읽어오는 에러 수정

                    'data = range.Cells(i, 14).Value
                    data = range.Cells(i, 16).Value

                    If data IsNot Nothing Then
                        If Len(data) > 0 And Not data.Equals("ORDER ENTRY CODE") Then
                            Dim data1 As ReadModel = New ReadModel(data)
                            ModelNamelist.Add(data1)
                        End If
                    End If
                Next

            Catch ex As Exception
                SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "ModelNameRead()", ex.Message)
            End Try
            Try
                excelApp.DisplayAlerts = False
            Catch __unusedException1__ As Exception
            Finally
                excelApp.Workbooks.Close()
                excelApp.Quit()
            End Try
        Else
            MsgBox(DateTime.Now.ToString("yyyy/MM/dd") + " 작업지시서파일이 존재하지않습니다")
        End If

        Return ModelNamelist

    End Function

    ''' <summary>
    ''' Master Data 수정
    ''' </summary>
    Public Sub MasterDataInputNew(sender As Object)

        Dim RtnData As String = Nothing
        Dim ErrMsg As String = Nothing
        Dim QDBWResult As Boolean
        'Dim Field As String() = {"RECNO", "MODEL", "COMPONENT_SET", "MAEDZUKE", "MAUNT", "LEAD_CUTTING", "VISUAL_EXAMINATION", "PICKUP", "ASSAMBLY", "M_FUNCTION_CHECK", "A_FUNCTION_CHECK", "PERSON_EXAMINE", "INSPECTION_EQUIPMENT", "SOFT_NAME", "SOFT_VERSION", "REVISE_DATE"}

        Dim MasterDataList As List(Of DataCenter) = _userControl.masterDatalist
        Dim MasterDataListSuffix As List(Of DataCenterSuffix) = _userControl.masterDatalistSuffix
        Dim MasterDataListCarrier As List(Of DataCenterCarrier) = _userControl.masterDatalistCarrier
        Dim MasterDataListLimit As List(Of DataCenterLimit) = _userControl.masterDatalistLimit

        Dim Field As String() = {}
        Dim WrData As String() = {}
        Dim rdData As String() = {}
        Dim ResultData As String = Nothing

        If sender.Text.IndexOf("Suffix") >= 0 Then
            Field = {"RECNO", "SUFFIX", "ADDITIONAL_MAOUNTING", "ADDITIONAL_ASSEMBLY", "SOFT_NAME", "SOFT_VERSION", "REVISE_DATE"}

            Console.WriteLine(DateTime.Now.ToString("hh:mm:ss"))
            Try
                For i As Integer = 0 To MasterDataListSuffix.Count() - 1 ' 
                    Dim MxMnResult = EtherUty.EtherMXMN(_hostIP, Convert.ToInt32(_hostPort), _tableSuffix, "RECNO", RtnData) 'max 값을 왜 읽을까
                    Select Case MxMnResult
                        Case True
                            WrData = {"", MasterDataListSuffix(i).Suffix, MasterDataListSuffix(i).AdditionalMount, MasterDataListSuffix(i).AdditionalAssembly,
                                         "PTPP", "1.0.0", DateTime.Now.ToString("yyyy/MM/dd")} ' ptpp, 1.0.0 자동 변경 필요함

                            Dim ChkResult = EtherUty.QDBRead(_hostIP, Convert.ToInt32(_hostPort), _tableSuffix, "SUFFIX", MasterDataListSuffix(i).Suffix, Field, rdData, ErrMsg)
                            'QDBR   테이블명, 대상컬럼명 = 대상키, 취득할 컬럼명.

                            If ChkResult = False Then
                                WrData(0) = CStr(CInt(RtnData) + 1)
                                QDBWResult = EtherUty.QDBWrite(_hostIP, Convert.ToInt32(_hostPort), _tableSuffix, Field, WrData, ErrMsg)
                            Else
                                WrData(0) = rdData(0)
                                QDBWResult = EtherUty.QDBWrite(_hostIP, Convert.ToInt32(_hostPort), _tableSuffix, Field, WrData, ErrMsg, "U") 'U가 뭐냐 >> 갱신임ㅋ, I : 신규등록
                            End If

                    End Select
                Next
            Catch ex As Exception
                SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "RegistWorker()", ex.Message)
            End Try
            Console.WriteLine(DateTime.Now.ToString("hh:mm:ss"))

            'className = GetType(DataCenter).Name
        ElseIf sender.Text.IndexOf("Carrier") >= 0 Then

            Field = {"RECNO", "MODEL", "CARRIER", "SOFT_NAME", "SOFT_VERSION", "REVISE_DATE"}
            Console.WriteLine(DateTime.Now.ToString("hh:mm:ss"))
            Try
                For i As Integer = 0 To MasterDataListCarrier.Count() - 1
                    Dim MxMnResult = EtherUty.EtherMXMN(_hostIP, Convert.ToInt32(_hostPort), _tableCarrier, "RECNO", RtnData)
                    Select Case MxMnResult
                        Case True
                            WrData = {"", MasterDataListCarrier(i).CarrierModel, MasterDataListCarrier(i).CarrierCarrier,
                                         "PTPP", "1.0.0", DateTime.Now.ToString("yyyy/MM/dd")}

                            Dim ChkResult = EtherUty.QDBRead(_hostIP, Convert.ToInt32(_hostPort), _tableCarrier, "MODEL", MasterDataListCarrier(i).CarrierModel, Field, rdData, ErrMsg)

                            If ChkResult = False Then
                                WrData(0) = CStr(CInt(RtnData) + 1)
                                QDBWResult = EtherUty.QDBWrite(_hostIP, Convert.ToInt32(_hostPort), _tableCarrier, Field, WrData, ErrMsg)
                            Else
                                WrData(0) = rdData(0)
                                QDBWResult = EtherUty.QDBWrite(_hostIP, Convert.ToInt32(_hostPort), _tableCarrier, Field, WrData, ErrMsg, "U")
                            End If

                    End Select
                Next
            Catch ex As Exception
                SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "RegistWorker()", ex.Message)
            End Try
            Console.WriteLine(DateTime.Now.ToString("hh:mm:ss"))

        ElseIf sender.Text.IndexOf("Limit") >= 0 Then
            Field = {"RECNO", "CARRIER", "LIMIT", "QUANTITY", "SOFT_NAME", "SOFT_VERSION", "REVISE_DATE"}
            Console.WriteLine(DateTime.Now.ToString("hh:mm:ss"))
            Try
                For i As Integer = 0 To MasterDataListLimit.Count() - 1
                    Dim MxMnResult = EtherUty.EtherMXMN(_hostIP, Convert.ToInt32(_hostPort), _tableLimit, "RECNO", RtnData)
                    Select Case MxMnResult
                        Case True
                            WrData = {"", MasterDataListLimit(i).Carrier, MasterDataListLimit(i).Limit, MasterDataListLimit(i).Quantity,
                                         "PTPP", "1.0.0", DateTime.Now.ToString("yyyy/MM/dd")}

                            Dim ChkResult = EtherUty.QDBRead(_hostIP, Convert.ToInt32(_hostPort), _tableLimit, "CARRIER", MasterDataListLimit(i).Carrier, Field, rdData, ErrMsg)

                            If ChkResult = False Then
                                WrData(0) = CStr(CInt(RtnData) + 1)
                                QDBWResult = EtherUty.QDBWrite(_hostIP, Convert.ToInt32(_hostPort), _tableLimit, Field, WrData, ErrMsg)
                            Else
                                WrData(0) = rdData(0)
                                QDBWResult = EtherUty.QDBWrite(_hostIP, Convert.ToInt32(_hostPort), _tableLimit, Field, WrData, ErrMsg, "U")
                            End If

                    End Select
                Next
            Catch ex As Exception
                SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "RegistWorker()", ex.Message)
            End Try
            Console.WriteLine(DateTime.Now.ToString("hh:mm:ss"))

        ElseIf sender.Text.IndexOf("Master") >= 0 Then
            Field = {"RECNO", "MODEL", "ACCESSORY", "COMPONENT_SET", "MAEDZUKE", "MAUNT", "LEAD_CUTTING", "VISUAL_EXAMINATION", "PICKUP", "ASSEMBLY", "M_FUNCTION_CHECK", "A_FUNCTION_CHECK", "PERSON_EXAMINE", "INSPECTION_EQUIPMENT", "SOFT_NAME", "SOFT_VERSION", "REVISE_DATE"}
            Console.WriteLine(DateTime.Now.ToString("hh:mm:ss"))
            Try
                For i As Integer = 0 To MasterDataList.Count() - 1
                    Dim MxMnResult = EtherUty.EtherMXMN(_hostIP, Convert.ToInt32(_hostPort), _table, "RECNO", RtnData)
                    Select Case MxMnResult
                        Case True
                            WrData = {"", MasterDataList(i).Model, MasterDataList(i).Accessory, MasterDataList(i).ComponentSet, MasterDataList(i).Maedzuke, MasterDataList(i).Mount,
                                     MasterDataList(i).LeadCutting, MasterDataList(i).VisualExamination, MasterDataList(i).PickUp, MasterDataList(i).Assambly,
                                     MasterDataList(i).MFunctionCheck, MasterDataList(i).AFunctionCheck, MasterDataList(i).PersonalExamine, MasterDataList(i).ExamineEquipment,
                                     "PTPP", "1.0.0", DateTime.Now.ToString("yyyy/MM/dd")}

                            Dim ChkResult = EtherUty.QDBRead(_hostIP, Convert.ToInt32(_hostPort), _table, "MODEL", MasterDataList(i).Model, Field, rdData, ErrMsg)

                            If ChkResult = False Then
                                WrData(0) = CStr(CInt(RtnData) + 1)
                                QDBWResult = EtherUty.QDBWrite(_hostIP, Convert.ToInt32(_hostPort), _table, Field, WrData, ErrMsg)
                            Else
                                WrData(0) = rdData(0)
                                QDBWResult = EtherUty.QDBWrite(_hostIP, Convert.ToInt32(_hostPort), _table, Field, WrData, ErrMsg, "U")
                            End If

                    End Select
                Next
            Catch ex As Exception
                SystemLogger.Instance.ErrorLog(ProgramEnum.LogType.File, "RegistWorker()", ex.Message)
            End Try
            Console.WriteLine(DateTime.Now.ToString("hh:mm:ss"))
        End If

    End Sub

    ''' <summary>
    ''' Excel Save
    ''' </summary>
    ''' <param name="sheetList"></param>
    ''' <param name="tableList"></param>
    Public Sub Save_Excel(sheetList As List(Of String), tableList As List(Of System.Data.DataTable))

        Dim TargetPath = ".\" + ProgramConfig.ReadIniUserSetting("OutputFileName") + DateTime.Now.ToString("yyyyMMdd") + ".xlsx"
        If File.Exists(ProgramConfig.ReadIniUserSetting("SaveFilePath") + TargetPath) Then File.Delete(ProgramConfig.ReadIniUserSetting("SaveFilePath") + TargetPath)

        Using cn As New OleDb.OleDbConnection With {.ConnectionString = Connection.WriteConnectionString(ProgramConfig.ReadIniUserSetting("SaveFilePath") + TargetPath)}
            cn.Open()

            For i As Integer = 0 To 2

                Dim ColNames As String = Nothing
                Dim ColParams As String = Nothing
                Dim ColNamesTypes As String = Nothing

                For Each DCol As DataColumn In tableList(i).Columns
                    ColNames &= "[" & DCol.ColumnName & "],"
                    ColParams &= "@" & DCol.ColumnName & ","
                    ColNamesTypes &= "[" & DCol.ColumnName & "]" & " String,"
                Next

                ColNames = ColNames.Substring(0, ColNames.Length - 1)
                ColParams = ColParams.Substring(0, ColParams.Length - 1)
                ColNamesTypes = ColNamesTypes.Substring(0, ColNamesTypes.Length - 1)

                Using CreateTableCMD As New OleDb.OleDbCommand("CREATE TABLE " & sheetList(i) &
                                                                     "(" & ColNamesTypes & ")", cn)
                    CreateTableCMD.ExecuteNonQuery()
                End Using

                Dim TotalRows As Integer = tableList(i).Rows.Count

                For Each Drow As DataRow In tableList(i).Rows
                    Using InsertCMD As New OleDb.OleDbCommand("INSERT INTO " & sheetList(i) & " (" & ColNames & ") VALUES (" &
                                                      ColParams & ")", cn)
                        For Each Dcol As DataColumn In tableList(i).Columns
                            InsertCMD.Parameters.AddWithValue("@" & Dcol.ColumnName, Drow(Dcol.ColumnName).ToString)
                        Next

                        InsertCMD.ExecuteNonQuery()

                    End Using
                Next
            Next

            cn.Close()
            MessageBox.Show("저장이 완료되었습니다.")
        End Using
    End Sub

End Class
