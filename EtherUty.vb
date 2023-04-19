Public Class EtherUty
    'イーサネットアクセス関連共通ルーチン
    '   Ver     Data        By
    '   0.00    99.05.27    T.Nakagiri    新規作成
    '   0.01    99.06.30    T.Nakagiri    リトライ処理変更
    '   0.02    99.08.04    T.Nakagiri    リトライルーチン変更、ByVal 追加
    '   0.03    99.09.09    T.Nakagiri    エラーメッセージ送信先の変更
    '   0.03    99.09.27    T.Nakagiri    再送処理のＭａｉｌアドレスの追加
    '   0.04    99.10.25    T.Nakagiri    エラー復帰後のＭａｉｌ送信追加
    '   0.05    99.11.08    T.Nakagiri    再送処理の修正
    '   1.00    04.02.02    M.Otawa       ＥＪＸデータベース対応
    '   1.01    04.03.01    M.Otawa       バグ修正
    '   1.02    04.05.11    M.Otawa       QDBWrite書き込みエラーはすべてリトライ対象にする。
    '   1.03    04.09.16    M.Otawa       バグ修正
    '   1.04    06.07.05    M.Otawa       QDBWrite_NotSaveText作成
    '   1.05    08.02.26    T.Yoshihara   QDBReadで"ERROR:B0,C0"はリトライ対象から外す
    '   1.06    08.03.12    T.Yoshihara   QDBReadで"ERROR:C0"はログSaveはしない
    '   1.07    09.04.24    M.Otawa       QDBWrite,QDBWrite_NotSaveText 引数の配列を保護
    '   2.00    13.02.12    T.Yoshihara   .Net対応
    '   2.01    14.03.10    T.Yoshihara   QDBWrite,EtherRetryバグ修正
    '   2.02    15.04.20    T.Yoshihara   QDBReadでセカンダリサーバー指定可能
    '   2.03    15.06.18    T.Yoshihara   QDBWriteのバックアップ保存 & EtherRtryバグ修正
    '   2.04    17.09.21    T.Yoshihara   QDBReadのError時バグ修正
    '                                     QDBWriteでセカンダリサーバー指定可能
    '   2.05    20.01.16    J.Ninomiya    VB2019に対応
    Private Const EtherUtyVer As String = "Ver:" + "2.05"
    Private Const EtherUtyDate As String = "Date:" + "20.01.16"

    Public Ethercom As New EtherCom.Dll
    Public MailCount%                      'エラーメッセージ送信の為の異常検出回数
    Public MailSendFlg As Boolean
    Private EtherErrorFile$                'Ethercom 転送失敗データ保存先
    Public EtherRetryMax%                  'Ｅｔｈｅｒｃｏｍ再処理回数
    Public FileRetryMax%                   'FileAccess再処理回数
    Public ErrMsgDispFlg As Boolean        'Error Massage 表示 On/Off（True:有り False:無し）
    Private LogPath$
    '
    '                   C0              　C1              C2            C3              C4
    Private Const Cmsg$ = "レコードが無い。|プログラムエラー|オラクルエラー|送信フレーム異常|データ保存失敗"
    Public Function EtherQicWrite(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal Key$, ByVal Data$) As Boolean
        'ＱＩＣデータ保存
        On Error Resume Next
        Dim s$, i%
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$, SerialNo$
        Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない

        EtherQicWrite = False
        Label1.Text = "EtherQicWrite"
        If yLen(Key$) = 0 Then Exit Function
        If yLen(Data$) = 0 Then Exit Function

        '送信データの作成
        SendDt$ = "QSV-"
        i = yInStr(Key$, "=")
        If i > 0 Then
            SerialNo$ = yLeft(yLeft(Key$, i - 1) + ySpace(12), 12)
        Else
            SerialNo$ = yLeft(Key$ + ySpace(12), 12)
        End If
        SendDt$ = SendDt$ + SerialNo$ + Data$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = Ethercom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If (yInStr(RtnMsg$, "ERROR") > 0 Or yInStr(RtnMsg$, yTrim(SerialNo$)) = 0) Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＱＩＣデータ　転送異常" + yvbCrLf() _
                    + "データ保存できない可能性があります！！"
                Label3.Text = ServerIP$ + "," + CStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If CBool(yInStr(RtnMsg$, yTrim(SerialNo$))) Then
            '正常終了
            EtherQicWrite = True
        Else
            '異常終了
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            DataSendError(s$)
            LogSave("QICData Write Error :" + RtnMsg$ + ":" + s$)
            EtherQicWrite = True
        End If
        Me.Hide()
    End Function
    Public Function EtherDBWrite(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal NewDB$, ByVal Table$, ByVal Key$, ByVal Data$) As Boolean
        'ＤＢデータ新規保存
        On Error Resume Next
        Dim s$
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$
        Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない
        EtherDBWrite = False
        Label1.Text = "EtherDBWrite"
        If yLen(FileName$) = 0 Then Exit Function
        If yLen(NewDB$) = 0 Then Exit Function
        If yLen(Table$) = 0 Then Exit Function
        If yLen(Key$) = 0 Then Exit Function
        If yLen(Data$) = 0 Then Exit Function

        '送信データの作成
        SendDt$ = "DBWT-" + FileName$ + "," + NewDB$ + "," + Table$ + "," + Key$ + "," + Data$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = EtherCom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If yInStr(RtnMsg$, "ERROR") > 0 And ErrCount < EtherRetryMax Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＤＢデータ　転送異常" + yvbCrLf() _
                    + "データ保存できない可能性があります！！"
                Label3.Text = ServerIP$ + "," + CStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            EtherDBWrite = True
        Else
            '異常終了
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            DataSendError(s$)
            LogSave("DBData Write Error :" + RtnMsg$ + ":" + s$)
            EtherDBWrite = True
        End If
        Me.Hide()
    End Function
    Public Function EtherDBRead(ByVal ServerIP$, ByVal ServerSIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal Table$, ByVal Key$, ByRef Data$) As Boolean
        'ＤＢデータ読込
        On Error Resume Next
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$

        EtherDBRead = False
        Label1.Text = "EtherDBRead"
        If yLen(FileName$) = 0 Then Exit Function
        If yLen(Table$) = 0 Then Exit Function
        If yLen(Key$) = 0 Then Exit Function
        If yLen(Data$) = 0 Then Exit Function

        '送信データの作成
        SendDt$ = "DBRD-" + FileName$ + "," + Table$ + "," + Key$ + "," + Data$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = Ethercom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If yInStr(RtnMsg$, "ERROR") > 0 Then
            If yInStr(RtnMsg$, "ERROR:31") = 0 Then  'Data not found
                Wait(1)
                Debug.Print(CStr(ErrCount) + " " + SendDt$)
                ErrCount += 1
                If ErrCount < EtherRetryMax Then GoTo Exec
            End If
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            Data$ = RtnMsg$
            EtherDBRead = True
        Else
            '異常終了
            Data$ = RtnMsg$
            LogSave("DBData Read Error :" + RtnMsg$ + ":" + SendDt$)
            EtherDBRead = True
        End If
        Me.Hide()
    End Function
    Public Function EtherDBDelete(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal Table$, ByVal Key$) As Boolean
        'ＤＢデータ削除
        On Error Resume Next
        Dim s$
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$
        Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない

        EtherDBDelete = False
        Label1.Text = "EtherDBDelete"
        If yLen(FileName$) = 0 Then Exit Function
        If yLen(Table$) = 0 Then Exit Function
        If yLen(Key$) = 0 Then Exit Function

        '送信データの作成
        SendDt$ = "DBDL-" + FileName$ + "," + Table$ + "," + Key$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = EtherCom.ComHost(ServerIP$, "", ServerPort, SendDt$)
        If yInStr(RtnMsg$, "ERROR") > 0 Then
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            EtherDBDelete = True
        Else
            '異常終了
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            DataSendError(s$)
            LogSave("DBData Delete Error :" + RtnMsg$ + ":" + s$)
            EtherDBDelete = True
        End If
        Me.Hide()
    End Function
    Public Function EtherDBReplace(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal NewDB$, ByVal Table$, ByVal Key$, ByVal Data$) As Boolean
        'ＤＢデータ部分保存
        On Error Resume Next
        Dim s$
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$
        Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない

        EtherDBReplace = False
        Label1.Text = "EtherDBReplace"
        If yLen(FileName$) = 0 Then Exit Function
        If yLen(NewDB$) = 0 Then Exit Function
        If yLen(Table$) = 0 Then Exit Function
        If yLen(Key$) = 0 Then Exit Function
        If yLen(Data$) = 0 Then Exit Function

        '送信データの作成
        SendDt$ = "DBRP-" + FileName$ + "," + NewDB$ + "," + Table$ + "," + Key$ + "," + Data$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = EtherCom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If yInStr(RtnMsg$, "ERROR") > 0 Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＤＢデータ　転送異常" + yvbCrLf() _
                    + "データ保存できない可能性があります！！"
                Label3.Text = ServerIP$ + "," + CStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            EtherDBReplace = True
        Else
            '異常終了
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            DataSendError(s$)
            LogSave("DBData Append Error :" + RtnMsg$ + ":" + s$)
            EtherDBReplace = True
        End If
        Me.Hide()
    End Function
    Public Function EtherFlRead(ByVal ServerIP$, ByVal ServerSIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal Data$) As Boolean
        'Ｔｅｘｔデータ追加保存
        On Error Resume Next
        Dim s$
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$

        EtherFlRead = False
        Label1.Text = "EtherFlRead"
        If yLen(FileName$) = 0 Then Exit Function

        '送信データの作成
        SendDt$ = "FLRD-" + FileName$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = EtherCom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If CBool(yInStr(RtnMsg$, "ERROR")) Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＴＥＸＴデータ　読込異常"
                Label3.Text = ServerIP$ + "," + CStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            Data$ = RtnMsg$
            EtherFlRead = True
        Else
            '異常終了
            Data$ = RtnMsg$
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            LogSave("TextData Read Error :" + RtnMsg$ + ":" + s$)
            EtherFlRead = True
        End If
        Me.Hide()
    End Function
    Public Function EtherFlWrite(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal Data$) As Boolean
        'Ｔｅｘｔデータ新規保存
        On Error Resume Next
        Dim s$
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$
        Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない
        EtherFlWrite = False
        Label1.Text = "EtherFlWrite"
        If yLen(FileName$) = 0 Then Exit Function
        If yLen(Data$) = 0 Then Exit Function
        '送信データの作成
        SendDt$ = "FLWT-" + FileName$ + "," + Data$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = EtherCom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If CBool(yInStr(RtnMsg$, "ERROR")) Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＴＥＸＴデータ　転送異常" + yvbCrLf() _
                    + "データ保存できない可能性があります！！"
                Label3.Text = ServerIP$ + "," + CStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            EtherFlWrite = True
        Else
            '異常終了
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            DataSendError(s$)
            LogSave("TextData Write Error :" + RtnMsg$ + ":" + s$)
            EtherFlWrite = True
        End If
        Me.Hide()
    End Function
    Public Function EtherTxAppend(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal Data$) As Boolean
        'Ｔｅｘｔデータ追加保存
        On Error Resume Next
        Dim s$
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$
        Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない

        EtherTxAppend = False
        Label1.Text = "EtherTxAppend"
        If yLen(FileName$) = 0 Then Exit Function
        If yLen(Data$) = 0 Then Exit Function

        '送信データの作成
        SendDt$ = "TXAP-" + FileName$ + "," + Data$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = Ethercom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If CBool(yInStr(RtnMsg$, "ERROR")) Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＴＥＸＴデータ　転送異常" + yvbCrLf() _
                    + "データ保存できない可能性があります！！"
                Label3.Text = ServerIP$ + "," + CStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            EtherTxAppend = True
        Else
            '異常終了
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            DataSendError(s$)
            LogSave("TextData Append Error :" + RtnMsg$ + ":" + s$)
            EtherTxAppend = True
        End If
        Me.Hide()
    End Function
    Public Function EtherTxRead(ByVal ServerIP$, ByVal ServerSIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal Data$) As Boolean
        'Ｔｅｘｔデータ追加保存
        On Error Resume Next
        Dim s$
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$

        EtherTxRead = False
        Label1.Text = "EtherTxRead"
        If yLen(FileName$) = 0 Then Exit Function
        If yLen(Data$) = 0 Then Exit Function
        '送信データの作成
        SendDt$ = "TXRD-" + FileName$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = EtherCom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If CBool(yInStr(RtnMsg$, "ERROR")) Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＴＥＸＴデータ　読込異常"
                Label3.Text = ServerIP$ + "," + CStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            Data$ = RtnMsg$
            EtherTxRead = True
        Else
            '異常終了
            Data$ = RtnMsg$
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            LogSave("TextData Read Error :" + RtnMsg$ + ":" + s$)
            EtherTxRead = True
        End If
        Me.Hide()
    End Function
    Public Function EtherTxWrite(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal FileName$, ByVal Data$) As Boolean
        'Ｔｅｘｔデータ新規保存
        On Error Resume Next
        Dim s$
        Dim RtnMsg$
        Dim ErrCount As Integer
        Dim SendDt$
        Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない
        EtherTxWrite = False
        Label1.Text = "EtherTxWrite"
        If yLen(FileName$) = 0 Then Exit Function
        If yLen(Data$) = 0 Then Exit Function
        '送信データの作成
        SendDt$ = "TXWT-" + FileName$ + "," + Data$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = Ethercom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If CBool(yInStr(RtnMsg$, "ERROR")) Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＴＥＸＴデータ　転送異常" + yvbCrLf() _
                    + "データ保存できない可能性があります！！"
                Label3.Text = ServerIP$ + "," + CStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(CStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            EtherTxWrite = True
        Else
            '異常終了
            s$ = ServerIP$ + "," + ServerSIP$ + "," + CStr(ServerPort) + "," + SendDt$
            DataSendError(s$)
            LogSave("TextData Write Error :" + RtnMsg$ + ":" + s$)
            EtherTxWrite = True
        End If
        Me.Hide()
    End Function
    Public Function EtherRetry(ByVal MailServerIP$, ByVal MailServerSIP$, ByVal MailServerPort As Integer, ByVal MyPcName$, ByVal MailFrom$, ByVal MailSend$) As Integer
        'データ転送失敗したデータを再転送
        On Error Resume Next
        Dim i%, j%, k%, iStart%, iEnd%
        Const RetryMax As Integer = 10000
        Dim a$(RetryMax), b$(RetryMax), s$
        Dim RtnMsg$, ErrCount%
        Dim RetryIP$ = ""
        Dim RetrySIP$ = ""
        Dim RetryPort As Integer
        Dim SendDt$
        Dim ErrorCount%, DataMax%, SendCount%
        EtherRetry = 0  '書込残件
        SendCount = 0   '送信完了件数
        ErrorCount = 0  '送信失敗件数
        Label1.Text = "EtherError 未保存データを転送中です。"
        Label2.Text = "しばらくお待ち下さい！"
        Me.Show()
        If yDir(EtherErrorFile$) = "" Then
            Me.Hide()
            Exit Function
        Else
            If yFileLen(EtherErrorFile$) = 0 Then
                Me.Hide()
                Exit Function
            End If
        End If
        FileLRead(EtherErrorFile$, a$, DataMax)
        i = 1
        Do Until i > RetryMax Or i > DataMax
            If yLeft(a$(i - 1), 1) = "@" Then
                iStart = i
                s$ = yMid(a$(i - 1), 2)      '＠マークを削除する。
                Do Until a$.Count >= i OrElse yLeft(a$(i), 1) = "@"
                    '次の＠マークまで加算する。
                    s$ = s$ + yvbCrLf() + a$(i)
                    i += 1
                Loop
                iEnd = i
                j = yInStr(s$, ",")  '日付、時刻
                If j > 0 Then s$ = yMid(s$, j + 1)
                j = yInStr(s$, ",")  'ＩＰアドレス
                If j > 0 Then RetryIP$ = yLeft(s$, j - 1) : s$ = yMid(s$, j + 1)
                j = yInStr(s$, ",")  'ＩＰアドレス
                If j > 0 Then
                    RetrySIP$ = yLeft(s$, j - 1) : s$ = yMid(s$, j + 1)
                    j = yInStr(s$, ",")  'ポート番号、転送データ
                    If j > 0 Then
                        RetryPort = CInt(yLeft(s$, j - 1))
                        SendDt$ = yMid(s$, j + 1)
                        If RetryPort <> 0 And SendDt$ <> "" Then
                            ErrCount = 0
                            'データベースパソコンへ転送します。
                            Label3.Text = CStr(i) + "/" + CStr(DataMax) + " :" + yLeft(SendDt$, 30)
                            yDoEvents()
                            RtnMsg$ = Ethercom.ComHost(RetryIP$, RetrySIP$, RetryPort, SendDt$)
                            Label4.Text = CStr(i) & " --> " & RtnMsg$
                            If yInStr(RtnMsg$, "ERROR") = 0 Then
                                ' 正常終了
                                SendCount += 1
                                k = iStart
                                Do Until k > iEnd
                                    a$(k - 1) = ""
                                    k += 1
                                Loop
                                If SendCount > 20 Then Exit Do
                            Else
                                ErrorCount += 1
                                If ErrorCount > 5 Then Exit Do
                            End If
                        End If
                    End If
                End If
                yDoEvents()
                Wait(0.5)
            End If
            i += 1
        Loop

        '書込ミスのデータをファイルに戻す
        i = 1 : j = 0 : k = 0
        Do Until i > RetryMax Or i > DataMax
            yDoEvents()
            If a$(i - 1) <> "" Then
                b$(j) = a$(i - 1)
                j += 1
                If CBool(YInStr(a$(i - 1), "@")) Then k += 1
            End If
            i += 1
        Loop
        ReDim Preserve b$(yUBound(a$))
        If j > 0 Then
            EtherRetry = j
            FileLWrite(EtherErrorFile$, b$)
            If k >= MailCount And MailSendFlg = False Then
                SendDt$ = "Mail-" + MyPcName + "," + MailFrom + "," + MailSend _
                    + "," + MyPcName + " EtherRetry Error,データ転送失敗件数が多くなりました。" + yvbCrLf() _
                    + "エラーログを調べて下さい。" + yvbCrLf() _
                    + "     MyPCName     :" + MyPcName + yvbCrLf() _
                    + "     転送先サーバー:" + RetryIP$ + " " + RetrySIP$ + yvbCrLf() _
                    + "     エラー件数    :" + yStr(j) + yvbCrLf() _
                    + "     ログ内容"
                For i = 1 To j
                    SendDt$ = SendDt$ + yvbCrLf() + b$(i)
                Next i
                RtnMsg$ = EtherCom.ComHost(MailServerIP$, MailServerSIP$, MailServerPort, SendDt$)
                MailSendFlg = True
            End If
        Else
            If MailSendFlg = True Then
                SendDt$ = "Mail-" + MyPcName + "," + MailFrom + "," + MailSend _
                    + "," + MyPcName + " EtherRetry Error Recovery,データ転送エラーが復旧しました。"
                RtnMsg$ = EtherCom.ComHost(MailServerIP$, MailServerSIP$, MailServerPort, SendDt$)
            End If
            EtherRetry = 0
            yKill(EtherErrorFile$)
            MailSendFlg = False
        End If
        Me.Hide()
        yDoEvents()
    End Function
    Public Sub DataSendError(ByVal Data$)
        '転送失敗データの保存
        On Error Resume Next
        Dim a$(0)
        a$(0) = "@" + DateTime.Now.ToString("yy/MM/dd HH:mm:ss") + "," + Data$
        FileLAppend(EtherErrorFile$, a$)
    End Sub
    Private Function FileLRead(ByVal FileName$, ByRef Data$(), ByRef Num%) As Boolean
        Try
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
            Dim Lines As String() = System.IO.File.ReadAllLines(FileName, enc)
            Data = Lines
            Num = Lines.Length
            FileLRead = True
        Catch
            FileLRead = False
        End Try
    End Function
    Sub FileLWrite(ByVal FileName$, ByVal Data$())
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
        System.IO.File.WriteAllLines(FileName, Data$, enc)
    End Sub
    Private Function FileLAppend(ByVal FileName$, ByVal Data$()) As Boolean
        Try
            If FileName$ <> "" Then
                Call MakeDir(FileName$)
                Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
                System.IO.File.AppendAllLines(FileName, Data$, enc)
            End If
            FileLAppend = True
        Catch
            FileLAppend = False
        End Try
    End Function
    Private Sub LogSave(ByVal Msg$)
        Dim FileName$
        MakeDir(LogPath$)
        FileName$ = LogPath$ + DateTime.Now.ToString("yyMMdd") + ".log"
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
        System.IO.File.AppendAllText(FileName$, DateTime.Now.ToString("HH:mm:ss") & " " & Msg$ & Environment.NewLine, enc)
    End Sub
    Sub MakeDir(ByVal FileName$)
        If IO.Path.GetFileName(FileName$) <> "" Then      'ファイル名は取り除く
            FileName$ = FileName$.Substring(0, FileName$.IndexOf(IO.Path.GetFileName(FileName$)))
        End If
        IO.Directory.CreateDirectory(FileName$)
    End Sub
    Private Sub Wait(ByVal Wt As Single)
        On Error Resume Next
        '   Wt秒 waitting!!
        Dim WtTimes%
        Dim SleepTm%
        Dim i%
        SleepTm = 50
        WtTimes = CInt(Wt * 1000 / SleepTm * 0.9)
        For i = 1 To WtTimes
            yDoEvents()
            ySleep(SleepTm)
        Next i
    End Sub
    Private Sub Form_Load()
        On Error Resume Next
        Me.Text = "Ethercom File Access Utility  : " + EtherUtyVer + " " + EtherUtyDate
        LogPath$ = AppPath() + "\Log\"
        EtherErrorFile$ = AppPath() + "\Data\EtherErrorData.txt"
        ErrMsgDispFlg = True
        Label1.Text = ""
        Label2.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        'デフォルト値
        EtherCom.WaitTimeOut = 5
        MailCount = 10            'エラーメッセージ Mail 送信の為の異常検出回数
        EtherRetryMax = 2         'Ｅｔｈｅｒｃｏｍ再処理回数
        FileRetryMax = 10         'FileAccess再処理回数
    End Sub
    Public Function QDBWrite(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal Table$, ByVal WrField$(), ByVal WrData$(), ByRef ErrMsg$, Optional ServerSIP$ = "") As Boolean
        'ＤＢデータ新規保存
        On Error Resume Next
        'Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない
        ErrMsg$ = ""
        QDBWrite = False
        Label1.Text = "QDBWrite"
        '
        Dim Field$, Data$
        Dim ErrCode$    'C0,C1,C2,C3,C4
        Dim i As Integer
        Dim SendDt$
        Dim ErrCount As Integer
        Dim RtnMsg$
        Dim s$
        Dim WrDataSet$()            'V1.07 add
        Dim WrFieldSet$()           'V1.07 add
        '
        If yTrim(ServerIP$) = "" Then ErrMsg$ = "サーバーの指定が無い。" : Exit Function
        If yLen(Table$) = 0 Then ErrMsg$ = "テーブル指定が無い。" : Exit Function
        If yUBound(WrField$) = 0 Then ErrMsg$ = "フィールド指定が無い。" : Exit Function
        If yUBound(WrData$) = 0 Then ErrMsg$ = "データ指定が無い。" : Exit Function
        If yUBound(WrField$) <> yUBound(WrData$) Then ErrMsg$ = "フィールド指定とデータ指定の数不一致" : Exit Function 'フィ－ルドの数とデ－タの数が違った場合

        '送信データの作成

        ReDim WrDataSet$(yUBound(WrField$))                                                          'V1.07 add
        ReDim WrFieldSet$(yUBound(WrField$))                                                         'V1.07 add

        '   For i = 0 To UBound(WrData$)
        '        WrData$(i) = Replace(WrData$(i), ",", "\c")                                            'V1.07 delete
        '        WrData$(i) = Replace(WrData$(i), "'", "\s")                                            'V1.07 delete
        '        WrData$(i) = Replace(WrData$(i), """", "\d")                                           'V1.07 delete

        For i = 0 To yUBound(WrField$)                                                               'V1.07 add
            WrDataSet$(i) = WrData$(i)                                                              'V1.07 add
            WrDataSet$(i) = yReplace(WrDataSet$(i), ",", "\c")                                       'V1.07 add
            WrDataSet$(i) = yReplace(WrDataSet$(i), "'", "\s")                                       'V1.07 add
            WrDataSet$(i) = yReplace(WrDataSet$(i), """", "\d")                                      'V1.07 add

            WrFieldSet$(i) = WrField$(i)                                                            'V1.07 add
        Next i
        '
        '    Field$ = WrField$(0): Data$ = "'" + WrData$(0) + "'"                                       'V1.07 delete
        '    '
        '    For i = 1 To UBound(WrField$)                                                              'V1.07 delete
        '        Field$ = Field$ + "," + WrField$(i): Data$ = Data$ + "," + "'" + WrData$(i) + "'"      'V1.07 delete
        '    Next i                                                                                     'V1.07 delete

        Field$ = WrFieldSet$(0) : Data$ = "'" + WrDataSet$(0) + "'"                                  'V1.07 add
        '
        For i = 1 To yUBound(WrField$)                                                               'V1.07 add
            Field$ = Field$ + "," + WrFieldSet$(i) : Data$ = Data$ + "," + "'" + WrDataSet$(i) + "'" 'V1.07 add
        Next i                                                                                      'V1.07 add

        SendDt$ = "QDBW-" + Table$ + "|" + Field$ + "|" + Data$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = Ethercom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If yInStr(RtnMsg$, "ERROR") > 0 And ErrCount < EtherRetryMax Then
            If ErrCount >= EtherRetryMax - 1 And ErrMsgDispFlg = True Then
                Label2.Text = "ＤＢデータ　転送異常" + yvbCrLf() _
                    + "データ保存リトライします！！"
                Label3.Text = ServerIP$ + "," + yStr(ServerPort) + "," + yLeft(SendDt$, 30)
                Me.Show()
            End If
            Wait(1)
            Debug.Print(yStr(ErrCount) + " " + SendDt$)
            ErrCount += 1
            If ErrCount < EtherRetryMax Then GoTo Exec
        End If
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            QDBWrite = True
        Else
            '異常終了
            ErrCode$ = yUCase(yMid(RtnMsg$, yInStr(RtnMsg$, "ERROR:") + 6, 2))
            ErrMsg$ = Ethercom_Error_text(ErrCode$)                                 'V1.04 追加
            'ErrMsg$ = ""                                                            'V1.04 削除
            'If InStr(ErrCode$, "C") <> 0 Then                                       'V1.04 削除
            '    c$() = Split(Cmsg$, "|")                                            'V1.04 削除
            '    ErrMsg$ = c$(Val(Right(ErrCode$, 1)))       'エラーの内容           'V1.04 削除
            'End If                                                                  'V1.04 削除
            '
            s$ = ServerIP$ + "," + ServerSIP$ + "," + yStr(ServerPort) + "," + SendDt$  'リトライファイルに保存するデータ
            '
            DataSendError(s$)
            LogSave("DBData Retry File  :" + RtnMsg$ + ":" + s$)
            QDBWrite = True     'リトライファイルに保存したためTRUEを返す。
        End If
        Me.Hide()
    End Function
    Public Function QDBRead(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal Table$, ByVal RdKey$, ByVal RdKeyData$, ByRef RdField$(), ByRef RdData$(), ByRef ErrMsg$, Optional ServerIP2$ = "") As Boolean
        'ＤＢデータ読込
        On Error Resume Next
        '
        Dim ServerSIP$
        Dim Key$
        Dim Field$
        Dim i As Integer
        Dim SendDt$
        Dim ErrCount As Integer
        Dim RtnMsg$
        Dim RtnField$()
        Dim ErrCode$
        '
        ServerSIP$ = ServerIP2$
        QDBRead = False
        Label1.Text = "QDBRead"
        If yLen(yTrim(Table$)) = 0 Then ErrMsg$ = "テーブル指定無し。" : Exit Function
        If yLen(yTrim(RdKey$)) = 0 Then ErrMsg$ = "キーフィールド指定無し。" : Exit Function
        If yLen(yTrim(RdKeyData$)) = 0 Then ErrMsg$ = "キー指定無し。" : Exit Function
        If RdField$(0) = "" Then ErrMsg$ = "フィールド指定無し。" : Exit Function
        '
        Key$ = RdKey$ + "='" + RdKeyData$ + "'"     ' key='*****' 形式に変換
        '
        Field$ = RdField$(0)
        For i = 1 To yUBound(RdField$)
            Field$ = Field$ + "," + RdField$(i)
        Next i
        '
        '送信データの作成
        SendDt$ = "QDBR-" + Table$ + "|" + Key$ + "|" + Field$
        ErrCount = 0
Exec:
        'データベースパソコンへ転送します。
        RtnMsg$ = Ethercom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)
        If yInStr(RtnMsg$, "ERROR") > 0 Then
            'If InStr(RtnMsg$, "ERROR:31") = 0 Then  'Data not found
            If yInStr("ERROR:31 ERROR:B0 ERROR:C0", yLeft(RtnMsg$, 8)) = 0 Then   'Data not found
                Wait(1)
                Debug.Print(yStr(ErrCount) + " " + SendDt$)
                ErrCount += 1
                If ErrCount < EtherRetryMax Then GoTo Exec
            End If
        End If
        '
        'ReDim RdData$(yUBound(RdField$))
        'For i = 0 To yUBound(RdField$)
        ' RdData$(i) = ""
        'Next i
        If yInStr(RtnMsg$, "ERROR") = 0 Then
            '正常終了
            RtnField$ = ySplit(yLeft(RtnMsg$, yInStr(RtnMsg$, "|") - 1), ",")
            RdData$ = ySplit(yMid(RtnMsg$, yInStr(RtnMsg$, "|") + 1, yLen(RtnMsg$) - yInStr(RtnMsg$, "|") + 1), ",")

            For i = 0 To yUBound(RdData$)
                RdData$(i) = yReplace(RdData$(i), "\c", ",")
                RdData$(i) = yReplace(RdData$(i), "\s", "'")
                RdData$(i) = yReplace(RdData$(i), "\d", """")
            Next i
            QDBRead = True
        Else
            '異常終了
            ErrCode$ = yUCase(yMid(RtnMsg$, yInStr(RtnMsg$, "ERROR:") + 6, 2))
            ErrMsg$ = Ethercom_Error_text(ErrCode$)                                 'V1.04  追加
            'c$() = Split(Cmsg$, "|")                                               'V1.04　削除
            'ErrMsg$ = c$(Val(Right(ErrCode$, 1)))       'エラーの内容              'V1.04　削除
            '
            If yInStr(RtnMsg$, "ERROR:C0") = 0 Then       'C0エラーは除く
                LogSave("QDBData Read Error :" + RtnMsg$ + ":" + SendDt$)
                'QDBRead = True
            End If
        End If
        Me.Hide()
    End Function
    '                       /// データベース保存処理 ///　保存に失敗した場合はエラーとして保存しようとした内容は廃棄。
    Public Function QDBWrite_NotSaveText(ByVal ServerIP$, ByVal ServerPort As Integer, ByVal Table$, ByVal WrField$(), ByVal WrData$(), ByRef ErrMsg$, Optional ServerSIP$ = "") As Boolean
        'ＤＢデータ新規保存
        On Error Resume Next
        'Const ServerSIP$ = ""   'セカンダリーのデータベースへは書き込まない
        ErrMsg$ = ""
        QDBWrite_NotSaveText = False
        Label1.Text = "QDBWrite_NotSaveText"
        '
        Dim Field$, Data$               'データベースに書き込むフィールドとデータを配列変数⇒変数に変換
        Dim ErrCode$                    'データベースアクセスのエラーコード
        Dim i As Integer                'データベースに書き込み不可な文字を変換するときのポインタ
        Dim SendDt$                     'データベースに送るデータ
        Dim ErrCount As Integer         'データベースアクセスエラー時のリトライカウンタ
        Dim RtnMsg$                     'データベースアクセス後の返り値
        '  Dim c$()
        '  Dim S$
        Dim WrDataSet$()            'V1.07 add
        Dim WrFieldSet$()           'V1.07 add

        If yTrim(ServerIP$) = "" Then ErrMsg$ = "サーバーの指定が無い。" : Exit Function
        If yLen(Table$) = 0 Then ErrMsg$ = "テーブル指定が無い。" : Exit Function
        If yUBound(WrField$) = 0 Then ErrMsg$ = "フィールド指定が無い。" : Exit Function
        If yUBound(WrData$) = 0 Then ErrMsg$ = "データ指定が無い。" : Exit Function
        If yUBound(WrField$) <> yUBound(WrData$) Then ErrMsg$ = "フィールド指定とデータ指定の数不一致" : Exit Function 'フィ－ルドの数とデ－タの数が違った場合

        '送信データの作成

        ReDim WrDataSet$(yUBound(WrField$))                                                          'V1.07 add
        ReDim WrFieldSet$(yUBound(WrField$))                                                         'V1.07 add

        '    For i = 0 To UBound(WrData$)
        '        WrData$(i) = Replace(WrData$(i), ",", "\c")                                            'V1.07 delete 'データベース使用禁止文字を変換
        '        WrData$(i) = Replace(WrData$(i), "'", "\s")                                            'V1.07 delete 'データベース使用禁止文字を変換
        '        WrData$(i) = Replace(WrData$(i), """", "\d")                                           'V1.07 delete 'データベース使用禁止文字を変換

        For i = 0 To yUBound(WrField$)                                                               'V1.07 add
            WrDataSet$(i) = WrData$(i)                                                              'V1.07 add
            WrDataSet$(i) = yReplace(WrDataSet$(i), ",", "\c")                                       'V1.07 add
            WrDataSet$(i) = yReplace(WrDataSet$(i), "'", "\s")                                       'V1.07 add
            WrDataSet$(i) = yReplace(WrDataSet$(i), """", "\d")                                      'V1.07 add

            WrFieldSet$(i) = WrField$(i)                                                            'V1.07 add
        Next i
        '
        '    '
        '    Field$ = WrField$(0): Data$ = "'" & WrData$(0) & "'"                                       'V1.07 delete '一番目のデータをセット
        '
        '    For i = 1 To UBound(WrField$)                                                              'V1.07 delete
        '        Field$ = Field$ & "," & WrField$(i): Data$ = Data$ & "," & "'" & WrData$(i) & "'"      'V1.07 delete '残りのデータをループしてセット
        '    Next i                                                                                     'V1.07 delete

        Field$ = WrFieldSet$(0) : Data$ = "'" + WrDataSet$(0) + "'"                                  'V1.07 add
        '
        For i = 1 To yUBound(WrField$)                                                               'V1.07 add
            Field$ = Field$ + "," + WrFieldSet$(i) : Data$ = Data$ + "," + "'" + WrDataSet$(i) + "'" 'V1.07 add
        Next i                                                                                      'V1.07 add

        SendDt$ = "QDBW-" & Table$ & "|" & Field$ & "|" & Data$                                     'データベースに保存する形にセット
        ErrCount = 0

        Do
            'データベースパソコンへ転送します。
            RtnMsg$ = Ethercom.ComHost(ServerIP$, ServerSIP$, ServerPort, SendDt$)                 'データ保存へ　RtnMsg$に"ERROR"を含んでいたら失敗

            If yInStr(RtnMsg$, "ERROR") = 0 Then                                                     'データ保存成功した場合はループから抜ける
                QDBWrite_NotSaveText = True
                Exit Do
            End If

            ErrCount += 1                                                                 'リトライカウント　カウントアップ
            If ErrCount <= EtherRetryMax Then                                                       'リトライ処理を行う。
                If ErrMsgDispFlg = True Then
                    Label2.Text = "ＤＢデータ　転送異常" + yvbCrLf() + "データ保存リトライします！！"
                    Label3.Text = ServerIP$ + "," + yStr(ServerPort) + "," + yLeft(SendDt$, 30)
                    Me.Show()
                End If
            Else
                '異常終了
                ErrCode$ = yUCase(yMid(RtnMsg$, yInStr(RtnMsg$, "ERROR:") + 6, 2))                     'エラーコード取り出し
                ErrMsg$ = Ethercom_Error_text(ErrCode$)                                             'エラ－コードからエラーメッセージを取得
            End If
            '
            Wait(1)
        Loop While ErrCount <= EtherRetryMax

        Me.Hide()
    End Function
    '           /// データベースアクセルエラーを文字列に変換  ///
    Private Function Ethercom_Error_text(ByVal Ethercom_ErrorNo As String) As String

        Dim Ethercom_err() As String
        Dim ErrMax As Integer
        Dim i As Integer

        Ethercom_Error_text = "エラー内容不明"

        ReDim Ethercom_err(100)

        Ethercom_err(0) = "00,初期化出来ない"
        Ethercom_err(1) = "01,ソケットを作成出来ない"
        Ethercom_err(2) = "02,サーバーに接続出来ない"
        Ethercom_err(3) = "03,サーバーから切断された"
        Ethercom_err(4) = "04,通信エラー"
        Ethercom_err(5) = "05,チェックサムエラー"
        Ethercom_err(6) = "06,指定機能が存在しない"

        Ethercom_err(7) = "11,サーバー接続後にデータを取得出来ない"
        Ethercom_err(8) = "12,サーバービジー"
        Ethercom_err(9) = "13,ホストＰＣダウン"
        Ethercom_err(10) = "14,サーバーソフトダウン"
        Ethercom_err(11) = "15,サーバー接続後、返答前に接続（連続送信）"

        Ethercom_err(12) = "20,ＦＴＰファイル転送失敗"
        Ethercom_err(13) = "21,iniFile データ保存失敗"
        Ethercom_err(14) = "22,iniFile データ取得失敗"
        Ethercom_err(15) = "23,Mail転送失敗"
        Ethercom_err(16) = "24,CapsuleNo取得失敗"

        Ethercom_err(17) = "30,DB データ保存失敗"
        Ethercom_err(18) = "31,DB 指定されたデータ（ｋｅｙ）が存在しないエラー"
        Ethercom_err(19) = "32,DB 指定されたフィールド名が存在しないエラー"
        Ethercom_err(20) = "33,DB データ取得失敗"
        Ethercom_err(21) = "34,DB その他のエラー"

        Ethercom_err(22) = "40,Text データ保存失敗"
        Ethercom_err(23) = "41,Text データ取得失敗"
        Ethercom_err(24) = "42,Text データサイズオーバー（８ｋＢｙｔｅ以内）"
        Ethercom_err(25) = "43,Text ファイル拡張子エラー（ＴＸＴ，ＣＳＶ，ＩＮＩのみ）"

        Ethercom_err(26) = "50,Yewmacとの通信エラー"

        Ethercom_err(27) = "A0,荷札がない（Axx:新ＰＬＡＳＭＡＰＣサーバ）"
        Ethercom_err(28) = "A1,変更前の荷札"
        Ethercom_err(29) = "A2,プログラムエラー（生産情報取得）"
        Ethercom_err(30) = "A3,プログラムエラー（データ数取得）"
        Ethercom_err(31) = "A4,プログラムエラー（ダミーデータ書き込み）"
        Ethercom_err(32) = "A5,ORACLEに関連したエラー（通信エラー）"

        Ethercom_err(33) = "B0,レコードが無い（Picking)"
        Ethercom_err(34) = "B1,ＭＳ＿ＣＯＤＥが無い（Picking)"
        Ethercom_err(35) = "B2,ＭＯＤＥＬが無い（Picking)"
        Ethercom_err(36) = "B3,PARTS_NOが無い（Picking)"
        Ethercom_err(37) = "B4,PARTS数量が無い（Picking)"
        Ethercom_err(38) = "B5,PARENT_IDが無い(Picking)"

        Ethercom_err(39) = "C0,QDB レコードが無い（ラインデータ）"
        Ethercom_err(40) = "C1,QDB プログラムエラー（ラインデータ）"
        Ethercom_err(41) = "C2,QDB ORACLEに関連したエラー（ラインデータ）"
        Ethercom_err(42) = "C3,QDB 送信フレーム異常"
        Ethercom_err(43) = "C4,QDB データ保存失敗"
        Ethercom_err(44) = "C5,QDB パスワードエラー"
        Ethercom_err(45) = "C6,QDB サーバー日時異常"

        Ethercom_err(46) = "D0,RS-232C ポートオープンエラー(Atomos,Temp,Humd)"
        Ethercom_err(47) = "D1,RS-232C 通信エラー(Atomos,Temp,Humd)"
        Ethercom_err(48) = "D2,MU パラメータエラー(Atomos)"

        ErrMax = 48

        For i = 0 To ErrMax
            If yLeft(Ethercom_err(i), 2) = Ethercom_ErrorNo Then Ethercom_Error_text = Ethercom_err(i) : Exit For
        Next i

    End Function
    Private Sub EtherUty_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = "Ethercom File Access Utility  : " + EtherUtyVer + " " + EtherUtyDate
        LogPath$ = AppPath() & "\Log\"
        EtherErrorFile$ = AppPath() & "\Data\EtherErrorData.txt"
        ErrMsgDispFlg = True
        Label1.Text = ""
        Label2.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        'デフォルト値
        Ethercom.WaitTimeOut = 5
        MailCount = 10            'エラーメッセージ Mail 送信の為の異常検出回数
        EtherRetryMax = 2         'Ｅｔｈｅｒｃｏｍ再処理回数
        FileRetryMax = 10         'FileAccess再処理回数
    End Sub
End Class