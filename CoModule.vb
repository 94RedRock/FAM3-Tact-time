Imports System.IO
Imports System.Text
Imports System.Net.NetworkInformation
Module CoModule
    '------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------
    '共通モジュール、関数群、等々
    '   Ver     Data        By
    '   0.00    13.02.08    T.Yoshihara     新規作成
    '   0.01      .02.22    T.Yoshihara     旧CoModule包括
    '   0.02      .02.25    T.Yoshihara     yLng追加
    '   0.03      .03.06    T.Yoshihara     ySplit変更
    '   0.04      .04.23    T.Yoshihara     yVBCr,yVBLf
    '   0.05      .06.28    T.Yoshihara     DataSelバグ修正
    '   0.06      .08.27    T.Yoshihara     IpAddress最新方式採用
    '   0.07      .09.03    T.Yoshihara     yChooseの引数Object対応
    '   0.08      .09.13    T.Yoshihara     GetCurFileのTrim判断廃止
    '   0.09      .09.26    T.Yoshihara     ySplitバグ修正(１文字以上のセパレータ対応)
    '                                       SaveText変更(上書き保存対応)
    '   0.10      .10.02    T.Yoshihara     yMid取得文字数デフォルト値変更999→8192,yKill2追加
    '   0.11      .10.04    T.Yoshihara     ySplit改修(セパレータをnullStringに対応)
    '   0.12      .10.10    T.Yoshihara     yTrimにNothing対応、yValにヌル許可
    '   0.13      .11.12    T.Yoshihara     MsToItemにCPA,CPX追加(EJXと同等の分解)
    '   0.14    14.06.10    T.Yoshihara     yWeekday,yvbSunday他追加
    '   0.15      .06.27    T.Yoshihara     ネットワークドライブ割り付け時の引数変更
    '   0.16    15.03.06    T.Yoshihara     MsToItemにBodyKind追加,GetPatNo等追加
    '   0.17      .03.07    T.Yoshihara     MsToItemからBodyKind削除
    '   0.18      .03.09    T.Yoshihara     yStr,yFormat(オーバーロード)追加
    '   0.19      .04.03    T.Yoshihara     IsNumeric追加
    '   0.20      .04.07    T.Yoshihara     yDir機能向上(引数によりフォルダ名も取得可能)
    '   0.21      .04.20    T.Yoshihara     yLenB追加
    '   0.22      .04.13    K.Sadakata      VB6共通モジュールV0.29取り込み　組み合わせ製品対応
    '   0.23      .08.27    K.Sadakata      VB6共通モジュールV0.30取り込み　組み合わせ型名変更対応
    '   0.24      .11.11    T.Yoshihara     Sort降順のバグ修正
    '   0.25      .11.16    T.Yoshihara     GetCurFile,GetCurFileNullの取り込み数量増
    '   0.26      .11.20    K.Sadakata      C10FRコード変更対応
    '   0.27    16.06.01    T.Yoshihara     fncNetDriveConnectArr内のコメント削除
    '   0.28      .09.06    K.Sadakata      組合せPhase2対応
    '   0.29    17.01.25    T.Yoshihara     yLeft,yMid,yRightのNothing時に非Error化
    '   0.30      .03.15    T.Yoshihara     fncNetDriveConnectArr内に"i"の宣言漏れ対応
    '	0.31	  .04.26    T.Yoshihara		yVal各型対応(yValD,yValS,yValL,yValI)
    '	0.32	  .05.11    K.Maekawa		C10SA対応,二桁Z仕様変更対応
    '   0.33      .08.15    T.Yoshihara     yDateDiff("ww")のバグ修正
    '   0.34      .12.22    T.Yoshihara     yKill修正(エラー回避)
    '   0.35    18.03.13    S.Shiya         yKillの誤記訂正
    '   0.36      .05.10    T.Yoshihara     yLeft,yRightの訂正
    '   0.37    19.11.21    T.Yoshihara     VB2019対応
    '------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------
    'ネットワークドライブへの接続を行うめの関数宣言
    Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer
    'ネットワークドライブの切断を行うための関数宣言(第2引数は、Windows終了時に接続を回復するかどうかを表す.Falseは、回復することを意味する)
    Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer
    'ネットワークリソース情報構造体
    Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As Integer
        Public lpProvider As Integer
    End Structure
    'リソースタイプ定数
    Public Const RESOURCE_CONNECTED As Integer = &H1
    '接続スコープ定数
    Public Const RESOURCETYPE_ANY As Integer = &H0
    '表示タイプ定数
    Public Const RESOURCEDISPLAYTYPE_DOMAIN As Integer = &H1
    '接続オプション定数
    Public Const CONNECT_UPDATE_PROFILE As Integer = &H1
    '****************************************************
    Public g_strDriveLetter() As String '空きドライブ文字
    '****************************************************
    Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Declare Function GetPrivateProfileSectionNames Lib "Kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    '--------------------------------------------------------------
    Structure TypeDecodeModel    ' ＭＳコード分解
        Public Series As String
        Public MsCode As String
        Public Model As String
        Public OutPut As String           '出力信号
        Public CapSpn As String           '測定スパン
        Public Matral As String           '接液部材質
        Public FlSize As String           'フランジ規格
        Public FlMatl As String           'フランジ材質
        Public BtMatl As String           '締め付けボルト材質
        Public Oilkid As String           '封入液
        Public Instal As String           '取り付け
        Public Econct As String           '電源接続口
        Public LCDkid As String           'LCD
        Public Braket As String           '取り付けブラケット
        Public Pconct As String           'プロセス接続口
        Public AmpHig As String           'アンプケース
        Public Caplry As String           'キャピラリー長
        Public WtMatl As String           'ＤＦＳ、レベル計　接液部材質
        Public FlsgRg As String           'フラッシングコネクションリング
        Public Ctsize As String           '取り付けサイズ
        Public X2kind As String           '突き出し型
        Public SelFac As String           'ガスケット座表面処理
        Public PrType As String           'プロセス接続構造 突き出し、フラッシュ形
        Public PrTemp As String           'プロセス温度入力(マルバリのみ)
        Public OpCode As String           '付加仕様(マルバリのみ)
        Public ProductName As String      '機種名（FP、EJX・・・等）
        Public ProductName2 As String     '機種名（FP、EJX、EJA、NEWEJA・・・等）上記機種名にEJAとNEWEJAの区別追加         V0.24 ADD
        Public AgType As String           '測定圧タイプ（FPのみ）
        Public MeType As String           '測定種類(組合せ製品用)                                                          V0.21 ADD
        Public PrCfg1 As String           '製品構成①(組合せ製品用)                                                        V0.21 ADD
        Public PrCfg2 As String           '製品構成②(組合せ製品用)                                                        V0.21 ADD
        Public HPrTyp As String           '高圧側接続形状(組合せ製品用)                                                    V0.21 ADD
        Public HWtMtl As String           '高圧側接液材質(組合せ製品用)                                                    V0.21 ADD
        Public LPrTyp As String           '低圧側接続形状(組合せ製品用)                                                    V0.21 ADD
        Public LWtMtl As String           '低圧側接液材質(組合せ製品用)                                                    V0.21 ADD
        Public HAcsry As String           '高圧側付属品(組合せ製品用)                                                      V0.21 ADD
        Public LAcsry As String           '低圧側付属品(組合せ製品用)                                                      V0.21 ADD
    End Structure

    Structure TypeDecodeModelC   ' ＭＳコード分解(組合せ製品用)                                                            V0.21 ADD
        Public MsCode As String                                                                                           'V0.21 ADD
        Public Model As String                                                                                            'V0.21 ADD
        Public HLSide As String           '製品識別                                                                        V0.21 ADD
        Public OilKind As String          '封入液種類                                                                      V0.21 ADD
        Public CnctMthd As String         '伝送器との接続形状                                                              V0.21 ADD
        Public CplyLngth As String        'キャピラリ長さ                                                                  V0.21 ADD
        Public CplyDmtr As String         'キャピラリ内径                                                                  V0.21 ADD
        Public CplyCvr As String          'キャピラリ被覆種類                                                              V0.21 ADD
        Public CplyPipe As String         'キャピラリPIPE構造                                                              V0.21 ADD
        Public GsktSize As String         'ガスケット座サイズ                                                              V0.21 ADD
        Public GsktStd As String          'ガスケット座面規格                                                              V0.21 ADD
        Public Srrtn As String            'セレーション                                                                    V0.21 ADD
        Public DphrgmMtrl As String       'ダイアフラム材質                                                                V0.21 ADD
        Public WtMtrl As String           'その他接液部材質                                                                V0.21 ADD
        Public DphrgmDmtr As String       'ダイアフラム径                                                                  V0.21 ADD
        Public DphrgmTrtmnt As String     'ダイアフラム処理                                                                V0.21 ADD
        Public WtTrtmnt As String         '接液部処理                                                                      V0.21 ADD
        Public CplyPstn As String         'キャピラリ取出構造                                                              V0.21 ADD
        Public FlngSize As String         'フランジサイズ                                                                  V0.21 ADD
        Public FlngStd As String          'フランジ規格                                                                    V0.21 ADD
        Public FlngRate As String         'フランジ定格                                                                    V0.21 ADD
        Public FlngBrngStd As String      'フランジ座面規格                                                                V0.21 ADD
        Public FlngMtrl As String         'フランジ材質                                                                    V0.21 ADD
        Public X2Dmtr As String           '突出し部外径                                                                    V0.21 ADD
        Public X2Lngth As String          '突出し部長さ                                                                    V0.21 ADD
        Public JckupBlt As String         'ジャッキアップボルト                                                            V0.21 ADD
        Public FlngKind As String         'フランジ構造                                                                    V0.21 ADD
        Public FCRKind As String          'フラッシュリング種類                                                            V0.21 ADD
        Public FCRSize As String          'フラッシュリングサイズ                                                          V0.21 ADD
        Public FCRMtrl As String          'フラッシュリング材質                                                            V0.21 ADD
        Public VntPlgQty As String        'ベントプラグ個数                                                                V0.21 ADD
        Public VntScrwStd As String       'ベントプラグねじ規格                                                            V0.21 ADD
        Public VntPlgKind As String       'ベントプラグ種類                                                                V0.21 ADD
        '        Public FCRTrtmnt As String        '-                                                                               V0.26 Change         V0.21 ADD   V0.32 DELETE
        Public GsktKind As String         'ガスケット種類                                                                  V0.21 ADD
        Public HpflrMtrl As String        'ガスケット材質　　　                                                            V0.21 ADD
        Public ProductName As String                                                                                      'V0.21 ADD
    End Structure                                                                                                         'V0.21 ADD

    Public DecodeModel As TypeDecodeModel
    Public DecodeModelC As TypeDecodeModelC                                                                               'V0.21 ADD
    Sub MsToItem(ByVal MsCode$, ByVal ItemIni$, ByRef DModel As TypeDecodeModel)
        '
        Dim Item() As Integer
        Dim Ipnt() As Integer
        Dim Pt As Integer
        Dim Turn$
        Dim Dm$
        Dim i As Integer
        'Dim DmIpnt As Integer
        Dim intMscodeLen As Integer 'MSコードの基本仕様部バイト数                                                                             V0.32 ADD
        Dim intDmLen As Integer     'Dmのバイト数                                                                                             V0.32 ADD
        '
        Erase Item
        Erase Ipnt
        '
        Turn$ = "PLFNSQJDRBHKAGMETUCWXYIabcdefghi"                                                                                          'V0.21 ADD

        ReDim Item(Turn$.Length)
        ReDim Ipnt(Turn$.Length)
        '
        DModel.Series = YMid(MsCode, 7, 1)       'EJX専用
        DModel.AmpHig = ""          'AMP housing ---------- P ------- アンプケース
        DModel.Braket = ""          'BRAKET---------------- L ------- 取り付けブラケット
        DModel.BtMatl = ""          'Bolt Material--------- F ------- 締め付けボルト材質
        DModel.Caplry = ""          'CAPILARY ------------- N ------- キャピラリー長
        DModel.CapSpn = ""          'SPAN------------------ S ------- 測定スパン
        DModel.Ctsize = ""          'Connection Size ------ Q ------- 取り付けサイズ
        DModel.Econct = ""          'Erectric Connect ----- J ------- 電源接続口
        DModel.FlMatl = ""          'Flange Material------- D ------- フランジ材質
        DModel.FlsgRg = ""          'Flushing C Ring ------ R ------- フラッシングコネクションリング
        DModel.FlSize = ""          'Flange Size----------- B ------- フランジ規格
        DModel.Instal = ""          'Inst------------------ H ------- 取り付け
        DModel.LCDkid = ""          'LCD ------------------ K ------- ＬＣＤ
        DModel.Matral = ""          'Wet Material---------- A ------- 接液部材質
        DModel.MsCode = MsCode
        DModel.Oilkid = ""          'Oil------------------- G ------- 封入液
        DModel.OutPut = ""          'OUTPUT---------------- M ------- 出力信号
        DModel.Pconct = ""          'P-Connection---------- E ------- プロセス接続口
        DModel.SelFac = ""          'Sealing Face --------- T ------- ガスケット座表面処理
        DModel.WtMatl = ""          'Wet Parts Material --- U ------- 接液部材質　ＤＦＳ、レベルのみ
        DModel.X2kind = ""          'X2-------------------- C ------- ダイアフラム突き出し
        DModel.PrType = ""          'Prosess Type---------- W ------- プロセス接続構造 突き出し、フラッシュ形
        DModel.PrTemp = ""          'Process Temperature -- X ------- プロセス温度入力(マルバリのみ)
        DModel.OpCode = ""          'Optional Codes ------- Y ------- 付加仕様コード(マルバリのみ)
        DModel.AgType = ""
        DModel.MeType = ""          'Measure Type---------- a ------- 測定種類(組合せ製品用)                                                 V0.21 ADD
        DModel.PrCfg1 = ""          'Product Configuration- b ------- 製品構成①(組合せ製品用)                                               V0.21 ADD
        DModel.PrCfg2 = ""          'Product Configuration- c ------- 製品構成②(組合せ製品用)                                               V0.21 ADD
        DModel.HPrTyp = ""          'H Side Process Type--- d ------- 高圧側接続形状(組合せ製品用)                                           V0.21 ADD
        DModel.HWtMtl = ""          'H Side Wet Material--- e ------- 高圧側接液材質(組合せ製品用)                                           V0.21 ADD
        DModel.LPrTyp = ""          'L Side Process Type--- f ------- 低圧側接続形状(組合せ製品用)                                           V0.21 ADD
        DModel.LWtMtl = ""          'L Side Wet Material--- g ------- 低圧側接液材質(組合せ製品用)                                           V0.21 ADD
        DModel.HAcsry = ""          'H Side Flushing C Ring h ------- 高圧側付属品(組合せ製品用)                                             V0.21 ADD
        DModel.LAcsry = ""          'L Side Flushing C Ring i ------- 低圧側付属品(組合せ製品用)                                             V0.21 ADD
        DModel.Model = ""                                               'v0.17 初期化追加
        DModel.ProductName = ""                                         'v0.17 初期化追加
        DModel.ProductName2 = ""                                            'V0.24 ADD
        '
        If YLeft(MsCode, 3) = "EJX" Or YLeft(MsCode, 7) Like "EJA???[EJ]" Or YLeft(MsCode, 3) Like "CP[AX]" Then   '【EJX,NewEJA,CPA,CPX】 'V0.13 ADD
            DModel.Model = YMid(MsCode, 4, 3)
            DModel.ProductName = YLeft(MsCode, 3)
            DModel.ProductName2 = DModel.ProductName                                                        'V0.24 ADD
            If DModel.ProductName2 = "EJA" Then                                                             'V0.24 ADD
                DModel.ProductName2 = "NEWEJA"                                                              'V0.24 ADD
            End If                                                                                          'V0.24 ADD
            Dm$ = GetCurFile("ItemPosition", DModel.Model, "", ItemIni$) : If YTrim(Dm$) = "" Then GoTo MsToItemEnd 'V0.21 ADD
            Dm$ = YSpace(YInStr(MsCode, "-") - 1) & Dm$       'EJX***■(7文字)をスペースで埋める。                                            'V0.21 ADD
        ElseIf YLeft(MsCode, 2) = "EJ" And YMid(MsCode, 3, 1) Like "[AB]" Then '【EJA,EJB】
            DModel.Model = YMid(MsCode, 4, YInStr(MsCode, "-") - 4)
            If YRight(DModel.Model, 1) = "A" Then DModel.Model = YLeft(DModel.Model, YLen(DModel.Model) - 1) 'Omitto A series
            DModel.ProductName = YLeft(MsCode, 3)
            DModel.ProductName2 = DModel.ProductName                                                        'V0.24 ADD
            Dm$ = GetCurFile("EJAItemPosition", DModel.Model, "", ItemIni$) : If YTrim(Dm$) = "" Then GoTo MsToItemEnd 'V0.21 ADD
            Dm$ = YSpace(YInStr(MsCode, "-") - 1) & Dm$       'EJ□***をスペースで埋める。                                                    'V0.21 ADD
        ElseIf YLeft(MsCode, 2) = "EJ" And YMid(MsCode, 3, 1) Like "[0-9]" Then   '【EJ】
            DModel.Model = YMid(MsCode, 3, YInStr(MsCode, "-") - 3)
            DModel.ProductName = YLeft(MsCode, 2)
            DModel.ProductName2 = DModel.ProductName                                                        'V0.24 ADD
            Dm$ = GetCurFile("EJAItemPosition", DModel.Model, "", ItemIni$) : If YTrim(Dm$) = "" Then GoTo MsToItemEnd 'V0.21 ADD
            Dm$ = YSpace(YInStr(MsCode, "-") - 1) & Dm$       'EJ***をスペースで埋める。
        ElseIf YLeft(MsCode, 2) = "FP" Or YLeft(MsCode, 2) = "JP" Or YLeft(MsCode, 2) = "VS" Then
            DModel.Model = YMid(MsCode, 3, 3)
            DModel.ProductName = YLeft(MsCode, 2)
            DModel.ProductName2 = DModel.ProductName                                                        'V0.24 ADD
            If YLeft(MsCode, 5) Like "FP??1" = True Then    'MSCODE旧体系
                Dm$ = GetCurFile("ItemPosition", "FP1", "", ItemIni$) : If YTrim(Dm$) = "" Then GoTo MsToItemEnd 'V0.21 ADD
            Else
                Dm$ = GetCurFile("ItemPosition", "FP2", "", ItemIni$) : If YTrim(Dm$) = "" Then GoTo MsToItemEnd 'V0.21 ADD
            End If
            Dm$ = YSpace(YInStr(MsCode, "-") - 1) & Dm$                                                                                       'V0.21 ADD

        ElseIf YLeft(MsCode, 3) = "FVX" Then  '【FVX】
            DModel.Model = YMid(MsCode, 4, 3)
            DModel.ProductName = YLeft(MsCode, 3)
            DModel.ProductName2 = DModel.ProductName                                                        'V0.24 ADD
            Dm$ = GetCurFile("FVXItemPosition", DModel.Model, "", ItemIni$) : If YTrim(Dm$) = "" Then GoTo MsToItemEnd 'V0.21 ADD
            Dm$ = YSpace(YInStr(MsCode, "-") - 1) & Dm$       'FVX***■(7文字)をスペースで埋める。                                            'V0.21 ADD

        Else
            GoTo MsToItemEnd
        End If

        intMscodeLen = YInStr(MsCode$ & "/", "/") - 1                                                                                           'V0.32 ADD
        intDmLen = YLen(Dm$)                                                                                                                    'V0.32 ADD

        If intMscodeLen < intDmLen Then                                                                                                         'V0.32 ADD
            '   Dm$修正
            For i = 1 To YLen(Dm$)                                                           'V0.15 追加
                If YLen(Dm$) >= i Then                                                       'V0.15 追加　"Z"仕様でMScodeが短くなる場合があるため
                    If YMid(Dm$, i, 1) <> " " And YInStr(Turn$, YMid(Dm$, i, 1)) <> 0 Then   'V0.15 追加　スペース以外で登録文字の場合
                        If YMid(Dm$, i, 1) = YMid(Dm$, i + 1, 1) Then                        'V0.15 追加　同じITEMが２つ続いたら
                            If YMid(MsCode$, i, 1) = "Z" Then                                'V0.15 追加　"Z"の場合は一桁詰める。
                                Dm$ = YLeft(Dm$, i - 1) + YMid(Dm$, i + 1)                   'V0.15 追加
                            End If                                                           'V0.15 追加
                        End If                                                               'V0.15 追加
                    End If                                                                   'V0.15 追加
                End If                                                                       'V0.15 追加
            Next i                                                                           'V0.15 追加
        End If                                                                                                                                  'V0.32 ADD
        '
        For i = 1 To YLen(Turn$)
            Item(i) = 0                                              'ITEMが無い
            Ipnt(i) = YInStr(Dm$, YMid(Turn$, i, 1))
            If Ipnt(i) <> 0 Then
                Item(i) = 1                                          'ITEMが一桁
                If YMid(Dm$, Ipnt(i) + 1, 1) = YMid(Turn$, i, 1) Then  '同じITEMが２つ続いたら
                    Item(i) = 2                                      'ITEMが二桁
                End If
            End If
        Next i
        '           "PLFNSQJDRBHKAGMETUCWXY"
        With DModel
            Pt = YInStr(Turn$, "P") : If Item(Pt) <> 0 Then .AmpHig = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "L") : If Item(Pt) <> 0 Then .Braket = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "F") : If Item(Pt) <> 0 Then .BtMatl = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "N") : If Item(Pt) <> 0 Then .Caplry = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "S") : If Item(Pt) <> 0 Then .CapSpn = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "Q") : If Item(Pt) <> 0 Then .Ctsize = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "J") : If Item(Pt) <> 0 Then .Econct = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "D") : If Item(Pt) <> 0 Then .FlMatl = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "R") : If Item(Pt) <> 0 Then .FlsgRg = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "B") : If Item(Pt) <> 0 Then .FlSize = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "H") : If Item(Pt) <> 0 Then .Instal = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "K") : If Item(Pt) <> 0 Then .LCDkid = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "A") : If Item(Pt) <> 0 Then .Matral = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "G") : If Item(Pt) <> 0 Then .Oilkid = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "M") : If Item(Pt) <> 0 Then .OutPut = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "E") : If Item(Pt) <> 0 Then .Pconct = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "T") : If Item(Pt) <> 0 Then .SelFac = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "U") : If Item(Pt) <> 0 Then .WtMatl = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "C") : If Item(Pt) <> 0 Then .X2kind = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "W") : If Item(Pt) <> 0 Then .PrType = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "X") : If Item(Pt) <> 0 Then .PrTemp = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "Y") : If Item(Pt) <> 0 Then .OpCode = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "I") : If Item(Pt) <> 0 Then .AgType = YMid(MsCode$, Ipnt(Pt), Item(Pt))
            Pt = YInStr(Turn$, "a") : If Item(Pt) <> 0 Then .MeType = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD
            Pt = YInStr(Turn$, "b") : If Item(Pt) <> 0 Then .PrCfg1 = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD
            Pt = YInStr(Turn$, "c") : If Item(Pt) <> 0 Then .PrCfg2 = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD
            Pt = YInStr(Turn$, "d") : If Item(Pt) <> 0 Then .HPrTyp = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD
            Pt = YInStr(Turn$, "e") : If Item(Pt) <> 0 Then .HWtMtl = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD
            Pt = YInStr(Turn$, "f") : If Item(Pt) <> 0 Then .LPrTyp = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD
            Pt = YInStr(Turn$, "g") : If Item(Pt) <> 0 Then .LWtMtl = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD
            Pt = YInStr(Turn$, "h") : If Item(Pt) <> 0 Then .HAcsry = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD
            Pt = YInStr(Turn$, "i") : If Item(Pt) <> 0 Then .LAcsry = YMid(MsCode$, Ipnt(Pt), Item(Pt)) 'V0.21 ADD

            If .AmpHig = "ZZ" Then .AmpHig = "Z" 'V0.32 ADD
            If .Braket = "ZZ" Then .Braket = "Z" 'V0.32 ADD
            If .BtMatl = "ZZ" Then .BtMatl = "Z" 'V0.32 ADD
            If .Caplry = "ZZ" Then .Caplry = "Z" 'V0.32 ADD
            If .CapSpn = "ZZ" Then .CapSpn = "Z" 'V0.32 ADD
            If .Ctsize = "ZZ" Then .Ctsize = "Z" 'V0.32 ADD
            If .Econct = "ZZ" Then .Econct = "Z" 'V0.32 ADD
            If .FlMatl = "ZZ" Then .FlMatl = "Z" 'V0.32 ADD
            If .FlsgRg = "ZZ" Then .FlsgRg = "Z" 'V0.32 ADD
            If .FlSize = "ZZ" Then .FlSize = "Z" 'V0.32 ADD
            If .Instal = "ZZ" Then .Instal = "Z" 'V0.32 ADD
            If .LCDkid = "ZZ" Then .LCDkid = "Z" 'V0.32 ADD
            If .Matral = "ZZ" Then .Matral = "Z" 'V0.32 ADD
            If .Oilkid = "ZZ" Then .Oilkid = "Z" 'V0.32 ADD
            If .OutPut = "ZZ" Then .OutPut = "Z" 'V0.32 ADD
            If .Pconct = "ZZ" Then .Pconct = "Z" 'V0.32 ADD
            If .SelFac = "ZZ" Then .SelFac = "Z" 'V0.32 ADD
            If .WtMatl = "ZZ" Then .WtMatl = "Z" 'V0.32 ADD
            If .X2kind = "ZZ" Then .X2kind = "Z" 'V0.32 ADD
            If .PrType = "ZZ" Then .PrType = "Z" 'V0.32 ADD
            If .PrTemp = "ZZ" Then .PrTemp = "Z" 'V0.32 ADD
            If .OpCode = "ZZ" Then .OpCode = "Z" 'V0.32 ADD
            If .AgType = "ZZ" Then .AgType = "Z" 'V0.32 ADD
            If .MeType = "ZZ" Then .MeType = "Z" 'V0.32 ADD
            If .PrCfg1 = "ZZ" Then .PrCfg1 = "Z" 'V0.32 ADD
            If .PrCfg2 = "ZZ" Then .PrCfg2 = "Z" 'V0.32 ADD
            If .HPrTyp = "ZZ" Then .HPrTyp = "Z" 'V0.32 ADD
            If .HWtMtl = "ZZ" Then .HWtMtl = "Z" 'V0.32 ADD
            If .LPrTyp = "ZZ" Then .LPrTyp = "Z" 'V0.32 ADD
            If .LWtMtl = "ZZ" Then .LWtMtl = "Z" 'V0.32 ADD
            If .HAcsry = "ZZ" Then .HAcsry = "Z" 'V0.32 ADD
            If .LAcsry = "ZZ" Then .LAcsry = "Z" 'V0.32 ADD
        End With
MsToItemEnd:
    End Sub

    Sub MsToItemC(strMsCode As String, strItemIni As String, ByRef DModel As TypeDecodeModelC)                                                    'V0.21 ADD
        ' ''Dim usrDummy As TypeDecodeModelC '構造体初期化用
        Dim strItemPos As String           'Item Position
        Dim i As Integer          'LOOP COUNTER
        Dim strItemLetters As String           'Item Letters
        Dim strItemChar As String           'Item Character
        Dim intItemOrder As Integer          'Item Order
        Dim strItemContents As String           'Item Contents

        '構造体初期化
        ' ''DModel = usrDummy
        DModel = Nothing

        'MSCode代入
        DModel.MsCode = strMsCode

        'If yLeft(strMsCode, 5) Like "C[0-9][0-9]F[WE]" Then  '【DFS】                                                                        'V0.23 change
        If YLeft(strMsCode, 5) Like "C[2-9][0-9][A-Z][A-Z]" Then '【DFS】                                                                     'V0.28 change
            DModel.Model = YMid(strMsCode, 4, 2)
            DModel.ProductName = YLeft(strMsCode, 3)
            strItemPos = GetCurFile("DFSItemPosition", YLeft(strMsCode, 5), "", strItemIni) : If YTrim(strItemPos) = "" Then GoTo MsToItemEnd 'V0.23 change
            strItemPos = YSpace(YInStr(strMsCode, "-") - 1) & strItemPos  'CF***■をスペースで埋める。
            '<<<DFS Item Position Table for Get >>>
            'HL Side -------------- A ------- '製品識別
            'Oil Kind ------------- B ------- '封入液種類
            'Connection Method ---- C ------- '伝送器との接続形状
            'CAPILARY Length ------ D ------- 'キャピラリ長さ
            'CAPILARY Diameter ---- E ------- 'キャピラリ内径
            'CAPILARY Cover ------- F ------- 'キャピラリ被覆種類
            'CAPILARY Pipe -------- G ------- 'キャピラリPIPE構造,ステイ構造
            'Gasket Size ---------- H ------- 'ガスケット座サイズ,取付規格・サイズ,導圧管レス形
            'Gasket Standard ------ I ------- 'ガスケット座面規格,プロセス接続構造
            'Serration ------------ J ------- 'セレーション,ガスケット当り面
            'Diaphragm Material --- K ------- 'ダイアフラム材質
            'Wet Material --------- L ------- 'その他接液部材質
            'Diaphragm Diameter --- M ------- 'ダイアフラム径
            'Diaphragm Treatment -- N ------- 'ダイアフラム処理
            'Wet Treatment -------- O ------- '接液部処理
            'CAPILARY Position ---- P ------- 'キャピラリ取出構造
            'Flange Size ---------- Q ------- 'フランジサイズ,取付サイズ
            'Flange Standard ------ R ------- 'フランジ規格,取付規格および接続構造
            'Flange Rate ---------- S ------- 'フランジ定格
            'Flange Bearing ------- T ------- 'フランジ座面規格
            'Flange Material ------ U ------- 'フランジ材質,ガスケット材質
            'X2 Diameter ---------- V ------- '突出し部外径,フランジ仕様
            'X2 Length ------------ W ------- '突出し部長さ,ベントプラグ
            'Jack Up Bolt---------- X ------- 'ジャッキアップボルト
            strItemLetters = "ABCDEFGHIJKLMNOPQRSTUVWX"
            For i = 1 To YLen(strItemPos)
                strItemChar = YMid(strItemPos, i, 1)
                intItemOrder = YInStr(strItemPos, strItemChar)
                If intItemOrder > 0 Then
                    strItemContents = YMid(strMsCode, intItemOrder, 1)
                    With DModel
                        Select Case strItemChar
                            Case "A" : .HLSide = strItemContents         '製品識別
                            Case "B" : .OilKind = strItemContents        '封入液種類
                            Case "C" : .CnctMthd = strItemContents       '伝送器との接続形状
                            Case "D" : .CplyLngth = strItemContents      'キャピラリ長さ
                            Case "E" : .CplyDmtr = strItemContents       'キャピラリ内径
                            Case "F" : .CplyCvr = strItemContents        'キャピラリ被覆種類
                            Case "G" : .CplyPipe = strItemContents       'キャピラリPIPE構造,ステイ構造
                            Case "H" : .GsktSize = strItemContents       'ガスケット座サイズ,取付規格・サイズ,導圧管レス形
                            Case "I" : .GsktStd = strItemContents        'ガスケット座面規格,プロセス接続構造
                            Case "J" : .Srrtn = strItemContents          'セレーション,ガスケット当り面
                            Case "K" : .DphrgmMtrl = strItemContents     'ダイアフラム材質
                            Case "L" : .WtMtrl = strItemContents         'その他接液部材質
                            Case "M" : .DphrgmDmtr = strItemContents     'ダイアフラム径
                            Case "N" : .DphrgmTrtmnt = strItemContents   'ダイアフラム処理
                            Case "O" : .WtTrtmnt = strItemContents       '接液部処理
                            Case "P" : .CplyPstn = strItemContents       'キャピラリ取出構造
                            Case "Q" : .FlngSize = strItemContents       'フランジサイズ,取付サイズ
                            Case "R" : .FlngStd = strItemContents        'フランジ規格,取付規格および接続構造
                            Case "S" : .FlngRate = strItemContents       'フランジ定格
                            Case "T" : .FlngBrngStd = strItemContents    'フランジ座面規格
                            Case "U" : .FlngMtrl = strItemContents       'フランジ材質,ガスケット材質
                            Case "V" : .X2Dmtr = strItemContents         '突出し部外径,フランジ仕様
                            Case "W" : .X2Lngth = strItemContents        '突出し部長さ,ベントプラグ
                            Case "X" : .JckupBlt = strItemContents       'ジャッキアップボルト
                        End Select
                    End With
                End If
            Next i
            'ElseIf yLeft(strMsCode, 5) Like "CFR[0-9][0-9]" Then '【FlushConnectionRing】
            'ElseIf yLeft(strMsCode, 5) Like "C[0-9][0-9]FR" Then '【FlushConnectionRing】                                                        'V0.23 change
            'ElseIf yLeft(strMsCode, 5) Like "C1[0-9]FR" Then '【FlushConnectionRing】                                                                  'V0.28 change V0.32 DELETE
        ElseIf YLeft(strMsCode, 5) Like "C1[0-9][A-Z][A-Z]" Then '【FlushConnectionRing】                                                             'V0.32 ADD

            DModel.Model = YMid(strMsCode, 4, 2)
            DModel.ProductName = YLeft(strMsCode, 3)
            strItemPos = GetCurFile("FCRItemPosition", YLeft(strMsCode, 5), "", strItemIni) : If YTrim(strItemPos) = "" Then GoTo MsToItemEnd 'V0.23 change
            strItemPos = YSpace(YInStr(strMsCode, "-") - 1) & strItemPos  'CF***■をスペースで埋める。
            '<<<FCR Item Position Table for Get >>>
            'HL Side -------------- A ------- '製品識別
            'Flange Kind ---------- B ------- '構造
            'Flush Ring Kind ------ C ------- '用途
            'Flush Ring Size ------ D ------- 'プロセス接続サイズ
            'Flush Ring Material -- E ------- '接液部材質
            'Flange Standard ------ F ------- 'フランジ規格
            'Flange Rate ---------- G ------- 'フランジ定格
            'Gasket Standard ------ H ------- 'ガスケット座面形状
            'Serration ------------ I ------- 'ガスケット当たり面
            'Vent Plug Quantity --- J ------- 'ベントプラグ個数
            'Vent Screw Standard -- K ------- 'ベントプラグねじ規格
            'Vent Plug Kind ------- L ------- 'ベントプラグ種類
            'Gasket Kind ---------- M ------- 'ガスケット仕様
            'Gasket Size ---------- N ------- 'ガスケットサイズ
            'Foop Filler Material - O ------- 'ガスケット材質

            strItemLetters = "ABCDEFGHIJKLMN"
            For i = 1 To YLen(strItemPos)
                strItemChar = YMid(strItemPos, i, 1)
                intItemOrder = YInStr(strItemPos, strItemChar)
                If intItemOrder > 0 Then
                    strItemContents = YMid(strMsCode, intItemOrder, 1)
                    With DModel
                        Select Case strItemChar
                            Case "A" : .HLSide = strItemContents         '製品識別
                            Case "B" : .FlngKind = strItemContents       'フランジ構造
                            Case "C" : .FCRKind = strItemContents        'フラッシュリング種類
                            Case "D" : .FCRSize = strItemContents        'フラッシュリングサイズ
                            Case "E" : .FCRMtrl = strItemContents        'フラッシュリング材質
                            Case "F" : .FlngStd = strItemContents        'フランジ規格            V0.26 Change
                            Case "G" : .FlngRate = strItemContents       'フランジ定格            V0.26 Change
                            Case "H" : .GsktStd = strItemContents        'ガスケット座面形状      V0.26 Change
                            Case "I" : .Srrtn = strItemContents          'ガスケット当たり面      V0.26 Change
                            Case "J" : .VntPlgQty = strItemContents      'ベントプラグ個数        V0.26 Change
                            Case "K" : .VntScrwStd = strItemContents     'ベントプラグねじ規格    V0.26 Change
                            Case "L" : .VntPlgKind = strItemContents     'ベントプラグ種類        V0.26 Change
                            Case "M" : .GsktKind = strItemContents       'ガスケット種類          V0.26 Change
                            Case "N" : .GsktSize = strItemContents       'ガスケットサイズ        V0.26 Change
                            Case "O" : .HpflrMtrl = strItemContents      'ガスケット材質          V0.26 Change
                        End Select
                    End With
                End If
            Next i
        Else
            GoTo MsToItemEnd
        End If
MsToItemEnd:
    End Sub


    Function OptionCheck(ByVal MsCode$, ByVal OpList$) As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim L As Integer
        Dim OptionList As String
        Dim a As String
        '
        Const Sep As String = "/"

        OptionList = OpList & Sep
        OptionCheck = False
        L = YLen(OptionList)
        i = 1
        While i < L
            j = YInStr(i + 1, OptionList, Sep)
            a = YMid(OptionList, i, j - i) & "/"
            If a <> "" And YInStr(MsCode & "/", a) <> 0 Then OptionCheck = True
            i = j
        End While
    End Function
    ''' <summary>
    ''' ネットドライブ接続処理
    ''' </summary>
    ''' <param name="strFolder">共有フォルダ名</param>
    ''' <param name="strUser">共有フォルダユーザ名</param>
    ''' <param name="strPass">共有フォルダユーザワード</param>
    ''' <param name="arrNo">複数ドライブ割付用</param>
    ''' <returns>0:正常終了  -1:既ドライブ文字がある  -2:ドライブ文字取得失敗</returns>
    ''' <remarks>Output : g_strDriveLetter(arrNo)</remarks>
    Public Function FncNetDriveConnectArr(ByVal strFolder As String, ByVal strUser As String, ByVal strPass As String, ByVal arrNo As Integer) As Integer
        Dim typNetResource As NETRESOURCE
        Dim intRet As Integer
        Dim pass As String
        Dim user As String
        Dim strUNC As String
        Dim LetterNo() As Integer
        Dim UseDrive As String
        Dim AsciiCode As Integer
        Dim i As Integer
        '
        pass = strPass
        user = strUser
        strUNC = strFolder
        '
        ReDim Preserve g_strDriveLetter(arrNo)
        If g_strDriveLetter(arrNo) <> "" Then
            FncNetDriveConnectArr = -1                      '既ドライブ文字がある
            Exit Function
        End If
        UseDrive = ""
        For Each strDrive As String In Environment.GetLogicalDrives()
            UseDrive &= strDrive.Substring(0, 1)
        Next
        LetterNo = New Integer(26) {}
        '使用されている文字番号配列に１を入れる
        For i = 0 To UseDrive.Length - 1
            AsciiCode = YAsc(UseDrive.Substring(i, 1))
            If AsciiCode >= YAsc("C") And AsciiCode <= YAsc("Z") Then
                LetterNo(AsciiCode - &H40) = 1
            End If
        Next
        '使用されていない最初の文字を割り付ける
        For i = 3 To 26
            If LetterNo(i) = 0 Then
                g_strDriveLetter(arrNo) = YChr(i + &H40) & ":"
                Exit For
            End If
        Next
        If g_strDriveLetter(arrNo) = "" Then
            FncNetDriveConnectArr = -2                      'ドライブ文字取得失敗
            Exit Function
        End If
        With typNetResource
            .dwScope = RESOURCE_CONNECTED
            .dwType = RESOURCETYPE_ANY
            .dwDisplayType = RESOURCEDISPLAYTYPE_DOMAIN
            .lpLocalName = g_strDriveLetter(arrNo)
            .lpRemoteName = strUNC
        End With
        intRet = WNetAddConnection2(typNetResource, pass, user, 0) '次回ログオン時に再接続しない場合
        FncNetDriveConnectArr = intRet
    End Function
    ''' <summary>
    ''' ネットドライブ切断処理
    ''' </summary>
    ''' <returns>0:正常終了   -1:ドライブ割付未完了   -2:ドライブ切断失敗   -3:元々割り付いていない</returns>
    Public Function FncNetDriveDisconnect() As Integer
        Dim lngRet As Integer
        Dim strDrive As String   'ドライブ名
        Dim i As Integer
        '
        FncNetDriveDisconnect = 0
        If g_strDriveLetter IsNot Nothing Then
            '既にドライブが割りついているかの確認
            For i = 0 To g_strDriveLetter.GetUpperBound(0)
                If g_strDriveLetter(i) = "" Then
                    FncNetDriveDisconnect = -1  'ドライブ割付未完了がある。
                Else
                    'ネットワークドライブの切断
                    strDrive = g_strDriveLetter(i)
                    lngRet = WNetCancelConnection2(strDrive, CONNECT_UPDATE_PROFILE, CInt(True))

                    If lngRet <> 0 Then
                        FncNetDriveDisconnect = -2  'ドライブ切断失敗
                    End If
                End If
            Next
        Else
            FncNetDriveDisconnect = -3        '元々割り付いていない
        End If
        Erase g_strDriveLetter
    End Function
    ''' <summary>
    ''' ディレクトリパスを取得
    ''' </summary>
    ''' <param name="Ft">0:Windowsディレクトリのパス
    '''                  1:システムディレクトリのパス</param>
    Public Function GetWinDir(ByVal Ft%) As String
        If Ft = 0 Then
            Return System.Environment.GetEnvironmentVariable("windir")
        Else
            Return System.Environment.SystemDirectory
        End If
    End Function
    ''' <summary>
    ''' iniFileのセクション名だけを取得
    ''' </summary>
    ''' <param name="sFileName">iniファイル名</param>
    ''' <returns>セクション名</returns>
    ''' <remarks>CChar("")区切りで返信</remarks>
    Public Function GetCurFileSection(ByVal sFileName As String) As String
        Dim lsReturn As String = YSpace(2048)
        GetPrivateProfileSectionNames(lsReturn, 8192, sFileName)
        Return lsReturn.ToString
    End Function
    Public Function GetCurFile(ByVal ApName As String, ByVal KeyName As String, ByVal Defaults As String, ByVal Filename As String) As String
        'INIファイルから参照したいキーの値を取得する
        'ApName   : セクション名
        'KeyName  : 項目名
        'Default  : 項目が存在しない場合の初期値
        'FileName : 参照ファイル名
        '****************************************
        Dim BufDefaults As String = Defaults
        Dim strResult As String = " ".PadLeft(61440, " "c)
        Call GetPrivateProfileString(ApName, KeyName, Defaults, strResult, strResult.Length, Filename)
        GetCurFile = strResult.Substring(0, strResult.IndexOf(YChr(0)))
        If GetCurFile = "" Then GetCurFile = BufDefaults
    End Function
    Public Function GetCurFileNull(ByVal ApName As String, ByVal KeyName As String, ByVal Defaults As String, ByVal Filename As String) As String
        'INIファイルから参照したいキーの値を取得する(Null値も許可)
        'ApName   : セクション名
        'KeyName  : 項目名
        'Default  : 項目が存在しない場合の初期値
        'FileName : 参照ファイル名
        '****************************************
        Dim strResult As String = " ".PadLeft(81920, " "c)
        Call GetPrivateProfileString(ApName, KeyName, Defaults, strResult, strResult.Length, Filename)
        GetCurFileNull = strResult
    End Function
    Public Sub SaveCurFile(ByVal ApName As String, ByVal KeyName As String, ByVal Param As String, ByVal Filename As String)
        'INIファイルに新たなキーの値を書込む
        '   ※既存のキーがあれば更新・なければ新規作成する
        'ApName   : セクション名
        'KeyName  : 項目名
        'Param    : 更新する値
        'FileName : 書出ファイル名
        '****************************************
        Call WritePrivateProfileString(ApName, KeyName, Param, Filename)
    End Sub
    Public Function AppPath() As String
        Return Application.StartupPath()
        'Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function
    Function BarExpand(ByVal ChBar As String) As String
        'Barcode No. Expanding
        Dim IndkeyV As Integer
        Dim IndkeyN As Integer
        Dim DkeyNo As Double
        Dim DkeyPo As Double
        '
        If ChBar <> "" Then
            DkeyNo = 0
            DkeyPo = 1
            For IndkeyV = 5 To 0 Step -1
                IndkeyN = YAsc(ChBar.Substring(IndkeyV, 1))
                'DkeyNo = DkeyNo + DkeyPo * ((IndkeyN - 48) * Abs((48 <= IndkeyN And IndkeyN <= 57)) + (IndkeyN - 55) * Abs((65 <= IndkeyN And IndkeyN <= 86)))
                DkeyNo += DkeyPo * ((IndkeyN - YAsc("0")) * Math.Abs(CInt(YAsc("0") <= IndkeyN And IndkeyN <= YAsc("9"))) + (IndkeyN - YAsc("7")) * Math.Abs(CInt(YAsc("A") <= IndkeyN And IndkeyN <= YAsc("V"))))
                DkeyPo *= 32
            Next IndkeyV
            BarExpand = YRight("0000000000" & YMid(YStr(CInt(DkeyNo)), 2), 10)
        Else
            BarExpand = "0000000000"
        End If
    End Function
    Function Cvt32(ByVal CvtBarNo As String) As String
        'Make Barcode 6 string
        Dim Cvt As Long
        Dim Dm As String
        Dim N As Integer
        Dim i As Integer
        '
        Cvt = CInt(CvtBarNo)
        Cvt32 = ""
        '
        For N = 1 To 6
            i = CInt(Cvt Mod 32)
            Cvt = YInt(Cvt / 32)
            If i < 16 Then
                Dm = YRight(YHex(i), 1)
            Else
                Dm = YChr(&H41 + i - 10)
            End If
            Cvt32 = Dm + Cvt32
        Next N
    End Function
    Function DataSel(ByVal Nifuda As String, ByVal Key As String) As String
        'Data Select
        Dim Dm As String
        Dim P As Long
        '
        Dm = ""
        If Key <> "" Then
            Key = "$" & Key
            P = YInStr(Nifuda, Key)                'Search "$***"
            If P <> 0 Then
                Dm = YMid(Nifuda, CInt(P + YLen(Key)))    'Cutting BeforeData
                P = YInStr(Dm, "$")
                If P <> 0 Then
                    Dm = YLeft(Dm, CInt(P - 1))            'Cutting BehindData
                End If
            End If
        End If
        DataSel = YTrim(Dm)
    End Function
    Function DataSelC(strMsCode As String, strKey As String, intTargetNo As Integer) As String  'V0.21 ADD
        Select Case intTargetNo
            Case 0
                DataSelC = DataSel(strMsCode, strKey)
            Case Is > 0
                DataSelC = DataSel(strMsCode, YFormat(intTargetNo) & "_" & strKey)
            Case Else
                DataSelC = ""   '???
        End Select
    End Function
    Function ComputerName() As String
        Return Environment.MachineName
    End Function
    Function DecToHex(ByVal decdata As Double) As String
        Dim X As Double
        Dim z As Integer
        Dim c As Integer
        Dim D As Long
        Dim Res As Double
        Dim Bbb1 As String
        Dim Bbb2 As String
        '
        X = decdata
        If X = 0 Then DecToHex = "00000000" : Exit Function
        '
        If X < 0 Then z = 1 Else z = 0
        X = YAbs(X)
        Res = YLog(X) / YLog(2) 'Don't Erase "Res"
        c = CInt(YInt(Res + 1))
        D = CLng(X / 2 ^ (c - 24))
        c += &H40
        If z = 1 Then c += &H80
        Bbb1 = YHex(c)
        Bbb2 = YHex(CInt(D))
        DecToHex = Bbb1 + Bbb2
    End Function
    Function HexToDec(ByVal hexdata As String) As Double
        Dim Y As String
        Dim c As Long
        Dim D As Long
        Dim X As Double
        Dim z As Integer
        '
        HexToDec = 0
        '
        Y = YTrim(hexdata)
        If YLen(Y) <> 8 Then YBeep() : Exit Function
        c = CLng(YVal("&H" + YLeft(Y, 2)))
        D = CLng(YVal("&H" + YMid(Y, 3, 6)))
        If c > &H80 Then
            c -= &H80
            z = 1
        Else
            z = 0
        End If
        c -= &H40
        X = D * (2 ^ (c - 24))
        If z = 1 Then X = -X
        HexToDec = X
    End Function
    Function DecToHartChr(ByVal ChC As Double) As String  '１０進 → ＩＥＥＥ変換
        'IEEE754 format                         :E ;index
        'SEEEEEEE EMMMMMMM MMMMMMMM MMMMMMMM    :M ;coefficient（below decimal point）
        'd=(-1)^s*(1+M)*2^E-127                 :S ;"+-"
        Dim s As Double
        Dim E As Double
        Dim M As Double
        Dim Re1 As Double
        Dim Re2 As Double
        '
        If ChC = 0 Then DecToHartChr = "00000000" : Exit Function
        If YSgn(ChC) < 0 Then s = &H80000000 Else s = &H0
        ChC = YAbs(ChC)
        Re1 = YLog(ChC) / YLog(2)
        E = YInt(Re1) + 127
        Re2 = (ChC / (2 ^ YInt(Re1)) - 1) * &H800000
        M = YInt(Re2)
        DecToHartChr = YHex(CInt((CLng(E * &H800000) Or CLng(M)) Or CLng(s)))
    End Function
    Function HartChrToDec(ByVal ChC As String) As Double  'ＩＥＥＥ → １０進変換
        'IEEE754 format                         :E ;index
        'SEEEEEEE EMMMMMMM MMMMMMMM MMMMMMMM    :M ;coefficient（below decimal point）
        'd=(-1)^S*(1+M)*2^E-127                 :S ;"+-"
        Dim s As Double
        Dim E As Double
        Dim M As Double
        Dim Zrdt As Double
        '
        On Error Resume Next
        ChC = YRight("00000000" + ChC, 8)
        If ChC = "00000000" Then HartChrToDec = 0 : Exit Function
        s = (CLng(YVal("&h" + YLeft(ChC, 1)) * 2) And &H10) / &H10
        E = YVal("&h" + YLeft(YHex(CInt((CLng(YVal("&h" + YLeft(ChC, 3))) And &H7FF) * 2)), 2))
        If (CLng(YVal("&h" + YMid(ChC, 3, 2))) And &HFF) = 0 And (CLng(YVal("&h" + YMid(ChC, 5, 1))) And &H8) <> 0 Then
            M = (&H810000 + YVal("&h" + YMid(ChC, 3, 6))) * 2 ^ -23
            Zrdt = (-1) ^ s * (M) * 2 ^ (E - 127)
        Else
            M = (CLng(YVal("&h" + YMid(ChC, 3, 6))) And &H7FFFFF) / &H800000
            Zrdt = (-1) ^ s * (1 + M) * 2 ^ (E - 127)
        End If
        HartChrToDec = YVal(YFormat(Zrdt, "0.000000E+000"))
    End Function
    Function Fan(ByVal Dm As String) As String      'バイト列入れ替え
        'In ) "ASDFGH"
        'OUT) "GHDFAS"
        Dim Dm2 As String
        Dim i As Long
        Dim k As Long
        '
        Dm2 = ""
        k = YLen(Dm)
        For i = k To 2 Step -2
            Dm2 &= YMid(Dm, CInt(i - 1), 2)
        Next
        Fan = Dm2
    End Function
    Public Function IPaddress() As String
        'Dim adrList As IPAddress() = Dns.GetHostAddresses(Dns.GetHostName)    却下１
        'For Each Address As IPAddress In adrList
        '    If Address.ToString.Substring(0, 3) = "10." Then
        '        Return Address.ToString
        '    End If
        'Next
        'For Each hostAdr In Dns.GetHostEntry(hostName).AddressList()                                却下２
        '    If hostAdr.ToString().StartsWith("10.") Then
        '        hostIP = hostAdr.ToString()                                     'IP Address
        '        Exit For
        '    End If
        'Next
        Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()             '採用
        For Each adapter As NetworkInterface In adapters
            If adapter.OperationalStatus = OperationalStatus.Up Then                                'ネットワーク接続状態が UP のアダプタのみ表示 (adapter.Name , adapter.Description)
                Dim ip_prop As IPInterfaceProperties = adapter.GetIPProperties()
                Dim addrs As UnicastIPAddressInformationCollection = ip_prop.UnicastAddresses()     'ユニキャスト IP アドレスの取得
                For Each addr As UnicastIPAddressInformation In addrs                               '******************************************************************
                    If ip_prop.DnsSuffix <> "" AndAlso addr.IsDnsEligible = True Then               'DNSサフィックスが有り、DNSに表示されるものが有効と判断 ***********
                        Return addr.Address.ToString()                                              '******************************************************************
                    End If                                                                          '******************************************************************
                Next
            End If
        Next
        Return ""
    End Function
    '     引数なし                このメッセージを表示します (-? と同じです)
    '    -i                      GUI インターフェイスを表示します。このオプションは最初に指定する必要があります
    '    -l                      ログオフ (-m オプションとは併用できません)
    '    -s                      コンピュータをシャットダウンします
    '    -r                      コンピュータをシャットダウンして再起動します
    '    -a                      システム シャットダウンを中止します
    '    -m \\コンピュータ名     シャットダウン/再起動/中止するリモート コンピュータの名前です
    '    -t xx                   シャットダウンのタイムアウトを xx 秒に設定します
    '    -c "コメント"           シャットダウンのコメントです (127 文字まで)
    '    -f                      実行中のアプリケーションを警告なしに閉じます
    '    -d [u][p]:xx:yy         シャットダウンの理由コードです
    '                            u = ユーザー コード
    '                            p = 計画されたシャットダウンのコード
    '                            xx = 重大な理由コード (255 以下の正の整数)
    '                            yy = 重大ではない理由コード (65535 以下の正の整数)
    Sub ExWindow(ByVal Ret As Integer)
        Dim psi As New System.Diagnostics.ProcessStartInfo() With {.FileName = "shutdown.exe"}
        'psi.FileName = "shutdown.exe"
        'コマンドラインを指定
        Select Case Ret
            Case 0          'シャットダウン
                psi.Arguments = "-s -f -t 00"
            Case 1          '再起動
                psi.Arguments = "-r"
            Case 2          'ログオフ
                psi.Arguments = "-l"
        End Select
        'ウィンドウを表示しないようにする（こうしても表示される）
        psi.CreateNoWindow = True
        '起動
        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(psi)
    End Sub
    Function GetFileDateOrTime(ByVal FileName As String, ByVal WhichValue As Integer) As String
        'FileInfoオブジェクトを作成
        Dim fi As New FileInfo(FileName)
        Select Case WhichValue
            Case 1
                '作成日時の取得
                Return fi.CreationTime.ToString("yyyy/MM/dd")
            Case 2
                'アクセス日時の取得
                Return fi.LastAccessTime.ToString("yyyy/MM/dd")
            Case 3
                '更新日時の取得
                Return fi.LastWriteTime.ToString("yyyy/MM/dd")
            Case 4
                '作成日時の取得
                Return fi.CreationTime.ToString("HH:mm:ss")
            Case 5
                'アクセス日時の取得
                Return fi.LastAccessTime.ToString("HH:mm:ss")
            Case 6
                '更新日時の取得
                Return fi.LastWriteTime.ToString("HH:mm:ss")
            Case Else
                Return fi.LastWriteTime.ToString()
        End Select
    End Function
    Sub LogSave(ByVal Msg As String)  '自分へLogSave
        Dim LogFile As String
        '
        LogFile = AppPath() & "\Log\" & DateTime.Now.ToString("yyyy") & "\" & DateTime.Now.ToString("yyMMdd") & ".log"
        Call MakeDirectory(LogFile)

        Dim enc As Encoding = Encoding.GetEncoding("shift_jis")
        Msg$ = Msg$.Replace(YvbCrLf, "")
        File.AppendAllText(LogFile, DateTime.Now.ToString("HH:mm:ss") & "[" & Msg$ & "]" & YvbCrLf(), enc)
    End Sub
    Function MakeDirectory(ByVal strFileName As String) As Boolean
        If IO.Path.GetFileName(strFileName) <> "" Then      'ファイル名は取り除く
            strFileName = strFileName.Substring(0, strFileName.IndexOf(IO.Path.GetFileName(strFileName)))
        End If
        Try
            IO.Directory.CreateDirectory(strFileName)
            Return True
        Catch
            Return False
        End Try
    End Function
    Public Function IsWindowsNT(Optional ByRef OSType As String = "") As Boolean
        ' OSの情報を取得
        Dim osInfo As OperatingSystem = Environment.OSVersion
        Dim windowsName As String = "不明" ' Windows名
        Select Case osInfo.Platform
            Case PlatformID.Win32Windows  ' Windows 9x系
                If osInfo.Version.Major = 4 Then
                    Select Case osInfo.Version.Minor
                        Case 0  ' Win95は、.NET FrameworkのサポートOSではない
                            windowsName = "Windows 95"
                        Case 10
                            windowsName = "Windows 98"
                        Case 90
                            windowsName = "Windows Me"
                    End Select
                End If
                If OSType <> "" Then OSType = windowsName
                Return False
            Case PlatformID.Win32NT  ' Windows NT系
                If osInfo.Version.Major = 4 Then
                    windowsName = "Windows NT 4.0"
                ElseIf osInfo.Version.Major = 5 Then
                    Select Case osInfo.Version.Minor
                        Case 0
                            windowsName = "Windows 2000"
                        Case 1
                            windowsName = "Windows XP"
                        Case 2
                            windowsName = "Windows Server 2003"
                    End Select
                End If
                If OSType <> "" Then OSType = windowsName
                Return True
        End Select
        If OSType <> "" Then OSType = windowsName
        Return False
    End Function
    Function FileLAppend(ByVal FileName As String, ByRef Data() As String) As Boolean
        Try
            Dim enc As Encoding = Encoding.GetEncoding("shift_jis")
            File.AppendAllLines(FileName, Data, enc)
            Return True
        Catch
            Return False
        End Try
    End Function
    Function FileLRead(ByVal FileName As String, ByRef Data() As String, ByRef Num%) As Boolean
        Try
            Dim enc As Encoding = Encoding.GetEncoding("shift_jis")
            Dim Lines As String() = File.ReadAllLines(FileName, enc)
            Data = Lines
            Num = Lines.Length
            Return True
        Catch
            Return False
        End Try
    End Function
    Function FileLWrite(ByVal FileName$, ByRef Data() As String, ByRef Num%) As Boolean
        Try
            If Num = 0 Then

            End If
            Dim enc As Encoding = Encoding.GetEncoding("shift_jis")
            File.WriteAllLines(FileName, Data, enc)
            Return True
        Catch
            Return False
        End Try
    End Function
    Sub SaveText(ByVal strFileName As String, ByVal strMsg As String, Optional ByVal strHeader As String = "", Optional ByVal blnWT As Boolean = False)
        'ディレクトリーの作成
        Call MakeDirectory(strFileName)
        Try
            'データ追加
            Dim enc As Encoding = Encoding.GetEncoding("shift_jis")
            If YDir(strFileName) = "" AndAlso strHeader <> "" Then
                File.AppendAllText(strFileName, strHeader & YvbCrLf(), enc)   'ファイルが無く、ヘッダ指定ある場合はヘッダーを付けて作成
            End If
            If blnWT = False Then                                           '追加保存
                strMsg = strMsg.Replace(YvbCrLf, "") & YvbCrLf()
                File.AppendAllText(strFileName, strMsg, enc)
            Else                                                            '上書き保存（改行コードも含む）
                File.WriteAllText(strFileName, strMsg, enc)
            End If
        Catch
            Call MessageBox.Show("ファイル書込みできません。" & YvbCrLf() & strFileName & YvbCrLf() & "を確認してください。", "書込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' 文字列バイト数取得
    ''' </summary>
    ''' <param name="stTarget">対象文字列</param>
    ''' <returns>文字列のバイト数</returns>
    Function LenByte(ByVal stTarget As String) As Integer
        Return Encoding.GetEncoding(932).GetByteCount(stTarget)
    End Function
    ''' <summary>
    ''' 文字列の左からの切り出し
    ''' </summary>
    ''' <param name="strTarget">対象文字列</param>
    ''' <param name="intByteSize">切り出すバイト数</param>
    ''' <returns>切り出した文字列</returns>
    Function LeftByte(ByVal strTarget As String, ByVal intByteSize As Integer) As String
        Dim strTempString As String   '文字列
        strTempString = strTarget
        Do
            If LenByte(strTempString) <= intByteSize Then
                Return strTempString
            Else
                strTempString = YLeft(strTempString, YLen(strTempString) - 1)
            End If
        Loop
    End Function
    ''' <summary>
    ''' 待ち(s)
    ''' </summary>
    ''' <param name="Wt">秒</param>
    Sub Wait(ByVal Wt As Single)
        Dim WtTime As Double
        Dim Day As Double
        '   Wt(s) waitting!!
        Day = 24 * 60.0# * 60.0#        'Max
        If Wt <> 0 Then
            If Wt < Day Then
                If Wt < 0.4 Then
                    YSleep(CInt(Wt * 1000))
                    YDoEvents()
                Else
                    WtTime = YTimer()
                    Do Until Math.Abs(YTimer() - WtTime) >= Wt
                        YSleep(50)
                        If YTimer() < WtTime Then WtTime -= Day
                        YDoEvents()
                    Loop
                End If
            End If
        End If
    End Sub
    Function Marume(ByVal z As Double, ByVal MK%) As String
        Dim W As String
        '
        W = "0." + YString(MK - 1, "0") + "E-00"
        If YInStr(YFormat(z), "E") > 0 Then           '浮動小数点のものは浮動小数点のまま表示
            Marume = (YFormat(z, W))
        Else
            Marume = YFormat(YVal(YFormat(z, W)))      '固定小数点
        End If
    End Function
    '***********************************************************************************************************/
    '* #! Discription    -: パターン番号取得処理(V0.26 ADD)                                     :-             */
    '* #! Input          -: strItemInfo() As String     Item情報配列                            :-             */
    '* #!                -: strItemDelm As String       ITEMのデリミタ                          :-             */
    '* #!                -: strInputTypeCode As String  入力タイプコード(ex.118-W-M-SW-3-J1-A)  :-             */
    '* #!                -: strTypeCodeDelm As String   TypeCodeのデリミタ                      :-             */
    '* #!                -: strCodeDelm As String       Codeのデリミタ                          :-             */
    '* #! Output         -: なし                                                                :-             */
    '* #! ReturnValue    -: GetPatNo As Integer         パターン番号(-1のときは無し)            :-             */
    '***********************************************************************************************************/
    Function GetPatNo(strItemInfo() As String, strItemDelm As String, strInputTypeCode As String,
                      strTypeCodeDelm As String, strCodeDelm As String) As Integer
        Dim i As Integer  'ループカウンタ
        Dim j As Integer  'ループカウンタ
        Dim k As Integer  'ループカウンタ
        Dim L As Integer  'ループカウンタ
        Dim strItemInfoContents As String   'Item情報内容
        Dim strFieldContents() As String   'フィールド内容
        Dim strTypeCodeFieldContents() As String   'TypeCodeフィールド内容
        Dim strCodeContents() As String   'Code内容
        Dim strInputTypeCodeContents() As String   '入力Code内容
        Dim intCodeMatchCount As Integer  'Code一致数

        GetPatNo = -1

        '入力タイプコードの分解
        strInputTypeCodeContents = YSplit(strInputTypeCode, strTypeCodeDelm)
        For i = 0 To YUBound(strInputTypeCodeContents)
            strInputTypeCodeContents(i) = YTrim(strInputTypeCodeContents(i))
        Next i

        For i = 0 To YUBound(strItemInfo)    'ITEM数分繰り返す
            'Item情報読み込み(ex."118-SW|HW|TW|UW-M-SW-3-J1-A,118-SW|HW|TW|UW-M-SW-3-J1-B")
            strItemInfoContents = YTrim(strItemInfo(i))

            If strItemInfoContents <> "" Then
                strFieldContents = YSplit(strItemInfoContents, strItemDelm)
                For j = 0 To YUBound(strFieldContents)               'フィールド数分繰り返す
                    'フィールド内容(ex."118-SW|HW|TW|UW-M-SW-3-J1-A")取得
                    strFieldContents(j) = YTrim(strFieldContents(j))

                    'Code一致数初期化
                    intCodeMatchCount = 0
                    strTypeCodeFieldContents = YSplit(strFieldContents(j), strTypeCodeDelm)
                    For k = 0 To YUBound(strTypeCodeFieldContents)   'TypeCodeのフィールド数分繰り返す
                        'TypeCodeフィールド内容(ex."SW|HW|TW|UW")取得
                        strTypeCodeFieldContents(k) = YTrim(strTypeCodeFieldContents(k))

                        strCodeContents = YSplit(strTypeCodeFieldContents(k), strCodeDelm)
                        For L = 0 To YUBound(strCodeContents)    'Codeのフィールド数分繰り返す
                            'Code内容(ex."HW")取得
                            strCodeContents(L) = YTrim(strCodeContents(L))

                            '一致確認
                            If YTrim(strInputTypeCodeContents(k)) Like YTrim(strCodeContents(L)) Then
                                'Code一致数インクリメント
                                intCodeMatchCount += 1
                            End If
                        Next L
                    Next k
                    If intCodeMatchCount = YUBound(strTypeCodeFieldContents) + 1 Then
                        'パターン番号設定
                        GetPatNo = i
                        Exit For
                    End If
                Next j
            End If

            If GetPatNo >= 0 Then Exit For
        Next i
    End Function
    ''' <summary>
    ''' パスワード（暗号化／複合化）
    ''' </summary>
    ''' <param name="Section">Section名</param>
    ''' <param name="Key">Key値</param>
    ''' <param name="strPWD">暗号化する文字</param>
    ''' <param name="FilePath">保存ファイル名</param>
    ''' <returns>RD:暗号 WT:OK</returns>
    ''' <remarks>暗号文字を指定しない場合は、複合</remarks>
    Function PassWordEncDec(ByVal Section As String, ByVal Key As String, ByVal strPWD As String, ByVal FilePath As String) As String
        If strPWD <> "" Then                                                     'エンコード
            '文字列からUTF8のバイト列に変換
            Dim datas() As Byte = Encoding.UTF8.GetBytes(strPWD)
            Dim hexText As String = BitConverter.ToString(datas).Replace("-", "")
            Call SaveCurFile(Section, Key, hexText, FilePath)
            Return "OK"
        Else                                                                    'デコード
            strPWD = GetCurFile(Section, Key, "", FilePath)
            Dim hexChars(CInt(strPWD.Length / 2) - 1) As String
            For i As Integer = 0 To CInt(strPWD.Length / 2) - 1
                hexChars(i) = strPWD.Substring(i * 2, 2)
            Next
            '16進文字列をbyteに変換
            Dim decData(hexChars.Length) As Byte
            For i As Integer = 0 To hexChars.Length - 1
                decData(i) = Convert.ToByte(hexChars(i), 16)
            Next
            'UTF8のバイト列からstringに変換
            Dim decText As String = Encoding.UTF8.GetString(decData)
            decText = decText.Substring(0, decText.Length - 1)  'ヌルカット
            Return decText
        End If
    End Function
    ''' <summary>
    ''' 半角から全角へ
    ''' </summary>
    ''' <param name="dblOrg">半角</param>
    ''' <returns>全角</returns>
    Function ToWide(ByVal dblOrg As Double) As String               'オーバーロード　引数に数値
        Dim strWide As String = ""
        Dim intAsc As Integer
        For i As Integer = 0 To dblOrg.ToString.Length - 1
            intAsc = YAsc(dblOrg.ToString.Substring(i, 1))
            Select Case intAsc
                Case 33 To 126 : strWide &= YChr(65248 + intAsc)        '半角英数字記号 → 全角英数字記号
                Case Else : strWide &= YChr(intAsc)
            End Select
        Next
        Return strWide
    End Function
    ''' <summary>
    ''' 半角から全角へ
    ''' </summary>
    ''' <param name="strOrg">半角</param>
    ''' <returns>全角</returns>
    Function ToWide(ByVal strOrg As String) As String               'オーバーロード　引数に文字列
        Dim strWide As String = ""
        Dim intAsc As Integer
        For i As Integer = 0 To strOrg.Length - 1
            intAsc = YAsc(strOrg.Substring(i, 1))
            Select Case intAsc
                Case 33 To 126 : strWide &= YChr(65248 + intAsc)        '半角英数字記号 → 全角英数字記号
                Case Else : strWide &= YChr(intAsc)
            End Select
        Next
        Return strWide
    End Function
    ''' <summary>
    ''' 配列の初期化
    ''' </summary>
    ''' <param name="strArr">初期化される配列</param>
    Public Sub ArrErase(ByRef strArr() As String)
        Array.Clear(strArr, 0, strArr.Length)
    End Sub
    ''' <summary>
    ''' 配列のソート
    ''' </summary>
    ''' <param name="strArr">変換される配列</param>
    ''' <param name="i">0:昇順ソート
    '''                 1:降順ソート</param>
    Public Sub ArrSort(ByVal strArr() As String, Optional ByVal i As Integer = 0)
        Array.Sort(strArr)
        If i = 1 Then
            Array.Reverse(strArr)
        End If
    End Sub
    ''' <summary>
    ''' ＯＳ言語の種類判別
    ''' </summary>
    ''' <returns> 0:不明
    '''           1:日本語
    '''           2:英語
    '''           3:中国語</returns>
    Public Function GetLanguage() As Integer
        'CultureName にOSの言語を格納
        Dim CultureName As String = System.Globalization.CultureInfo.CurrentCulture.Name
        Dim intLanguage As Integer
        Select Case CultureName
            Case "ja-JP"
                intLanguage = 1 '日本語
            Case "en-"
                intLanguage = 2 '英語
            Case "zh-CHS"
                intLanguage = 3 '中国語
            Case Else
                intLanguage = 0
        End Select
        Return intLanguage
    End Function
    '*************************************************************************************************************************
    '*************************************************************************************************************************
    '******************************************* <<< 以下、自作関数 >>> *******************************************************
    '*************************************************************************************************************************
    '*************************************************************************************************************************
    Function YAbs(ByVal num As Double) As Double    'オーバーロード　　引数にDoubleを指定するとDouble型
        Return Math.Abs(num)
    End Function
    Function YAbs(ByVal num As Long) As Long        'オーバーロード　　引数にLongを指定するとLong型
        Return Math.Abs(num)
    End Function
    Function YAbs(ByVal num As Integer) As Integer  'オーバーロード　　引数にIntegerを指定するとInteger型
        Return Math.Abs(num)
    End Function
    Function YAsc(ByVal OrgStr As String) As Integer
        Dim strCode As Char = Convert.ToChar(OrgStr)
        Dim intCode As Integer = Convert.ToInt32(strCode)
        Return intCode
    End Function
    Sub YBeep()
        System.Media.SystemSounds.Beep.Play()
    End Sub
    Function YChoose(ByVal IntNum As Integer, ByVal Sel1 As Object, ByVal Sel2 As Object, Optional ByVal Sel3 As Object = Nothing, Optional ByVal Sel4 As Object = Nothing, Optional ByVal Sel5 As Object = Nothing, Optional ByVal Sel6 As Object = Nothing, Optional ByVal Sel7 As Object = Nothing, Optional ByVal Sel8 As Object = Nothing, Optional ByVal Sel9 As Object = Nothing, Optional ByVal Sel10 As Object = Nothing) As Object 'オーバーロード　引数にObjectを指定された時
        Dim Ary As String() = New String() {"", CStr(Sel1), CStr(Sel2), CStr(Sel3), CStr(Sel4), CStr(Sel5), CStr(Sel6), CStr(Sel7), CStr(Sel8), CStr(Sel9), CStr(Sel10)}
        Return Ary(IntNum)
    End Function
    Function YChoose(ByVal IntNum As Integer, ByVal Sel1 As String, ByVal Sel2 As String, Optional ByVal Sel3 As String = "", Optional ByVal Sel4 As String = "", Optional ByVal Sel5 As String = "", Optional ByVal Sel6 As String = "", Optional ByVal Sel7 As String = "", Optional ByVal Sel8 As String = "", Optional ByVal Sel9 As String = "", Optional ByVal Sel10 As String = "") As String 'オーバーロード　引数に文字列を指定された時
        Dim Ary As String() = New String() {"", Sel1, Sel2, Sel3, Sel4, Sel5, Sel6, Sel7, Sel8, Sel9, Sel10}
        Return Ary(IntNum)
    End Function
    Function YChr(ByVal OrgInt As Integer) As String
        Dim strCode As Char = Convert.ToChar(OrgInt)
        Return strCode
    End Function
    Function YCount(ByVal OrgArr() As String) As Integer
        Return OrgArr.Length
    End Function
    Function YDate() As String
        Return System.DateTime.Now.ToString("yyyy/MM/dd")
    End Function
    Function YDateAdd(ByVal strInterval As String, ByVal intNum As Long, ByVal strDate As String) As String
        Select Case strInterval
            Case "yyyy" : Return Date.Parse(strDate).AddYears(Convert.ToInt16(intNum)).ToString("yyyy/MM/dd")
            Case "m" : Return Date.Parse(strDate).AddMonths(Convert.ToInt16(intNum)).ToString("yyyy/MM/dd")
            Case "d" : Return Date.Parse(strDate).AddDays(intNum).ToString("yyyy/MM/dd")
            Case "h" : Return Date.Parse(strDate).AddHours(intNum).ToString("HH:mm:ss")
            Case "n" : Return Date.Parse(strDate).AddMinutes(intNum).ToString("HH:mm:ss")
            Case "s" : Return Date.Parse(strDate).AddSeconds(intNum).ToString("HH:mm:ss")
            Case Else
                Call MessageBox.Show("yDateAddの使い方が間違っています！", "関数エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return ""
        End Select
    End Function
    Function YDateDiff(ByVal interval As String, ByVal date1 As Date, ByVal date2 As Date) As String        'オーバーロード　引数にDate型
        Dim bYY As Integer
        Dim bMM As Integer
        'Dim bDD As Integer
        '
        bYY = (Convert.ToInt16(date2.Year) - Convert.ToInt16(date1.Year)) * 12
        bMM = (Convert.ToInt16(date2.Month) - Convert.ToInt16(date1.Month)) + bYY
        'If date1.Day > date2.Day Then                   '比較日以前だったら
        'Dim dt As System.DateTime = date2
        'dt = dt.AddMonths(-1)
        'bDD = System.DateTime.DaysInMonth(date2.Year, dt.Month)            '前月の末日を求める
        'bDD -= date1.Day                       '末日までの日数を求める
        'bDD += date2.Day                       '当月の日数をプラスする'1ヶ月未満の日数を求める
        'Else
        'bDD = date2.Day - date1.Day                 '1ヶ月未満の日数を求める
        'End If
        bYY = bMM \ 12                                  '年数を求める 
        bMM = bMM Mod 12                                '月数を求める
        '-------------------------------------------------------------------------------------------
        Dim TimeD As TimeSpan = date2.Subtract(date1)   'TimeSpan構造体を利用(d,h,n,sの時)
        '-------------------------------------------------------------------------------------------
        Select Case interval
            Case "yyyy" : Return Convert.ToString(bYY)
            Case "m" : Return Convert.ToString(bMM)
            Case "d" : Return Convert.ToString(TimeD.TotalDays)
            Case "h" : Return Convert.ToString(TimeD.TotalHours)
            Case "n" : Return Convert.ToString(TimeD.TotalMinutes)
            Case "s" : Return Convert.ToString(TimeD.TotalSeconds)
            Case "ww" : Return CStr(Math.Truncate(CInt(CInt(Convert.ToString(TimeD.TotalDays)) + YWeekday(YRight(YFormat(date1.Year), 2) & "/01/01") - 1) / 7))
            Case Else
                Call MessageBox.Show("要対応！", "DateDiff関数", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return ""
        End Select
    End Function
    Function YDateDiff(ByVal interval As String, ByVal date1 As String, ByVal date2 As String) As String    'オーバーロード　引数にString型
        Dim bYY As Integer
        Dim bMM As Integer
        'Dim bDD As Integer
        '
        bYY = (Convert.ToInt16(DateTime.Parse(date2).Year) - Convert.ToInt16(DateTime.Parse(date1).Year)) * 12
        bMM = (Convert.ToInt16(DateTime.Parse(date2).Month) - Convert.ToInt16(DateTime.Parse(date1).Month)) + bYY
        'If DateTime.Parse(date1).Day > DateTime.Parse(date2).Day Then                   '比較日以前だったら
        'Dim dt As System.DateTime = CDate(date2)
        'dt = dt.AddMonths(-1)
        'bDD = System.DateTime.DaysInMonth(DateTime.Parse(date2).Year, dt.Month)            '前月の末日を求める
        'bDD -= DateTime.Parse(date1).Day                       '末日までの日数を求める
        'bDD += DateTime.Parse(date2).Day                       '当月の日数をプラスする'1ヶ月未満の日数を求める
        'Else
        'bDD = DateTime.Parse(date2).Day - DateTime.Parse(date1).Day                 '1ヶ月未満の日数を求める
        'End If
        bYY = bMM \ 12                                  '年数を求める 
        bMM = bMM Mod 12                                '月数を求める
        '-------------------------------------------------------------------------------------------
        Dim TimeD As TimeSpan = DateTime.Parse(date2).Subtract(DateTime.Parse(date1))   'TimeSpan構造体を利用(d,h,n,sの時)
        '-------------------------------------------------------------------------------------------
        Select Case interval
            Case "yyyy" : Return Convert.ToString(bYY)
            Case "m" : Return Convert.ToString(bMM)
            Case "d" : Return Convert.ToString(TimeD.TotalDays)
            Case "h" : Return Convert.ToString(TimeD.TotalHours)
            Case "n" : Return Convert.ToString(TimeD.TotalMinutes)
            Case "s" : Return Convert.ToString(TimeD.TotalSeconds)
            Case "ww" : Return CStr(Math.Truncate(CInt(CInt(Convert.ToString(TimeD.TotalDays)) + YWeekday(YFormat(date1, "yy") & "/01/01") - 1) / 7))
            Case Else
                Call MessageBox.Show("要対応！", "DateDiff関数", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return ""
        End Select
    End Function
    Function YDir(ByVal strDirFile As String, Optional ByVal intDum As Integer = 0) As String
        If IO.File.Exists(strDirFile) Then
            If intDum = 0 Then
                Return IO.Path.GetFileName(strDirFile)          'ファイル名 "c:\windows\system32\notepad.exe" → "notepad.exe"
            Else
                Return IO.Path.GetDirectoryName(strDirFile)     'フォルダ名 "c:\windows\system32\notepad.exe" → "c:\windows\system32"
            End If
        Else
            Return ""
        End If
    End Function
    Sub YDoEvents()
        Application.DoEvents()
    End Sub
    Function YFileLen(ByVal FileName As String) As Long
        Dim fi As New FileInfo(FileName)
        Return fi.Length
    End Function
    Function YFix(ByVal num As Double) As Long
        Return Convert.ToInt64(Math.Truncate(num))
    End Function
    Function YFormat(ByVal OrgStr As Integer, Optional ByVal strForm As String = "") As String 'オーバーロード　　引数にIntegerを指定された時
        If strForm = "" Then
            Return OrgStr.ToString.Trim
        Else
            Return OrgStr.ToString(strForm)
        End If
    End Function
    Function YFormat(ByVal OrgStr As Double, Optional ByVal strForm As String = "") As String 'オーバーロード　　引数にDoubleを指定された時
        If strForm = "" Then
            Return OrgStr.ToString.Trim
        Else
            Return OrgStr.ToString(strForm)
        End If
    End Function
    Function YFormat(ByVal OrgStr As String, Optional ByVal strForm As String = "") As String 'オーバーロード　　引数にStringを指定された時
        If strForm = "" Then
            Return Convert.ToString(OrgStr)
        Else
            Return Date.Parse(OrgStr).ToString(strForm)
        End If
    End Function
    Function YFormat(ByVal OrgObj As Object) As String                                          'オーバーロード　　引数にObjectを指定された時
        Return Convert.ToString(OrgObj)
    End Function
    Function YHex(ByVal OrgInt As Integer) As String
        Dim strHex As String = Convert.ToString(OrgInt, 16)
        Return strHex.ToUpper
    End Function
    Function YHex(ByVal OrgLng As Long) As String               'オーバーロード　　引数にLongを指定された時
        Dim strHex As String = Convert.ToString(OrgLng, 16)
        Return strHex.ToUpper
    End Function
    Function YIIf(ByVal TestEX As Boolean, ByVal TruePart As Object, ByVal FalsePart As Object) As Object 'オーバーロード　　引数にObjectを指定された時
        Return If(TestEX, TruePart, FalsePart)
    End Function
    Function YIIf(ByVal TestEX As Boolean, ByVal TruePart As String, ByVal FalsePart As String) As String 'オーバーロード　　引数にStringを指定された時
        Return If(TestEX, TruePart, FalsePart)
    End Function
    Function YInputBox(ByVal Prompt As String, ByVal Titl As String, Optional ByVal Def As String = "") As String
        'InputBoxだけは、とりあえずこのようにしてしまった。。
        Return Microsoft.VisualBasic.InputBox(Prompt, Titl, Def)
    End Function
    Function YInStr(ByVal OrgStr As String, ByVal FindStr As String) As Integer                             'オーバーロード　　引数にiStartが無い時
        If OrgStr = Nothing Then OrgStr = ""    'エラー防止
        Return OrgStr.IndexOf(FindStr) + 1
    End Function
    Function YInStr(ByVal iStart As Integer, ByVal OrgStr As String, ByVal FindStr As String) As Integer    'オーバーロード　　引数にiStartがある時
        If OrgStr = Nothing Then OrgStr = ""    'エラー防止
        Return OrgStr.IndexOf(FindStr, iStart - 1) + 1
    End Function
    Function YInt(ByVal num As Double) As Integer
        Return Convert.ToInt32(Math.Floor(num))
    End Function
    Function YIsDate(ByVal ChkStr As String) As Boolean
        Dim Dt As DateTime
        If DateTime.TryParse(ChkStr, Dt) Then
            Return True
        Else
            Return False
        End If
    End Function
    Function YIsNumeric(ByVal ChkStr As String) As Boolean
        Dim d As Double
        Try
            d = CDbl(ChkStr)
            ChkStr = CStr(d)
            If Double.TryParse(ChkStr, d) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Function YJoin(ByVal OrgStr As String(), ByVal Delim As String) As String
        Return String.Join(Delim, OrgStr)
    End Function
    Sub YKill(ByVal FileName As String)
        'If System.IO.File.Exists(IO.Path.GetDirectoryName(FileName)) = True Then       'V0.35 Del ファイル存在確認
        If File.Exists(FileName) = True Then                                  'V0.35 Add ファイル存在確認
            My.Computer.FileSystem.DeleteFile(FileName)
        End If
    End Sub
    Sub YKill2(ByVal FileName As String)    'ワイルドカードの指定可能
        If Directory.Exists(IO.Path.GetDirectoryName(FileName)) = True Then       'ディレクトリ存在確認
            If Directory.GetFiles(IO.Path.GetDirectoryName(FileName), IO.Path.GetFileName(FileName), SearchOption.AllDirectories).Length > 0 Then   'ファイル存在確認
                Call Microsoft.VisualBasic.FileSystem.Kill(FileName)
            End If
        End If
    End Sub
    Function YLBound(ByVal OrgArr() As String) As Integer
        Return OrgArr.GetLowerBound(0)
    End Function
    Function YLCase(ByVal OrgStr As String) As String
        Return OrgStr.ToLower
    End Function
    Function YLeft(ByVal OrgStr As String, ByVal iLength As Integer) As String
        If OrgStr <> Nothing AndAlso iLength <= OrgStr.Length Then
            Return OrgStr.Substring(0, iLength)
        End If
        Return OrgStr
    End Function
    Function YLeftB(ByVal OrgStr As String, ByVal iLength As Integer) As String
        Dim hEncode As Encoding = Encoding.GetEncoding("Shift_JIS")
        Dim btBytes As Byte() = hEncode.GetBytes(OrgStr)
        If iLength <= btBytes.Length Then
            Return hEncode.GetString(btBytes, 0, iLength)
        End If
        Return OrgStr
    End Function
    Function YLen(ByVal OrgStr As String) As Integer
        If OrgStr <> Nothing Then
            Return OrgStr.Length
        End If
        Return 0
    End Function
    Function YLenB(ByVal OrgStr As String) As Integer
        If OrgStr <> Nothing Then
            Return Encoding.GetEncoding("Shift_JIS").GetByteCount(OrgStr)
        End If
        Return 0
    End Function
    Function YLng(ByVal num As Double) As Long
        Return Convert.ToInt64(Math.Floor(num))
    End Function
    Function YLog(ByVal num As Double) As Double
        Return Math.Log(num)
    End Function
    Function YMsgBox(ByVal Msg As String, ByVal Button As MessageBoxButtons, ByVal Header As String, Optional ByVal Mark As MessageBoxIcon = MessageBoxIcon.None) As DialogResult
        Dim msgRes As DialogResult = MessageBox.Show(Msg, Header, Button, Mark)
        Return msgRes
    End Function
    Function YMid(ByVal OrgStr As String, ByVal iStart As Integer, Optional ByVal iLength As Integer = -1) As String
        If OrgStr <> Nothing AndAlso iStart <= OrgStr.Length Then
            If iLength = -1 Then iLength = OrgStr.Length     '長さ指定の無い場合は、文字列の長さとする
            If iStart + iLength - 1 <= OrgStr.Length Then
                Return OrgStr.Substring(iStart - 1, iLength)
            End If
            Return OrgStr.Substring(iStart - 1)
        End If
        Return String.Empty
    End Function
    Function YMidB(ByVal OrgStr As String, ByVal iStart As Integer, Optional ByVal iLength As Integer = 2147483647) As String
        Dim hEncode As Encoding = Encoding.GetEncoding("Shift_JIS")
        Dim btBytes As Byte() = hEncode.GetBytes(OrgStr)
        If iStart <= btBytes.Length Then
            If (btBytes.Length - iStart) < iLength Then
                iLength = btBytes.Length - iStart + 1
            End If
            Return hEncode.GetString(btBytes, iStart - 1, iLength)
        End If
        Return String.Empty
    End Function
    Function YNow(ByVal sType As Integer) As Date   'オーバーロード　　引数に(0)を指定するとDate型 *苦肉の策*
        If sType = 0 Then

        End If
        Return System.DateTime.Now
    End Function
    Function YNow() As String                       'オーバーロード　　引数を指定しないとString型
        Return System.DateTime.Now.ToString
    End Function
    Function YReplace(ByVal OrgStr As String, ByVal oldValue As String, ByVal newValue As String) As String
        Return OrgStr.Replace(oldValue, newValue)
    End Function
    Function YRight(ByVal OrgStr As String, ByVal iLength As Integer) As String
        If OrgStr <> Nothing AndAlso iLength <= OrgStr.Length Then
            Return OrgStr.Substring(OrgStr.Length - iLength)
        End If
        Return OrgStr
    End Function
    Function YRightB(ByVal OrgStr As String, ByVal iLength As Integer) As String
        Dim hEncode As Encoding = Encoding.GetEncoding("Shift_JIS")
        Dim btBytes As Byte() = hEncode.GetBytes(OrgStr)
        If iLength <= btBytes.Length Then
            Return hEncode.GetString(btBytes, btBytes.Length - iLength, iLength)
        End If
        Return OrgStr
    End Function
    Function YSgn(ByVal num As Double) As Integer
        Return Math.Sign(num)
    End Function
    ''' <summary>
    ''' 待ち
    ''' </summary>
    ''' <param name="setTime">（ミリ秒）</param>
    Sub YSleep(ByVal setTime As Integer)
        System.Threading.Thread.Sleep(setTime)
    End Sub
    Function YSpace(ByVal Num As Integer) As String
        Return " ".PadLeft(Num, " "c)
    End Function
    Function YSplit(ByVal OrgStr As String, ByVal Delim As Object) As String()
        Dim sp(0) As String
        Dim Temp() As String
        If Delim Is Nothing Then
            Temp = OrgStr.Split(CType(Delim, Char))
        Else
            sp(0) = CStr(Delim)
            Temp = OrgStr.Split(sp, StringSplitOptions.None)
        End If
        Return Temp
    End Function
    Function YStr(ByVal OrgInt As Double) As String
        Return " " & Convert.ToString(OrgInt)
    End Function
    Function YStr(ByVal OrgObj As Object) As String
        Return " " & Convert.ToString(OrgObj)
    End Function
    Function YStr(ByVal OrgStr As String) As String
        Return " " & Convert.ToString(OrgStr)
    End Function
    Function YString(ByVal iNum As Integer, ByVal OrgStr As String) As String
        Dim Buf As String = ""
        For i As Integer = 1 To iNum
            Buf &= OrgStr
        Next
        Return Buf
    End Function
    Function YTime() As String
        Return System.DateTime.Now.ToString("HH:mm:ss")
    End Function
    Function YTimer() As Double
        Return Convert.ToDouble(System.DateTime.Now.TimeOfDay.TotalSeconds.ToString)
    End Function
    Function YTrim(ByVal OrgStr As String) As String
        If OrgStr Is Nothing Then
            Return ""
        Else
            Return OrgStr.Trim
        End If
    End Function
    Function YUBound(ByVal OrgArr() As String) As Integer
        Return OrgArr.GetUpperBound(0)
    End Function
    Function YUCase(ByVal OrgStr As String) As String
        Return OrgStr.ToUpper
    End Function
    Function YVal(ByVal OrgStr As String) As Double
        If OrgStr = "" Then
            Return 0
        Else
            Return CDbl(OrgStr)
        End If
    End Function
    Function YValD(ByVal OrgStr As String) As Double
        If OrgStr = "" Then
            Return 0
        Else
            Return CDbl(OrgStr)
        End If
    End Function
    Function YValS(ByVal OrgStr As String) As Single
        If OrgStr = "" Then
            Return 0
        Else
            Return CShort(OrgStr)
        End If
    End Function
    Function YValL(ByVal OrgStr As String) As Long
        If OrgStr = "" Then
            Return 0
        Else
            Return CLng(OrgStr)
        End If
    End Function
    Function YValI(ByVal OrgStr As String) As Integer
        If OrgStr = "" Then
            Return 0
        Else
            Return CInt(OrgStr)
        End If
    End Function
    Function YvbCr() As String
        Return Microsoft.VisualBasic.ControlChars.Cr
    End Function
    Function YvbCrLf() As String
        Return Environment.NewLine
    End Function
    Function YvbLf() As String
        Return Microsoft.VisualBasic.ControlChars.Lf
    End Function
    Function YvbNullString() As String
        Return Nothing
    End Function
    Function YvbSunday() As Integer
        Return DayOfWeek.Sunday + 1
    End Function
    Function YvbMonday() As Integer
        Return DayOfWeek.Monday + 1
    End Function
    Function YvbTuesday() As Integer
        Return DayOfWeek.Tuesday + 1
    End Function
    Function YvbWednesday() As Integer
        Return DayOfWeek.Wednesday + 1
    End Function
    Function YvbThursday() As Integer
        Return DayOfWeek.Thursday + 1
    End Function
    Function YvbFriday() As Integer
        Return DayOfWeek.Friday + 1
    End Function
    Function YvbSaturday() As Integer
        Return DayOfWeek.Saturday + 1
    End Function
    Function YWeekday(Dt As String) As Integer
        Dim WDt As Date = Date.Parse(Dt)
        Return "日月火水木金土".IndexOf(WDt.ToString("ddd")) + 1
    End Function
End Module
