Imports System.Threading
Imports System.IO

Public Class UserManual
    Public Sub New()

        ' 디자이너에서 이 호출이 필요합니다.
        InitializeComponent()


    End Sub

    Public Sub 계획공수계산()
        AxAcroPDF1.src = Directory.GetCurrentDirectory() + "\FAM3계획공수계산프로그램메뉴얼V2(계획공수계산).pdf"
    End Sub

    Public Sub 마스터데이터관리()
        AxAcroPDF1.src = Directory.GetCurrentDirectory() + "\FAM3계획공수계산프로그램메뉴얼V2(마스터데이터관리).pdf"
    End Sub
End Class