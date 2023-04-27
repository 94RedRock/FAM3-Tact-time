<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainView
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainView))
        Me.RibbonButton2 = New System.Windows.Forms.RibbonButton()
        Me.RibbonButton3 = New System.Windows.Forms.RibbonButton()
        Me.ProductTime = New System.Windows.Forms.RibbonTab()
        Me.RibbonTab1 = New System.Windows.Forms.RibbonTab()
        Me.RibbonPanel2 = New System.Windows.Forms.RibbonPanel()
        Me.RibbonButton4 = New System.Windows.Forms.RibbonButton()
        Me.RibbonButton5 = New System.Windows.Forms.RibbonButton()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.RibbonTab2 = New System.Windows.Forms.RibbonTab()
        Me.UserControl1 = New PTPP.userControl() '초기화에 들어가는 부분이 맞는가?, prg에 기본적으로 필요한 부분일수도 
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'RibbonButton2
        '
        Me.RibbonButton2.Image = CType(resources.GetObject("RibbonButton2.Image"), System.Drawing.Image)
        Me.RibbonButton2.LargeImage = CType(resources.GetObject("RibbonButton2.LargeImage"), System.Drawing.Image)
        Me.RibbonButton2.Name = "RibbonButton2"
        Me.RibbonButton2.SmallImage = CType(resources.GetObject("RibbonButton2.SmallImage"), System.Drawing.Image)
        Me.RibbonButton2.Text = "RibbonButton2"
        '
        'RibbonButton3
        '
        Me.RibbonButton3.Image = CType(resources.GetObject("RibbonButton3.Image"), System.Drawing.Image)
        Me.RibbonButton3.LargeImage = CType(resources.GetObject("RibbonButton3.LargeImage"), System.Drawing.Image)
        Me.RibbonButton3.Name = "RibbonButton3"
        Me.RibbonButton3.SmallImage = CType(resources.GetObject("RibbonButton3.SmallImage"), System.Drawing.Image)
        Me.RibbonButton3.Text = "RibbonButton3"
        '
        'ProductTime
        '
        Me.ProductTime.Name = "ProductTime"
        Me.ProductTime.Text = "ProductTime"
        '
        'RibbonTab1
        '
        Me.RibbonTab1.Name = "RibbonTab1"
        Me.RibbonTab1.Panels.Add(Me.RibbonPanel2)
        Me.RibbonTab1.Text = "Management"
        '
        'RibbonPanel2
        '
        Me.RibbonPanel2.Items.Add(Me.RibbonButton4)
        Me.RibbonPanel2.Items.Add(Me.RibbonButton5)
        Me.RibbonPanel2.Name = "RibbonPanel2"
        Me.RibbonPanel2.Text = "Menu"
        '
        'RibbonButton4
        '
        Me.RibbonButton4.Image = CType(resources.GetObject("RibbonButton4.Image"), System.Drawing.Image)
        Me.RibbonButton4.LargeImage = CType(resources.GetObject("RibbonButton4.LargeImage"), System.Drawing.Image)
        Me.RibbonButton4.Name = "RibbonButton4"
        Me.RibbonButton4.SmallImage = CType(resources.GetObject("RibbonButton4.SmallImage"), System.Drawing.Image)
        Me.RibbonButton4.Text = "RibbonButton4"
        '
        'RibbonButton5
        '
        Me.RibbonButton5.Image = CType(resources.GetObject("RibbonButton5.Image"), System.Drawing.Image)
        Me.RibbonButton5.LargeImage = CType(resources.GetObject("RibbonButton5.LargeImage"), System.Drawing.Image)
        Me.RibbonButton5.Name = "RibbonButton5"
        Me.RibbonButton5.SmallImage = CType(resources.GetObject("RibbonButton5.SmallImage"), System.Drawing.Image)
        Me.RibbonButton5.Text = "RibbonButton5"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.UserControl1, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1578, 829)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'RibbonTab2
        '
        Me.RibbonTab2.Name = "RibbonTab2"
        Me.RibbonTab2.Text = "RibbonTab2"
        '
        'UserControl1
        '
        Me.UserControl1.Location = New System.Drawing.Point(3, 3)
        Me.UserControl1.Name = "UserControl1"
        Me.UserControl1.Size = New System.Drawing.Size(1546, 823)
        Me.UserControl1.TabIndex = 0
        '
        'MainView
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(1578, 829)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.KeyPreview = True
        Me.Name = "MainView"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FAM3 계획공수 계산 프로그램"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ProductTime As RibbonTab
    Friend WithEvents RibbonButton2 As RibbonButton
    Friend WithEvents RibbonButton3 As RibbonButton
    Friend WithEvents RibbonTab1 As RibbonTab
    Friend WithEvents RibbonPanel2 As RibbonPanel
    Friend WithEvents RibbonButton4 As RibbonButton
    Friend WithEvents RibbonButton5 As RibbonButton
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents RibbonTab2 As RibbonTab
    Friend WithEvents UserControl1 As userControl
End Class
