Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
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

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl4 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl5 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.DropDown1 = Me.Factory.CreateRibbonDropDown
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Chk1 = Me.Factory.CreateRibbonCheckBox
        Me.Chk2 = Me.Factory.CreateRibbonCheckBox
        Me.Chk3 = Me.Factory.CreateRibbonCheckBox
        Me.Chk4 = Me.Factory.CreateRibbonCheckBox
        Me.Chk5 = Me.Factory.CreateRibbonCheckBox
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "计算书"
        Me.Tab1.Name = "Tab1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Label = "初始化变量"
        Me.Group2.Name = "Group2"
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Image = Global.Excel计算书工具箱1._0.My.Resources.Resources.加载_loading
        Me.Button2.Label = "单元格命名"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.DropDown1)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Button5)
        Me.Group1.Items.Add(Me.Chk1)
        Me.Group1.Items.Add(Me.Chk2)
        Me.Group1.Items.Add(Me.Chk3)
        Me.Group1.Items.Add(Me.Chk4)
        Me.Group1.Items.Add(Me.Chk5)
        Me.Group1.Items.Add(Me.Button9)
        Me.Group1.Items.Add(Me.Button6)
        Me.Group1.Label = "智能单元格"
        Me.Group1.Name = "Group1"
        '
        'DropDown1
        '
        RibbonDropDownItemImpl1.Label = "0"
        RibbonDropDownItemImpl1.Tag = ""
        RibbonDropDownItemImpl2.Label = "1"
        RibbonDropDownItemImpl3.Label = "2"
        RibbonDropDownItemImpl4.Label = "3"
        RibbonDropDownItemImpl5.Label = "4"
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl1)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl2)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl3)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl4)
        Me.DropDown1.Items.Add(RibbonDropDownItemImpl5)
        Me.DropDown1.Label = "ROUND小数位数"
        Me.DropDown1.Name = "DropDown1"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = Global.Excel计算书工具箱1._0.My.Resources.Resources.切换按钮_switch_button
        Me.Button1.Label = "开"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Button5
        '
        Me.Button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button5.Image = Global.Excel计算书工具箱1._0.My.Resources.Resources.切换按钮_switch_button_黑白
        Me.Button5.Label = "关"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Chk1
        '
        Me.Chk1.Checked = True
        Me.Chk1.Label = "变量名->单元格名"
        Me.Chk1.Name = "Chk1"
        '
        'Chk2
        '
        Me.Chk2.Checked = True
        Me.Chk2.Label = "显示用变量表示的公式"
        Me.Chk2.Name = "Chk2"
        '
        'Chk3
        '
        Me.Chk3.Checked = True
        Me.Chk3.Label = "显示用数字表示的公式"
        Me.Chk3.Name = "Chk3"
        '
        'Chk4
        '
        Me.Chk4.Checked = True
        Me.Chk4.Label = "应用ROUND函数"
        Me.Chk4.Name = "Chk4"
        '
        'Chk5
        '
        Me.Chk5.Checked = True
        Me.Chk5.Label = "量纲计算"
        Me.Chk5.Name = "Chk5"
        '
        'Button9
        '
        Me.Button9.Image = Global.Excel计算书工具箱1._0.My.Resources.Resources.搜索_search
        Me.Button9.Label = "可识别的函数"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Button6
        '
        Me.Button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button6.Image = Global.Excel计算书工具箱1._0.My.Resources.Resources.刷新_refresh
        Me.Button6.Label = "刷新"
        Me.Button6.Name = "Button6"
        Me.Button6.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button3)
        Me.Group3.Label = "发送到Word"
        Me.Group3.Name = "Group3"
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.Image = Global.Excel计算书工具箱1._0.My.Resources.Resources.传出3_efferent_three
        Me.Button3.Label = "发送到Word"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button4)
        Me.Group4.Items.Add(Me.Button7)
        Me.Group4.Label = "其他"
        Me.Group4.Name = "Group4"
        '
        'Button4
        '
        Me.Button4.Image = Global.Excel计算书工具箱1._0.My.Resources.Resources.数据库网络节点_database_network_point
        Me.Button4.Label = "打开建标库"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Button7
        '
        Me.Button7.Image = Global.Excel计算书工具箱1._0.My.Resources.Resources.数据库网络节点_database_network_point
        Me.Button7.Label = "打开项目地址"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Chk1 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Chk2 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Chk3 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Chk4 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents DropDown1 As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents Chk5 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
