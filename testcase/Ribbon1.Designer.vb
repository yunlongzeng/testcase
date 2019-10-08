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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl4 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl5 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl6 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl7 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl8 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl9 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl10 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl11 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl12 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl13 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl14 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl15 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl16 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl17 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.EditBox1 = Me.Factory.CreateRibbonEditBox
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Projects = Me.Factory.CreateRibbonGroup
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.work = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.EditBox2 = Me.Factory.CreateRibbonEditBox
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.ComboBox1 = Me.Factory.CreateRibbonComboBox
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.EditBox3 = Me.Factory.CreateRibbonEditBox
        Me.EditBox4 = Me.Factory.CreateRibbonEditBox
        Me.ComboBox2 = Me.Factory.CreateRibbonComboBox
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.ComboBox3 = Me.Factory.CreateRibbonComboBox
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Projects.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group7.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Projects)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group7)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "Sort"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Button3
        '
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Label = "Unhide"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Items.Add(Me.EditBox1)
        Me.Group2.Items.Add(Me.Button4)
        Me.Group2.Label = "Group2"
        Me.Group2.Name = "Group2"
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "fill_steps"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'EditBox1
        '
        Me.EditBox1.Label = "filename"
        Me.EditBox1.Name = "EditBox1"
        Me.EditBox1.Text = Nothing
        '
        'Button4
        '
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Label = "SaveAs"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button5)
        Me.Group3.Items.Add(Me.Button6)
        Me.Group3.Label = "database"
        Me.Group3.Name = "Group3"
        '
        'Button5
        '
        Me.Button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.Label = "ExcelData"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Button6
        '
        Me.Button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.Label = "UpdateAccess"
        Me.Button6.Name = "Button6"
        Me.Button6.ShowImage = True
        '
        'Projects
        '
        Me.Projects.Items.Add(Me.Button7)
        Me.Projects.Items.Add(Me.Button8)
        Me.Projects.Items.Add(Me.Button9)
        Me.Projects.Label = "Projects"
        Me.Projects.Name = "Projects"
        '
        'Button7
        '
        Me.Button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button7.Image = CType(resources.GetObject("Button7.Image"), System.Drawing.Image)
        Me.Button7.Label = "Sort"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        '
        'Button8
        '
        Me.Button8.Label = "Process"
        Me.Button8.Name = "Button8"
        '
        'Button9
        '
        Me.Button9.Label = "Color"
        Me.Button9.Name = "Button9"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.work)
        Me.Group4.Items.Add(Me.Button10)
        Me.Group4.Items.Add(Me.Button11)
        Me.Group4.Label = "Group4"
        Me.Group4.Name = "Group4"
        '
        'work
        '
        Me.work.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.work.Image = CType(resources.GetObject("work.Image"), System.Drawing.Image)
        Me.work.Label = "Work"
        Me.work.Name = "work"
        Me.work.ShowImage = True
        '
        'Button10
        '
        Me.Button10.Label = "Summery"
        Me.Button10.Name = "Button10"
        '
        'Button11
        '
        Me.Button11.Label = "Line"
        Me.Button11.Name = "Button11"
        '
        'Group7
        '
        Me.Group7.Items.Add(Me.EditBox2)
        Me.Group7.Items.Add(Me.Button17)
        Me.Group7.Items.Add(Me.Button19)
        Me.Group7.Items.Add(Me.Button16)
        Me.Group7.Items.Add(Me.Button18)
        Me.Group7.Items.Add(Me.ComboBox1)
        Me.Group7.Items.Add(Me.Button12)
        Me.Group7.Items.Add(Me.Button13)
        Me.Group7.Items.Add(Me.Button14)
        Me.Group7.Label = "Work"
        Me.Group7.Name = "Group7"
        '
        'EditBox2
        '
        Me.EditBox2.Label = "ite_id"
        Me.EditBox2.Name = "EditBox2"
        Me.EditBox2.Text = Nothing
        '
        'Button17
        '
        Me.Button17.Label = "获取所有项目"
        Me.Button17.Name = "Button17"
        '
        'Button19
        '
        Me.Button19.Label = "工时"
        Me.Button19.Name = "Button19"
        '
        'Button16
        '
        Me.Button16.Label = "单个迭代情况"
        Me.Button16.Name = "Button16"
        '
        'Button18
        '
        Me.Button18.Label = "获取所有信息"
        Me.Button18.Name = "Button18"
        '
        'ComboBox1
        '
        RibbonDropDownItemImpl1.Label = "All"
        RibbonDropDownItemImpl2.Label = "测试"
        RibbonDropDownItemImpl3.Label = "前端"
        RibbonDropDownItemImpl4.Label = "全渠道"
        RibbonDropDownItemImpl5.Label = "开放平台"
        RibbonDropDownItemImpl6.Label = "CRM"
        RibbonDropDownItemImpl7.Label = "产品"
        RibbonDropDownItemImpl8.Label = "UI"
        RibbonDropDownItemImpl9.Label = "没有部门"
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl1)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl2)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl3)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl4)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl5)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl6)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl7)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl8)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl9)
        Me.ComboBox1.Label = "部门"
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Text = Nothing
        '
        'Button12
        '
        Me.Button12.Label = "日报"
        Me.Button12.Name = "Button12"
        '
        'Button13
        '
        Me.Button13.Label = "周报"
        Me.Button13.Name = "Button13"
        '
        'Button14
        '
        Me.Button14.Label = "灰度表"
        Me.Button14.Name = "Button14"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.EditBox3)
        Me.Group5.Items.Add(Me.EditBox4)
        Me.Group5.Items.Add(Me.ComboBox2)
        Me.Group5.Items.Add(Me.ComboBox3)
        Me.Group5.Items.Add(Me.Button15)
        Me.Group5.Label = "Story"
        Me.Group5.Name = "Group5"
        '
        'EditBox3
        '
        Me.EditBox3.Label = "CreateTime"
        Me.EditBox3.Name = "EditBox3"
        Me.EditBox3.Text = Nothing
        '
        'EditBox4
        '
        Me.EditBox4.Label = "FinishTime"
        Me.EditBox4.Name = "EditBox4"
        Me.EditBox4.Text = Nothing
        '
        'ComboBox2
        '
        RibbonDropDownItemImpl10.Label = "待开发"
        RibbonDropDownItemImpl11.Label = "开发中"
        RibbonDropDownItemImpl12.Label = "已实现"
        RibbonDropDownItemImpl13.Label = "挂起"
        RibbonDropDownItemImpl14.Label = "拒绝"
        Me.ComboBox2.Items.Add(RibbonDropDownItemImpl10)
        Me.ComboBox2.Items.Add(RibbonDropDownItemImpl11)
        Me.ComboBox2.Items.Add(RibbonDropDownItemImpl12)
        Me.ComboBox2.Items.Add(RibbonDropDownItemImpl13)
        Me.ComboBox2.Items.Add(RibbonDropDownItemImpl14)
        Me.ComboBox2.Label = "Status"
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Text = Nothing
        '
        'Button15
        '
        Me.Button15.Label = "Search"
        Me.Button15.Name = "Button15"
        '
        'ComboBox3
        '
        RibbonDropDownItemImpl15.Label = "商城"
        RibbonDropDownItemImpl16.Label = "分销"
        RibbonDropDownItemImpl17.Label = "微信"
        Me.ComboBox3.Items.Add(RibbonDropDownItemImpl15)
        Me.ComboBox3.Items.Add(RibbonDropDownItemImpl16)
        Me.ComboBox3.Items.Add(RibbonDropDownItemImpl17)
        Me.ComboBox3.Label = "Category"
        Me.ComboBox3.Name = "ComboBox3"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Projects.ResumeLayout(False)
        Me.Projects.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents EditBox1 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Projects As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents work As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents EditBox2 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ComboBox1 As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents EditBox3 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ComboBox2 As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditBox4 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ComboBox3 As Microsoft.Office.Tools.Ribbon.RibbonComboBox
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
