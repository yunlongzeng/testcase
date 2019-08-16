Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
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
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Projects.SuspendLayout()
        Me.Group4.SuspendLayout()
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
        'Group5
        '
        Me.Group5.Items.Add(Me.Button12)
        Me.Group5.Items.Add(Me.Button13)
        Me.Group5.Label = "Time"
        Me.Group5.Name = "Group5"
        '
        'Button12
        '
        Me.Button12.Label = "Format"
        Me.Button12.Name = "Button12"
        '
        'Button13
        '
        Me.Button13.Label = "CopyData"
        Me.Button13.Name = "Button13"
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
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
