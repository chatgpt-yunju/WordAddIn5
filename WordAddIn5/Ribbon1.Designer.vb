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
        Me.components = New System.ComponentModel.Container()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Box1 = Me.Factory.CreateRibbonBox
        Me.EditBox1 = Me.Factory.CreateRibbonEditBox
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Box2 = Me.Factory.CreateRibbonBox
        Me.EditBox2 = Me.Factory.CreateRibbonEditBox
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.作者 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button22 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.Group7.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Box1.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Box2.SuspendLayout()
        Me.Group5.SuspendLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Groups.Add(Me.Group7)
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "图片处理"
        Me.Tab1.Name = "Tab1"
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.Button10)
        Me.Group6.Items.Add(Me.作者)
        Me.Group6.Items.Add(Me.Button11)
        Me.Group6.Items.Add(Me.Button13)
        Me.Group6.Items.Add(Me.Button14)
        Me.Group6.Label = "校对"
        Me.Group6.Name = "Group6"
        '
        'Group7
        '
        Me.Group7.Items.Add(Me.Button12)
        Me.Group7.Items.Add(Me.Button18)
        Me.Group7.Items.Add(Me.Button19)
        Me.Group7.Items.Add(Me.Button20)
        Me.Group7.Items.Add(Me.Button21)
        Me.Group7.Items.Add(Me.Button23)
        Me.Group7.Items.Add(Me.Button15)
        Me.Group7.Items.Add(Me.Button24)
        Me.Group7.Items.Add(Me.Button16)
        Me.Group7.Items.Add(Me.Button25)
        Me.Group7.Label = "修正"
        Me.Group7.Name = "Group7"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button17)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Label = "设置"
        Me.Group1.Name = "Group1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Label = "批量更换"
        Me.Group2.Name = "Group2"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button3)
        Me.Group3.Items.Add(Me.Box1)
        Me.Group3.Label = "单张更换"
        Me.Group3.Name = "Group3"
        '
        'Box1
        '
        Me.Box1.Items.Add(Me.EditBox1)
        Me.Box1.Items.Add(Me.Button6)
        Me.Box1.Items.Add(Me.Button5)
        Me.Box1.Name = "Box1"
        '
        'EditBox1
        '
        Me.EditBox1.Label = "x  ="
        Me.EditBox1.MaxLength = 3
        Me.EditBox1.Name = "EditBox1"
        Me.EditBox1.SizeString = "10"
        Me.EditBox1.Text = Nothing
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button4)
        Me.Group4.Items.Add(Me.Box2)
        Me.Group4.Label = "插入图片"
        Me.Group4.Name = "Group4"
        '
        'Box2
        '
        Me.Box2.Items.Add(Me.EditBox2)
        Me.Box2.Items.Add(Me.Button8)
        Me.Box2.Items.Add(Me.Button7)
        Me.Box2.Name = "Box2"
        '
        'EditBox2
        '
        Me.EditBox2.Label = "y  ="
        Me.EditBox2.MaxLength = 3
        Me.EditBox2.Name = "EditBox2"
        Me.EditBox2.SizeString = "10"
        Me.EditBox2.Text = Nothing
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.Button9)
        Me.Group5.Label = "帮助"
        Me.Group5.Name = "Group5"
        '
        'Button10
        '
        Me.Button10.Label = "标题"
        Me.Button10.Name = "Button10"
        '
        '作者
        '
        Me.作者.Label = "作者"
        Me.作者.Name = "作者"
        '
        'Button11
        '
        Me.Button11.Label = "摘要"
        Me.Button11.Name = "Button11"
        '
        'Button13
        '
        Me.Button13.Label = "关键词"
        Me.Button13.Name = "Button13"
        '
        'Button14
        '
        Me.Button14.Label = "参考文献"
        Me.Button14.Name = "Button14"
        '
        'Button12
        '
        Me.Button12.Label = "标题"
        Me.Button12.Name = "Button12"
        '
        'Button18
        '
        Me.Button18.Label = "副标题"
        Me.Button18.Name = "Button18"
        '
        'Button19
        '
        Me.Button19.Label = "一级标题"
        Me.Button19.Name = "Button19"
        '
        'Button20
        '
        Me.Button20.Label = "二级标题"
        Me.Button20.Name = "Button20"
        '
        'Button21
        '
        Me.Button21.Label = "三级标题"
        Me.Button21.Name = "Button21"
        '
        'Button23
        '
        Me.Button23.Label = "四级标题"
        Me.Button23.Name = "Button23"
        '
        'Button15
        '
        Me.Button15.Label = "参考文献"
        Me.Button15.Name = "Button15"
        '
        'Button24
        '
        Me.Button24.Label = "正文"
        Me.Button24.Name = "Button24"
        '
        'Button16
        '
        Me.Button16.Label = "全文（Input）"
        Me.Button16.Name = "Button16"
        '
        'Button25
        '
        Me.Button25.Label = "全文（Input）"
        Me.Button25.Name = "Button25"
        '
        'Button17
        '
        Me.Button17.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button17.Image = Global.WordAddIn5.My.Resources.Resources.help
        Me.Button17.Label = "设置文本格式"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = Global.WordAddIn5.My.Resources.Resources.setings
        Me.Button1.Label = "设置图片路径"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Image = Global.WordAddIn5.My.Resources.Resources.images
        Me.Button2.Label = "更换全部图片"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Button3
        '
        Me.Button3.Image = Global.WordAddIn5.My.Resources.Resources.image
        Me.Button3.Label = "更换第x张图片"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Button6
        '
        Me.Button6.Label = "x++"
        Me.Button6.Name = "Button6"
        Me.Button6.ShowImage = True
        '
        'Button5
        '
        Me.Button5.Label = "x--"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Button4
        '
        Me.Button4.Image = Global.WordAddIn5.My.Resources.Resources.image
        Me.Button4.Label = "在光标处插入第y张图片"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Button8
        '
        Me.Button8.Label = "y++"
        Me.Button8.Name = "Button8"
        Me.Button8.ShowImage = True
        '
        'Button7
        '
        Me.Button7.Label = "y--"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        '
        'Button9
        '
        Me.Button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button9.Image = Global.WordAddIn5.My.Resources.Resources.help
        Me.Button9.Label = "检查更新"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Button22
        '
        Me.Button22.Label = "二级标题"
        Me.Button22.Name = "Button22"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Box1.ResumeLayout(False)
        Me.Box1.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Box2.ResumeLayout(False)
        Me.Box2.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents EditBox1 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditBox2 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Box1 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Box2 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents 作者 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BindingSource1 As Windows.Forms.BindingSource
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
