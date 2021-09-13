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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.Group8 = Me.Factory.CreateRibbonGroup
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Box1 = Me.Factory.CreateRibbonBox
        Me.EditBox1 = Me.Factory.CreateRibbonEditBox
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Box2 = Me.Factory.CreateRibbonBox
        Me.EditBox2 = Me.Factory.CreateRibbonEditBox
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Button28 = Me.Factory.CreateRibbonButton
        Me.Button41 = Me.Factory.CreateRibbonButton
        Me.Button42 = Me.Factory.CreateRibbonButton
        Me.Button43 = Me.Factory.CreateRibbonButton
        Me.Button44 = Me.Factory.CreateRibbonButton
        Me.Button47 = Me.Factory.CreateRibbonButton
        Me.Button45 = Me.Factory.CreateRibbonButton
        Me.Button46 = Me.Factory.CreateRibbonButton
        Me.Button29 = Me.Factory.CreateRibbonButton
        Me.Button30 = Me.Factory.CreateRibbonButton
        Me.Button31 = Me.Factory.CreateRibbonButton
        Me.Button32 = Me.Factory.CreateRibbonButton
        Me.Button33 = Me.Factory.CreateRibbonButton
        Me.Button34 = Me.Factory.CreateRibbonButton
        Me.Button35 = Me.Factory.CreateRibbonButton
        Me.Button36 = Me.Factory.CreateRibbonButton
        Me.Button37 = Me.Factory.CreateRibbonButton
        Me.Button38 = Me.Factory.CreateRibbonButton
        Me.Button39 = Me.Factory.CreateRibbonButton
        Me.Button40 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Button27 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Button26 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
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
        Me.Group1.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.Group8.SuspendLayout()
        Me.Group7.SuspendLayout()
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
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Groups.Add(Me.Group8)
        Me.Tab1.Groups.Add(Me.Group7)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "图文排版"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button17)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Label = "设置"
        Me.Group1.Name = "Group1"
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.Button10)
        Me.Group6.Items.Add(Me.Button11)
        Me.Group6.Items.Add(Me.Button13)
        Me.Group6.Items.Add(Me.Button14)
        Me.Group6.Items.Add(Me.Button28)
        Me.Group6.Items.Add(Me.Button41)
        Me.Group6.Items.Add(Me.Button42)
        Me.Group6.Items.Add(Me.Button43)
        Me.Group6.Items.Add(Me.Button44)
        Me.Group6.Items.Add(Me.Button47)
        Me.Group6.Items.Add(Me.Button45)
        Me.Group6.Items.Add(Me.Button46)
        Me.Group6.Label = "插入论文要素"
        Me.Group6.Name = "Group6"
        '
        'Group8
        '
        Me.Group8.Items.Add(Me.Button29)
        Me.Group8.Items.Add(Me.Button30)
        Me.Group8.Items.Add(Me.Button31)
        Me.Group8.Items.Add(Me.Button32)
        Me.Group8.Items.Add(Me.Button33)
        Me.Group8.Items.Add(Me.Button34)
        Me.Group8.Items.Add(Me.Button35)
        Me.Group8.Items.Add(Me.Button36)
        Me.Group8.Items.Add(Me.Button37)
        Me.Group8.Items.Add(Me.Button38)
        Me.Group8.Items.Add(Me.Button39)
        Me.Group8.Items.Add(Me.Button40)
        Me.Group8.Label = "段落样式"
        Me.Group8.Name = "Group8"
        '
        'Group7
        '
        Me.Group7.Items.Add(Me.Button12)
        Me.Group7.Items.Add(Me.Button23)
        Me.Group7.Items.Add(Me.Button27)
        Me.Group7.Items.Add(Me.Button18)
        Me.Group7.Items.Add(Me.Button19)
        Me.Group7.Items.Add(Me.Button20)
        Me.Group7.Items.Add(Me.Button21)
        Me.Group7.Items.Add(Me.Button24)
        Me.Group7.Items.Add(Me.Button15)
        Me.Group7.Items.Add(Me.Button26)
        Me.Group7.Items.Add(Me.Button25)
        Me.Group7.Items.Add(Me.Button16)
        Me.Group7.Label = "自动修正"
        Me.Group7.Name = "Group7"
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
        'Button17
        '
        Me.Button17.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button17.Image = CType(resources.GetObject("Button17.Image"), System.Drawing.Image)
        Me.Button17.Label = "设置文本格式"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = Global.WordAddIn5.My.Resources.Resources.image
        Me.Button1.Label = "设置图片路径"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Button10
        '
        Me.Button10.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button10.Label = "标题"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
        '
        'Button11
        '
        Me.Button11.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button11.Label = "摘要"
        Me.Button11.Name = "Button11"
        Me.Button11.ShowImage = True
        '
        'Button13
        '
        Me.Button13.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button13.Label = "关键词"
        Me.Button13.Name = "Button13"
        Me.Button13.ShowImage = True
        '
        'Button14
        '
        Me.Button14.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button14.Label = "作者信息"
        Me.Button14.Name = "Button14"
        Me.Button14.ShowImage = True
        '
        'Button28
        '
        Me.Button28.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button28.Label = "一级标题"
        Me.Button28.Name = "Button28"
        Me.Button28.ShowImage = True
        '
        'Button41
        '
        Me.Button41.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button41.Label = "二级标题"
        Me.Button41.Name = "Button41"
        Me.Button41.ShowImage = True
        '
        'Button42
        '
        Me.Button42.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button42.Label = "三级标题"
        Me.Button42.Name = "Button42"
        Me.Button42.ShowImage = True
        '
        'Button43
        '
        Me.Button43.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button43.Label = "正文"
        Me.Button43.Name = "Button43"
        Me.Button43.ShowImage = True
        '
        'Button44
        '
        Me.Button44.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button44.Label = "参考文献"
        Me.Button44.Name = "Button44"
        Me.Button44.ShowImage = True
        '
        'Button47
        '
        Me.Button47.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button47.Label = "图注"
        Me.Button47.Name = "Button47"
        Me.Button47.ShowImage = True
        '
        'Button45
        '
        Me.Button45.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button45.Label = "公式"
        Me.Button45.Name = "Button45"
        Me.Button45.ShowImage = True
        '
        'Button46
        '
        Me.Button46.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button46.Label = "图表"
        Me.Button46.Name = "Button46"
        Me.Button46.ShowImage = True
        '
        'Button29
        '
        Me.Button29.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button29.Label = "标题"
        Me.Button29.Name = "Button29"
        Me.Button29.ShowImage = True
        '
        'Button30
        '
        Me.Button30.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button30.Label = "摘要"
        Me.Button30.Name = "Button30"
        Me.Button30.ShowImage = True
        '
        'Button31
        '
        Me.Button31.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button31.Label = "关键词"
        Me.Button31.Name = "Button31"
        Me.Button31.ShowImage = True
        '
        'Button32
        '
        Me.Button32.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button32.Label = "作者信息"
        Me.Button32.Name = "Button32"
        Me.Button32.ShowImage = True
        '
        'Button33
        '
        Me.Button33.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button33.Label = "一级标题"
        Me.Button33.Name = "Button33"
        Me.Button33.ShowImage = True
        '
        'Button34
        '
        Me.Button34.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button34.Label = "二级标题"
        Me.Button34.Name = "Button34"
        Me.Button34.ShowImage = True
        '
        'Button35
        '
        Me.Button35.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button35.Label = "三级标题"
        Me.Button35.Name = "Button35"
        Me.Button35.ShowImage = True
        '
        'Button36
        '
        Me.Button36.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button36.Label = "正文"
        Me.Button36.Name = "Button36"
        Me.Button36.ShowImage = True
        '
        'Button37
        '
        Me.Button37.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button37.Label = "参考文献"
        Me.Button37.Name = "Button37"
        Me.Button37.ShowImage = True
        '
        'Button38
        '
        Me.Button38.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button38.Label = "图注"
        Me.Button38.Name = "Button38"
        Me.Button38.ShowImage = True
        '
        'Button39
        '
        Me.Button39.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button39.Label = "公式"
        Me.Button39.Name = "Button39"
        Me.Button39.ShowImage = True
        '
        'Button40
        '
        Me.Button40.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button40.Label = "图表"
        Me.Button40.Name = "Button40"
        Me.Button40.ShowImage = True
        '
        'Button12
        '
        Me.Button12.Image = CType(resources.GetObject("Button12.Image"), System.Drawing.Image)
        Me.Button12.Label = "标题"
        Me.Button12.Name = "Button12"
        Me.Button12.ShowImage = True
        '
        'Button23
        '
        Me.Button23.Image = CType(resources.GetObject("Button23.Image"), System.Drawing.Image)
        Me.Button23.Label = "摘要"
        Me.Button23.Name = "Button23"
        Me.Button23.ShowImage = True
        '
        'Button27
        '
        Me.Button27.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button27.Label = "关键词"
        Me.Button27.Name = "Button27"
        Me.Button27.ShowImage = True
        '
        'Button18
        '
        Me.Button18.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button18.Label = "作者信息"
        Me.Button18.Name = "Button18"
        Me.Button18.ShowImage = True
        '
        'Button19
        '
        Me.Button19.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button19.Label = "一级标题"
        Me.Button19.Name = "Button19"
        Me.Button19.ShowImage = True
        '
        'Button20
        '
        Me.Button20.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button20.Label = "二级标题"
        Me.Button20.Name = "Button20"
        Me.Button20.ShowImage = True
        '
        'Button21
        '
        Me.Button21.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button21.Label = "三级标题"
        Me.Button21.Name = "Button21"
        Me.Button21.ShowImage = True
        '
        'Button24
        '
        Me.Button24.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button24.Label = "正文"
        Me.Button24.Name = "Button24"
        Me.Button24.ShowImage = True
        '
        'Button15
        '
        Me.Button15.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button15.Label = "参考文献"
        Me.Button15.Name = "Button15"
        Me.Button15.ShowImage = True
        '
        'Button25
        '
        Me.Button25.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button25.Label = "图注"
        Me.Button25.Name = "Button25"
        Me.Button25.ShowImage = True
        '
        'Button26
        '
        Me.Button26.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button26.Label = "删空白行"
        Me.Button26.Name = "Button26"
        Me.Button26.ShowImage = True
        '
        'Button16
        '
        Me.Button16.Image = Global.WordAddIn5.My.Resources.Resources.校对
        Me.Button16.Label = "全文"
        Me.Button16.Name = "Button16"
        Me.Button16.ShowImage = True
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
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.Group8.ResumeLayout(False)
        Me.Group8.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
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
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BindingSource1 As Windows.Forms.BindingSource
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button26 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button27 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group8 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button29 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button30 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button31 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button32 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button33 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button34 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button35 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button36 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button37 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button38 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button39 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button40 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button28 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button41 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button42 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button43 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button44 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button47 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button45 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button46 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SaveFileDialog1 As Windows.Forms.SaveFileDialog
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
