Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Public Class Form2
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox63_TextChanged(sender As Object, e As EventArgs) Handles TextBoxt3.TextChanged

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox55_TextChanged(sender As Object, e As EventArgs) Handles TextBoxr7.TextChanged

    End Sub

    Private Sub TextBox56_TextChanged(sender As Object, e As EventArgs) Handles TextBoxr6.TextChanged

    End Sub

    Private Sub TextBox57_TextChanged(sender As Object, e As EventArgs) Handles TextBoxr5.TextChanged

    End Sub

    Private Sub TextBox58_TextChanged(sender As Object, e As EventArgs) Handles TextBoxr4.TextChanged

    End Sub

    Private Sub TextBox59_TextChanged(sender As Object, e As EventArgs) Handles TextBoxr3.TextChanged

    End Sub

    Private Sub TextBox60_TextChanged(sender As Object, e As EventArgs) Handles TextBoxr2.TextChanged

    End Sub

    Private Sub TextBox68_TextChanged(sender As Object, e As EventArgs) Handles TextBoxr1.TextChanged

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

    End Sub

    Private Sub TextBox67_TextChanged(sender As Object, e As EventArgs) Handles TextBoxt7.TextChanged

    End Sub

    Private Sub TextBoxf10_TextChanged(sender As Object, e As EventArgs) Handles TextBoxf10.TextChanged

    End Sub

    Private Sub Label39_Click(sender As Object, e As EventArgs) Handles Label39.Click

    End Sub

    Private Sub Label38_Click(sender As Object, e As EventArgs) Handles Label38.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz10.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz9.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz8.TextChanged

    End Sub

    Private Sub Label37_Click(sender As Object, e As EventArgs) Handles Label37.Click

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz7.TextChanged

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz6.TextChanged

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz5.TextChanged

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz4.TextChanged

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz3.TextChanged

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz2.TextChanged

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBoxtz1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form2.ActiveForm.Visible = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '标题
        TextBoxt00.Text = “0”
        TextBoxt0.Text = “0.5”
        TextBoxt1.Text = “0”
        TextBoxt2.Text = “0”
        TextBoxt3.Text = “1”
        TextBoxt4.Text = “1”
        TextBoxt5.Text = “Arial”
        TextBoxt6.Text = “14”
        TextBoxt7.Text = “True”
        '副标题
        TextBoxs00.Text = “0.5”
        TextBoxs0.Text = “0”
        TextBoxs1.Text = “0”
        TextBoxs2.Text = “0”
        TextBoxs3.Text = “1”
        TextBoxs4.Text = “1”
        TextBoxs5.Text = “Arial”
        TextBoxs6.Text = “11”
        TextBoxs7.Text = “False”
        '摘要
        TextBoxAb00.Text = “1.5”
        TextBoxAb0.Text = “0”
        TextBoxAb1.Text = “0”
        TextBoxAb2.Text = “0”
        TextBoxAb3.Text = “1”
        TextBoxAb4.Text = “3”
        TextBoxAb5.Text = “Times New Roman”
        TextBoxAb6.Text = “12”
        TextBoxAb7.Text = “False”
        '关键词
        TextBoxKey00.Text = “1.5”
        TextBoxKey0.Text = “0”
        TextBoxKey1.Text = “0”
        TextBoxKey2.Text = “0”
        TextBoxKey3.Text = “1”
        TextBoxKey4.Text = “3”
        TextBoxKey5.Text = “Arial”
        TextBoxKey6.Text = “11”
        TextBoxKey7.Text = “False”
        TextBoxKey8.Text = “1”
        TextBoxKey9.Text = “8”
        TextBoxKey10.Text = “Keywords”


        '一级标题
        TextBoxf00.Text = “1.5”
        TextBoxf0.Text = “0.5”
        TextBoxf1.Text = “0”
        TextBoxf2.Text = “0”
        TextBoxf3.Text = “1”
        TextBoxf4.Text = “3”
        TextBoxf5.Text = “Times New Roman”
        TextBoxf6.Text = “12”
        TextBoxf7.Text = “True”
        '二级标题
        TextBoxsec00.Text = “0.5”
        TextBoxsec0.Text = “0.5”
        TextBoxsec1.Text = “0”
        TextBoxsec2.Text = “0”
        TextBoxsec3.Text = “1”
        TextBoxsec4.Text = “3”
        TextBoxsec5.Text = “Times New Roman”
        TextBoxsec6.Text = “12”
        TextBoxsec7.Text = “True”
        '三级标题
        TextBoxthd00.Text = “0.5”
        TextBoxthd0.Text = “0.5”
        TextBoxthd1.Text = “0”
        TextBoxthd2.Text = “0”
        TextBoxthd3.Text = “1”
        TextBoxthd4.Text = “3”
        TextBoxthd5.Text = “Times New Roman”
        TextBoxthd6.Text = “12”
        TextBoxthd7.Text = “True”
        '正文
        TextBoxc00.Text = “0”
        TextBoxc0.Text = “0”
        TextBoxc1.Text = “0”
        TextBoxc2.Text = “0.5”
        TextBoxc3.Text = “1”
        TextBoxc4.Text = “3”
        TextBoxc5.Text = “Times New Roman”
        TextBoxc6.Text = “12”
        TextBoxc7.Text = “False”
        '参考文献
        TextBoxr00.Text = “0”
        TextBoxr0.Text = “0”
        TextBoxr1.Text = “0”
        TextBoxr2.Text = “0”
        TextBoxr3.Text = “1”
        TextBoxr4.Text = “3”
        TextBoxr5.Text = “Times New Roman”
        TextBoxr6.Text = “12”
        TextBoxr7.Text = “False”
        '图注
        TextBoxtz00.Text = “0.5”
        TextBoxtz0.Text = “0.5”
        TextBoxtz1.Text = “0”
        TextBoxtz2.Text = “0”
        TextBoxtz3.Text = “1”
        TextBoxtz4.Text = “1”
        TextBoxtz5.Text = “Times New Roman”
        TextBoxtz6.Text = “12”
        TextBoxtz7.Text = “False”
    End Sub

    Private Sub TextBoxt0_TextChanged(sender As Object, e As EventArgs) Handles TextBoxt0.TextChanged

    End Sub

    Private Sub TextBoxc00_TextChanged(sender As Object, e As EventArgs) Handles TextBoxc00.TextChanged

    End Sub

    Private Sub Label43_Click(sender As Object, e As EventArgs) Handles Label43.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '标题
        TextBoxt00.Text = “”
        TextBoxt0.Text = “”
        TextBoxt1.Text = “”
        TextBoxt2.Text = “”
        TextBoxt3.Text = “”
        TextBoxt4.Text = “”
        TextBoxt5.Text = “”
        TextBoxt6.Text = “”
        TextBoxt7.Text = “”
        TextBoxt00.Enabled = False
        TextBoxt0.Enabled = False
        TextBoxt1.Enabled = False
        TextBoxt2.Enabled = False
        TextBoxt3.Enabled = False
        TextBoxt4.Enabled = False
        TextBoxt5.Enabled = False
        TextBoxt6.Enabled = False
        TextBoxt7.Enabled = False
        '副标题
        TextBoxs00.Text = “”
        TextBoxs0.Text = “”
        TextBoxs1.Text = “”
        TextBoxs2.Text = “”
        TextBoxs3.Text = “”
        TextBoxs4.Text = “”
        TextBoxs5.Text = “”
        TextBoxs6.Text = “”
        TextBoxs7.Text = “”
        TextBoxs00.Enabled = False
        TextBoxs0.Enabled = False
        TextBoxs1.Enabled = False
        TextBoxs2.Enabled = False
        TextBoxs3.Enabled = False
        TextBoxs4.Enabled = False
        TextBoxs5.Enabled = False
        TextBoxs6.Enabled = False
        TextBoxs7.Enabled = False
        '摘要
        TextBoxAb00.Text = “”
        TextBoxAb0.Text = “”
        TextBoxAb1.Text = “”
        TextBoxAb2.Text = “”
        TextBoxAb3.Text = “”
        TextBoxAb4.Text = “”
        TextBoxAb5.Text = “”
        TextBoxAb6.Text = “”
        TextBoxAb7.Text = “”
        TextBoxAb00.Enabled = False
        TextBoxAb0.Enabled = False
        TextBoxAb1.Enabled = False
        TextBoxAb2.Enabled = False
        TextBoxAb3.Enabled = False
        TextBoxAb4.Enabled = False
        TextBoxAb5.Enabled = False
        TextBoxAb6.Enabled = False
        TextBoxAb7.Enabled = False
        TextBoxAb8.Enabled = False
        TextBoxAb9.Enabled = False
        TextBoxAb10.Enabled = False
        '关键词
        TextBoxKey00.Text = “0”
        TextBoxKey0.Text = “0”
        TextBoxKey1.Text = “0”
        TextBoxKey2.Text = “0”
        TextBoxKey3.Text = “20/12”
        TextBoxKey4.Text = “3”
        TextBoxKey5.Text = “Arial”
        TextBoxKey6.Text = “11”
        TextBoxKey7.Text = “False”
        TextBoxKey8.Text = “1”
        TextBoxKey9.Text = “3”
        TextBoxKey10.Text = “关键词”


        '一级标题
        TextBoxf00.Text = “”
        TextBoxf0.Text = “”
        TextBoxf1.Text = “”
        TextBoxf2.Text = “”
        TextBoxf3.Text = “”
        TextBoxf4.Text = “”
        TextBoxf5.Text = “”
        TextBoxf6.Text = “”
        TextBoxf7.Text = “”
        TextBoxf00.Enabled = False
        TextBoxf0.Enabled = False
        TextBoxf1.Enabled = False
        TextBoxf2.Enabled = False
        TextBoxf3.Enabled = False
        TextBoxf4.Enabled = False
        TextBoxf5.Enabled = False
        TextBoxf6.Enabled = False
        TextBoxf7.Enabled = False
        TextBoxf8.Enabled = False
        TextBoxf9.Enabled = False
        TextBoxf10.Enabled = False
        '二级标题
        TextBoxsec00.Text = “2”
        TextBoxsec0.Text = “0.5”
        TextBoxsec1.Text = “0”
        TextBoxsec2.Text = “0”
        TextBoxsec3.Text = “1”
        TextBoxsec4.Text = “3”
        TextBoxsec5.Text = “黑体”
        TextBoxsec6.Text = “14”
        TextBoxsec7.Text = “True”
        '三级标题
        TextBoxthd00.Text = “2”
        TextBoxthd0.Text = “0.5”
        TextBoxthd1.Text = “0”
        TextBoxthd2.Text = “0”
        TextBoxthd3.Text = “1”
        TextBoxthd4.Text = “3”
        TextBoxthd5.Text = “黑体”
        TextBoxthd6.Text = “12”
        TextBoxthd7.Text = “True”
        '正文
        TextBoxc00.Text = “0”
        TextBoxc0.Text = “0”
        TextBoxc1.Text = “0”
        TextBoxc2.Text = “0.5”
        TextBoxc3.Text = “20/12”
        TextBoxc4.Text = “3”
        TextBoxc5.Text = “宋体”
        TextBoxc6.Text = “12”
        TextBoxc7.Text = “False”
        '参考文献
        TextBoxr00.Text = “0”
        TextBoxr0.Text = “0”
        TextBoxr1.Text = “0”
        TextBoxr2.Text = “0”
        TextBoxr3.Text = “20/12”
        TextBoxr4.Text = “3”
        TextBoxr5.Text = “Times New Roman”
        TextBoxr6.Text = “12”
        TextBoxr7.Text = “False”
        '图注
        TextBoxtz00.Text = “0.5”
        TextBoxtz0.Text = “1”
        TextBoxtz1.Text = “0”
        TextBoxtz2.Text = “0”
        TextBoxtz3.Text = “1”
        TextBoxtz4.Text = “1”
        TextBoxtz5.Text = “宋体”
        TextBoxtz6.Text = “10.5”
        TextBoxtz7.Text = “False”
    End Sub

    Private Sub bt4_Click(sender As Object, e As EventArgs) Handles bt_zwqk.Click
        '标题
        TextBoxt00.Text = “0”
        TextBoxt0.Text = “0”
        TextBoxt1.Text = “0”
        TextBoxt2.Text = “0”
        TextBoxt3.Text = “1.5”
        TextBoxt4.Text = “1”
        TextBoxt5.Text = “宋体”
        TextBoxt6.Text = “22”
        TextBoxt7.Text = “True”
        '副标题
        TextBoxs00.Text = “0”
        TextBoxs0.Text = “0”
        TextBoxs1.Text = “0”
        TextBoxs2.Text = “0”
        TextBoxs3.Text = “1.5”
        TextBoxs4.Text = “1”
        TextBoxs5.Text = “楷体”
        TextBoxs6.Text = “10.5”
        TextBoxs7.Text = “False”
        '摘要
        TextBoxAb00.Text = “0”
        TextBoxAb0.Text = “0”
        TextBoxAb1.Text = “0”
        TextBoxAb2.Text = “0.74”
        TextBoxAb3.Text = “1.5”
        TextBoxAb4.Text = “3”
        TextBoxAb5.Text = “楷体”
        TextBoxAb6.Text = “9”
        TextBoxAb7.Text = “False”
        TextBoxAb8.Text = “1”
        TextBoxAb9.Text = “2”
        TextBoxAb10.Text = “摘要”
        '关键词
        TextBoxKey00.Text = “0”
        TextBoxKey0.Text = “0”
        TextBoxKey1.Text = “0”
        TextBoxKey2.Text = “0”
        TextBoxKey3.Text = “1.5”
        TextBoxKey4.Text = “3”
        TextBoxKey5.Text = “楷体”
        TextBoxKey6.Text = “9”
        TextBoxKey7.Text = “False”
        TextBoxKey8.Text = “1”
        TextBoxKey9.Text = “3”
        TextBoxKey10.Text = “关键词”


        '一级标题
        TextBoxf00.Text = “0”
        TextBoxf0.Text = “0”
        TextBoxf1.Text = “0”
        TextBoxf2.Text = “0”
        TextBoxf3.Text = “1.5”
        TextBoxf4.Text = “3”
        TextBoxf5.Text = “宋体”
        TextBoxf6.Text = “12”
        TextBoxf7.Text = “True”
        '二级标题
        TextBoxsec00.Text = “0”
        TextBoxsec0.Text = “0”
        TextBoxsec1.Text = “0”
        TextBoxsec2.Text = “0”
        TextBoxsec3.Text = “1.5”
        TextBoxsec4.Text = “3”
        TextBoxsec5.Text = “宋体”
        TextBoxsec6.Text = “10.5”
        TextBoxsec7.Text = “True”
        '三级标题
        TextBoxthd00.Text = “0”
        TextBoxthd0.Text = “0”
        TextBoxthd1.Text = “0”
        TextBoxthd2.Text = “0”
        TextBoxthd3.Text = “1.5”
        TextBoxthd4.Text = “3”
        TextBoxthd5.Text = “宋体”
        TextBoxthd6.Text = “10.5”
        TextBoxthd7.Text = “False”
        '正文
        TextBoxc00.Text = “0”
        TextBoxc0.Text = “0”
        TextBoxc1.Text = “0”
        TextBoxc2.Text = “0.74”
        TextBoxc3.Text = “1.5”
        TextBoxc4.Text = “3”
        TextBoxc5.Text = “Times New Roman”
        TextBoxc6.Text = “12”
        TextBoxc7.Text = “False”
        '参考文献
        TextBoxr00.Text = “0”
        TextBoxr0.Text = “0”
        TextBoxr1.Text = “0”
        TextBoxr2.Text = “0.74”
        TextBoxr3.Text = “1.5”
        TextBoxr4.Text = “3”
        TextBoxr5.Text = “宋体”
        TextBoxr6.Text = “10.5”
        TextBoxr7.Text = “False”
        '图注
        TextBoxtz00.Text = “0”
        TextBoxtz0.Text = “0”
        TextBoxtz1.Text = “0”
        TextBoxtz2.Text = “0”
        TextBoxtz3.Text = “1.5”
        TextBoxtz4.Text = “1”
        TextBoxtz5.Text = “黑体”
        TextBoxtz6.Text = “9”
        TextBoxtz7.Text = “False”
        TextBoxtz8.Text = “1”
        TextBoxtz9.Text = “1”
        TextBoxtz10.Text = “图”
    End Sub
End Class