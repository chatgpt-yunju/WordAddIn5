Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
    Public fm1 As Form1
    Public fm2 As Form2
    Public x As String
    Public y As String
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        EditBox1.Text = "1"
        EditBox2.Text = "1"
        x = "1"
        y = "1"
        If fm1 Is Nothing Then
            fm1 = New Form1
        End If
        If fm2 Is Nothing Then
            fm2 = New Form2
        End If
        '中文期刊论文模板初始化
        '标题
        fm2.TextBoxt00.Text = “0”
        fm2.TextBoxt0.Text = “0”
        fm2.TextBoxt1.Text = “0”
        fm2.TextBoxt2.Text = “0”
        fm2.TextBoxt3.Text = “1.5”
        fm2.TextBoxt4.Text = “1”
        fm2.TextBoxt5.Text = “宋体”
        fm2.TextBoxt6.Text = “22”
        fm2.TextBoxt7.Text = “True”
        '副标题
        fm2.TextBoxs00.Text = “0”
        fm2.TextBoxs0.Text = “0”
        fm2.TextBoxs1.Text = “0”
        fm2.TextBoxs2.Text = “0”
        fm2.TextBoxs3.Text = “1.5”
        fm2.TextBoxs4.Text = “1”
        fm2.TextBoxs5.Text = “楷体”
        fm2.TextBoxs6.Text = “10.5”
        fm2.TextBoxs7.Text = “False”
        '摘要
        fm2.TextBoxAb00.Text = “0”
        fm2.TextBoxAb0.Text = “0”
        fm2.TextBoxAb1.Text = “0”
        fm2.TextBoxAb2.Text = “0.74”
        fm2.TextBoxAb3.Text = “1.5”
        fm2.TextBoxAb4.Text = “3”
        fm2.TextBoxAb5.Text = “楷体”
        fm2.TextBoxAb6.Text = “9”
        fm2.TextBoxAb7.Text = “False”
        fm2.TextBoxAb8.Text = “1”
        fm2.TextBoxAb9.Text = “2”
        fm2.TextBoxAb10.Text = “摘要”
        '关键词
        fm2.TextBoxKey00.Text = “0”
        fm2.TextBoxKey0.Text = “0”
        fm2.TextBoxKey1.Text = “0”
        fm2.TextBoxKey2.Text = “0”
        fm2.TextBoxKey3.Text = “1.5”
        fm2.TextBoxKey4.Text = “3”
        fm2.TextBoxKey5.Text = “楷体”
        fm2.TextBoxKey6.Text = “9”
        fm2.TextBoxKey7.Text = “False”
        fm2.TextBoxKey8.Text = “1”
        fm2.TextBoxKey9.Text = “3”
        fm2.TextBoxKey10.Text = “关键词”


        '一级标题
        fm2.TextBoxf00.Text = “0”
        fm2.TextBoxf0.Text = “0”
        fm2.TextBoxf1.Text = “0”
        fm2.TextBoxf2.Text = “0”
        fm2.TextBoxf3.Text = “1.5”
        fm2.TextBoxf4.Text = “3”
        fm2.TextBoxf5.Text = “宋体”
        fm2.TextBoxf6.Text = “12”
        fm2.TextBoxf7.Text = “True”
        '二级标题
        fm2.TextBoxsec00.Text = “0”
        fm2.TextBoxsec0.Text = “0”
        fm2.TextBoxsec1.Text = “0”
        fm2.TextBoxsec2.Text = “0”
        fm2.TextBoxsec3.Text = “1.5”
        fm2.TextBoxsec4.Text = “3”
        fm2.TextBoxsec5.Text = “宋体”
        fm2.TextBoxsec6.Text = “10.5”
        fm2.TextBoxsec7.Text = “True”
        '三级标题
        fm2.TextBoxthd00.Text = “0”
        fm2.TextBoxthd0.Text = “0”
        fm2.TextBoxthd1.Text = “0”
        fm2.TextBoxthd2.Text = “0”
        fm2.TextBoxthd3.Text = “1.5”
        fm2.TextBoxthd4.Text = “3”
        fm2.TextBoxthd5.Text = “宋体”
        fm2.TextBoxthd6.Text = “10.5”
        fm2.TextBoxthd7.Text = “False”
        '正文
        fm2.TextBoxc00.Text = “0”
        fm2.TextBoxc0.Text = “0”
        fm2.TextBoxc1.Text = “0”
        fm2.TextBoxc2.Text = “0.74”
        fm2.TextBoxc3.Text = “1.5”
        fm2.TextBoxc4.Text = “3”
        fm2.TextBoxc5.Text = “Times New Roman”
        fm2.TextBoxc6.Text = “12”
        fm2.TextBoxc7.Text = “False”
        '参考文献
        fm2.TextBoxr00.Text = “0”
        fm2.TextBoxr0.Text = “0”
        fm2.TextBoxr1.Text = “0”
        fm2.TextBoxr2.Text = “0.74”
        fm2.TextBoxr3.Text = “1.5”
        fm2.TextBoxr4.Text = “3”
        fm2.TextBoxr5.Text = “宋体”
        fm2.TextBoxr6.Text = “10.5”
        fm2.TextBoxr7.Text = “False”
        '图注
        fm2.TextBoxtz00.Text = “0”
        fm2.TextBoxtz0.Text = “0”
        fm2.TextBoxtz1.Text = “0”
        fm2.TextBoxtz2.Text = “0”
        fm2.TextBoxtz3.Text = “1.5”
        fm2.TextBoxtz4.Text = “1”
        fm2.TextBoxtz5.Text = “黑体”
        fm2.TextBoxtz6.Text = “9”
        fm2.TextBoxtz7.Text = “False”
        fm2.TextBoxtz8.Text = “1”
        fm2.TextBoxtz9.Text = “1”
        fm2.TextBoxtz10.Text = “图”
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        If Len(fm1.TextBox1.Text) > 0 Then ReplaceFig(fm1.TextBox1.Text, 1)
        If Len(fm1.TextBox2.Text) > 0 Then ReplaceFig(fm1.TextBox2.Text, 2)
        If Len(fm1.TextBox3.Text) > 0 Then ReplaceFig(fm1.TextBox3.Text, 3)
        If Len(fm1.TextBox4.Text) > 0 Then ReplaceFig(fm1.TextBox4.Text, 4)
        If Len(fm1.TextBox5.Text) > 0 Then ReplaceFig(fm1.TextBox5.Text, 5)
        If Len(fm1.TextBox6.Text) > 0 Then ReplaceFig(fm1.TextBox6.Text, 6)
        If Len(fm1.TextBox7.Text) > 0 Then ReplaceFig(fm1.TextBox7.Text, 7)
        If Len(fm1.TextBox8.Text) > 0 Then ReplaceFig(fm1.TextBox8.Text, 8)
        If Len(fm1.TextBox9.Text) > 0 Then ReplaceFig(fm1.TextBox9.Text, 9)
        If Len(fm1.TextBox10.Text) > 0 Then ReplaceFig(fm1.TextBox10.Text, 10)
        If Len(fm1.TextBox11.Text) > 0 Then ReplaceFig(fm1.TextBox11.Text, 11)
        If Len(fm1.TextBox12.Text) > 0 Then ReplaceFig(fm1.TextBox12.Text, 12)
        If Len(fm1.TextBox13.Text) > 0 Then ReplaceFig(fm1.TextBox13.Text, 13)
        If Len(fm1.TextBox14.Text) > 0 Then ReplaceFig(fm1.TextBox14.Text, 14)
        If Len(fm1.TextBox15.Text) > 0 Then ReplaceFig(fm1.TextBox15.Text, 15)
        If Len(fm1.TextBox16.Text) > 0 Then ReplaceFig(fm1.TextBox16.Text, 16)
        If Len(fm1.TextBox17.Text) > 0 Then ReplaceFig(fm1.TextBox17.Text, 17)
        If Len(fm1.TextBox18.Text) > 0 Then ReplaceFig(fm1.TextBox18.Text, 18)
        If Len(fm1.TextBox19.Text) > 0 Then ReplaceFig(fm1.TextBox19.Text, 19)
        If Len(fm1.TextBox20.Text) > 0 Then ReplaceFig(fm1.TextBox20.Text, 20)
        If Len(fm1.TextBox21.Text) > 0 Then ReplaceFig(fm1.TextBox21.Text, 21)
        If Len(fm1.TextBox22.Text) > 0 Then ReplaceFig(fm1.TextBox22.Text, 22)
        If Len(fm1.TextBox23.Text) > 0 Then ReplaceFig(fm1.TextBox23.Text, 23)
        If Len(fm1.TextBox24.Text) > 0 Then ReplaceFig(fm1.TextBox24.Text, 24)
        If Len(fm1.TextBox25.Text) > 0 Then ReplaceFig(fm1.TextBox25.Text, 25)
        If Len(fm1.TextBox26.Text) > 0 Then ReplaceFig(fm1.TextBox26.Text, 26)
        If Len(fm1.TextBox27.Text) > 0 Then ReplaceFig(fm1.TextBox27.Text, 27)
        If Len(fm1.TextBox28.Text) > 0 Then ReplaceFig(fm1.TextBox28.Text, 28)
        If Len(fm1.TextBox29.Text) > 0 Then ReplaceFig(fm1.TextBox29.Text, 29)
        If Len(fm1.TextBox30.Text) > 0 Then ReplaceFig(fm1.TextBox30.Text, 30)
        If Len(fm1.TextBox31.Text) > 0 Then ReplaceFig(fm1.TextBox31.Text, 31)
        If Len(fm1.TextBox32.Text) > 0 Then ReplaceFig(fm1.TextBox32.Text, 32)
        If Len(fm1.TextBox33.Text) > 0 Then ReplaceFig(fm1.TextBox33.Text, 33)
        If Len(fm1.TextBox34.Text) > 0 Then ReplaceFig(fm1.TextBox34.Text, 34)
        If Len(fm1.TextBox35.Text) > 0 Then ReplaceFig(fm1.TextBox35.Text, 35)
        If Len(fm1.TextBox36.Text) > 0 Then ReplaceFig(fm1.TextBox36.Text, 36)
        If Len(fm1.TextBox37.Text) > 0 Then ReplaceFig(fm1.TextBox37.Text, 37)
        If Len(fm1.TextBox38.Text) > 0 Then ReplaceFig(fm1.TextBox38.Text, 38)
        If Len(fm1.TextBox39.Text) > 0 Then ReplaceFig(fm1.TextBox39.Text, 39)
        If Len(fm1.TextBox40.Text) > 0 Then ReplaceFig(fm1.TextBox40.Text, 40)
        If Len(fm1.TextBox41.Text) > 0 Then ReplaceFig(fm1.TextBox41.Text, 41)
        If Len(fm1.TextBox42.Text) > 0 Then ReplaceFig(fm1.TextBox42.Text, 42)
        If Len(fm1.TextBox43.Text) > 0 Then ReplaceFig(fm1.TextBox43.Text, 43)
        If Len(fm1.TextBox44.Text) > 0 Then ReplaceFig(fm1.TextBox44.Text, 44)
        If Len(fm1.TextBox45.Text) > 0 Then ReplaceFig(fm1.TextBox45.Text, 45)
        If Len(fm1.TextBox46.Text) > 0 Then ReplaceFig(fm1.TextBox46.Text, 46)
        If Len(fm1.TextBox47.Text) > 0 Then ReplaceFig(fm1.TextBox47.Text, 47)
        If Len(fm1.TextBox48.Text) > 0 Then ReplaceFig(fm1.TextBox48.Text, 48)
        If Len(fm1.TextBox49.Text) > 0 Then ReplaceFig(fm1.TextBox49.Text, 49)
        If Len(fm1.TextBox50.Text) > 0 Then ReplaceFig(fm1.TextBox50.Text, 50)
        If Len(fm1.TextBox51.Text) > 0 Then ReplaceFig(fm1.TextBox51.Text, 51)
        If Len(fm1.TextBox52.Text) > 0 Then ReplaceFig(fm1.TextBox52.Text, 52)
        If Len(fm1.TextBox53.Text) > 0 Then ReplaceFig(fm1.TextBox53.Text, 53)
        If Len(fm1.TextBox54.Text) > 0 Then ReplaceFig(fm1.TextBox54.Text, 54)
        If Len(fm1.TextBox55.Text) > 0 Then ReplaceFig(fm1.TextBox55.Text, 55)
        If Len(fm1.TextBox56.Text) > 0 Then ReplaceFig(fm1.TextBox56.Text, 56)
        If Len(fm1.TextBox57.Text) > 0 Then ReplaceFig(fm1.TextBox57.Text, 57)
        If Len(fm1.TextBox58.Text) > 0 Then ReplaceFig(fm1.TextBox58.Text, 58)
        If Len(fm1.TextBox59.Text) > 0 Then ReplaceFig(fm1.TextBox59.Text, 59)
        If Len(fm1.TextBox60.Text) > 0 Then ReplaceFig(fm1.TextBox60.Text, 60)
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        If fm1.Visible = False Then
            fm1.Show()
        Else

        End If
    End Sub



    Function InsertFig(url)
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        If Len(url) > 0 Then
            wdapp.Selection.InlineShapes.AddPicture(url)
        End If
    End Function

    Function ReplaceFig(url, i)
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        If Len(url) > 0 Then
            wdapp.ActiveDocument.InlineShapes(i).Select()
            wdapp.Selection.InlineShapes.AddPicture(url)
        End If
    End Function

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        If Convert.ToInt16(x) = 1 Then ReplaceFig(fm1.TextBox1.Text, 1)
        If Convert.ToInt16(x) = 2 Then ReplaceFig(fm1.TextBox2.Text, 2)
        If Convert.ToInt16(x) = 3 Then ReplaceFig(fm1.TextBox3.Text, 3)
        If Convert.ToInt16(x) = 4 Then ReplaceFig(fm1.TextBox4.Text, 4)
        If Convert.ToInt16(x) = 5 Then ReplaceFig(fm1.TextBox5.Text, 5)
        If Convert.ToInt16(x) = 6 Then ReplaceFig(fm1.TextBox6.Text, 6)
        If Convert.ToInt16(x) = 7 Then ReplaceFig(fm1.TextBox7.Text, 7）
        If Convert.ToInt16(x) = 8 Then ReplaceFig(fm1.TextBox8.Text, 8)
        If Convert.ToInt16(x) = 9 Then ReplaceFig(fm1.TextBox9.Text, 9)
        If Convert.ToInt16(x) = 10 Then ReplaceFig(fm1.TextBox10.Text, 10)
        If Convert.ToInt16(x) = 11 Then ReplaceFig(fm1.TextBox11.Text, 11)
        If Convert.ToInt16(x) = 12 Then ReplaceFig(fm1.TextBox12.Text, 12)
        If Convert.ToInt16(x) = 13 Then ReplaceFig(fm1.TextBox13.Text, 13)
        If Convert.ToInt16(x) = 14 Then ReplaceFig(fm1.TextBox14.Text, 14)
        If Convert.ToInt16(x) = 15 Then ReplaceFig(fm1.TextBox15.Text, 15)
        If Convert.ToInt16(x) = 16 Then ReplaceFig(fm1.TextBox16.Text, 16)
        If Convert.ToInt16(x) = 17 Then ReplaceFig(fm1.TextBox17.Text, 17)
        If Convert.ToInt16(x) = 18 Then ReplaceFig(fm1.TextBox18.Text, 18)
        If Convert.ToInt16(x) = 19 Then ReplaceFig(fm1.TextBox19.Text, 19)
        If Convert.ToInt16(x) = 20 Then ReplaceFig(fm1.TextBox20.Text, 20)
        If Convert.ToInt16(x) = 21 Then ReplaceFig(fm1.TextBox21.Text, 21)
        If Convert.ToInt16(x) = 22 Then ReplaceFig(fm1.TextBox22.Text, 22)
        If Convert.ToInt16(x) = 23 Then ReplaceFig(fm1.TextBox23.Text, 23)
        If Convert.ToInt16(x) = 24 Then ReplaceFig(fm1.TextBox24.Text, 24)
        If Convert.ToInt16(x) = 25 Then ReplaceFig(fm1.TextBox25.Text, 25)
        If Convert.ToInt16(x) = 26 Then ReplaceFig(fm1.TextBox26.Text, 26)
        If Convert.ToInt16(x) = 27 Then ReplaceFig(fm1.TextBox27.Text, 27)
        If Convert.ToInt16(x) = 28 Then ReplaceFig(fm1.TextBox28.Text, 28)
        If Convert.ToInt16(x) = 29 Then ReplaceFig(fm1.TextBox29.Text, 29)
        If Convert.ToInt16(x) = 30 Then ReplaceFig(fm1.TextBox30.Text, 30)
        If Convert.ToInt16(x) = 31 Then ReplaceFig(fm1.TextBox31.Text, 31)
        If Convert.ToInt16(x) = 32 Then ReplaceFig(fm1.TextBox32.Text, 32)
        If Convert.ToInt16(x) = 33 Then ReplaceFig(fm1.TextBox33.Text, 33)
        If Convert.ToInt16(x) = 34 Then ReplaceFig(fm1.TextBox34.Text, 34)
        If Convert.ToInt16(x) = 35 Then ReplaceFig(fm1.TextBox35.Text, 35)
        If Convert.ToInt16(x) = 36 Then ReplaceFig(fm1.TextBox36.Text, 36)
        If Convert.ToInt16(x) = 37 Then ReplaceFig(fm1.TextBox37.Text, 37)
        If Convert.ToInt16(x) = 38 Then ReplaceFig(fm1.TextBox38.Text, 38)
        If Convert.ToInt16(x) = 39 Then ReplaceFig(fm1.TextBox39.Text, 39)
        If Convert.ToInt16(x) = 40 Then ReplaceFig(fm1.TextBox40.Text, 40)
        If Convert.ToInt16(x) = 41 Then ReplaceFig(fm1.TextBox41.Text, 41)
        If Convert.ToInt16(x) = 42 Then ReplaceFig(fm1.TextBox42.Text, 42)
        If Convert.ToInt16(x) = 43 Then ReplaceFig(fm1.TextBox43.Text, 43)
        If Convert.ToInt16(x) = 44 Then ReplaceFig(fm1.TextBox44.Text, 44)
        If Convert.ToInt16(x) = 45 Then ReplaceFig(fm1.TextBox45.Text, 45)
        If Convert.ToInt16(x) = 46 Then ReplaceFig(fm1.TextBox46.Text, 46)
        If Convert.ToInt16(x) = 47 Then ReplaceFig(fm1.TextBox47.Text, 47)
        If Convert.ToInt16(x) = 48 Then ReplaceFig(fm1.TextBox48.Text, 48)
        If Convert.ToInt16(x) = 49 Then ReplaceFig(fm1.TextBox49.Text, 49)
        If Convert.ToInt16(x) = 50 Then ReplaceFig(fm1.TextBox50.Text, 50)
        If Convert.ToInt16(x) = 51 Then ReplaceFig(fm1.TextBox51.Text, 51)
        If Convert.ToInt16(x) = 52 Then ReplaceFig(fm1.TextBox52.Text, 52)
        If Convert.ToInt16(x) = 53 Then ReplaceFig(fm1.TextBox53.Text, 53)
        If Convert.ToInt16(x) = 54 Then ReplaceFig(fm1.TextBox54.Text, 54)
        If Convert.ToInt16(x) = 55 Then ReplaceFig(fm1.TextBox55.Text, 55)
        If Convert.ToInt16(x) = 56 Then ReplaceFig(fm1.TextBox56.Text, 56)
        If Convert.ToInt16(x) = 57 Then ReplaceFig(fm1.TextBox57.Text, 57)
        If Convert.ToInt16(x) = 58 Then ReplaceFig(fm1.TextBox58.Text, 58)
        If Convert.ToInt16(x) = 59 Then ReplaceFig(fm1.TextBox59.Text, 59)
        If Convert.ToInt16(x) = 60 Then ReplaceFig(fm1.TextBox60.Text, 60)
    End Sub

    Private Sub EditBox1_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles EditBox1.TextChanged
        x = EditBox1.Text

    End Sub

    Private Sub EditBox2_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles EditBox2.TextChanged
        y = EditBox2.Text
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        If Convert.ToInt16(y) = 1 Then InsertFig(fm1.TextBox1.Text)
        If Convert.ToInt16(y) = 2 Then InsertFig(fm1.TextBox2.Text)
        If Convert.ToInt16(y) = 3 Then InsertFig(fm1.TextBox3.Text)
        If Convert.ToInt16(y) = 4 Then InsertFig(fm1.TextBox4.Text)
        If Convert.ToInt16(y) = 5 Then InsertFig(fm1.TextBox5.Text)
        If Convert.ToInt16(y) = 6 Then InsertFig(fm1.TextBox6.Text)
        If Convert.ToInt16(y) = 7 Then InsertFig(fm1.TextBox7.Text）
        If Convert.ToInt16(y) = 8 Then InsertFig(fm1.TextBox8.Text)
        If Convert.ToInt16(y) = 9 Then InsertFig(fm1.TextBox9.Text)
        If Convert.ToInt16(y) = 10 Then InsertFig(fm1.TextBox10.Text)
        If Convert.ToInt16(y) = 11 Then InsertFig(fm1.TextBox11.Text)
        If Convert.ToInt16(y) = 12 Then InsertFig(fm1.TextBox12.Text)
        If Convert.ToInt16(y) = 13 Then InsertFig(fm1.TextBox13.Text)
        If Convert.ToInt16(y) = 14 Then InsertFig(fm1.TextBox14.Text)
        If Convert.ToInt16(y) = 15 Then InsertFig(fm1.TextBox15.Text)
        If Convert.ToInt16(y) = 16 Then InsertFig(fm1.TextBox16.Text)
        If Convert.ToInt16(y) = 17 Then InsertFig(fm1.TextBox17.Text)
        If Convert.ToInt16(y) = 18 Then InsertFig(fm1.TextBox18.Text)
        If Convert.ToInt16(y) = 19 Then InsertFig(fm1.TextBox19.Text)
        If Convert.ToInt16(y) = 20 Then InsertFig(fm1.TextBox20.Text)
        If Convert.ToInt16(y) = 21 Then InsertFig(fm1.TextBox21.Text)
        If Convert.ToInt16(y) = 22 Then InsertFig(fm1.TextBox22.Text)
        If Convert.ToInt16(y) = 23 Then InsertFig(fm1.TextBox23.Text)
        If Convert.ToInt16(y) = 24 Then InsertFig(fm1.TextBox24.Text)
        If Convert.ToInt16(y) = 25 Then InsertFig(fm1.TextBox25.Text)
        If Convert.ToInt16(y) = 26 Then InsertFig(fm1.TextBox26.Text)
        If Convert.ToInt16(y) = 27 Then InsertFig(fm1.TextBox27.Text)
        If Convert.ToInt16(y) = 28 Then InsertFig(fm1.TextBox28.Text)
        If Convert.ToInt16(y) = 29 Then InsertFig(fm1.TextBox29.Text)
        If Convert.ToInt16(y) = 30 Then InsertFig(fm1.TextBox30.Text)
        If Convert.ToInt16(y) = 31 Then InsertFig(fm1.TextBox31.Text)
        If Convert.ToInt16(y) = 32 Then InsertFig(fm1.TextBox32.Text)
        If Convert.ToInt16(y) = 33 Then InsertFig(fm1.TextBox33.Text)
        If Convert.ToInt16(y) = 34 Then InsertFig(fm1.TextBox34.Text)
        If Convert.ToInt16(y) = 35 Then InsertFig(fm1.TextBox35.Text)
        If Convert.ToInt16(y) = 36 Then InsertFig(fm1.TextBox36.Text)
        If Convert.ToInt16(y) = 37 Then InsertFig(fm1.TextBox37.Text)
        If Convert.ToInt16(y) = 38 Then InsertFig(fm1.TextBox38.Text)
        If Convert.ToInt16(y) = 39 Then InsertFig(fm1.TextBox39.Text)
        If Convert.ToInt16(y) = 40 Then InsertFig(fm1.TextBox40.Text)
        If Convert.ToInt16(y) = 41 Then InsertFig(fm1.TextBox41.Text)
        If Convert.ToInt16(y) = 42 Then InsertFig(fm1.TextBox42.Text)
        If Convert.ToInt16(y) = 43 Then InsertFig(fm1.TextBox43.Text)
        If Convert.ToInt16(y) = 44 Then InsertFig(fm1.TextBox44.Text)
        If Convert.ToInt16(y) = 45 Then InsertFig(fm1.TextBox45.Text)
        If Convert.ToInt16(y) = 46 Then InsertFig(fm1.TextBox46.Text)
        If Convert.ToInt16(y) = 47 Then InsertFig(fm1.TextBox47.Text)
        If Convert.ToInt16(y) = 48 Then InsertFig(fm1.TextBox48.Text)
        If Convert.ToInt16(y) = 49 Then InsertFig(fm1.TextBox49.Text)
        If Convert.ToInt16(y) = 50 Then InsertFig(fm1.TextBox50.Text)
        If Convert.ToInt16(y) = 51 Then InsertFig(fm1.TextBox51.Text)
        If Convert.ToInt16(y) = 52 Then InsertFig(fm1.TextBox52.Text)
        If Convert.ToInt16(y) = 53 Then InsertFig(fm1.TextBox53.Text)
        If Convert.ToInt16(y) = 54 Then InsertFig(fm1.TextBox54.Text)
        If Convert.ToInt16(y) = 55 Then InsertFig(fm1.TextBox55.Text)
        If Convert.ToInt16(y) = 56 Then InsertFig(fm1.TextBox56.Text)
        If Convert.ToInt16(y) = 57 Then InsertFig(fm1.TextBox57.Text)
        If Convert.ToInt16(y) = 58 Then InsertFig(fm1.TextBox58.Text)
        If Convert.ToInt16(y) = 59 Then InsertFig(fm1.TextBox59.Text)
        If Convert.ToInt16(y) = 60 Then InsertFig(fm1.TextBox60.Text)
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click

        If (Convert.ToInt16(x) - 1) >= 1 Then
            x = Convert.ToString(Convert.ToInt16(x) - 1)
            EditBox1.Text = x
        Else
            MsgBox("目前版本x的范围是1~60哦")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        If (Convert.ToInt16(x) + 1) <= 60 Then
            x = Convert.ToString(Convert.ToInt16(x) + 1)
            EditBox1.Text = x
        Else
            MsgBox("目前版本x的范围是1~60哦")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        If (Convert.ToInt16(y) - 1) >= 1 Then
            y = Convert.ToString(Convert.ToInt16(y) - 1)
            EditBox2.Text = y
        Else
            MsgBox("目前版本y的范围是1~60哦")
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        If (Convert.ToInt16(y) + 1) <= 60 Then
            y = Convert.ToString(Convert.ToInt16(y) + 1)
            EditBox2.Text = y
        Else
            MsgBox("目前版本y的范围是1~60哦")
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        System.Diagnostics.Process.Start("http://home.ustc.edu.cn/~yunju/IP/")
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs)
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim pcount As Int16
        pcount = 0
        For Each oP In wdapp.ActiveDocument.Paragraphs
            pcount = pcount + 1
        Next
        Dim i As Long
        For i = 1 To pcount
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Alignment = 1 And oP.Range.Characters.Count > 1) Then
                If (oP.Range.Font.Bold = True) Then
                    If (oP.Range.Characters(1).Text <> " ") Then
                        If (oP.Range.Font.Color <> RGB(250, 64, 6)) Then
                            MsgBox("这是总标题")

                            If (Math.Round(wdapp.PointsToCentimeters(oP.LeftIndent), 2) = 0) Then
                                'MsgBox("总标题左缩进为0.63cm√")
                                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                                    oP.Range.Font.Underline = False
                                End If
                            Else
                                'MsgBox("总标题左缩进不是0.63cm×，已使用蓝色下划线标识")
                                oP.Range.Font.Underline = True
                                oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                            End If

                            If (oP.FirstLineIndent = 0) Then
                                'MsgBox("二级标题首行缩进为2字符√")
                                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                                    oP.Range.Font.Underline = False
                                End If
                            Else
                                'MsgBox("二级标题首行缩进不是2字符×，已使用紫色下划线标出")
                                oP.Range.Font.Underline = True
                                oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen
                            End If

                            If (oP.LineSpacing = 12) Then
                                'MsgBox("总标题行距为36磅√")
                                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange) Then
                                    oP.Range.Font.Underline = False
                                End If
                            Else
                                'MsgBox("总标题行距不是36磅×，已使用绿色下划线标识")
                                oP.Range.Font.Underline = True
                                oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange
                            End If

                            If (oP.Alignment = 1) Then
                                'MsgBox("总标题已居中对齐√")
                                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                                    oP.Range.Font.Underline = False
                                End If
                            Else
                                'MsgBox("总标题未居中对齐×，已使用橘黄色下划线标识")
                                oP.Range.Font.Underline = True
                                oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                            End If

                            Dim j As Long
                            Dim cName, cSize, cBold As Int16
                            cName = 0
                            cSize = 0
                            cBold = 0
                            For j = 1 To oP.Range.Characters.Count
                                If (oP.Range.Characters(j).Font.Name = "Times New Roman") Then
                                    cName = 0
                                Else
                                    cName = 1
                                End If
                                If (oP.Range.Characters(j).Font.Size = 11) Then
                                    cSize = 0
                                Else
                                    cSize = 1
                                End If
                                If (oP.Range.Characters(j).Font.Bold = True) Then
                                    cBold = 0
                                Else
                                    cBold = 1
                                End If
                                If (cName = 0 And cSize = 0 And cBold = 0) Then
                                    oP.Range.Characters(j).HighlightColorIndex = 0
                                Else
                                    oP.Range.Characters(j).HighlightColorIndex = 7
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("总标题已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Alignment = 1 And oP.Range.Characters.Count > 1) Then
                If (oP.Range.Font.Bold = True And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxtz8.Text), CInt(fm2.TextBoxtz9.Text)) <> fm2.TextBoxtz10.Text.Trim.ToString) Then
                    'MsgBox("这是总标题")
                    '段前距
                    If Len(fm2.TextBoxt00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxt00.Text))
                    '段后距
                    If Len(fm2.TextBoxt0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxt0.Text))
                    '左侧进
                    If Len(fm2.TextBoxt1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxt1.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If Len(fm2.TextBoxt2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxt2.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxt3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxt3.Text))
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                        oP.Range.Font.Underline = False
                    End If
                    '对齐
                    If Len(fm2.TextBoxt4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxt4.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                        oP.Range.Font.Underline = False
                    End If

                    '字型'
                    If Len(fm2.TextBoxt5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxt5.Text)
                    If Len(fm2.TextBoxt6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxt6.Text)
                    If Len(fm2.TextBoxt7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxt7.Text)
                    oP.Range.HighlightColorIndex = 0
                    If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                        MsgBox("已修正完毕", 0, "消息提示")
                    End If
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("总标题已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub 作者_Click(sender As Object, e As RibbonControlEventArgs)
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim pcount As Int16
        pcount = 0
        For Each oP In wdapp.ActiveDocument.Paragraphs
            pcount = pcount + 1
        Next
        Dim i As Long
        For i = 1 To pcount
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Alignment = 1 And oP.Range.Characters.Count > 1) Then
                If (oP.Range.Font.Bold = False) Then
                    If (oP.Range.Characters(1).Text <> " ") Then
                        If (oP.Range.Font.Color <> RGB(250, 64, 6)) Then
                            'MsgBox("这是作者信息")

                            If (Math.Round(wdapp.PointsToCentimeters(oP.LeftIndent), 2) = 0) Then
                                'MsgBox("总标题左缩进为0.63cm√")
                                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                                    oP.Range.Font.Underline = False
                                End If
                            Else
                                'MsgBox("总标题左缩进不是0.63cm×，已使用蓝色下划线标识")
                                oP.Range.Font.Underline = True
                                oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                            End If

                            If (oP.FirstLineIndent = 0) Then
                                'MsgBox("二级标题首行缩进为2字符√")
                                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                                    oP.Range.Font.Underline = False
                                End If
                            Else
                                'MsgBox("二级标题首行缩进不是2字符×，已使用紫色下划线标出")
                                oP.Range.Font.Underline = True
                                oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen
                            End If

                            If (oP.LineSpacing = 12) Then
                                'MsgBox("总标题行距为36磅√")
                                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange) Then
                                    oP.Range.Font.Underline = False
                                End If
                            Else
                                'MsgBox("总标题行距不是36磅×，已使用绿色下划线标识")
                                oP.Range.Font.Underline = True
                                oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange
                            End If

                            If (oP.Alignment = 1) Then
                                'MsgBox("总标题已居中对齐√")
                                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                                    oP.Range.Font.Underline = False
                                End If
                            Else
                                'MsgBox("总标题未居中对齐×，已使用橘黄色下划线标识")
                                oP.Range.Font.Underline = True
                                oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                            End If

                            Dim j As Long
                            Dim cName, cSize, cBold As Int16
                            cName = 0
                            cSize = 0
                            cBold = 0
                            For j = 1 To oP.Range.Characters.Count
                                If (oP.Range.Characters(j).Font.Name = "Times New Roman") Then
                                    cName = 0
                                Else
                                    cName = 1
                                End If
                                If (oP.Range.Characters(j).Font.Size = 11) Then
                                    cSize = 0
                                Else
                                    cSize = 1
                                End If
                                If (oP.Range.Characters(j).Font.Bold = False) Then
                                    cBold = 0
                                Else
                                    cBold = 1
                                End If
                                If (cName = 0 And cSize = 0 And cBold = 0) Then
                                    oP.Range.Characters(j).HighlightColorIndex = 0
                                Else
                                    oP.Range.Characters(j).HighlightColorIndex = 7
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("作者信息已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs)
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim pcount As Int16
        pcount = 0
        For Each oP In wdapp.ActiveDocument.Paragraphs
            pcount = pcount + 1
        Next
        Dim i As Long
        For i = 1 To pcount
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Alignment <> 1 And oP.Range.Characters.Count > 1) Then
                If (Mid(Trim(oP.Range.Text.ToString), 1, 8) = "Abstract") Then
                    If (oP.Range.Font.Color <> RGB(250, 64, 6)) Then
                        'MsgBox("这是摘要")

                        If (Math.Round(wdapp.PointsToCentimeters(oP.LeftIndent), 2) = 0) Then
                            'MsgBox("总标题左缩进为0.63cm√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题左缩进不是0.63cm×，已使用蓝色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                        End If

                        If (oP.FirstLineIndent = 0) Then
                            'MsgBox("二级标题首行缩进为2字符√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("二级标题首行缩进不是2字符×，已使用紫色下划线标出")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen
                        End If

                        If (oP.LineSpacing = 12) Then
                            'MsgBox("总标题行距为36磅√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题行距不是36磅×，已使用绿色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange
                        End If

                        If (oP.Alignment = 3) Then
                            'MsgBox("总标题已居中对齐√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题未居中对齐×，已使用橘黄色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                        End If

                        Dim j As Long
                        Dim cName, cSize, cBold As Int16
                        cName = 0
                        cSize = 0
                        cBold = 0
                        For j = 10 To oP.Range.Characters.Count
                            If (oP.Range.Characters(j).Font.Name = "Times New Roman") Then
                                cName = 0
                            Else
                                cName = 1
                            End If
                            If (oP.Range.Characters(j).Font.Size = 11) Then
                                cSize = 0
                            Else
                                cSize = 1
                            End If
                            If (oP.Range.Characters(j).Font.Bold = False) Then
                                cBold = 0
                            Else
                                cBold = 1
                            End If
                            If (cName = 0 And cSize = 0 And cBold = 0) Then
                                oP.Range.Characters(j).HighlightColorIndex = 0
                            Else
                                oP.Range.Characters(j).HighlightColorIndex = 7
                            End If
                        Next
                    End If
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("摘要已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs)
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim pcount As Int16
        pcount = 0
        For Each oP In wdapp.ActiveDocument.Paragraphs
            pcount = pcount + 1
        Next
        Dim i As Long
        For i = 1 To pcount
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Alignment <> 1 And oP.Range.Characters.Count > 1) Then
                If (Mid(Trim(oP.Range.Text.ToString), 1, 3) = "Key") Then
                    If (oP.Range.Font.Color <> RGB(250, 64, 6)) Then
                        'MsgBox("这是关键词")

                        If (Math.Round(wdapp.PointsToCentimeters(oP.LeftIndent), 2) = 0) Then
                            'MsgBox("总标题左缩进为0.63cm√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题左缩进不是0.63cm×，已使用蓝色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                        End If

                        If (oP.FirstLineIndent = 0) Then
                            'MsgBox("二级标题首行缩进为2字符√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("二级标题首行缩进不是2字符×，已使用紫色下划线标出")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen
                        End If

                        If (oP.LineSpacing = 12) Then
                            'MsgBox("总标题行距为36磅√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题行距不是36磅×，已使用绿色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange
                        End If

                        If (oP.Alignment = 3) Then
                            'MsgBox("总标题已居中对齐√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题未居中对齐×，已使用橘黄色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                        End If

                        Dim j As Long
                        Dim cName, cSize, cBold As Int16
                        cName = 0
                        cSize = 0
                        cBold = 0
                        For j = 10 To oP.Range.Characters.Count
                            If (oP.Range.Characters(j).Font.Name = "Times New Roman") Then
                                cName = 0
                            Else
                                cName = 1
                            End If
                            If (oP.Range.Characters(j).Font.Size = 11) Then
                                cSize = 0
                            Else
                                cSize = 1
                            End If
                            If (oP.Range.Characters(j).Font.Bold = False) Then
                                cBold = 0
                            Else
                                cBold = 1
                            End If
                            If (cName = 0 And cSize = 0 And cBold = 0) Then
                                oP.Range.Characters(j).HighlightColorIndex = 0
                            Else
                                oP.Range.Characters(j).HighlightColorIndex = 7
                            End If
                        Next
                    End If
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("关键词已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs)
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim pcount As Int16
        pcount = 0
        For Each oP In wdapp.ActiveDocument.Paragraphs
            pcount = pcount + 1
        Next
        Dim i As Long
        For i = 1 To pcount
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Alignment <> 1 And oP.Range.Characters.Count > 1) Then
                If (Mid(Trim(oP.Range.Text.ToString), 1, 1) = "[") Then
                    If (oP.Range.Font.Color <> RGB(250, 64, 6)) Then
                        'MsgBox("这是参考文献")

                        If (Math.Round(wdapp.PointsToCentimeters(oP.LeftIndent), 2) = 0) Then
                            'MsgBox("总标题左缩进为0.63cm√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题左缩进不是0.63cm×，已使用蓝色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                        End If

                        If (oP.FirstLineIndent = 0) Then
                            'MsgBox("二级标题首行缩进为2字符√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("二级标题首行缩进不是2字符×，已使用紫色下划线标出")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen
                        End If

                        If (oP.LineSpacing = 12) Then
                            'MsgBox("总标题行距为36磅√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题行距不是36磅×，已使用绿色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorOrange
                        End If

                        If (oP.Alignment = 3) Then
                            'MsgBox("总标题已居中对齐√")
                            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                                oP.Range.Font.Underline = False
                            End If
                        Else
                            'MsgBox("总标题未居中对齐×，已使用橘黄色下划线标识")
                            oP.Range.Font.Underline = True
                            oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                        End If

                        Dim j As Long
                        Dim cName, cSize, cBold As Int16
                        cName = 0
                        cSize = 0
                        cBold = 0
                        For j = 1 To oP.Range.Characters.Count
                            If (oP.Range.Characters(j).Font.Name = "Times New Roman") Then
                                cName = 0
                            Else
                                cName = 1
                            End If
                            If (oP.Range.Characters(j).Font.Size = 11) Then
                                cSize = 0
                            Else
                                cSize = 1
                            End If
                            If (oP.Range.Characters(j).Font.Bold = False) Then
                                cBold = 0
                            Else
                                cBold = 1
                            End If
                            If (cName = 0 And cSize = 0 And cBold = 0) Then
                                oP.Range.Characters(j).HighlightColorIndex = 0
                            Else
                                oP.Range.Characters(j).HighlightColorIndex = 7
                            End If
                        Next
                    End If
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("参考文献已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Range.Characters.Count > 3) Then
                If (Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxr8.Text), CInt(fm2.TextBoxr9.Text)) = fm2.TextBoxr10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxr11.Text), CInt(fm2.TextBoxr12.Text)) = fm2.TextBoxr13.Text.Trim.ToString) Then
                    'MsgBox("这是参考文献")
                    '段前距
                    If Len(fm2.TextBoxr00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxr00.Text))
                    '段后距
                    If Len(fm2.TextBoxr0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxr0.Text))
                    '左侧进
                    If Len(fm2.TextBoxr1.Text) > 0 Then oP.LeftIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxr1.Text))
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If Len(fm2.TextBoxr2.Text) > 0 Then oP.FirstLineIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxr2.Text)）
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxr3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxr3.Text))
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                        oP.Range.Font.Underline = False
                    End If
                    '对齐
                    If Len(fm2.TextBoxr4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxr4.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                        oP.Range.Font.Underline = False
                    End If
                    '字型'
                    If Len(fm2.TextBoxr5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxr5.Text)
                    If Len(fm2.TextBoxr6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxr6.Text)
                    If Len(fm2.TextBoxr7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxr7.Text)
                    oP.Range.HighlightColorIndex = 0
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("参考文献已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs)
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            '左侧进
            'If Len(fm1.TextBox61.Text) > 0 Then oP.LeftIndent = CSng(fm1.TextBox61.Text)
            'If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
            '    oP.Range.Font.Underline = False
            'End If
            ''特殊格式
            'If Len(fm1.TextBox62.Text) > 0 Then oP.FirstLineIndent = CSng(fm1.TextBox62.Text)
            'If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
            '    oP.Range.Font.Underline = False
            'End If
            ''行距
            'If Len(fm1.TextBox63.Text) > 0 Then oP.LineSpacing = CSng(fm1.TextBox63.Text)
            'If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
            '    oP.Range.Font.Underline = False
            'End If
            ''对齐
            'If Len(fm1.TextBox64.Text) > 0 Then oP.Alignment = CInt(fm1.TextBox64.Text)
            'If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
            '    oP.Range.Font.Underline = False
            'End If

            ''字型'
            'If Len(fm1.TextBox65.Text) > 0 Then oP.Range.Font.Name = CStr(fm1.TextBox65.Text)
            'If Len(fm1.TextBox66.Text) > 0 Then oP.Range.Font.Size = CSng(fm1.TextBox66.Text)
            'If Len(fm1.TextBox67.Text) > 0 Then oP.Range.Font.Bold = CBool(fm1.TextBox67.Text)
            'oP.Range.HighlightColorIndex = 0

        Next
    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click
        If fm2.Visible = False Then
            fm2.Show()
        Else

        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As RibbonControlEventArgs) Handles Button18.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Alignment = 1 And oP.Range.Characters.Count > 1) Then
                If (oP.Range.Font.Bold = False And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxtz8.Text), CInt(fm2.TextBoxtz9.Text)) <> fm2.TextBoxtz10.Text.Trim.ToString) Then
                    'MsgBox("这是副标题")
                    '段前距
                    If Len(fm2.TextBoxs00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxs00.Text))
                    '段后距
                    If Len(fm2.TextBoxs0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxs0.Text))
                    '左侧进
                    If Len(fm2.TextBoxs1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxs1.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If Len(fm2.TextBoxs2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxs2.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxs3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxs3.Text))
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                        oP.Range.Font.Underline = False
                    End If
                    '对齐
                    If Len(fm2.TextBoxs4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxs4.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                        oP.Range.Font.Underline = False
                    End If

                    '字型'
                    If Len(fm2.TextBoxs5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxs5.Text)
                    If Len(fm2.TextBoxs6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxs6.Text)
                    If Len(fm2.TextBoxs7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxs7.Text)
                    oP.Range.HighlightColorIndex = 0
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("副标题已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (oP.Range.Characters.Count > 2 And oP.Alignment <> 1) Then
                If (Asc(Mid(Trim(oP.Range.Text.ToString), 1, 1)) > 48 And Asc(Mid(Trim(oP.Range.Text.ToString), 1, 1)) < 57 And Asc(Mid(Trim(oP.Range.Text.ToString), 2, 1)) = 32) Then
                    'MsgBox("这是一级标题1")
                    '段前距
                    If Len(fm2.TextBoxf00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxf00.Text))
                    '段后距
                    If Len(fm2.TextBoxf0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxf0.Text))
                    '左侧进
                    If Len(fm2.TextBoxf1.Text) > 0 Then oP.LeftIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxf1.Text))
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxf3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxf3.Text))
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                        oP.Range.Font.Underline = False
                    End If
                    '对齐
                    If Len(fm2.TextBoxf4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxf4.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                        oP.Range.Font.Underline = False
                    End If
                    '字型'
                    If Len(fm2.TextBoxf5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxf5.Text)
                    If Len(fm2.TextBoxf6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxf6.Text)
                    If Len(fm2.TextBoxf7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxf7.Text)
                    oP.Range.HighlightColorIndex = 0
                End If

                If (oP.Alignment <> 1 And Asc(Mid(Trim(oP.Range.Text.ToString), 1, 1)) > 48 And Asc(Mid(Trim(oP.Range.Text.ToString), 1, 1)) < 57 And Asc(Mid(Trim(oP.Range.Text.ToString), 2, 1)) = 46) Then
                    'MsgBox("这是一级标题2")
                    If (Asc(Mid(Trim(oP.Range.Text.ToString), 3, 1)) > 48 And Asc(Mid(Trim(oP.Range.Text.ToString), 3, 1)) < 57) Then

                    Else
                        '段前距
                        If Len(fm2.TextBoxf00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxf00.Text))
                        '段后距
                        If Len(fm2.TextBoxf0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxf0.Text))
                        '悬挂侧进
                        If Len(fm2.TextBoxf1.Text) > 0 Then oP.LeftIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxf1.Text))
                        If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                            oP.Range.Font.Underline = False
                        End If
                        '特殊格式
                        If Len(fm2.TextBoxf2.Text) > 0 Then oP.FirstLineIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxf2.Text)）
                        If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                            oP.Range.Font.Underline = False
                        End If
                        '行距
                        If Len(fm2.TextBoxf3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxf3.Text))
                        If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                            oP.Range.Font.Underline = False
                        End If
                        '对齐
                        If Len(fm2.TextBoxf4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxf4.Text)
                        If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                            oP.Range.Font.Underline = False
                        End If
                        '字型'
                        If Len(fm2.TextBoxf5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxf5.Text)
                        If Len(fm2.TextBoxf6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxf6.Text)
                        If Len(fm2.TextBoxf7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxf7.Text)
                        oP.Range.HighlightColorIndex = 0
                    End If
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("一级标题已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button20_Click(sender As Object, e As RibbonControlEventArgs) Handles Button20.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxsec8.Text), CInt(fm2.TextBoxsec9.Text)) = fm2.TextBoxsec10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxsec11.Text), CInt(fm2.TextBoxsec12.Text)) <> fm2.TextBoxsec13.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), 6, 1) <> ".") Then
                'MsgBox("这是二级标题")
                If (Asc(Mid(Trim(oP.Range.Text.ToString), 3, 1)) > 48 And Asc(Mid(Trim(oP.Range.Text.ToString), 3, 1)) < 57) Then
                    '段前距 你你
                    If Len(fm2.TextBoxsec00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxsec00.Text))
                    '段后距
                    If Len(fm2.TextBoxsec0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxsec0.Text))
                    '左侧进
                    If Len(fm2.TextBoxsec1.Text) > 0 Then oP.LeftIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxsec1.Text))
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If Len(fm2.TextBoxsec2.Text) > 0 Then oP.FirstLineIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxsec2.Text)）
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxsec3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxsec3.Text))
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                        oP.Range.Font.Underline = False
                    End If
                    '对齐
                    If Len(fm2.TextBoxsec4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxsec4.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                        oP.Range.Font.Underline = False
                    End If

                    '字型'
                    If Len(fm2.TextBoxsec5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxsec5.Text)
                    If Len(fm2.TextBoxsec6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxsec6.Text)
                    If Len(fm2.TextBoxsec7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxsec7.Text)
                    oP.Range.HighlightColorIndex = 0
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("二级标题已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button21_Click(sender As Object, e As RibbonControlEventArgs) Handles Button21.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            'If (Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxthd8.Text), CInt(fm2.TextBoxthd9.Text)) = fm2.TextBoxthd10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxthd11.Text), CInt(fm2.TextBoxthd12.Text)) = fm2.TextBoxthd13.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), Val(fm2.TextBoxthd14.Text), Val(fm2.TextBoxthd15.Text)) = fm2.TextBoxthd16.Text.ToString.Trim) Then
            If (Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxthd8.Text), CInt(fm2.TextBoxthd9.Text)) = fm2.TextBoxthd10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxthd11.Text), CInt(fm2.TextBoxthd12.Text)) = fm2.TextBoxthd13.Text.Trim.ToString) Then
                'MsgBox("这是三级标题")
                '段前距
                If Len(fm2.TextBoxthd00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxthd00.Text))
                '段后距
                If Len(fm2.TextBoxthd0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxthd0.Text))
                '左侧进
                If Len(fm2.TextBoxthd1.Text) > 0 Then oP.LeftIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxthd1.Text))
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                    oP.Range.Font.Underline = False
                End If
                '特殊格式
                If Len(fm2.TextBoxthd2.Text) > 0 Then oP.FirstLineIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxthd2.Text)）
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                    oP.Range.Font.Underline = False
                End If
                '行距
                If Len(fm2.TextBoxthd3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxthd3.Text))
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                    oP.Range.Font.Underline = False
                End If
                '对齐
                If Len(fm2.TextBoxthd4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxthd4.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                    oP.Range.Font.Underline = False
                End If
                '字型'
                If Len(fm2.TextBoxthd5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxthd5.Text)
                If Len(fm2.TextBoxthd6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxthd6.Text)
                If Len(fm2.TextBoxthd7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxthd7.Text)
                oP.Range.HighlightColorIndex = 0
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("三级标题已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button24_Click(sender As Object, e As RibbonControlEventArgs) Handles Button24.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            'If (oP.Alignment <> 1 And oP.Range.Characters.Count > 1 And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxf8.Text), CInt(fm2.TextBoxf9.Text)) <> fm2.TextBoxf10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxsec8.Text), CInt(fm2.TextBoxsec9.Text)) <> fm2.TextBoxsec10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxthd8.Text), CInt(fm2.TextBoxthd9.Text)) <> fm2.TextBoxthd10.Text.Trim.ToString) Then
            If (oP.Alignment <> 1 And Asc(Mid(Trim(oP.Range.Text.ToString), 1, 1)) <> 91) Then
                If (Asc(Mid(Trim(oP.Range.Text.ToString), 1, 1)) > 48 And Asc(Mid(Trim(oP.Range.Text.ToString), 1, 1)) < 57) Then

                Else
                    If (Mid(Trim(oP.Range.Text.ToString), Val(fm2.TextBoxAb8.Text), Val(fm2.TextBoxAb9.Text)) <> fm2.TextBoxAb10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxKey8.Text), CInt(fm2.TextBoxKey9.Text)) <> fm2.TextBoxKey10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxtz8.Text), CInt(fm2.TextBoxtz9.Text)) <> fm2.TextBoxtz10.Text.Trim.ToString) Then
                        'MsgBox("这是正文")
                        '段前距
                        If Len(fm2.TextBoxc00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxc00.Text))
                        '段后距
                        If Len(fm2.TextBoxc0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxc0.Text))
                        '左侧进
                        If Len(fm2.TextBoxc1.Text) > 0 Then oP.LeftIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxc1.Text))
                        If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                            oP.Range.Font.Underline = False
                        End If
                        '特殊格式
                        If Len(fm2.TextBoxc2.Text) > 0 Then oP.FirstLineIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxc2.Text))
                        If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                            oP.Range.Font.Underline = False
                        End If
                        '行距
                        If Len(fm2.TextBoxc3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CDbl(fm2.TextBoxc3.Text))
                        If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                            oP.Range.Font.Underline = False
                        End If
                        '对齐
                        If Len(fm2.TextBoxc4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxc4.Text)
                        If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                            oP.Range.Font.Underline = False
                        End If

                        '字型'
                        If Len(fm2.TextBoxc5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxc5.Text)
                        If Len(fm2.TextBoxc6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxc6.Text)
                        If Len(fm2.TextBoxc7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxc7.Text)
                        oP.Range.HighlightColorIndex = 0
                    End If
                End If
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("正文已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button25_Click(sender As Object, e As RibbonControlEventArgs) Handles Button25.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxtz8.Text), CInt(fm2.TextBoxtz9.Text)) = fm2.TextBoxtz10.Text.Trim.ToString) Then
                '这是图注
                '段前距
                If Len(fm2.TextBoxtz00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxtz00.Text))
                '段后距
                If Len(fm2.TextBoxtz0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxtz0.Text))
                '左侧进
                If Len(fm2.TextBoxtz1.Text) > 0 Then oP.LeftIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxtz1.Text)）
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                    oP.Range.Font.Underline = False
                End If
                '特殊格式
                If Len(fm2.TextBoxtz2.Text) > 0 Then oP.FirstLineIndent = wdapp.CentimetersToPoints(CSng(fm2.TextBoxtz2.Text)）
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                    oP.Range.Font.Underline = False
                End If
                '行距
                If Len(fm2.TextBoxtz3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxtz3.Text))
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                    oP.Range.Font.Underline = False
                End If
                '对齐
                If Len(fm2.TextBoxtz4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxtz4.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                    oP.Range.Font.Underline = False
                End If
                '字型'
                If Len(fm2.TextBoxtz5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxtz5.Text)
                If Len(fm2.TextBoxtz6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxtz6.Text)
                If Len(fm2.TextBoxtz7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxtz7.Text)
                oP.Range.HighlightColorIndex = 0
            End If
            For Each shape In wdapp.ActiveDocument.InlineShapes
                shape.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            Next
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("图注已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button23_Click(sender As Object, e As RibbonControlEventArgs) Handles Button23.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (Mid(Trim(oP.Range.Text.ToString), Val(fm2.TextBoxAb8.Text), Val(fm2.TextBoxAb9.Text)) = fm2.TextBoxAb10.Text.Trim.ToString) Then
                'MsgBox("这是摘要")

                '段前距
                If Len(fm2.TextBoxAb00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxAb00.Text))
                '段后距
                If Len(fm2.TextBoxAb0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxAb0.Text))
                '左侧进
                If Len(fm2.TextBoxAb1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxAb1.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                    oP.Range.Font.Underline = False
                End If
                '特殊格式
                If Len(fm2.TextBoxAb2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxAb2.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                    oP.Range.Font.Underline = False
                End If
                '行距
                If Len(fm2.TextBoxAb3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxAb3.Text))
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                    oP.Range.Font.Underline = False
                End If
                '对齐
                If Len(fm2.TextBoxAb4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxAb4.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                    oP.Range.Font.Underline = False
                End If

                '字型'
                If Len(fm2.TextBoxAb5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxAb5.Text)
                If Len(fm2.TextBoxAb6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxAb6.Text)
                If Len(fm2.TextBoxAb7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxAb7.Text)
                oP.Range.HighlightColorIndex = 0
                '标识符加粗
                Dim j As Long
                For j = Val(fm2.TextBoxAb8.Text) To Val(fm2.TextBoxAb9.Text)
                    oP.Range.Characters(j).Font.Bold = True
                Next
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("摘要已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button26_Click(sender As Object, e As RibbonControlEventArgs) Handles Button26.Click
        MsgBox("功能开发中", 0, "消息提示")
    End Sub

    Private Sub Button27_Click(sender As Object, e As RibbonControlEventArgs) Handles Button27.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Dim oP As Microsoft.Office.Interop.Word.Paragraph
        Dim i As Long
        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxKey8.Text), CInt(fm2.TextBoxKey9.Text)) = fm2.TextBoxKey10.Text.Trim.ToString) Then
                'MsgBox("这是关键词")

                '段前距
                If Len(fm2.TextBoxKey00.Text) > 0 Then oP.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxKey00.Text))
                '段后距
                If Len(fm2.TextBoxKey0.Text) > 0 Then oP.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxKey0.Text))
                '左侧进
                If Len(fm2.TextBoxKey1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxKey1.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                    oP.Range.Font.Underline = False
                End If
                '特殊格式
                If Len(fm2.TextBoxKey2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxKey2.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                    oP.Range.Font.Underline = False
                End If
                '行距
                If Len(fm2.TextBoxKey3.Text) > 0 Then oP.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxKey3.Text))
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange) Then
                    oP.Range.Font.Underline = False
                End If
                '对齐
                If Len(fm2.TextBoxKey4.Text) > 0 Then oP.Alignment = CInt(fm2.TextBoxKey4.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorRed) Then
                    oP.Range.Font.Underline = False
                End If
                '字型'
                If Len(fm2.TextBoxKey5.Text) > 0 Then oP.Range.Font.Name = CStr(fm2.TextBoxKey5.Text)
                If Len(fm2.TextBoxKey6.Text) > 0 Then oP.Range.Font.Size = CSng(fm2.TextBoxKey6.Text)
                If Len(fm2.TextBoxKey7.Text) > 0 Then oP.Range.Font.Bold = CBool(fm2.TextBoxKey7.Text)
                oP.Range.HighlightColorIndex = 0
                '标识符加粗
                Dim j As Long
                For j = Val(fm2.TextBoxKey8.Text) To Val(fm2.TextBoxKey9.Text)
                    oP.Range.Characters(j).Font.Bold = True
                Next
            End If
            If (i = wdapp.ActiveDocument.Paragraphs.Count) Then
                MsgBox("关键词已校对完毕", 0, "消息提示")
            End If
        Next
    End Sub

    Private Sub Button28_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button16_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        Call Button12_Click(sender, e)
        Call Button18_Click(sender, e)
        Call Button23_Click(sender, e)
        Call Button27_Click(sender, e)
        Call Button19_Click(sender, e)

        Call Button20_Click(sender, e)
        Call Button21_Click(sender, e)
        Call Button24_Click(sender, e)
        Call Button15_Click(sender, e)
        Call Button25_Click(sender, e)
        MsgBox("全文已校对完毕", 0, "消息提示")
    End Sub

    Private Sub Button29_Click(sender As Object, e As RibbonControlEventArgs) Handles Button29.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是总标题")
        '段前距
        If Len(fm2.TextBoxt00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxt00.Text))
        '段后距
        If Len(fm2.TextBoxt0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxt0.Text))
        '左侧进
        If Len(fm2.TextBoxt1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxt1.Text)
        '特殊格式
        If Len(fm2.TextBoxt2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxt2.Text)
        '行距
        If Len(fm2.TextBoxt3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.CentimetersToPoints(CSng(fm2.TextBoxt3.Text))
        '对齐
        If Len(fm2.TextBoxt4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxt4.Text)
        '字型'
        If Len(fm2.TextBoxt5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxt5.Text)
        If Len(fm2.TextBoxt6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxt6.Text)
        If Len(fm2.TextBoxt7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxt7.Text)
    End Sub

    Private Sub Button38_Click(sender As Object, e As RibbonControlEventArgs) Handles Button38.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是图注")
        '段前距
        If Len(fm2.TextBoxtz00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxtz00.Text))
        '段后距
        If Len(fm2.TextBoxtz0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxtz0.Text))
        '左侧进
        If Len(fm2.TextBoxtz1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxtz1.Text)
        '特殊格式
        If Len(fm2.TextBoxtz2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxtz2.Text)
        '行距
        If Len(fm2.TextBoxtz3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.CentimetersToPoints(CSng(fm2.TextBoxtz3.Text))
        '对齐
        If Len(fm2.TextBoxtz4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxtz4.Text)
        '字型'
        If Len(fm2.TextBoxtz5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxtz5.Text)
        If Len(fm2.TextBoxtz6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxtz6.Text)
        If Len(fm2.TextBoxtz7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxtz7.Text)
    End Sub

    Private Sub Button30_Click(sender As Object, e As RibbonControlEventArgs) Handles Button30.Click

        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是摘要")
        '段前距
        If Len(fm2.TextBoxAb00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxAb00.Text))
        '段后距
        If Len(fm2.TextBoxAb0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxAb0.Text))
        '左侧进
        If Len(fm2.TextBoxAb1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxAb1.Text)
        '特殊格式
        If Len(fm2.TextBoxAb2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxAb2.Text)
        '行距
        If Len(fm2.TextBoxAb3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxAb3.Text))
        '对齐
        If Len(fm2.TextBoxAb4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxAb4.Text)
        '字型'
        If Len(fm2.TextBoxAb5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxAb5.Text)
        If Len(fm2.TextBoxAb6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxAb6.Text)
        If Len(fm2.TextBoxAb7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxAb7.Text)
        wdapp.Selection.Characters(1).Bold = True
        wdapp.Selection.Characters(2).Bold = True
    End Sub

    Private Sub Button32_Click(sender As Object, e As RibbonControlEventArgs) Handles Button32.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是副标题")
        '段前距
        If Len(fm2.TextBoxs00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxs00.Text))
        '段后距
        If Len(fm2.TextBoxs0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxs0.Text))
        '左侧进
        If Len(fm2.TextBoxs1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxs1.Text)
        '特殊格式
        If Len(fm2.TextBoxs2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxs2.Text)
        '行距
        If Len(fm2.TextBoxs3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxs3.Text))
        '对齐
        If Len(fm2.TextBoxs4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxs4.Text)
        '字型'
        If Len(fm2.TextBoxs5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxs5.Text)
        If Len(fm2.TextBoxs6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxs6.Text)
        If Len(fm2.TextBoxs7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxs7.Text)

    End Sub

    Private Sub Button31_Click(sender As Object, e As RibbonControlEventArgs) Handles Button31.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是摘要")
        '段前距
        If Len(fm2.TextBoxKey00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxKey00.Text))
        '段后距
        If Len(fm2.TextBoxKey0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxKey0.Text))
        '左侧进
        If Len(fm2.TextBoxKey1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxKey1.Text)
        '特殊格式
        If Len(fm2.TextBoxKey2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxKey2.Text)
        '行距
        If Len(fm2.TextBoxKey3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxKey3.Text))
        '对齐
        If Len(fm2.TextBoxKey4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxKey4.Text)
        '字型'
        If Len(fm2.TextBoxKey5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxKey5.Text)
        If Len(fm2.TextBoxKey6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxKey6.Text)
        If Len(fm2.TextBoxKey7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxKey7.Text)
    End Sub

    Private Sub Button33_Click(sender As Object, e As RibbonControlEventArgs) Handles Button33.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是一级标题")
        '段前距
        If Len(fm2.TextBoxf00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxf00.Text))
        '段后距
        If Len(fm2.TextBoxf0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxf0.Text))
        '左侧进
        If Len(fm2.TextBoxf1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxf1.Text)
        '特殊格式
        If Len(fm2.TextBoxf2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxf2.Text)
        '行距
        If Len(fm2.TextBoxf3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxf3.Text))
        '对齐
        If Len(fm2.TextBoxf4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxf4.Text)
        '字型'
        If Len(fm2.TextBoxf5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxf5.Text)
        If Len(fm2.TextBoxf6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxf6.Text)
        If Len(fm2.TextBoxf7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxf7.Text)
        wdapp.Selection.Characters(1).Bold = True
        wdapp.Selection.Characters(2).Bold = True
        wdapp.Selection.Characters(3).Bold = True
    End Sub

    Private Sub Button34_Click(sender As Object, e As RibbonControlEventArgs) Handles Button34.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是二级标题")
        '段前距
        If Len(fm2.TextBoxsec00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxsec00.Text))
        '段后距
        If Len(fm2.TextBoxsec0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxsec0.Text))
        '左侧进
        If Len(fm2.TextBoxsec1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxsec1.Text)
        '特殊格式
        If Len(fm2.TextBoxsec2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxsec2.Text)
        '行距
        If Len(fm2.TextBoxsec3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxsec3.Text))
        '对齐
        If Len(fm2.TextBoxsec4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxsec4.Text)
        '字型'
        If Len(fm2.TextBoxsec5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxsec5.Text)
        If Len(fm2.TextBoxsec6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxsec6.Text)
        If Len(fm2.TextBoxsec7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxsec7.Text)
    End Sub

    Private Sub Button35_Click(sender As Object, e As RibbonControlEventArgs) Handles Button35.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是三级标题")
        '段前距
        If Len(fm2.TextBoxthd00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxthd00.Text))
        '段后距
        If Len(fm2.TextBoxthd0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxthd0.Text))
        '左侧进
        If Len(fm2.TextBoxthd1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxthd1.Text)
        '特殊格式
        If Len(fm2.TextBoxthd2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxthd2.Text)
        '行距
        If Len(fm2.TextBoxthd3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxthd3.Text))
        '对齐
        If Len(fm2.TextBoxthd4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxthd4.Text)
        '字型'
        If Len(fm2.TextBoxthd5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxthd5.Text)
        If Len(fm2.TextBoxthd6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxthd6.Text)
        If Len(fm2.TextBoxthd7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxthd7.Text)
    End Sub

    Private Sub Button36_Click(sender As Object, e As RibbonControlEventArgs) Handles Button36.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是正文")
        '段前距
        If Len(fm2.TextBoxc00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxc00.Text))
        '段后距
        If Len(fm2.TextBoxc0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxc0.Text))
        '左侧进
        If Len(fm2.TextBoxc1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxc1.Text)
        '特殊格式
        If Len(fm2.TextBoxc2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxc2.Text)
        '行距
        If Len(fm2.TextBoxc3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxc3.Text))
        '对齐
        If Len(fm2.TextBoxc4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxc4.Text)
        '字型'
        If Len(fm2.TextBoxc5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxc5.Text)
        If Len(fm2.TextBoxc6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxc6.Text)
        If Len(fm2.TextBoxc7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxc7.Text)
    End Sub

    Private Sub Button37_Click(sender As Object, e As RibbonControlEventArgs) Handles Button37.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        'MsgBox("这是参考文献")
        '段前距
        If Len(fm2.TextBoxr00.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceBefore = wdapp.LinesToPoints(CSng(fm2.TextBoxr00.Text))
        '段后距
        If Len(fm2.TextBoxr0.Text) > 0 Then wdapp.Selection.Range.Paragraphs.SpaceAfter = wdapp.LinesToPoints(CSng(fm2.TextBoxr0.Text))
        '左侧进
        If Len(fm2.TextBoxr1.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LeftIndent = CSng(fm2.TextBoxr1.Text)
        '特殊格式
        If Len(fm2.TextBoxr2.Text) > 0 Then wdapp.Selection.Range.Paragraphs.FirstLineIndent = CSng(fm2.TextBoxr2.Text)
        '行距
        If Len(fm2.TextBoxr3.Text) > 0 Then wdapp.Selection.Range.Paragraphs.LineSpacing = wdapp.LinesToPoints(CSng(fm2.TextBoxr3.Text))
        '对齐
        If Len(fm2.TextBoxr4.Text) > 0 Then wdapp.Selection.Range.Paragraphs.Alignment = CInt(fm2.TextBoxr4.Text)
        '字型'
        If Len(fm2.TextBoxr5.Text) > 0 Then wdapp.Selection.Range.Font.Name = CStr(fm2.TextBoxr5.Text)
        If Len(fm2.TextBoxr6.Text) > 0 Then wdapp.Selection.Range.Font.Size = CSng(fm2.TextBoxr6.Text)
        If Len(fm2.TextBoxr7.Text) > 0 Then wdapp.Selection.Range.Font.Bold = CBool(fm2.TextBoxr7.Text)
    End Sub

    Private Sub Button39_Click(sender As Object, e As RibbonControlEventArgs) Handles Button39.Click
        MsgBox("功能开发中", 0, "消息提示")
    End Sub

    Private Sub Button40_Click(sender As Object, e As RibbonControlEventArgs) Handles Button40.Click
        MsgBox("功能开发中", 0, "消息提示")
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application

        wdapp.Selection.InsertAfter（“论文标题（二号宋体，居中，加粗）"）
        Call Button29_Click(sender, e)


    End Sub

    Private Sub Button14_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        wdapp.Selection.InsertAfter（“作者1，作者2，作者3，.……（五号楷体，居中）"）
        Call Button32_Click(sender, e)
        wdapp.Selection.InsertParagraphAfter()
        wdapp.Selection.InsertAfter（“（1.学校院、系名，省份城市邮编；2.单位名称，省份城市邮编）（五号楷体，居中）"）
        Call Button32_Click(sender, e)
    End Sub

    Private Sub Button43_Click(sender As Object, e As RibbonControlEventArgs) Handles Button43.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        wdapp.Selection.InsertAfter（“ 正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正正文正文正文正文正文正文正文正文正，（五号宋体，段前缩进两格）"）
        Call Button36_Click(sender, e)

    End Sub

    Private Sub Button42_Click(sender As Object, e As RibbonControlEventArgs) Handles Button42.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        Call Button35_Click(sender, e)
        wdapp.Selection.InsertAfter（“3.1.1 三级标题（五号宋体，不加粗，顶格，序号和标题文字间有空格"）
    End Sub

    Private Sub Button11_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        wdapp.Selection.InsertAfter（“摘要:摘要内容摘要内容摘要内容摘要内容摘要内容摘要内容摘要内容摘要内容摘要内容摘内容摘要内容摘要内容摘要内容摘要内容摘要内容摘要内容摘要内容..…（小五号楷体）"）
        Call Button30_Click(sender, e)

    End Sub

    Private Sub Button13_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        wdapp.Selection.InsertAfter（“关键词；关键词；关键词；关键词（小五号楷体，全角分号隔开）"）
        Call Button31_Click(sender, e)


    End Sub

    Private Sub Button28_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button28.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        wdapp.Selection.InsertAfter（“1 一级标题（四号宋体，加粗，顶格，序号和标题文字间有空格"）
        Call Button33_Click(sender, e)

    End Sub

    Private Sub Button41_Click(sender As Object, e As RibbonControlEventArgs) Handles Button41.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        wdapp.Selection.InsertAfter（“2.1 二级标题（五号宋体，加粗，顶格，序号和标题文字间有空格"）
        Call Button34_Click(sender, e)

    End Sub

    Private Sub Button44_Click(sender As Object, e As RibbonControlEventArgs) Handles Button44.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        wdapp.Selection.InsertAfter（“[1]期刊——作者.题名[文献类型标志].刊名，出版年，卷（期）：起一止页码.（不要缺少页码）.（小五号宋体，缩进两格；序号和内容间空半格；内容中标点符号均使用半角，后空半格）"）
        Call Button37_Click(sender, e)

    End Sub

    Private Sub Button47_Click(sender As Object, e As RibbonControlEventArgs) Handles Button47.Click
        Dim wdapp As Word.Application = Globals.ThisAddIn.Application
        wdapp.Selection.InsertAfter（“图1. xxx示意图（图题使用小五号黑体，居中，列于图下）"）
        Call Button38_Click(sender, e)

    End Sub

    Private Sub Button45_Click(sender As Object, e As RibbonControlEventArgs) Handles Button45.Click
        MsgBox("功能开发中", 0, "消息提示")

    End Sub

    Private Sub Button46_Click(sender As Object, e As RibbonControlEventArgs) Handles Button46.Click
        MsgBox("功能开发中", 0, "消息提示")

    End Sub
End Class
