﻿Imports Microsoft.Office.Tools.Ribbon

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

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
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
                If (oP.Range.Font.Bold = True) Then
                    'MsgBox("这是总标题")
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
                    If Len(fm2.TextBoxt3.Text) > 0 Then oP.LineSpacing = CSng(fm2.TextBoxt3.Text)
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

    Private Sub 作者_Click(sender As Object, e As RibbonControlEventArgs) Handles 作者.Click
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

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
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

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
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

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click
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
        Dim qs, ws, pd As Int16
        If Len(fm2.TextBoxr8.Text) > 0 Then qs = CInt(fm2.TextBoxr8.Text)
        If Len(fm2.TextBoxr9.Text) > 0 Then ws = CInt(fm2.TextBoxr9.Text)
        If Len(fm2.TextBoxr10.Text) > 0 Then pd = CStr(fm2.TextBoxr10.Text)

        For i = 1 To wdapp.ActiveDocument.Paragraphs.Count
            oP = wdapp.ActiveDocument.Paragraphs(i)
            If (Mid(Trim(oP.Range.Text.ToString), qs, ws) = pd) Then
                If (oP.Range.Font.Bold = True) Then
                    'MsgBox("这是一级标题")
                    '左侧进
                    If Len(fm2.TextBoxr1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxr1.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If Len(fm2.TextBoxr2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxr2.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxr3.Text) > 0 Then oP.LineSpacing = CSng(fm2.TextBoxr3.Text)
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

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
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
                If (oP.Range.Font.Bold = False) Then
                    'MsgBox("这是副标题")
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
                    If Len(fm2.TextBoxs3.Text) > 0 Then oP.LineSpacing = CSng(fm2.TextBoxs3.Text)
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
            If (Mid(Trim(oP.Range.Text.ToString), Val(fm2.TextBoxf8.Text), Val(fm2.TextBoxf9.Text)) = fm2.TextBoxf10.Text.Trim.ToString) Then
                If (oP.Range.Font.Bold = True) Then
                    MsgBox("这是一级标题")
                    '左侧进
                    If Len(fm2.TextBoxf1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxf1.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If Len(fm2.TextBoxf2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxf2.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxf3.Text) > 0 Then oP.LineSpacing = CSng(fm2.TextBoxf3.Text)
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
            If (Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxsec8.Text), CInt(fm2.TextBoxsec9.Text)) = fm2.TextBoxsec10.Text.Trim.ToString) Then
                If (oP.Range.Font.Bold = True) Then
                    'MsgBox("这是二级标题")
                    '左侧进
                    If Len(fm2.TextBoxsec1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxsec1.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If Len(fm2.TextBoxsec2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxsec2.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxsec3.Text) > 0 Then oP.LineSpacing = CSng(fm2.TextBoxsec3.Text)
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
            If (Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxthd8.Text), CInt(fm2.TextBoxthd9.Text)) = fm2.TextBoxthd10.Text.Trim.ToString) Then
                If (oP.Range.Font.Bold = True) Then
                    'MsgBox("这是三级标题")
                    '左侧进
                    If Len(fm2.TextBoxthd1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxthd1.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                        oP.Range.Font.Underline = False
                    End If
                    '特殊格式
                    If Len(fm2.TextBoxthd2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxthd2.Text)
                    If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                        oP.Range.Font.Underline = False
                    End If
                    '行距
                    If Len(fm2.TextBoxthd3.Text) > 0 Then oP.LineSpacing = CSng(fm2.TextBoxthd3.Text)
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
            If (oP.Alignment <> 1 And oP.Range.Characters.Count > 1 And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxf8.Text), CInt(fm2.TextBoxf9.Text)) <> fm2.TextBoxf10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxsec8.Text), CInt(fm2.TextBoxsec9.Text)) <> fm2.TextBoxsec10.Text.Trim.ToString And Mid(Trim(oP.Range.Text.ToString), CInt(fm2.TextBoxthd8.Text), CInt(fm2.TextBoxthd9.Text)) <> fm2.TextBoxthd10.Text.Trim.ToString) Then

                'MsgBox("这是正文")
                '左侧进
                If Len(fm2.TextBoxc1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxc1.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                    oP.Range.Font.Underline = False
                End If
                '特殊格式
                If Len(fm2.TextBoxc2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxc2.Text)
                If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                    oP.Range.Font.Underline = False
                End If
                '行距
                If Len(fm2.TextBoxc3.Text) > 0 Then oP.LineSpacing = CSng(fm2.TextBoxc3.Text)
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
            '左侧进
            If Len(fm2.TextBoxc1.Text) > 0 Then oP.LeftIndent = CSng(fm2.TextBoxc1.Text)
            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorBlue) Then
                oP.Range.Font.Underline = False
            End If
            '特殊格式
            If Len(fm2.TextBoxc2.Text) > 0 Then oP.FirstLineIndent = CSng(fm2.TextBoxc2.Text)
            If (oP.Range.Font.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen) Then
                oP.Range.Font.Underline = False
            End If
            '行距
            If Len(fm2.TextBoxc3.Text) > 0 Then oP.LineSpacing = CSng(fm2.TextBoxc3.Text)
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

        Next
    End Sub

    Private Sub Button23_Click(sender As Object, e As RibbonControlEventArgs) Handles Button23.Click

    End Sub
End Class
