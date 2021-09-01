Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon

Public Class Form1
    Public Figurl1 As String
    Public xTextbox As TextBox
    Public v0 As String
    Dim fileOpendialog As OpenFileDialog

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = “http://home.ustc.edu.cn/~yunju/RAP/01.jpeg”
        TextBox2.Text = “http://home.ustc.edu.cn/~yunju/RAP/02.jpeg”
        TextBox3.Text = “http://home.ustc.edu.cn/~yunju/RAP/03.jpeg”
        TextBox4.Text = “http://home.ustc.edu.cn/~yunju/RAP/04.jpeg”
        TextBox5.Text = “http://home.ustc.edu.cn/~yunju/RAP/05.jpeg”
        TextBox6.Text = “http://home.ustc.edu.cn/~yunju/RAP/06.jpeg”
        TextBox7.Text = “http://home.ustc.edu.cn/~yunju/RAP/07.jpeg”
        TextBox8.Text = “http://home.ustc.edu.cn/~yunju/RAP/08.jpeg”
        TextBox9.Text = “http://home.ustc.edu.cn/~yunju/RAP/09.jpeg”
        TextBox10.Text = “http://home.ustc.edu.cn/~yunju/RAP/10.jpeg”
        TextBox11.Text = “http://home.ustc.edu.cn/~yunju/RAP/11.jpeg”
        TextBox12.Text = “http://home.ustc.edu.cn/~yunju/RAP/12.jpeg”
        TextBox13.Text = “http://home.ustc.edu.cn/~yunju/RAP/13.jpeg”
        TextBox14.Text = “http://home.ustc.edu.cn/~yunju/RAP/14.jpeg”
        TextBox15.Text = “http://home.ustc.edu.cn/~yunju/RAP/15.jpeg”
        TextBox16.Text = “http://home.ustc.edu.cn/~yunju/RAP/16.jpeg”
        TextBox17.Text = “http://home.ustc.edu.cn/~yunju/RAP/17.jpeg”
        TextBox18.Text = “http://home.ustc.edu.cn/~yunju/RAP/18.jpeg”
        TextBox19.Text = “http://home.ustc.edu.cn/~yunju/RAP/19.jpeg”
        TextBox20.Text = “http://home.ustc.edu.cn/~yunju/RAP/20.jpeg”
        TextBox21.Text = “http://home.ustc.edu.cn/~yunju/RAP/21.jpeg”
        TextBox22.Text = “http://home.ustc.edu.cn/~yunju/RAP/22.jpeg”
        TextBox23.Text = “http://home.ustc.edu.cn/~yunju/RAP/23.jpeg”
        TextBox24.Text = “http://home.ustc.edu.cn/~yunju/RAP/24.jpeg”
        TextBox25.Text = “http://home.ustc.edu.cn/~yunju/RAP/25.jpeg”
        TextBox26.Text = “http://home.ustc.edu.cn/~yunju/RAP/26.jpeg”
        TextBox27.Text = “http://home.ustc.edu.cn/~yunju/RAP/27.jpeg”
        TextBox28.Text = “http://home.ustc.edu.cn/~yunju/RAP/28.jpeg”
        TextBox29.Text = “http://home.ustc.edu.cn/~yunju/RAP/29.jpeg”
        TextBox30.Text = “http://home.ustc.edu.cn/~yunju/RAP/30.jpeg”
        TextBox31.Text = “http://home.ustc.edu.cn/~yunju/RAP/31.jpeg”
        TextBox32.Text = “http://home.ustc.edu.cn/~yunju/RAP/32.jpeg”
        TextBox33.Text = “http://home.ustc.edu.cn/~yunju/RAP/33.jpeg”
        TextBox34.Text = “http://home.ustc.edu.cn/~yunju/RAP/34.jpeg”
        TextBox35.Text = “http://home.ustc.edu.cn/~yunju/RAP/35.jpeg”
        TextBox36.Text = “http://home.ustc.edu.cn/~yunju/RAP/36.jpeg”
        TextBox37.Text = “http://home.ustc.edu.cn/~yunju/RAP/37.jpeg”
        TextBox38.Text = “http://home.ustc.edu.cn/~yunju/RAP/38.jpeg”
        TextBox39.Text = “http://home.ustc.edu.cn/~yunju/RAP/39.jpeg”
        TextBox40.Text = “http://home.ustc.edu.cn/~yunju/RAP/40.jpeg”
        TextBox41.Text = “http://home.ustc.edu.cn/~yunju/RAP/41.jpeg”
        TextBox42.Text = “http://home.ustc.edu.cn/~yunju/RAP/42.jpeg”
        TextBox43.Text = “http://home.ustc.edu.cn/~yunju/RAP/43.jpeg”
        TextBox44.Text = “http://home.ustc.edu.cn/~yunju/RAP/44.jpeg”
        TextBox45.Text = “http://home.ustc.edu.cn/~yunju/RAP/45.jpeg”
        TextBox46.Text = “http://home.ustc.edu.cn/~yunju/RAP/46.jpeg”
        TextBox47.Text = “http://home.ustc.edu.cn/~yunju/RAP/47.jpeg”
        TextBox48.Text = “http://home.ustc.edu.cn/~yunju/RAP/48.jpeg”
        TextBox49.Text = “http://home.ustc.edu.cn/~yunju/RAP/49.jpeg”
        TextBox50.Text = “http://home.ustc.edu.cn/~yunju/RAP/50.jpeg”
        TextBox51.Text = “http://home.ustc.edu.cn/~yunju/RAP/51.jpeg”
        TextBox52.Text = “http://home.ustc.edu.cn/~yunju/RAP/52.jpeg”
        TextBox53.Text = “http://home.ustc.edu.cn/~yunju/RAP/53.jpeg”
        TextBox54.Text = “http://home.ustc.edu.cn/~yunju/RAP/54.jpeg”
        TextBox55.Text = “http://home.ustc.edu.cn/~yunju/RAP/55.jpeg”
        TextBox56.Text = “http://home.ustc.edu.cn/~yunju/RAP/56.jpeg”
        TextBox57.Text = “http://home.ustc.edu.cn/~yunju/RAP/57.jpeg”
        TextBox58.Text = “http://home.ustc.edu.cn/~yunju/RAP/58.jpeg”
        TextBox59.Text = “http://home.ustc.edu.cn/~yunju/RAP/59.jpeg”
        TextBox60.Text = “http://home.ustc.edu.cn/~yunju/RAP/60.jpeg”
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.ActiveForm.Visible = False
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox1.Text = TextBox1.Text
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.Text = TextBox2.Text
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        TextBox3.Text = TextBox3.Text
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        TextBox4.Text = TextBox4.Text
    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        TextBox5.Text = TextBox5.Text
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        TextBox6.Text = TextBox6.Text
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        TextBox7.Text = TextBox7.Text
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        TextBox8.Text = TextBox8.Text
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        TextBox9.Text = TextBox9.Text
    End Sub
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        TextBox10.Text = TextBox10.Text
    End Sub
    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        TextBox11.Text = TextBox11.Text
    End Sub
    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        TextBox12.Text = TextBox12.Text
    End Sub
    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        TextBox13.Text = TextBox13.Text
    End Sub
    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        TextBox14.Text = TextBox14.Text
    End Sub
    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        TextBox15.Text = TextBox15.Text
    End Sub
    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        TextBox16.Text = TextBox16.Text
    End Sub
    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        TextBox17.Text = TextBox17.Text
    End Sub
    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        TextBox18.Text = TextBox18.Text
    End Sub
    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        TextBox19.Text = TextBox19.Text
    End Sub
    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged
        TextBox20.Text = TextBox20.Text
    End Sub
    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        TextBox21.Text = TextBox21.Text
    End Sub
    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged
        TextBox22.Text = TextBox22.Text
    End Sub
    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged
        TextBox23.Text = TextBox23.Text
    End Sub
    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        TextBox24.Text = TextBox24.Text
    End Sub
    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged
        TextBox25.Text = TextBox25.Text
    End Sub
    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged
        TextBox26.Text = TextBox26.Text
    End Sub
    Private Sub TextBox27_TextChanged(sender As Object, e As EventArgs) Handles TextBox27.TextChanged
        TextBox27.Text = TextBox27.Text
    End Sub
    Private Sub TextBox28_TextChanged(sender As Object, e As EventArgs) Handles TextBox28.TextChanged
        TextBox28.Text = TextBox28.Text
    End Sub
    Private Sub TextBox29_TextChanged(sender As Object, e As EventArgs) Handles TextBox29.TextChanged
        TextBox29.Text = TextBox29.Text
    End Sub
    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged
        TextBox30.Text = TextBox30.Text
    End Sub
    Private Sub TextBox31_TextChanged(sender As Object, e As EventArgs) Handles TextBox31.TextChanged
        TextBox31.Text = TextBox31.Text
    End Sub
    Private Sub TextBox32_TextChanged(sender As Object, e As EventArgs) Handles TextBox32.TextChanged
        TextBox32.Text = TextBox32.Text
    End Sub
    Private Sub TextBox33_TextChanged(sender As Object, e As EventArgs) Handles TextBox33.TextChanged
        TextBox33.Text = TextBox33.Text
    End Sub
    Private Sub TextBox34_TextChanged(sender As Object, e As EventArgs) Handles TextBox34.TextChanged
        TextBox34.Text = TextBox34.Text
    End Sub
    Private Sub TextBox35_TextChanged(sender As Object, e As EventArgs) Handles TextBox35.TextChanged
        TextBox35.Text = TextBox35.Text
    End Sub
    Private Sub TextBox36_TextChanged(sender As Object, e As EventArgs) Handles TextBox36.TextChanged
        TextBox36.Text = TextBox36.Text
    End Sub
    Private Sub TextBox37_TextChanged(sender As Object, e As EventArgs) Handles TextBox37.TextChanged
        TextBox37.Text = TextBox37.Text
    End Sub
    Private Sub TextBox38_TextChanged(sender As Object, e As EventArgs) Handles TextBox38.TextChanged
        TextBox38.Text = TextBox38.Text
    End Sub
    Private Sub TextBox39_TextChanged(sender As Object, e As EventArgs) Handles TextBox39.TextChanged
        TextBox39.Text = TextBox39.Text
    End Sub
    Private Sub TextBox40_TextChanged(sender As Object, e As EventArgs) Handles TextBox40.TextChanged
        TextBox40.Text = TextBox40.Text
    End Sub
    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles TextBox41.TextChanged
        TextBox41.Text = TextBox41.Text
    End Sub
    Private Sub TextBox42_TextChanged(sender As Object, e As EventArgs) Handles TextBox42.TextChanged
        TextBox42.Text = TextBox42.Text
    End Sub
    Private Sub TextBox43_TextChanged(sender As Object, e As EventArgs) Handles TextBox43.TextChanged
        TextBox43.Text = TextBox43.Text
    End Sub
    Private Sub TextBox44_TextChanged(sender As Object, e As EventArgs) Handles TextBox44.TextChanged
        TextBox44.Text = TextBox44.Text
    End Sub
    Private Sub TextBox45_TextChanged(sender As Object, e As EventArgs) Handles TextBox45.TextChanged
        TextBox45.Text = TextBox45.Text
    End Sub
    Private Sub TextBox46_TextChanged(sender As Object, e As EventArgs) Handles TextBox46.TextChanged
        TextBox46.Text = TextBox46.Text
    End Sub
    Private Sub TextBox47_TextChanged(sender As Object, e As EventArgs) Handles TextBox47.TextChanged
        TextBox47.Text = TextBox47.Text
    End Sub
    Private Sub TextBox48_TextChanged(sender As Object, e As EventArgs) Handles TextBox48.TextChanged
        TextBox48.Text = TextBox48.Text
    End Sub
    Private Sub TextBox49_TextChanged(sender As Object, e As EventArgs) Handles TextBox49.TextChanged
        TextBox49.Text = TextBox49.Text
    End Sub
    Private Sub TextBox50_TextChanged(sender As Object, e As EventArgs) Handles TextBox50.TextChanged
        TextBox50.Text = TextBox50.Text
    End Sub
    Private Sub TextBox51_TextChanged(sender As Object, e As EventArgs) Handles TextBox51.TextChanged
        TextBox51.Text = TextBox51.Text
    End Sub
    Private Sub TextBox52_TextChanged(sender As Object, e As EventArgs) Handles TextBox52.TextChanged
        TextBox52.Text = TextBox52.Text
    End Sub
    Private Sub TextBox53_TextChanged(sender As Object, e As EventArgs) Handles TextBox53.TextChanged
        TextBox53.Text = TextBox53.Text
    End Sub
    Private Sub TextBox54_TextChanged(sender As Object, e As EventArgs) Handles TextBox54.TextChanged
        TextBox54.Text = TextBox54.Text
    End Sub
    Private Sub TextBox55_TextChanged(sender As Object, e As EventArgs) Handles TextBox55.TextChanged
        TextBox55.Text = TextBox55.Text
    End Sub
    Private Sub TextBox56_TextChanged(sender As Object, e As EventArgs) Handles TextBox56.TextChanged
        TextBox56.Text = TextBox56.Text
    End Sub
    Private Sub TextBox57_TextChanged(sender As Object, e As EventArgs) Handles TextBox57.TextChanged
        TextBox57.Text = TextBox57.Text
    End Sub
    Private Sub TextBox58_TextChanged(sender As Object, e As EventArgs) Handles TextBox58.TextChanged
        TextBox58.Text = TextBox58.Text
    End Sub
    Private Sub TextBox59_TextChanged(sender As Object, e As EventArgs) Handles TextBox59.TextChanged
        TextBox59.Text = TextBox59.Text
    End Sub
    Private Sub TextBox60_TextChanged(sender As Object, e As EventArgs) Handles TextBox60.TextChanged
        TextBox60.Text = TextBox60.Text
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = “图01”
        Label2.Text = “图02”
        Label3.Text = “图03”
        Label4.Text = “图04”
        Label5.Text = “图05”
        Label6.Text = “图06”
        Label7.Text = “图07”
        Label8.Text = “图08”
        Label9.Text = “图09”
        Label10.Text = “图10”
        Label11.Text = “图11”
        Label12.Text = “图12”
        Label13.Text = “图13”
        Label14.Text = “图14”
        Label15.Text = “图15”
        Label16.Text = “图16”
        Label17.Text = “图17”
        Label18.Text = “图18”
        Label19.Text = “图19”
        Label20.Text = “图20”
        Label21.Text = “图21”
        Label22.Text = “图22”
        Label23.Text = “图23”
        Label24.Text = “图24”
        Label25.Text = “图25”
        Label26.Text = “图26”
        Label27.Text = “图27”
        Label28.Text = “图28”
        Label29.Text = “图29”
        Label30.Text = “图30”
        Label31.Text = “图31”
        Label32.Text = “图32”
        Label33.Text = “图33”
        Label34.Text = “图34”
        Label35.Text = “图35”
        Label36.Text = “图36”
        Label37.Text = “图37”
        Label38.Text = “图38”
        Label39.Text = “图39”
        Label40.Text = “图40”
        Label41.Text = “图41”
        Label42.Text = “图42”
        Label43.Text = “图43”
        Label44.Text = “图44”
        Label45.Text = “图45”
        Label46.Text = “图46”
        Label47.Text = “图47”
        Label48.Text = “图48”
        Label49.Text = “图49”
        Label50.Text = “图50”
        Label51.Text = “图51”
        Label52.Text = “图52”
        Label53.Text = “图53”
        Label54.Text = “图54”
        Label55.Text = “图55”
        Label56.Text = “图56”
        Label57.Text = “图57”
        Label58.Text = “图58”
        Label59.Text = “图59”
        Label60.Text = “图60”
        PictureBox1.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox2.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox3.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox4.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox5.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox6.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox7.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox8.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox9.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox10.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox11.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox12.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox13.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox14.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox15.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox16.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox17.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox18.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox19.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox20.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox21.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox22.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox23.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox24.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox25.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox26.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox27.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox28.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox29.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox30.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox31.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox32.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox33.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox34.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox35.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox36.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox37.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox38.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox39.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox40.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox41.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox42.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox43.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox44.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox45.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox46.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox47.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox48.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox49.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox50.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox51.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox52.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox53.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox54.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox55.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox56.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox57.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox58.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox59.BorderStyle = PictureBox1.BorderStyle.FixedSingle
        PictureBox60.BorderStyle = PictureBox1.BorderStyle.FixedSingle
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim excelApp As Microsoft.Office.Interop.Excel.Application
            Dim book As Microsoft.Office.Interop.Excel.Workbook '定义book为的工作簿
            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet '定义sheet为的工作表
            excelApp = New Microsoft.Office.Interop.Excel.Application
            book = excelApp.Workbooks.Add '新建book
            sheet = book.Sheets(1) '新建sheets(1)

            sheet.Range("A1").Value = Label1.Text
            sheet.Range("B1").Value = TextBox1.Text
            sheet.Range("A2").Value = Label2.Text
            sheet.Range("B2").Value = TextBox2.Text
            sheet.Range("A3").Value = Label3.Text
            sheet.Range("B3").Value = TextBox3.Text
            sheet.Range("A4").Value = Label4.Text
            sheet.Range("B4").Value = TextBox4.Text
            sheet.Range("A5").Value = Label5.Text
            sheet.Range("B5").Value = TextBox5.Text
            sheet.Range("A6").Value = Label6.Text
            sheet.Range("B6").Value = TextBox6.Text
            sheet.Range("A7").Value = Label7.Text
            sheet.Range("B7").Value = TextBox7.Text
            sheet.Range("A8").Value = Label8.Text
            sheet.Range("B8").Value = TextBox8.Text
            sheet.Range("A9").Value = Label9.Text
            sheet.Range("B9").Value = TextBox9.Text
            sheet.Range("A10").Value = Label10.Text
            sheet.Range("B10").Value = TextBox10.Text
            sheet.Range("A11").Value = Label11.Text
            sheet.Range("B11").Value = TextBox11.Text
            sheet.Range("A12").Value = Label12.Text
            sheet.Range("B12").Value = TextBox12.Text
            sheet.Range("A13").Value = Label13.Text
            sheet.Range("B13").Value = TextBox13.Text
            sheet.Range("A14").Value = Label14.Text
            sheet.Range("B14").Value = TextBox14.Text
            sheet.Range("A15").Value = Label15.Text
            sheet.Range("B15").Value = TextBox15.Text
            sheet.Range("A16").Value = Label16.Text
            sheet.Range("B16").Value = TextBox16.Text
            sheet.Range("A17").Value = Label17.Text
            sheet.Range("B17").Value = TextBox17.Text
            sheet.Range("A18").Value = Label18.Text
            sheet.Range("B18").Value = TextBox18.Text
            sheet.Range("A19").Value = Label19.Text
            sheet.Range("B19").Value = TextBox19.Text
            sheet.Range("A20").Value = Label20.Text
            sheet.Range("B20").Value = TextBox20.Text
            sheet.Range("A21").Value = Label21.Text
            sheet.Range("B21").Value = TextBox21.Text
            sheet.Range("A22").Value = Label22.Text
            sheet.Range("B22").Value = TextBox22.Text
            sheet.Range("A23").Value = Label23.Text
            sheet.Range("B23").Value = TextBox23.Text
            sheet.Range("A24").Value = Label24.Text
            sheet.Range("B24").Value = TextBox24.Text
            sheet.Range("A25").Value = Label25.Text
            sheet.Range("B25").Value = TextBox25.Text
            sheet.Range("A26").Value = Label26.Text
            sheet.Range("B26").Value = TextBox26.Text
            sheet.Range("A27").Value = Label27.Text
            sheet.Range("B27").Value = TextBox27.Text
            sheet.Range("A28").Value = Label28.Text
            sheet.Range("B28").Value = TextBox28.Text
            sheet.Range("A29").Value = Label29.Text
            sheet.Range("B29").Value = TextBox29.Text
            sheet.Range("A30").Value = Label30.Text
            sheet.Range("B30").Value = TextBox30.Text
            sheet.Range("A31").Value = Label31.Text
            sheet.Range("B31").Value = TextBox31.Text
            sheet.Range("A32").Value = Label32.Text
            sheet.Range("B32").Value = TextBox32.Text
            sheet.Range("A33").Value = Label33.Text
            sheet.Range("B33").Value = TextBox33.Text
            sheet.Range("A34").Value = Label34.Text
            sheet.Range("B34").Value = TextBox34.Text
            sheet.Range("A35").Value = Label35.Text
            sheet.Range("B35").Value = TextBox35.Text
            sheet.Range("A36").Value = Label36.Text
            sheet.Range("B36").Value = TextBox36.Text
            sheet.Range("A37").Value = Label37.Text
            sheet.Range("B37").Value = TextBox37.Text
            sheet.Range("A38").Value = Label38.Text
            sheet.Range("B38").Value = TextBox38.Text
            sheet.Range("A39").Value = Label39.Text
            sheet.Range("B39").Value = TextBox39.Text
            sheet.Range("A40").Value = Label40.Text
            sheet.Range("B40").Value = TextBox40.Text
            sheet.Range("A41").Value = Label41.Text
            sheet.Range("B41").Value = TextBox41.Text
            sheet.Range("A42").Value = Label42.Text
            sheet.Range("B42").Value = TextBox42.Text
            sheet.Range("A43").Value = Label43.Text
            sheet.Range("B43").Value = TextBox43.Text
            sheet.Range("A44").Value = Label44.Text
            sheet.Range("B44").Value = TextBox44.Text
            sheet.Range("A45").Value = Label45.Text
            sheet.Range("B45").Value = TextBox45.Text
            sheet.Range("A46").Value = Label46.Text
            sheet.Range("B46").Value = TextBox46.Text
            sheet.Range("A47").Value = Label47.Text
            sheet.Range("B47").Value = TextBox47.Text
            sheet.Range("A48").Value = Label48.Text
            sheet.Range("B48").Value = TextBox48.Text
            sheet.Range("A49").Value = Label49.Text
            sheet.Range("B49").Value = TextBox49.Text
            sheet.Range("A50").Value = Label50.Text
            sheet.Range("B50").Value = TextBox50.Text
            sheet.Range("A51").Value = Label51.Text
            sheet.Range("B51").Value = TextBox51.Text
            sheet.Range("A52").Value = Label52.Text
            sheet.Range("B52").Value = TextBox52.Text
            sheet.Range("A53").Value = Label53.Text
            sheet.Range("B53").Value = TextBox53.Text
            sheet.Range("A54").Value = Label54.Text
            sheet.Range("B54").Value = TextBox54.Text
            sheet.Range("A55").Value = Label55.Text
            sheet.Range("B55").Value = TextBox55.Text
            sheet.Range("A56").Value = Label56.Text
            sheet.Range("B56").Value = TextBox56.Text
            sheet.Range("A57").Value = Label57.Text
            sheet.Range("B57").Value = TextBox57.Text
            sheet.Range("A58").Value = Label58.Text
            sheet.Range("B58").Value = TextBox58.Text
            sheet.Range("A59").Value = Label59.Text
            sheet.Range("B59").Value = TextBox59.Text
            sheet.Range("A60").Value = Label60.Text
            sheet.Range("B60").Value = TextBox60.Text

            Dim filename As String
            Dim sfd As New SaveFileDialog()
            sfd.Filter = "表格文件|*.xls;*.xlsx"
            If sfd.ShowDialog <> DialogResult.OK Then
                Exit Sub
            Else
                filename = sfd.FileName
            End If
            sheet.SaveAs(filename)
            book.Close()
            excelApp.Quit()
            releaseObject(excelApp)
            releaseObject(book)
            releaseObject(sheet)
            MsgBox("所有图片路径已保存到" + filename)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        fileOpendialog = New OpenFileDialog
        fileOpendialog.ShowDialog() '打开文件选择框
        TextBox1.Text = fileOpendialog.FileName  '得到选择的文件
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label1.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label2.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label3.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label4.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label5.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label6.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label7.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label8.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label9.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label10.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label11.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label12.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label13.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label14.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label15.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label16.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label17.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label18.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label19.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label20.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label21.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label22.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label23.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label24.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label25.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label26.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label27.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label28_Click(sender As Object, e As EventArgs) Handles Label28.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label28.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label29_Click(sender As Object, e As EventArgs) Handles Label29.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label29.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label30.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label31_Click(sender As Object, e As EventArgs) Handles Label31.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label31.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label32_Click(sender As Object, e As EventArgs) Handles Label32.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label32.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label33_Click(sender As Object, e As EventArgs) Handles Label33.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label33.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label34.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label35_Click(sender As Object, e As EventArgs) Handles Label35.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label35.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label36.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label37_Click(sender As Object, e As EventArgs) Handles Label37.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label37.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label38_Click(sender As Object, e As EventArgs) Handles Label38.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label38.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label39_Click(sender As Object, e As EventArgs) Handles Label39.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label39.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label40_Click(sender As Object, e As EventArgs) Handles Label40.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label40.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label41_Click(sender As Object, e As EventArgs) Handles Label41.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label41.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label42_Click(sender As Object, e As EventArgs) Handles Label42.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label42.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label43_Click(sender As Object, e As EventArgs) Handles Label43.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label43.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label44_Click(sender As Object, e As EventArgs) Handles Label44.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label44.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label45_Click(sender As Object, e As EventArgs) Handles Label45.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label45.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label46_Click(sender As Object, e As EventArgs) Handles Label46.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label46.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label47_Click(sender As Object, e As EventArgs) Handles Label47.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label47.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label48_Click(sender As Object, e As EventArgs) Handles Label48.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label48.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label49_Click(sender As Object, e As EventArgs) Handles Label49.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label49.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label50_Click(sender As Object, e As EventArgs) Handles Label50.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label50.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label51_Click(sender As Object, e As EventArgs) Handles Label51.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label51.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label52_Click(sender As Object, e As EventArgs) Handles Label52.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label52.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label53_Click(sender As Object, e As EventArgs) Handles Label53.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label53.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label54_Click(sender As Object, e As EventArgs) Handles Label54.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label54.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label55_Click(sender As Object, e As EventArgs) Handles Label55.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label55.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label56_Click(sender As Object, e As EventArgs) Handles Label56.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label56.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label57_Click(sender As Object, e As EventArgs) Handles Label57.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label57.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label58_Click(sender As Object, e As EventArgs) Handles Label58.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label58.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label59_Click(sender As Object, e As EventArgs) Handles Label59.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label59.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub
    Private Sub Label60_Click(sender As Object, e As EventArgs) Handles Label60.Click
        v0 = InputBox("请输入图片备注", "数据输入框")
        If Len(v0) > 0 Then
            Label60.Text = v0
        Else
            MsgBox("图片备注不能为空")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim excelApp As Microsoft.Office.Interop.Excel.Application
            Dim book As Microsoft.Office.Interop.Excel.Workbook '定义book为的工作簿
            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet '定义sheet为的工作表
            Dim range As Microsoft.Office.Interop.Excel.Range

            Dim filename As String
            Dim sfd As New OpenFileDialog()
            sfd.Filter = "表格文件|*.xls;*.xlsx"
            If sfd.ShowDialog <> DialogResult.OK Then
                Exit Sub
            Else
                filename = sfd.FileName
            End If
            excelApp = New Microsoft.Office.Interop.Excel.Application
            book = excelApp.Workbooks.Open(filename)
            sheet = book.Worksheets(1)

            range = sheet.Range("A1")
            Label1.Text = range.Value
            range = sheet.Range("B1")
            TextBox1.Text = range.Value
            range = sheet.Range("A2")
            Label2.Text = range.Value
            range = sheet.Range("B2")
            TextBox2.Text = range.Value
            range = sheet.Range("A3")
            Label3.Text = range.Value
            range = sheet.Range("B3")
            TextBox3.Text = range.Value
            range = sheet.Range("A4")
            Label4.Text = range.Value
            range = sheet.Range("B4")
            TextBox4.Text = range.Value
            range = sheet.Range("A5")
            Label5.Text = range.Value
            range = sheet.Range("B5")
            TextBox5.Text = range.Value
            range = sheet.Range("A6")
            Label6.Text = range.Value
            range = sheet.Range("B6")
            TextBox6.Text = range.Value
            range = sheet.Range("A7")
            Label7.Text = range.Value
            range = sheet.Range("B7")
            TextBox7.Text = range.Value
            range = sheet.Range("A8")
            Label8.Text = range.Value
            range = sheet.Range("B8")
            TextBox8.Text = range.Value
            range = sheet.Range("A9")
            Label9.Text = range.Value
            range = sheet.Range("B9")
            TextBox9.Text = range.Value
            range = sheet.Range("A10")
            Label10.Text = range.Value
            range = sheet.Range("B10")
            TextBox10.Text = range.Value
            range = sheet.Range("A11")
            Label11.Text = range.Value
            range = sheet.Range("B11")
            TextBox11.Text = range.Value
            range = sheet.Range("A12")
            Label12.Text = range.Value
            range = sheet.Range("B12")
            TextBox12.Text = range.Value
            range = sheet.Range("A13")
            Label13.Text = range.Value
            range = sheet.Range("B13")
            TextBox13.Text = range.Value
            range = sheet.Range("A14")
            Label14.Text = range.Value
            range = sheet.Range("B14")
            TextBox14.Text = range.Value
            range = sheet.Range("A15")
            Label15.Text = range.Value
            range = sheet.Range("B15")
            TextBox15.Text = range.Value
            range = sheet.Range("A16")
            Label16.Text = range.Value
            range = sheet.Range("B16")
            TextBox16.Text = range.Value
            range = sheet.Range("A17")
            Label17.Text = range.Value
            range = sheet.Range("B17")
            TextBox17.Text = range.Value
            range = sheet.Range("A18")
            Label18.Text = range.Value
            range = sheet.Range("B18")
            TextBox18.Text = range.Value
            range = sheet.Range("A19")
            Label19.Text = range.Value
            range = sheet.Range("B19")
            TextBox19.Text = range.Value
            range = sheet.Range("A20")
            Label20.Text = range.Value
            range = sheet.Range("B20")
            TextBox20.Text = range.Value
            range = sheet.Range("A21")
            Label21.Text = range.Value
            range = sheet.Range("B21")
            TextBox21.Text = range.Value
            range = sheet.Range("A22")
            Label22.Text = range.Value
            range = sheet.Range("B22")
            TextBox22.Text = range.Value
            range = sheet.Range("A23")
            Label23.Text = range.Value
            range = sheet.Range("B23")
            TextBox23.Text = range.Value
            range = sheet.Range("A24")
            Label24.Text = range.Value
            range = sheet.Range("B24")
            TextBox24.Text = range.Value
            range = sheet.Range("A25")
            Label25.Text = range.Value
            range = sheet.Range("B25")
            TextBox25.Text = range.Value
            range = sheet.Range("A26")
            Label26.Text = range.Value
            range = sheet.Range("B26")
            TextBox26.Text = range.Value
            range = sheet.Range("A27")
            Label27.Text = range.Value
            range = sheet.Range("B27")
            TextBox27.Text = range.Value
            range = sheet.Range("A28")
            Label28.Text = range.Value
            range = sheet.Range("B28")
            TextBox28.Text = range.Value
            range = sheet.Range("A29")
            Label29.Text = range.Value
            range = sheet.Range("B29")
            TextBox29.Text = range.Value
            range = sheet.Range("A30")
            Label30.Text = range.Value
            range = sheet.Range("B30")
            TextBox30.Text = range.Value
            range = sheet.Range("A31")
            Label31.Text = range.Value
            range = sheet.Range("B31")
            TextBox31.Text = range.Value
            range = sheet.Range("A32")
            Label32.Text = range.Value
            range = sheet.Range("B32")
            TextBox32.Text = range.Value
            range = sheet.Range("A33")
            Label33.Text = range.Value
            range = sheet.Range("B33")
            TextBox33.Text = range.Value
            range = sheet.Range("A34")
            Label34.Text = range.Value
            range = sheet.Range("B34")
            TextBox34.Text = range.Value
            range = sheet.Range("A35")
            Label35.Text = range.Value
            range = sheet.Range("B35")
            TextBox35.Text = range.Value
            range = sheet.Range("A36")
            Label36.Text = range.Value
            range = sheet.Range("B36")
            TextBox36.Text = range.Value
            range = sheet.Range("A37")
            Label37.Text = range.Value
            range = sheet.Range("B37")
            TextBox37.Text = range.Value
            range = sheet.Range("A38")
            Label38.Text = range.Value
            range = sheet.Range("B38")
            TextBox38.Text = range.Value
            range = sheet.Range("A39")
            Label39.Text = range.Value
            range = sheet.Range("B39")
            TextBox39.Text = range.Value
            range = sheet.Range("A40")
            Label40.Text = range.Value
            range = sheet.Range("B40")
            TextBox40.Text = range.Value
            range = sheet.Range("A41")
            Label41.Text = range.Value
            range = sheet.Range("B41")
            TextBox41.Text = range.Value
            range = sheet.Range("A42")
            Label42.Text = range.Value
            range = sheet.Range("B42")
            TextBox42.Text = range.Value
            range = sheet.Range("A43")
            Label43.Text = range.Value
            range = sheet.Range("B43")
            TextBox43.Text = range.Value
            range = sheet.Range("A44")
            Label44.Text = range.Value
            range = sheet.Range("B44")
            TextBox44.Text = range.Value
            range = sheet.Range("A45")
            Label45.Text = range.Value
            range = sheet.Range("B45")
            TextBox45.Text = range.Value
            range = sheet.Range("A46")
            Label46.Text = range.Value
            range = sheet.Range("B46")
            TextBox46.Text = range.Value
            range = sheet.Range("A47")
            Label47.Text = range.Value
            range = sheet.Range("B47")
            TextBox47.Text = range.Value
            range = sheet.Range("A48")
            Label48.Text = range.Value
            range = sheet.Range("B48")
            TextBox48.Text = range.Value
            range = sheet.Range("A49")
            Label49.Text = range.Value
            range = sheet.Range("B49")
            TextBox49.Text = range.Value
            range = sheet.Range("A50")
            Label50.Text = range.Value
            range = sheet.Range("B50")
            TextBox50.Text = range.Value
            range = sheet.Range("A51")
            Label51.Text = range.Value
            range = sheet.Range("B51")
            TextBox51.Text = range.Value
            range = sheet.Range("A52")
            Label52.Text = range.Value
            range = sheet.Range("B52")
            TextBox52.Text = range.Value
            range = sheet.Range("A53")
            Label53.Text = range.Value
            range = sheet.Range("B53")
            TextBox53.Text = range.Value
            range = sheet.Range("A54")
            Label54.Text = range.Value
            range = sheet.Range("B54")
            TextBox54.Text = range.Value
            range = sheet.Range("A55")
            Label55.Text = range.Value
            range = sheet.Range("B55")
            TextBox55.Text = range.Value
            range = sheet.Range("A56")
            Label56.Text = range.Value
            range = sheet.Range("B56")
            TextBox56.Text = range.Value
            range = sheet.Range("A57")
            Label57.Text = range.Value
            range = sheet.Range("B57")
            TextBox57.Text = range.Value
            range = sheet.Range("A58")
            Label58.Text = range.Value
            range = sheet.Range("B58")
            TextBox58.Text = range.Value
            range = sheet.Range("A59")
            Label59.Text = range.Value
            range = sheet.Range("B59")
            TextBox59.Text = range.Value
            range = sheet.Range("A60")
            Label60.Text = range.Value
            range = sheet.Range("B60")
            TextBox60.Text = range.Value

            book.Close()
            excelApp.Quit()
            releaseObject(excelApp)
            releaseObject(book)
            releaseObject(sheet)
            MsgBox("所有图片路径已成功导入")
        Catch ex As Exception
        End Try
    End Sub
End Class