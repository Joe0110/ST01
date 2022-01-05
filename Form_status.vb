Public Class Form_status
    Public Log_File As String
    Public eColor_area(26) As Integer
    Public eColor_value(26) As Integer
    Public eColor_count(26) As Integer
    Public eColor_text(26) As String
    Public eColor2_text(26) As String
    Public total_mark As Integer = 0


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cal_Process()

        If Form_Total_Items.Created = True Then
            Form_Total_Items.ReNew_All()
            Form_Total_Items.ReSoft()
        End If

    End Sub
    Public Sub Cal_Process()
        Dim i As Integer
        Dim PageName As String
        Dim op_src As String

        Dim total_area As Integer

        total_mark = 0

        Debug.Print("Cal_Process:" & Now)

        Dim strArr2() As String
        Dim bWidth As Integer
        Dim bHeight As Integer
        Dim get_Color As Integer

        TB_Color.BackColor = Color.Red
        ' Form_Spec.TB_Page_Total.Text
        ' total_area = Form_Spec.Page_Default * Form_Spec_CHK_Top.PicBox_Default_width * Form_Spec_CHK_Top.PicBox_Default_Height
        total_area = Val(Form_Spec.TB_Page_Total.Text) * Form_Spec_CHK_Top.PicBox_Default_width * Form_Spec_CHK_Top.PicBox_Default_Height
        TB_Total.Text = total_area

        Label26.Text = Form_Spec.TB_Page_Total.Text

        For i = 0 To 25
            eColor_area(i) = 0
            eColor_value(i) = 0
            eColor_count(i) = 0
            eColor_text(i) = ""
        Next


        If Not IO.Directory.Exists(Form_Spec.FullName_Index_Path) Then
            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(Form_Spec.FullName_Index_Path)
        End If

        Log_File = Form_Spec.FullName_Index_Path & "\Process_Log.txt"
        '=================================================
        '   來源路徑                  =>   計算面積
        '=================================================
        'Form_Spec.FullName_item_Path => Full_Target_Path
        '   確認檔案是否存在
        '   讀檔
        '   計算
        '
        '   頁數: Form_Spec.Page_Default
        '-------------------------------------------------
        Debug.Print("cal:" & Form_Spec.FullName_item_Path)

        'For i = 1 To Form_Spec.Page_Default
        For i = 1 To Val(Form_Spec.TB_Page_Total.Text)

            PageName = "Page_" & i.ToString("000") & ".csv"
            op_src = Form_Spec.FullName_item_Path & "\" & PageName



            If (IO.File.Exists(op_src)) Then
                '------------------
                '   讀取檔案 - Start
                '------------------
                Dim Parameter_Line As String = ""



                '開啟證件
                Dim fs As IO.FileStream = New IO.FileStream(op_src, IO.FileMode.Open)
                Using sr As IO.StreamReader = New IO.StreamReader(fs, _
                                                            System.Text.Encoding.Default)






                    While (sr.EndOfStream <> True)

                        '送出相關設定
                        Parameter_Line = sr.ReadLine()                                    '讀取一行資料
                        strArr2 = Parameter_Line.Split(",")
                        bWidth = Val(strArr2(10))
                        bHeight = Val(strArr2(11))

                        get_Color = Val(strArr2(13))
                        eColor_area(get_Color) += bWidth * bHeight
                        eColor_count(get_Color) += 1

                        total_mark += bWidth * bHeight
                    End While




                    sr.Close()                                                          '關閉檔案

                End Using

                '------------------
                '   讀取檔案 - End
                '------------------

            End If
        Next
        TB_Mark.Text = total_mark
        TB_Color.BackColor = Color.YellowGreen
        LB_Percent.Text = ((total_mark / total_area) * 100).ToString("0.##") & "%"


        'Label0.Text = cal_each_color(eColor_count(0), eColor_area(0), total_mark)
        'Label1.Text = cal_each_color(eColor_count(1), eColor_area(1), total_mark)
        'Label2.Text = cal_each_color(eColor_count(2), eColor_area(2), total_mark)
        'Label3.Text = cal_each_color(eColor_count(3), eColor_area(3), total_mark)
        'Label4.Text = cal_each_color(eColor_count(4), eColor_area(4), total_mark)
        'Label5.Text = cal_each_color(eColor_count(5), eColor_area(5), total_mark)
        'Label6.Text = cal_each_color(eColor_count(6), eColor_area(6), total_mark)
        'Label7.Text = cal_each_color(eColor_count(7), eColor_area(7), total_mark)
        'Label8.Text = cal_each_color(eColor_count(8), eColor_area(8), total_mark)
        'Label9.Text = cal_each_color(eColor_count(9), eColor_area(9), total_mark)
        'Label10.Text = cal_each_color(eColor_count(10), eColor_area(10), total_mark)
        'Label11.Text = cal_each_color(eColor_count(11), eColor_area(11), total_mark)
        'Label12.Text = cal_each_color(eColor_count(12), eColor_area(12), total_mark)
        'Label13.Text = cal_each_color(eColor_count(13), eColor_area(13), total_mark)
        'Label14.Text = cal_each_color(eColor_count(14), eColor_area(14), total_mark)
        'Label15.Text = cal_each_color(eColor_count(15), eColor_area(15), total_mark)
        'Label16.Text = cal_each_color(eColor_count(16), eColor_area(16), total_mark)
        'Label17.Text = cal_each_color(eColor_count(17), eColor_area(17), total_mark)
        'Label18.Text = cal_each_color(eColor_count(18), eColor_area(18), total_mark)
        'Label19.Text = cal_each_color(eColor_count(19), eColor_area(19), total_mark)
        'Label20.Text = cal_each_color(eColor_count(20), eColor_area(20), total_mark)
        'Label21.Text = cal_each_color(eColor_count(21), eColor_area(21), total_mark)
        'Label22.Text = cal_each_color(eColor_count(22), eColor_area(22), total_mark)
        'Label23.Text = cal_each_color(eColor_count(23), eColor_area(23), total_mark)
        'Label24.Text = cal_each_color(eColor_count(24), eColor_area(24), total_mark)
        'Label25.Text = cal_each_color(eColor_count(25), eColor_area(25), total_mark)

        Label0.Text = cal_each_Xcolor(0)
        Label1.Text = cal_each_Xcolor(1)
        Label2.Text = cal_each_Xcolor(2)
        Label3.Text = cal_each_Xcolor(3)
        Label4.Text = cal_each_Xcolor(4)
        Label5.Text = cal_each_Xcolor(5)
        Label6.Text = cal_each_Xcolor(6)
        Label7.Text = cal_each_Xcolor(7)
        Label8.Text = cal_each_Xcolor(8)
        Label9.Text = cal_each_Xcolor(9)
        Label10.Text = cal_each_Xcolor(10)
        Label11.Text = cal_each_Xcolor(11)
        Label12.Text = cal_each_Xcolor(12)
        Label13.Text = cal_each_Xcolor(13)
        Label14.Text = cal_each_Xcolor(14)
        Label15.Text = cal_each_Xcolor(15)
        Label16.Text = cal_each_Xcolor(16)
        Label17.Text = cal_each_Xcolor(17)
        Label18.Text = cal_each_Xcolor(18)
        Label19.Text = cal_each_Xcolor(19)
        Label20.Text = cal_each_Xcolor(20)
        Label21.Text = cal_each_Xcolor(21)
        Label22.Text = cal_each_Xcolor(22)
        Label23.Text = cal_each_Xcolor(23)
        Label24.Text = cal_each_Xcolor(24)
        Label25.Text = cal_each_Xcolor(25)


        '=====================================
        '   保存Log
        '=====================================
        Dim filenum As Integer


        filenum = FreeFile()
        FileOpen(filenum, Log_File, OpenMode.Append)



        'Now.ToString("yyyyMMdd_HHmmss")
        'PrintLine(filenum, Now.ToString("yyyy/MM/dd HH:mm:ss") & "," & LB_Percent.Text)
        Print(filenum, Now.ToString("yyyy/MM/dd HH:mm:ss") & "," & LB_Percent.Text & ",")
        For i = 0 To 24
            Print(filenum, cal_each_Xcolor2(i) & ",")
            'Print(filenum, eColor_value(i) & "," & eColor_count(i) & ",")
            'eColor_value(i) = 0
            'eColor_count(i) = 0
        Next i
        PrintLine(filenum, cal_each_Xcolor2(25))
        'PrintLine(filenum, eColor_value(25) & "," & eColor_count(25))

        FileClose(filenum)
    End Sub
    Public Function cal_each_color(ByVal AA As Integer, ByVal BB As Integer, ByVal CC As Integer) As String
        ' Return AA.ToString("0000") & "_" & ((BB / CC) * 100).ToString("00.00") & "%"
        If BB = 0 Then
            Return (0).ToString("00.00") & "%" & " (" & AA & ")"
        Else
            Return ((BB / CC) * 100).ToString("00.00") & "%" & " (" & AA & ")"
        End If

    End Function
    Public Function cal_each_Xcolor(ByVal AA As Integer) As String

        If eColor_area(AA) = 0 Then
            eColor_text(AA) = (0).ToString("00.00") & "%" & " (" & eColor_count(AA) & ")"
        Else
            eColor_text(AA) = ((eColor_area(AA) / total_mark) * 100).ToString("00.00") & "%" & " (" & eColor_count(AA) & ")"
        End If

        Return eColor_text(AA)
    End Function
    Public Function cal_each_Xcolor2(ByVal AA As Integer) As String
        If eColor_area(AA) = 0 Then
            eColor2_text(AA) = (0).ToString("00.00") & "_" & eColor_count(AA)
        Else
            eColor2_text(AA) = ((eColor_area(AA) / total_mark) * 100).ToString("00.00") & "_" & eColor_count(AA)
        End If

        Return eColor2_text(AA)
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Open_X_file(Log_File)
    End Sub
    Public Sub Open_X_file(ByVal AAA As String)


        '===================
        '　開啟Report
        '===================

        If (IO.File.Exists(AAA)) Then
            '        ReportFileName
            If (AAA <> "") Then
                System.Diagnostics.Process.Start(AAA)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else
            MsgBox("找不到該檔案:" & AAA)
        End If
    End Sub


    Private Sub Form_status_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox0.Text = "無"
        'TextBox0.BackColor = Color.Blue
        TextBox1.BackColor = Color.Red
        TextBox2.BackColor = Color.Green
        TextBox3.BackColor = Color.DarkOrange
        TextBox4.BackColor = Color.Black
        TextBox5.BackColor = Color.Yellow
        TextBox6.BackColor = Color.YellowGreen
        TextBox7.BackColor = Color.Purple
        TextBox8.BackColor = Color.Blue
        TextBox9.BackColor = Color.Pink

        TextBox10.BackColor = Color.DodgerBlue

        TextBox11.BackColor = Color.Aqua
        TextBox12.BackColor = Color.Magenta
        TextBox13.BackColor = Color.LimeGreen

        TextBox14.BackColor = Color.Lime

        TextBox15.BackColor = Color.LightSkyBlue
        TextBox16.BackColor = Color.LightGreen

        TextBox17.BackColor = Color.Gold

        TextBox18.BackColor = Color.Olive
        TextBox19.BackColor = Color.RoyalBlue
        TextBox20.BackColor = Color.Violet
        TextBox21.BackColor = Color.DarkGray
        TextBox22.BackColor = Color.LightSalmon
        TextBox23.BackColor = Color.DeepPink
        TextBox24.BackColor = Color.Navy
        TextBox25.BackColor = Color.DarkGreen

        Cal_Process()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form_Total_Items.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim xfilename As String
        ' If Form_Spec.Pattern_Type <> "" Then
        'xfilename = Form_Spec.Get_file_Path_x & "\" & Form_Spec.Pattern_Type & "\All\Status_All_A.png"
        ' Else
        xfilename = Form_Spec.FullName_ALL & "\Status_All_A.png"
        ' End If

        Form_Total_Items.Capture_Screen_to_PNG(Me, xfilename)

        ' If Form_Spec.Pattern_Type <> "" Then
        'xfilename = Form_Spec.Get_file_Path_x & "\" & Form_Spec.Pattern_Type & "\All\Status_All_A_" & Now.ToString("yyyyMMdd_HHmmss") & ".Png"
        'Else
        xfilename = Form_Spec.FullName_ALL & "\Status_All_A_" & Now.ToString("yyyyMMdd_HHmmss") & ".Png"
        'End If

        Form_Total_Items.Capture_Screen_to_PNG(Me, xfilename)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form_Report.Show()
    End Sub
End Class