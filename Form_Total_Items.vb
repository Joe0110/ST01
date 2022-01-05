Public Class Form_Total_Items
    Public soft_Type As Integer = 0

    Public color_Note(26) As String
    Public color_Note_Path As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        ReNew_All()
        soft_Type = 0

    End Sub
    Public Sub ReSoft()
        If soft_Type <> 0 Then
            If soft_Type = 1 Then
                Soft_Type1()
            ElseIf soft_Type = 2 Then
                Soft_Type2()
            End If
        End If
    End Sub
    Public Sub ReNew_All()
        Dim xx_array() As String

        '===============================
        Dim xDataG As New DataTable


        xDataG.Columns.Add("No.")
        xDataG.Columns.Add("Color")
        xDataG.Columns.Add("Percent")
        xDataG.Columns.Add("Counter")
        xDataG.Columns.Add("Note")

        color_Note_Path = Form_Spec.Get_Source_DIR & "\APP\Color_Note.txt"

        If (IO.File.Exists(color_Note_Path)) Then


            Dim Parameter_Line As String = ""


            Dim fs As IO.FileStream = New IO.FileStream(color_Note_Path, IO.FileMode.Open)
            Using sr As IO.StreamReader = New IO.StreamReader(fs, _
                                                        System.Text.Encoding.Default)

                ' While (sr.EndOfStream <> True)

                For i = 0 To 25

                    '送出相關設定
                    Parameter_Line = sr.ReadLine()
                    xx_array = Parameter_Line.Split(",")
                    color_Note(i) = xx_array(1)
                Next i
                'End While

                sr.Close()                                                          '關閉檔案

            End Using

        End If


        Dim tTemp As String
        Dim tTemp_Array() As String



        For i = 0 To 25

            '    Form_NewID.ComboBox1.SelectedIndex = i
            '    tTemp = Form_NewID.Counter_Id(i, 0)

            Dim xRow As DataRow = xDataG.NewRow

            tTemp = Form_status.eColor_text(i)

            tTemp_Array = tTemp.Split(" ")

            tTemp_Array(1) = tTemp_Array(1).Replace("(", "")
            tTemp_Array(1) = tTemp_Array(1).Replace(")", "")

            xRow(0) = (i).ToString("00")
            xRow(1) = ""

            xRow(2) = tTemp_Array(0)
            xRow(3) = Val((tTemp_Array(1))).ToString("0000")
            xRow(4) = color_Note(i)



            xDataG.Rows.Add(xRow)

        Next i








        DGV.DataSource = xDataG

        ReNew_Grid_Color()

        DGV.Columns(0).Width = 25
        DGV.Columns(1).Width = 30
        DGV.Columns(2).Width = 45
        DGV.Columns(3).Width = 35
        DGV.Columns(4).Width = 120
    End Sub
    Public Function Get_text_color(ByVal AA As Integer) As Color
        Dim xxx As Color
        'Dim QQQ As TextBox

        'QQQ.Name = "TextBox" & AA

        Select Case AA
            Case 0
                xxx = Form_status.TextBox0.BackColor
            Case 1
                xxx = Form_status.TextBox1.BackColor
            Case 2
                xxx = Form_status.TextBox2.BackColor
            Case 3
                xxx = Form_status.TextBox3.BackColor
            Case 4
                xxx = Form_status.TextBox4.BackColor
            Case 5
                xxx = Form_status.TextBox5.BackColor
            Case 6
                xxx = Form_status.TextBox6.BackColor
            Case 7
                xxx = Form_status.TextBox7.BackColor
            Case 8
                xxx = Form_status.TextBox8.BackColor
            Case 9
                xxx = Form_status.TextBox9.BackColor
            Case 10
                xxx = Form_status.TextBox10.BackColor
            Case 11
                xxx = Form_status.TextBox11.BackColor
            Case 12
                xxx = Form_status.TextBox12.BackColor
            Case 13
                xxx = Form_status.TextBox13.BackColor
            Case 14
                xxx = Form_status.TextBox14.BackColor
            Case 15
                xxx = Form_status.TextBox15.BackColor
            Case 16
                xxx = Form_status.TextBox16.BackColor
            Case 17
                xxx = Form_status.TextBox17.BackColor
            Case 18
                xxx = Form_status.TextBox18.BackColor
            Case 19
                xxx = Form_status.TextBox19.BackColor
            Case 20
                xxx = Form_status.TextBox20.BackColor
            Case 21
                xxx = Form_status.TextBox21.BackColor
            Case 22
                xxx = Form_status.TextBox22.BackColor
            Case 23
                xxx = Form_status.TextBox23.BackColor
            Case 24
                xxx = Form_status.TextBox24.BackColor
            Case 25
                xxx = Form_status.TextBox25.BackColor

        End Select

        Return xxx

    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Soft_Type2()

        soft_Type = 2
    End Sub
    Public Sub Soft_Type1()
        DGV.Sort(DGV.Columns(2), 1)

        ReNew_Grid_Color()
    End Sub
    Public Sub Soft_Type2()
        DGV.Sort(DGV.Columns(3), 1)

        ReNew_Grid_Color()
    End Sub
    Public Sub ReNew_Grid_Color()
        For i = 0 To 25
            'DGV.Rows(i).DefaultCellStyle.BackColor = Get_text_color(i)
            ' DGV.CellBorderStyle(2, i).BackColor = Get_text_color(i)
            'DGV.Rows(i).Cells(3).Style.BackColor = Get_text_color(i)
            DGV.Rows(i).Cells(1).Style.BackColor = Get_text_color(Val(DGV.Rows(i).Cells(0).Value))
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Soft_Type1()
        soft_Type = 1

    End Sub

    Private Sub DGV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellContentClick
        ReNew_Grid_Color()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_status.Open_X_file(color_Note_Path)
    End Sub

    Private Sub Form_Total_Items_Load(sender As Object, e As EventArgs) Handles Me.Load
        color_Note_Path = Form_Spec.Get_Source_DIR & "\APP\Color_Note.txt"
        ReNew_All()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim xfilename As String
        'If Form_Spec.Pattern_Type <> "" Then
        'xfilename = Form_Spec.Get_file_Path_x & "\" & Form_Spec.Pattern_Type & "\All\Status_All_B.png"
        ' Else
        xfilename = Form_Spec.FullName_ALL & "\Status_All_B.png"
        ' End If

        Capture_Screen_to_PNG(Me, xfilename)

        ' If Form_Spec.Pattern_Type <> "" Then
        xfilename = Form_Spec.FullName_ALL & "\Status_All_B_" & Now.ToString("yyyyMMdd_HHmmss") & ".Png"
        ' Else
        ' xfilename = Form_Spec.Get_file_Path_x & "\All\Status_All_B_" & Now.ToString("yyyyMMdd_HHmmss") & ".Png"
        ' End If

        Capture_Screen_to_PNG(Me, xfilename)

    End Sub
    Public Sub Capture_Screen_to_PNG(ByRef XXX As Form, ByVal zz As String)
        Dim xx As Integer
        Dim yy As Integer
        Dim d1 As Integer
        Dim d2 As Integer


        d1 = 0
        d2 = 0

        xx = XXX.Width
        yy = XXX.Height


        Dim Screen_Img As New Bitmap(xx, yy)
        Dim g = Graphics.FromImage(Screen_Img)


        g.CopyFromScreen(New Point(XXX.Location.X + d1, XXX.Location.Y + d2), New Point(0, 0), New Size(xx, yy))
        'g.CopyFromScreen(New Point(Me.Location.X + Me.Left + d1, Me.Location.Y + Me.Top + d2), New Point(0, 0), New Size(xx, yy))

        Screen_Img.Save(zz, System.Drawing.Imaging.ImageFormat.Png)

        Screen_Img.Dispose()

    End Sub
End Class