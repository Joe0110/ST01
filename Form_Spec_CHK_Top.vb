Public Class Form_Spec_CHK_Top
    Public focus_Flag As Boolean = False
    Public first_flag As Boolean = False

    Public Const Default_Title As String = "Form_Spec_CHK_Top"
    Public Form_Ori_Width As Integer
    Public Form_Ori_X2_Width As Integer
    Public Form_Ori_Height As Integer

    Public Const Form_Default_width As Integer = 919
    Public Const Form_Default_X2_width As Integer = 1680

    Public Const Form_Default_Height As Integer = 1050

    Public Const PicBox_Default_width As Integer = 720
    Public Const PicBox_Default_Height As Integer = 960

    Public P1_Ori_Width As Integer
    Public P1_Ori_Height As Integer

    Public S_W_Ratio As Single
    Public S_H_Ratio As Single
    Public FormTOP_Opacity As Single = 0.7

    Public Const x_base As Integer = 26
    Public Const y_base As Integer = 22
    Public Const Picbox_Gap As Integer = 16

    Public Mouse_X As Integer
    Public Mouse_Y As Integer

    Dim dd1 As Integer
    Dim dd2 As Integer
    Dim dd3 As Integer
    Dim dd4 As Integer
    Public X_ReSize_flag As Boolean = False


    Public LB_Array(0) As Label
    Public LB_cnt As Integer = 0
    Public LB_cnt_Enable As Boolean = False



    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PicBox_TopMain.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BTN_Overlap.Click
        Overlap_windows()

        Form_Spec_CHK_Bottom.Activate()

        'Form_Spec_CHK2.Activate(vbNormalFocus)
        Form_Spec_CHK_Bottom.Focus()
        Me.Focus()
    End Sub
    Public Sub Overlap_windows()
        Me.Location = Form_Spec_CHK_Bottom.Location
        Me.Size = Form_Spec_CHK_Bottom.Size
        Me.Opacity = FormTOP_Opacity

    End Sub
    '-------------------
    '   擷取螢幕, 更改大小後儲存
    '-------------------
    Public Sub Capture_ReSize_save(ByVal xx As Integer, ByVal yy As Integer, ByVal tNewFileName As String)

        Dim Capture_Width As Integer
        Dim Capture_Height As Integer
        Dim Dpi_base As Single

        If Form_Spec.CKB_Lo_dpi.Checked = True Then
            Dpi_base = 0.5
        Else
            Dpi_base = 1
        End If


        If Form_Spec.CKB_compare.Checked = False Then
            Capture_Width = Int(720 * S_W_Ratio) * Dpi_base
            Capture_Height = Int(960 * S_H_Ratio) * Dpi_base
        Else
            Capture_Width = Int(1600 * S_W_Ratio) * Dpi_base
            Capture_Height = Int(960 * S_H_Ratio) * Dpi_base
        End If


        Dim Screen_Img As New Bitmap(Capture_Width, Capture_Height)
        Dim g = Graphics.FromImage(Screen_Img)
        Dim Small_bmp As New Bitmap(xx, yy)


        '==============================================
        '擷取螢幕
        g.CopyFromScreen(New Point(Me.Location.X + PicBox_TopMain.Left + 10, Me.Location.Y + PicBox_TopMain.Top + 30), New Point(0, 0), New Size(Capture_Width, Capture_Height))

        '==============================================
        '改大小
        Dim ScrnPB As New PictureBox

        ScrnPB.Width = Capture_Width
        ScrnPB.Height = Capture_Height

        ScrnPB.Image = Screen_Img

        ' img => bmp
        Dim gfx As Graphics = Graphics.FromImage(Small_bmp)
        gfx.DrawImage(ScrnPB.Image, 0, 0, xx, yy)

        gfx.Dispose()
        Small_bmp.Save(tNewFileName, System.Drawing.Imaging.ImageFormat.Png)
        Small_bmp.Dispose()
        ScrnPB.Dispose()
        '==============================================


        Screen_Img.Dispose()

    End Sub
    '-------------------
    '   擷取螢幕, 更改大小後儲存
    '-------------------
    Public Sub Capture_ReSize_New_save(ByVal xx As Integer, ByVal yy As Integer, ByVal zz As String, ByVal tNewFileName As String)

        Dim Capture_Width As Integer
        Dim Capture_Height As Integer
        Dim Dpi_base As Single

        If Form_Spec.CKB_Lo_dpi.Checked = True Then
            Dpi_base = 0.5
        Else
            Dpi_base = 1
        End If


        If Form_Spec.CKB_compare.Checked = False Then
            Capture_Width = Int(720 * S_W_Ratio) * Dpi_base
            Capture_Height = Int(960 * S_H_Ratio) * Dpi_base
        Else
            Capture_Width = Int(1600 * S_W_Ratio) * Dpi_base
            Capture_Height = Int(960 * S_H_Ratio) * Dpi_base
        End If


        Dim Screen_Img As New Bitmap(Capture_Width, Capture_Height)
        Dim g = Graphics.FromImage(Screen_Img)
        Dim Small_bmp As New Bitmap(xx, yy)


        '==============================================
        '擷取螢幕
        g.CopyFromScreen(New Point(Me.Location.X + PicBox_TopMain.Left + 10, Me.Location.Y + PicBox_TopMain.Top + 30), New Point(0, 0), New Size(Capture_Width, Capture_Height))

        '==============================================
        '改大小
        Dim ScrnPB As New PictureBox

        ScrnPB.Width = Capture_Width
        ScrnPB.Height = Capture_Height

        ScrnPB.Image = Screen_Img

        ' img => bmp
        Dim gfx As Graphics = Graphics.FromImage(Small_bmp)
        gfx.DrawImage(ScrnPB.Image, 0, 0, xx, yy)

        Debug.Print("target(small_bmp):" & tNewFileName)
        Dim xxx As String = ""
        xxx = keep_Path_remove_filename(tNewFileName)
        Debug.Print("x process:" & xxx)
        Create_All_Path2(xxx)

        gfx.Dispose()

        Small_bmp.Save(tNewFileName, System.Drawing.Imaging.ImageFormat.Png)

        Small_bmp.Dispose()
        ScrnPB.Dispose()
        '==============================================
        '儲存工作圖
        Debug.Print("Output:" & zz)

        xxx = keep_Path_remove_filename(zz)
        Debug.Print("x process:" & xxx)
        Create_All_Path2(xxx)

        Screen_Img.Save(zz, System.Drawing.Imaging.ImageFormat.Png)
        Screen_Img.Dispose()

    End Sub
    Public Function keep_Path_remove_filename(ByVal aaa As String) As String
        Dim x_leng As Integer
        Dim x_str As String
        x_leng = aaa.LastIndexOf("\")
        x_str = Mid(aaa, 1, x_leng)
        Return x_str
    End Function
    Public Sub Create_All_Path2(ByVal tFullPath As String)
        Dim full_path As String
        Dim x_path() As String
        full_path = tFullPath
        Dim temp_path As String = ""

        If full_path.Contains("\") Then
            x_path = full_path.Split("\")

            Debug.Print("LBound:" & LBound(x_path))
            Debug.Print("UBound:" & UBound(x_path))

            For i = 0 To (UBound(x_path)) Step 1
                'temp_path = temp_path & x_path(i) & "\" & x_path(i + 1)
                If i = 0 Then
                    temp_path = temp_path & x_path(i) & "\"
                ElseIf i = UBound(x_path) Then
                    temp_path = temp_path & x_path(i)
                Else
                    temp_path = temp_path & x_path(i) & "\"
                End If

                Debug.Print(i & ": " & temp_path)

                If Not IO.Directory.Exists(temp_path) Then

                    '¦p¤£¦s¦b¡A«Ø¥ß¸ê®Æ§¨
                    IO.Directory.CreateDirectory(temp_path)
                End If

            Next

        End If
    End Sub
    Public Sub PB3_setOpacity(ByVal opacity As Integer)
        PicBox_Drag.BackColor = Color.FromArgb(opacity, 255, 0, 0)
    End Sub

    Public Sub Capture_SavePNG_WorkPic(ByVal zz As String)

        Dim Capture_Width As Integer
        Dim Capture_Height As Integer
        Dim Dpi_base As Single

        If Form_Spec.CKB_Lo_dpi.Checked = True Then
            Dpi_base = 0.5
        Else
            Dpi_base = 1
        End If


        Try


            If Form_Spec.CKB_compare.Checked = False Then
                Capture_Width = 720 * S_W_Ratio * Dpi_base
                Capture_Height = 960 * S_H_Ratio * Dpi_base
            Else
                Capture_Width = 1600 * S_W_Ratio * Dpi_base
                Capture_Height = 960 * S_H_Ratio * Dpi_base
            End If


            Dim Screen_Img As New Bitmap(Capture_Width, Capture_Height)
            Dim g = Graphics.FromImage(Screen_Img)

            '==============================================
            '擷取螢幕
            g.CopyFromScreen(New Point(Me.Location.X + PicBox_TopMain.Left + 10, Me.Location.Y + PicBox_TopMain.Top + 30), New Point(0, 0), New Size(Capture_Width, Capture_Height))

            '==============================================
            '改大小

            'Dim ScrnPB As New PictureBox
            'ScrnPB.Width = 720
            'ScrnPB.Height = 960

            'ScrnPB.Image = Screen_Img

            ' img => bmp
            'Dim gfx As Graphics = Graphics.FromImage(Small_bmp)
            ' gfx.DrawImage(ScrnPB.Image, 0, 0, xx, yy)

            'gfx.Dispose()


            '==============================================

            Screen_Img.Save(zz, System.Drawing.Imaging.ImageFormat.Png)
            Screen_Img.Dispose()
            ' Small_bmp.Dispose()
            'ScrnPB.Dispose()

        Catch ex As Exception

        Finally

        End Try

    End Sub


    Public Sub Draw_RectangleMouse_UP(D1 As Integer, D2 As Integer, D3 As Integer, D4 As Integer)
        Dim g As Graphics


        'Print("123")

        g = Me.CreateGraphics

        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        g.DrawRectangle(Pens.Black, D1, D2, D3, D4)


        g.Dispose()
    End Sub
    Public Sub Graph_Clear()
        Dim g As Graphics

        g = Me.CreateGraphics

        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        g.DrawRectangle(Pens.Black, 0, 0, 1600, 1600)
        g.FillRectangle(Brushes.White, 0, 0, 1600, 1600)


        g.Dispose()
    End Sub
    '[DW01]
    '--------------------------------------
    '   D1: 原始X
    '   D2: 原始Y
    '   D3: 原始W
    '   D4: 原始H
    '--------------------------------------
    Public Sub Draw_RectangleNow(D1 As Integer, D2 As Integer, D3 As Integer, D4 As Integer, Dstr As String, FONT_Color_Index As Byte, BG_Color_Index As Byte)
        Dim g As Graphics
        Dim xbase As Single

        'Dim x_base As Integer = 26
        'Dim y_base As Integer = 22

        Dim opD1 As Integer
        Dim opD2 As Integer

        Dim opD3 As Integer
        Dim opD4 As Integer


        Dim opD1_plus As Integer
        'Dim opD2 As Integer

        Dim drawString As String = Dstr
        Dim tBG_Color As Color
        'Dim xFont_Size As Integer

        Dim Label_Font_Color As Color

        '-------------------
        '   雙圖形比較時的補償
        '-------------------
        If D1 > 765 Then
            If Form_Spec.CKB_Lo_dpi.Checked = True Then
                opD1_plus = -8
            Else
                opD1_plus = -16
            End If

            'opD1_plus = 0
        Else
            opD1_plus = 0

        End If


        If Form_Spec.CKB_Lo_dpi.Checked = True Then
            xbase = 0.5
        Else
            xbase = 1
        End If


        opD1 = D1 - x_base - opD1_plus
        opD2 = D2 - y_base
        opD3 = D3 - x_base
        opD4 = D4 - y_base


        D1 = (x_base + opD1_plus) + (opD1 * S_W_Ratio * xbase) 'X
        D2 = y_base + (opD2 * S_H_Ratio * xbase)              'Y
        D3 = D3 * S_W_Ratio * xbase 'W
        D4 = D4 * S_H_Ratio * xbase 'H


        'Cooper Black
        'Dim fnt As New Font("ＭＳ ゴシック", 12, FontStyle.Bold)

        'If (S_H_Ratio = 1) And (S_W_Ratio = 1) Then

        '    xFont_Size = 14
        'ElseIf (S_H_Ratio <= 0.5) And (S_W_Ratio <= 0.5) Then
        '    xFont_Size = 8
        'ElseIf (S_H_Ratio <= 0.75) And (S_W_Ratio <= 0.75) Then
        '    xFont_Size = 10
        'Else
        '    xFont_Size = 12
        'End If

        'Dim fnt As New Font("Cooper Black", xFont_Size, FontStyle.Bold)

        'Print("123")

        g = Me.CreateGraphics

        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        '[COLOR_00]
        '---------------
        '   背景色
        '---------------
        Select Case BG_Color_Index
            Case 1
                g.FillRectangle(Brushes.Red, D1, D2, D3, D4)
                tBG_Color = Color.Red
            Case 2
                g.FillRectangle(Brushes.Green, D1, D2, D3, D4)
                tBG_Color = Color.Green
            Case 3
                g.FillRectangle(Brushes.DarkOrange, D1, D2, D3, D4)
                tBG_Color = Color.DarkOrange
            Case 4
                g.FillRectangle(Brushes.Black, D1, D2, D3, D4)
                tBG_Color = Color.Black
            Case 5
                g.FillRectangle(Brushes.Yellow, D1, D2, D3, D4)
                tBG_Color = Color.Yellow
            Case 6
                g.FillRectangle(Brushes.YellowGreen, D1, D2, D3, D4)
                tBG_Color = Color.YellowGreen
            Case 7
                g.FillRectangle(Brushes.Purple, D1, D2, D3, D4)
                tBG_Color = Color.Purple
            Case 8
                g.FillRectangle(Brushes.Blue, D1, D2, D3, D4)
                tBG_Color = Color.Blue
            Case 9
                g.FillRectangle(Brushes.Pink, D1, D2, D3, D4)
                tBG_Color = Color.Pink
            Case 10
                g.FillRectangle(Brushes.DodgerBlue, D1, D2, D3, D4)
                tBG_Color = Color.DodgerBlue
            Case 11
                g.FillRectangle(Brushes.Aqua, D1, D2, D3, D4)
                tBG_Color = Color.Aqua
            Case 12
                g.FillRectangle(Brushes.Magenta, D1, D2, D3, D4)
                tBG_Color = Color.Magenta
            Case 13
                g.FillRectangle(Brushes.LimeGreen, D1, D2, D3, D4)
                tBG_Color = Color.LimeGreen
            Case 14
                g.FillRectangle(Brushes.Lime, D1, D2, D3, D4)
                tBG_Color = Color.Lime
            Case 15
                g.FillRectangle(Brushes.LightSkyBlue, D1, D2, D3, D4)
                tBG_Color = Color.LightSkyBlue
            Case 16
                g.FillRectangle(Brushes.LightGreen, D1, D2, D3, D4)
                tBG_Color = Color.LightGreen
            Case 17
                g.FillRectangle(Brushes.Gold, D1, D2, D3, D4)
                tBG_Color = Color.Gold
            Case 18
                g.FillRectangle(Brushes.Olive, D1, D2, D3, D4)
                tBG_Color = Color.Olive
            Case 19
                g.FillRectangle(Brushes.RoyalBlue, D1, D2, D3, D4)
                tBG_Color = Color.RoyalBlue
            Case 20
                g.FillRectangle(Brushes.Violet, D1, D2, D3, D4)
                tBG_Color = Color.Violet
            Case 21
                g.FillRectangle(Brushes.DarkGray, D1, D2, D3, D4)
                tBG_Color = Color.DarkGray
            Case 22
                g.FillRectangle(Brushes.LightSalmon, D1, D2, D3, D4)
                tBG_Color = Color.LightSalmon
            Case 23
                g.FillRectangle(Brushes.DeepPink, D1, D2, D3, D4)
                tBG_Color = Color.DeepPink
            Case 24
                g.FillRectangle(Brushes.Navy, D1, D2, D3, D4)
                tBG_Color = Color.Navy
            Case 25
                g.FillRectangle(Brushes.DarkGreen, D1, D2, D3, D4)
                tBG_Color = Color.DarkGreen
        End Select

        '[COLOR_01]
        '---------------
        '   填入字
        '---------------

        'Add_Button_Label(D1, D2, Dstr, xxC)

        'Dim rect As New RectangleF(D1, D2, D3, D4)

        Select Case FONT_Color_Index
            Case 0
                g.DrawRectangle(Pens.Blue, D1, D2, D3, D4)
                Label_Font_Color = Color.Blue
            Case 1
                g.DrawRectangle(Pens.Red, D1, D2, D3, D4)
                Label_Font_Color = Color.Red

            Case 2
                g.DrawRectangle(Pens.Green, D1, D2, D3, D4)
                Label_Font_Color = Color.Green

            Case 3
                g.DrawRectangle(Pens.Black, D1, D2, D3, D4)
                Label_Font_Color = Color.Black

            Case 4
                g.DrawRectangle(Pens.DarkOrange, D1, D2, D3, D4)
                Label_Font_Color = Color.DarkOrange

            Case 5
                g.DrawRectangle(Pens.Yellow, D1, D2, D3, D4)
                Label_Font_Color = Color.Yellow

            Case 6
                g.DrawRectangle(Pens.YellowGreen, D1, D2, D3, D4)
                Label_Font_Color = Color.YellowGreen

            Case 7
                g.DrawRectangle(Pens.Purple, D1, D2, D3, D4)
                Label_Font_Color = Color.Purple

            Case 8
                g.DrawRectangle(Pens.Magenta, D1, D2, D3, D4)
                Label_Font_Color = Color.Magenta

            Case 9
                g.DrawRectangle(Pens.Pink, D1, D2, D3, D4)
                Label_Font_Color = Color.Pink

            Case 10
                g.DrawRectangle(Pens.DodgerBlue, D1, D2, D3, D4)
                Label_Font_Color = Color.DodgerBlue

            Case 11
                g.DrawRectangle(Pens.Lime, D1, D2, D3, D4)
                Label_Font_Color = Color.Lime
            Case 12
                g.DrawRectangle(Pens.Gold, D1, D2, D3, D4)
                Label_Font_Color = Color.Gold
            Case 13
                g.DrawRectangle(Pens.LightSalmon, D1, D2, D3, D4)
                Label_Font_Color = Color.LightSalmon
            Case 14
                g.DrawRectangle(Pens.DeepPink, D1, D2, D3, D4)
                Label_Font_Color = Color.DeepPink
            Case 15
                g.DrawRectangle(Pens.Navy, D1, D2, D3, D4)
                Label_Font_Color = Color.Navy
            Case 16
                g.DrawRectangle(Pens.DarkGray, D1, D2, D3, D4)
                Label_Font_Color = Color.DarkGray

        End Select

        Add_Button_Label(D1, D2, Dstr, tBG_Color)
        Set_LB_Font_Color(Label_Font_Color)

        g.Dispose()
    End Sub

    Public Sub Set_LB_Font_Color(ByVal CC As Color)
        'Debug.Print(LB_cnt)

        If LB_Array(LB_cnt - 1).Created = True Then
            LB_Array(LB_cnt - 1).ForeColor = CC
        End If
    End Sub

    ' Public Sub Add_Button_Label(ByVal XX As Integer, ByVal YY As Integer, ByVal Font_Str As String, ByVal CC As Color, ByVal DD As Color)    '[BTN00]
    Public Sub Add_Button_Label(ByVal XX As Integer, ByVal YY As Integer, ByVal Font_Str As String, ByVal BG_Color As Color)    '[BTN00]
        ' Dim fnt As New Font("Cooper Black", 14, FontStyle.Bold)

        Dim xFont_Size As Integer

        '-----------------------------
        '   縮放時, 字形大小的微調
        '-----------------------------
        If Form_Spec.CKB_Lo_dpi.Checked = False Then
            If (S_H_Ratio = 1) And (S_W_Ratio = 1) Then
                xFont_Size = 12
            ElseIf (S_H_Ratio <= 0.5) And (S_W_Ratio <= 0.5) Then
                xFont_Size = 8
            ElseIf (S_H_Ratio <= 0.75) And (S_W_Ratio <= 0.75) Then
                xFont_Size = 10
            Else
                xFont_Size = 11
            End If
        Else
            xFont_Size = 8
        End If

        Dim fnt As New Font("Cooper Black", xFont_Size, FontStyle.Bold)



        Dim xxx As New Point
        'xxx.X = XX + 1
        'xxx.Y = YY + 1
        xxx.X = XX + 2
        xxx.Y = YY + 2

        LB_Array(LB_cnt) = New Label
        LB_Array(LB_cnt).Font = fnt
        LB_Array(LB_cnt).Name = "LBA_" & LB_cnt
        LB_Array(LB_cnt).Location = xxx

        If Font_Str <> "" Then
            LB_Array(LB_cnt).Text = Font_Str
            LB_Array(LB_cnt).BackColor = Color.White    '使用白色 (考慮增加一個使用背景色的選項)
        Else
            LB_Array(LB_cnt).Text = " "         '如果是空的情形, 轉成一格空白(才選的到)
            LB_Array(LB_cnt).BackColor = BG_Color
        End If

        LB_Array(LB_cnt).AutoSize = True

        'LB_Array(LB_cnt).BackColor = CC


        Me.Controls.Add(LB_Array(LB_cnt))

        '--------------------------
        '   按鍵按下行為處理程序的定義
        '--------------------------
        AddHandler LB_Array(LB_cnt).Click, AddressOf Form_Spec_CHK_Ctl.Btns_Click

        LB_cnt += 1



        If LB_cnt <> 0 Then
            ReDim Preserve LB_Array(LB_cnt)

        End If

        LB_cnt_Enable = True

    End Sub
    Public Sub Release_LB()

        For i = 0 To LB_cnt - 1
            LB_Array(i).Dispose()
        Next

        'Item_Counter = 0


        LB_cnt_Enable = False
        LB_cnt = 0
        ReDim Preserve LB_Array(LB_cnt)

    End Sub

    Private Sub Form_Spec_CHK_Top_Activated(sender As Object, e As EventArgs) Handles Me.Activated


        If focus_Flag = False Then
            focus_Flag = True
            Set_to_TOP()
        End If
        'Timer1.Enabled = True

    End Sub
    Public Sub Set_to_TOP()
        Form_Spec_CHK_Bottom.TopMost = True
        Me.TopMost = True

        Form_Spec_CHK_Bottom.TopMost = False
        Me.TopMost = False
        Debug.Print("Switch to TOP:" & Now)


    End Sub

    Private Sub Form_Spec_CHK_Top_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        Debug.Print("Form_Spec_CHK_Top_Deactivate:" & Now)
        focus_Flag = False
    End Sub


    Private Sub Form_Spec_CHK_Top_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        Form_Spec_CHK_Bottom.Close()
    End Sub

    Private Sub Form_Spec_CHK_Top_DragOver(sender As Object, e As DragEventArgs) Handles Me.DragOver
        Debug.Print("Form_Spec_CHK_Top_DragOver:" & Now)
    End Sub

    Private Sub Form_Spec_CHK_Top_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        Debug.Print("Got_Focus:" & Now)
        focus_Flag = True

    End Sub
    '[HT00]
    Private Sub Form_Spec_CHK_Top_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        '===================================
        ' [hotkey]
        '===================================
        If e.KeyChar = "r" Then
            '------------------
            '   改大小位置
            '------------------
            If Form_Spec_CHK_Ctl.BTN_Select_flag = True Then
                Debug.Print("Modify_Location & Size")
                'Form_Spec_CHK_Ctl.Modify_Location()
                Form_Spec_CHK_Ctl.CMB_ItemType.SelectedIndex = 21   '改大小位置

                Form_Spec_CHK_Ctl.Modify_BlockItem()
                Form_Spec_CHK_Ctl.BTN_Select_flag = False
            End If

        ElseIf e.KeyChar = "`" Then
            '------------------
            '   重劃
            '------------------
            Form_Spec.Page_Now()
            Debug.Print("NFS_5")
            Form_Spec.Update_for_Sync()

        ElseIf e.KeyChar = "1" Then
            '------------------
            '   XLINK ON/OFF 切換
            '------------------
            If Form_Spec_CHK_Ctl.CKB_XLINK.Checked = True Then
                Form_Spec_CHK_Ctl.CKB_XLINK.Checked = False
            Else
                Form_Spec_CHK_Ctl.CKB_XLINK.Checked = True
            End If

        ElseIf e.KeyChar = "2" Then
            '------------------
            '   XAppend ON/OFF 切換
            '------------------
            If Form_Spec_CHK_Ctl.CKB_append.Checked = True Then
                Form_Spec_CHK_Ctl.CKB_append.Checked = False
            Else
                Form_Spec_CHK_Ctl.CKB_append.Checked = True
            End If

        ElseIf e.KeyChar = "3" Then
            '------------------
            '   JUMP ON/OFF 切換
            '------------------
            If Form_Spec_CHK_Ctl.CKB_JUMP.Checked = True Then
                Form_Spec_CHK_Ctl.CKB_JUMP.Checked = False
            Else
                Form_Spec_CHK_Ctl.CKB_JUMP.Checked = True
            End If

        ElseIf e.KeyChar = "4" Then
            '------------------
            '   原規格/工作區 切換
            '------------------
            If Form_Spec_CHK_Ctl.CKB_Ori_Graphic.Checked = True Then
                Form_Spec_CHK_Ctl.CKB_Ori_Graphic.Checked = False
            Else
                Form_Spec_CHK_Ctl.CKB_Ori_Graphic.Checked = True
            End If

            Debug.Print("NFS_6")
            Form_Spec.Update_for_Sync()

        ElseIf e.KeyChar = "m" Then
            '------------------
            '   改位置
            '------------------
            If Form_Spec_CHK_Ctl.BTN_Select_flag = True Then
                Debug.Print("Modify_Location")
                'Form_Spec_CHK_Ctl.Modify_Location()
                Form_Spec_CHK_Ctl.CMB_ItemType.SelectedIndex = 21   '改大小位置

                Form_Spec_CHK_Ctl.TB_XX.Text = Mouse_X
                Form_Spec_CHK_Ctl.TB_YY.Text = Mouse_Y
                Form_Spec_CHK_Ctl.TB_WW.Text = Form_Spec_CHK_Ctl.TBw.Text
                Form_Spec_CHK_Ctl.TB_HH.Text = Form_Spec_CHK_Ctl.TBh.Text

                Form_Spec_CHK_Ctl.Modify_BlockItem()
                Form_Spec_CHK_Ctl.BTN_Select_flag = False
            End If

        ElseIf e.KeyChar = "c" Then
            '------------------
            '   改顏色
            '------------------
            Form_Color.Show()
            Form_Color.TopMost = True
            Form_Color.TopMost = False

        ElseIf e.KeyChar = "n" Then
            '------------------
            '   新方塊
            '------------------
            Form_Spec_CHK_Ctl.Create_New_Item()

        ElseIf e.KeyChar = "d" Then
            '------------------
            '   刪除
            '------------------
            Form_Spec_CHK_Ctl.Delete_Item()

        ElseIf e.KeyChar = "l" Then
            '------------------
            '   La
            '------------------
            Form_Spec_CHK_Ctl.Open_by_LA(Form_Spec_CHK_Ctl.Npp_file_name)

        ElseIf e.KeyChar = "s" Then
            '------------------
            '   產生小圖
            '------------------
            Form_Spec.Update_for_Sync_Create_SmallPic()

        ElseIf e.KeyChar = "p" Then
            '------------------
            '   最上層顯示
            '------------------
            Form_Spec.Set_Form_On_Top()

        ElseIf e.KeyChar = "o" Then
            '------------------
            '   開啟檔案
            '------------------
            Form_Spec_CHK_Ctl.Open_Xfile2()

        ElseIf e.KeyChar = "v" Then
            'Open_by_Npp()
            Form_Spec_CHK_Ctl.Open_by_Npp("AAA")

        ElseIf e.KeyChar = "x" Then
            '------------------
            '   Copy到工作區
            '------------------
            Form_Spec_CHK_Ctl.Copy_to_word_Area()

        ElseIf e.KeyChar = "z" Then
            '------------------
            '   切換Spec
            '------------------
            Form_Spec.Switch_temp_Spec()

        ElseIf e.KeyChar = "a" Then
            '------------------
            '   取得位置
            '------------------
            Form_Spec_CHK_Ctl.TB_XX.Text = Mouse_X
            Form_Spec_CHK_Ctl.TB_YY.Text = Mouse_Y

        ElseIf e.KeyChar = "q" Then
            '------------------
            '   取得位置
            '------------------
            If Form_status.Created = False Then
                Form_status.Show()
            Else
                Form_status.Cal_Process()
            End If

            Form_status.TopMost = True
            Form_status.TopMost = False

            If Form_Total_Items.Created = True Then
                Form_Total_Items.ReNew_All()
                Form_Total_Items.ReSoft()
                Form_Total_Items.TopMost = True
                Form_Total_Items.TopMost = False
            End If

            End If

            Debug.Print("Key:" & e.KeyChar)

    End Sub

    Private Sub Form_Spec_CHK_Top_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = Default_Title

        Debug.Print("Load Top")



        'S_W_Ratio = 1
        'S_H_Ratio = 1



        '------------------------------------------
        If Form_Spec.CKB_Lo_dpi.Checked = True Then
            Debug.Print("Load Top LO DPI")

            S_W_Ratio = 1
            S_H_Ratio = 1

            Me.Width = (Form_Default_width / 2) + 20
            Me.Height = (Form_Default_Height / 2) + 30

            Form_Ori_Width = Me.Width
            Form_Ori_X2_Width = Form_Default_X2_width / 2

            Form_Ori_Height = Me.Height





            PicBox_TopMain.Width = PicBox_Default_width / 2
            PicBox_TopMain.Height = PicBox_Default_Height / 2


            P1_Ori_Width = PicBox_TopMain.Width
            P1_Ori_Height = PicBox_TopMain.Height

            Form_Spec.TextBox6.Text = "1"
            Form_Spec.Manual_Switch_Size()

        Else
            Debug.Print("Load Top Default")
            S_W_Ratio = 1
            S_H_Ratio = 1

            Me.Width = Form_Default_width
            Me.Height = Form_Default_Height

            Form_Ori_Width = Me.Width
            Form_Ori_Height = Me.Height

            Debug.Print("Load Top Default w:" & Me.Width)
            Debug.Print("Load Top Default h:" & Me.Height)
            'Me.Width = Form_Default_width
            'Me.Height = Form_Default_Height

            'Form_Ori_Width = Form_Default_width
            'Form_Ori_Height = Form_Default_Height

            Form_Ori_X2_Width = Form_Default_X2_width

            PicBox_TopMain.Width = PicBox_Default_width
            PicBox_TopMain.Height = PicBox_Default_Height

            P1_Ori_Width = PicBox_TopMain.Width
            P1_Ori_Height = PicBox_TopMain.Height

            'Form_Spec.TextBox6.Text = "1"
            'Form_Spec.Manual_Switch_Size()
        End If


        '------------------------------------------
        Switch_Compare_Mode()

        Debug.Print("Load Top Default w:" & Me.Width)
        Debug.Print("Load Top Default h:" & Me.Height)

        Overlap_windows()
        PB3_setOpacity(40)
        X_ReSize_flag = True



    End Sub
    Public Sub Switch_Compare_Mode()
        Dim xy_P As Point

        If Form_Spec.CKB_compare.Checked = False Then
            '919, 1050
            Debug.Print("Switch_Compare_Mode w:" & Me.Width)
            Me.Width = get_Form_Width_Single()
            Debug.Print("Switch_Compare_Mode w:" & Me.Width)
            '-------------
            'Button1位置
            '-------------
            xy_P.X = PicBox_TopMain.Location.X + PicBox_TopMain.Width + 10
            xy_P.Y = BTN_Overlap.Location.Y

            BTN_Overlap.Location = xy_P




        Else

            '1900,1050
            'Me.Width = 1800 * S_W_Ratio
            'Me.Width = x_base + Picbox_Gap + (((2 * 720) - x_base) * S_W_Ratio) + 100
            Me.Width = get_Form_Width_Double()
            PicBox_TopSub.Size = PicBox_TopMain.Size

            '-------------
            'Button1位置
            '-------------
            xy_P.X = PicBox_TopSub.Location.X + PicBox_TopMain.Width + 10
            xy_P.Y = BTN_Overlap.Location.Y


            BTN_Overlap.Location = xy_P






        End If
    End Sub
    Public Function get_Form_Width_Single() As Integer

        'If Form_Spec.CheckBox1.Checked = True Then

        'Return (x_base + ((800 - x_base) * S_W_Ratio))
        ' Else

        Return (x_base + ((Form_Ori_Width - x_base) * S_W_Ratio))
        ' End If
        'Return x_base + PictureBox1.Width + 10 + BTN_Overlap.Width



    End Function
    Public Function get_Form_Width_Double() As Integer
        'If Form_Spec.CheckBox1.Checked = True Then
        'Return (x_base + ((1000 - x_base) * S_W_Ratio))
        ' Else
        Return (x_base + ((Form_Ori_X2_Width - x_base) * S_W_Ratio))
        ' End If
        ' Return x_base + PictureBox1.Width + 16 + PictureBox4.Width + 10 + BTN_Overlap.Width

    End Function

    Private Sub Form_Spec_CHK_Top_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        Timer_OnTOP.Enabled = False
        Timer_OnTOP.Enabled = True

    End Sub

    Private Sub Form_Spec_CHK_Top_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        focus_Flag = False
        'Debug.Print("Lost_Focus:" & Now)
    End Sub
    '============================
    ' [MS00] 按下滑鼠左鍵
    '============================
    Private Sub Form_Spec_CHK_Top_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        Dim p_XY As Point
        Dim x_gap As Integer = 0
        Dim op_X As Integer


        dd1 = Mouse_X
        dd2 = Mouse_Y

        p_XY.X = dd1 + 1
        p_XY.Y = dd2 + 1

        'If p_XY.X > 756 Then
        '    If S_W_Ratio <> 1 Then
        '        x_gap = 16
        '    End If
        'End If


        Dim xBase As Single
        If Form_Spec.CKB_Lo_dpi.Checked = True Then
            xBase = 0.5
        Else
            xBase = 1
        End If

        op_X = x_base + ((dd1 - x_base) / S_W_Ratio) / xBase

        If op_X > 756 Then
            x_gap = (1 - S_W_Ratio) * 32
        End If



        Form_Spec_CHK_Ctl.TB_XX.Text = (op_X + x_gap)
        Form_Spec_CHK_Ctl.TB_YY.Text = y_base + ((dd2 - y_base) / S_H_Ratio) / xBase

        '-----------------
        '   拖曳繪圖方塊
        '-----------------
        PicBox_Drag.Location = p_XY

        PicBox_Drag.Width = 1
        PicBox_Drag.Height = 1
        PicBox_Drag.Visible = True
    End Sub

    Private Sub Form_Spec_CHK_Top_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If (e.X <= 0) Then
            Return
        End If
        Mouse_X = e.X
        Mouse_Y = e.Y

        '-----------------
        '   拖曳繪圖方塊
        '-----------------
        PicBox_Drag.Width = Mouse_X - dd1 - 1
        PicBox_Drag.Height = Mouse_Y - dd2 - 1

        Form_Spec_CHK_Ctl.LB_XY.Text = e.X & " , " & e.Y
    End Sub
    '============================
    ' [MS01] 放開 滑鼠左鍵
    '============================
    Private Sub Form_Spec_CHK_Top_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        dd3 = Mouse_X - dd1
        dd4 = Mouse_Y - dd2


        Dim x_Base As Single
        If Form_Spec.CKB_Lo_dpi.Checked = True Then
            x_Base = 0.5
        Else
            x_Base = 1
        End If

        Form_Spec_CHK_Ctl.TB_WW.Text = (dd3 / S_W_Ratio) / x_Base
        Form_Spec_CHK_Ctl.TB_HH.Text = (dd4 / S_H_Ratio) / x_Base


        Draw_RectangleMouse_UP(dd1, dd2, dd3, dd4)


        PicBox_Arrow.Visible = False

        '-----------------
        '   拖曳繪圖方塊
        '-----------------
        PicBox_Drag.Visible = False

    End Sub

    Private Sub Form_Spec_CHK_Top_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        If e.Delta > 0 Then
            Form_Spec.Page_Dn()
        Else

            Form_Spec.Page_Up()
        End If
    End Sub


    Private Sub Form_Spec_CHK_Top_Move(sender As Object, e As EventArgs) Handles Me.Move
        Form_Spec_CHK_Bottom.Location = Me.Location
    End Sub

    Private Sub Form_Spec_CHK_Top_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Form_Spec_CHK_Bottom.Size = Me.Size

        'S_H_Ratio = Me.Height / Ori_Height
        'S_W_Ratio = Me.Width / Ori_Width

        'PictureBox1.Height = P1_Ori_Height * S_H_Ratio
        'PictureBox1.Width = P1_Ori_Width * S_W_Ratio
    End Sub

    Private Sub Form_Spec_CHK_Top_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Debug.Print("" & Me.Width)
    End Sub
    Public Sub Drag_Cal_Ratio_and_Resize()


        S_H_Ratio = (Me.Height - y_base) / (Form_Ori_Height - y_base)

        If Form_Spec.CKB_compare.Checked = True Then
            S_W_Ratio = (Me.Width - x_base) / (Form_Ori_X2_Width - x_base)
        Else
            Debug.Print("Now_w:" & Me.Width & ", ori_w:" & Form_Ori_Width)

            S_W_Ratio = (Me.Width - x_base) / (Form_Ori_Width - x_base)
        End If


        Debug.Print("New Ratio h:" & S_H_Ratio & ", w:" & S_W_Ratio)

        PicBox_TopMain.Height = P1_Ori_Height * S_H_Ratio
        PicBox_TopMain.Width = P1_Ori_Width * S_W_Ratio

        Form_Spec_CHK_Bottom.PicBox_Main.Height = PicBox_TopMain.Height
        Form_Spec_CHK_Bottom.PicBox_Main.Width = PicBox_TopMain.Width
    End Sub
    Public Sub XChange_Ratio_and_Resize(ByVal X_Value As Single)
        Dim op_tempX As Integer
        Dim op_tempY As Integer

        S_H_Ratio = X_Value
        S_W_Ratio = X_Value


        op_tempX = Form_Ori_Width - x_base
        op_tempY = Form_Ori_Height - y_base

        '--------------------
        '   視窗 - 宽度調整
        '--------------------

        Me.Width = x_base + (op_tempX * S_W_Ratio)

        Me.Height = y_base + (op_tempY * S_H_Ratio)

        ' Me.Width = Ori_Width * S_W_Ratio

        '--------------------
        '   圖形方塊 - 宽度調整
        '--------------------
        PicBox_TopMain.Height = P1_Ori_Height * S_H_Ratio
        PicBox_TopMain.Width = P1_Ori_Width * S_W_Ratio

        Form_Spec_CHK_Bottom.PicBox_Main.Height = PicBox_TopMain.Height
        Form_Spec_CHK_Bottom.PicBox_Main.Width = PicBox_TopMain.Width

    End Sub

    Private Sub Form_Spec_CHK_Top_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        'Drag_Cal_Ratio_and_Resize()
        Timer_Resize.Enabled = False
        Timer_Resize.Enabled = True

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer_OnTOP.Tick
        '--------------------
        '經過1秒後, On Top
        '--------------------

        Timer_OnTOP.Enabled = False
        Set_to_TOP()

        If first_flag = False Then
            first_flag = True
            Form_Spec.Page_Now()
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer_Resize.Tick
        Timer_Resize.Enabled = False

        If X_ReSize_flag = True Then
            Debug.Print("Tmr2_debug")
            Debug.Print("Now_w:" & Me.Width & ", ori_w:" & Form_Ori_Width)

            Drag_Cal_Ratio_and_Resize()

            Switch_Compare_Mode()

            Form_Spec.Page_Now()
            Debug.Print("NFS_7")
            Form_Spec.Update_for_Sync()

        End If

    End Sub
End Class