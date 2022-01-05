
'Public Shared Function GetWindowInfo(ByVal hwnd As IntPtr, ByRef pwi As WINDOWINFO) As Boolean

'End Function
Imports System.IO '新增此命名空間
Imports System.Drawing
Imports System.Drawing.Image




'usenamespace
'System.Drawing.Imaging



'Public Declare Function ShellExecute Lib "shell32.dll " Alias "ShellExecuteA " (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'Public Declare Auto Function GetWindowInfo Lib "user32" _(ByVal hwnd As IntPtr, ByRef pwi As WINDOWINFO) As Boolean

'Declare Function BitBlt Lib "gdi32" _(ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Class Form_Spec_CHK_Ctl
    Public Setup_first As Boolean = False

    Public USE_Color_Tool As Boolean = False


    Public ColorTool As Spec_Item_Color

    Public Npp_file_name As String

    Public Const Default_Title As String = "Form_Spec_CHK_Control"

    Public Compare_Page_List As New ListBox
    Public BTN_Action_Enable As Boolean = False
    ' Public BTN_SSWITCH_Enable As Boolean = False

    Public BTN_Select_flag As Boolean = False
    Public Sever_Type As Integer

    Public Select_Label As Integer
    Public Select_Label_Item As Integer
    Public Page_Items_Now As Integer


    'Public LB_Array(0) As Label
    'Public LB_cnt As Integer = 0
    'Public LB_cnt_Enable As Boolean = False

    Public Structure Block_Item
        Dim rect As RectangleF
        Dim Total_Area As Integer

        Dim Draw As Boolean     '1
        Dim Number As Integer   '2
        Dim Title1 As String        '3
        Dim Title1_sub As String    '4
        Dim Title2 As String        '5
        Dim Title2_sub As String    '6
        Dim Title3 As String        '7
        Dim Title3_sub As String    '8

        Dim XX As Integer       '9
        Dim YY As Integer       '10
        Dim Width As Integer    '11
        Dim Height As Integer   '12
        Dim Font_Color As Integer   '13
        Dim BG_Color As Integer   '14

        Dim Code_FileName As String   '15
        Dim Code_LineNumber As String   '16
        Dim Result As String   '17
        Dim Note As String   '18
        Dim Disp As Byte   '19

        Dim Date_Create As String   '20
        Dim Date_Modify As String   '21



    End Structure




    Public block_x(0) As Block_Item


    Public Shared Function GetWindowInfo(ByVal hwnd As IntPtr, ByRef pwi As WINDOWINFO) As Boolean
        Return True
    End Function


    Public Structure RECT
        Private _Left As Integer, _Top As Integer, _Right As Integer, _Bottom As Integer

        Public Sub New(ByVal Rectangle As Rectangle)
            Me.New(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
        End Sub
        Public Sub New(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer)
            _Left = Left
            _Top = Top
            _Right = Right
            _Bottom = Bottom
        End Sub

        Public Property X As Integer
            Get
                Return _Left
            End Get
            Set(ByVal value As Integer)
                _Right = _Right - _Left + value
                _Left = value
            End Set
        End Property
        Public Property Y As Integer
            Get
                Return _Top
            End Get
            Set(ByVal value As Integer)
                _Bottom = _Bottom - _Top + value
                _Top = value
            End Set
        End Property
        Public Property Left As Integer
            Get
                Return _Left
            End Get
            Set(ByVal value As Integer)
                _Left = value
            End Set
        End Property
        Public Property Top As Integer
            Get
                Return _Top
            End Get
            Set(ByVal value As Integer)
                _Top = value
            End Set
        End Property
        Public Property Right As Integer
            Get
                Return _Right
            End Get
            Set(ByVal value As Integer)
                _Right = value
            End Set
        End Property
        Public Property Bottom As Integer
            Get
                Return _Bottom
            End Get
            Set(ByVal value As Integer)
                _Bottom = value
            End Set
        End Property
        Public Property Height() As Integer
            Get
                Return _Bottom - _Top
            End Get
            Set(ByVal value As Integer)
                _Bottom = value + _Top
            End Set
        End Property
        Public Property Width() As Integer
            Get
                Return _Right - _Left
            End Get
            Set(ByVal value As Integer)
                _Right = value + _Left
            End Set
        End Property
        Public Property Location() As Point
            Get
                Return New Point(Left, Top)
            End Get
            Set(ByVal value As Point)
                _Right = _Right - _Left + value.X
                _Bottom = _Bottom - _Top + value.Y
                _Left = value.X
                _Top = value.Y
            End Set
        End Property
        Public Property Size() As Size
            Get
                Return New Size(Width, Height)
            End Get
            Set(ByVal value As Size)
                _Right = value.Width + _Left
                _Bottom = value.Height + _Top
            End Set
        End Property

        Public Shared Widening Operator CType(ByVal Rectangle As RECT) As Rectangle
            Return New Rectangle(Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height)
        End Operator
        Public Shared Widening Operator CType(ByVal Rectangle As Rectangle) As RECT
            Return New RECT(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
        End Operator
        Public Shared Operator =(ByVal Rectangle1 As RECT, ByVal Rectangle2 As RECT) As Boolean
            Return Rectangle1.Equals(Rectangle2)
        End Operator
        Public Shared Operator <>(ByVal Rectangle1 As RECT, ByVal Rectangle2 As RECT) As Boolean
            Return Not Rectangle1.Equals(Rectangle2)
        End Operator

        Public Overrides Function ToString() As String
            Return "{Left: " & _Left & "; " & "Top: " & _Top & "; Right: " & _Right & "; Bottom: " & _Bottom & "}"
        End Function

        Public Overloads Function Equals(ByVal Rectangle As RECT) As Boolean
            Return Rectangle.Left = _Left AndAlso Rectangle.Top = _Top AndAlso Rectangle.Right = _Right AndAlso Rectangle.Bottom = _Bottom
        End Function
        Public Overloads Overrides Function Equals(ByVal [Object] As Object) As Boolean
            If TypeOf [Object] Is RECT Then
                Return Equals(DirectCast([Object], RECT))
            ElseIf TypeOf [Object] Is Rectangle Then
                Return Equals(New RECT(DirectCast([Object], Rectangle)))
            End If

            Return False
        End Function
    End Structure

    'Public Type RECT
    ' Left As Long
    ' Top As Long
    ' Right As Long
    ' Bottom As Long
    'End Type

    Structure WINDOWINFO
        Dim cbSize As Integer
        Dim rcWindow As RECT
        Dim rcClient As RECT
        Dim dwStyle As Integer
        Dim dwExStyle As Integer
        Dim dwWindowStatus As UInt32
        Dim cxWindowBorders As UInt32
        Dim cyWindowBorders As UInt32
        Dim atomWindowType As UInt16
        Dim wCreatorVersion As Short
    End Structure

    Public Setup_mode_flag As Boolean
    Public Mouse_X As Integer
    Public Mouse_Y As Integer

    Dim dd1 As Integer
    Dim dd2 As Integer
    Dim dd3 As Integer
    Dim dd4 As Integer





    Private Sub Button1_Click(sender As Object, e As EventArgs)


    End Sub
    'Public Sub Overlap_windows()
    '    Me.Location = Form_Spec_CHK.Location
    '    Me.Size = Form_Spec_CHK.Size
    '    Me.Opacity = Val(TextBox1.Text)

    'End Sub


    Public Sub Draw_Rectangle_Ori(D1 As Integer, D2 As Integer, D3 As Integer, D4 As Integer)
        Dim g As Graphics


        'Print("123")

        g = Me.CreateGraphics

        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        g.DrawRectangle(Pens.Black, D1, D2, D3, D4)


        g.Dispose()
    End Sub
    Public Sub Draw_RectangleNoUse(D1 As Integer, D2 As Integer, D3 As Integer, D4 As Integer, Dstr As String)
        Dim g As Graphics

        Dim drawString As String = Dstr
        Dim fnt As New Font("ＭＳ ゴシック", 12, FontStyle.Bold)

        'Print("123")

        g = Me.CreateGraphics

        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        Dim rect As New RectangleF(D1, D2, D3, D4)
        g.DrawRectangle(Pens.Blue, D1, D2, D3, D4)
        g.DrawString(drawString, fnt, Brushes.Blue, rect)

        g.Dispose()
    End Sub

    'Public Sub Set_LB_Color(ByVal CC As Color)
    '    Debug.Print(LB_cnt)

    '    If LB_Array(LB_cnt - 1).Created = True Then
    '        LB_Array(LB_cnt - 1).ForeColor = CC
    '    End If
    'End Sub
    Public Sub Draw_RectangleTest(D1 As TextBox, D2 As TextBox, D3 As TextBox, D4 As TextBox)
        Dim xx1 As Integer
        Dim xx2 As Integer
        Dim xx3 As Integer
        Dim xx4 As Integer

        Dim g As Graphics

        Dim drawString As String = "智に働けば角が立つ。情に棹させば流される。"
        Dim fnt As New Font("ＭＳ ゴシック", 12)

        xx1 = Val(D1.Text)
        xx2 = Val(D2.Text)
        xx3 = Val(D3.Text)
        xx4 = Val(D4.Text)
        'Me.currentx = xx1

        'Me.Cursor.Position = New Point(10, 20)

        'Form_Spec_CHK2.Print("123")

        g = Me.CreateGraphics

        Dim rect As New RectangleF(xx1, xx2, xx3, xx4)

        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        g.DrawRectangle(Pens.Black, xx1, xx2, xx3, xx4)
        g.DrawString(drawString, fnt, Brushes.Black, rect)

        g.Dispose()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form_Spec_CHK_Top.Graph_Clear()

    End Sub
    'Public Sub Graph_Clear()
    '    Dim g As Graphics

    '    g = Me.CreateGraphics

    '    g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

    '    g.DrawRectangle(Pens.Black, 0, 0, 1600, 1600)
    '    g.FillRectangle(Brushes.White, 0, 0, 1600, 1600)


    '    g.Dispose()
    'End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        ' Move_Label(Label1, 1)   'UP
    End Sub
    Public Sub Move_Label(ByRef xLB As Label, ByVal dir_Value As Byte)
        Dim xx As Point
        Dim zz As Integer

        'zz = Val(TextBox5.Text)

        Select Case dir_Value
            Case 0
            Case 1

                '向上移
                xx.X = xLB.Location.X
                xx.Y = xLB.Location.Y - zz
            Case 2

                '向右移
                xx.X = xLB.Location.X + zz
                xx.Y = xLB.Location.Y
            Case 3
                '向下移
                xx.X = xLB.Location.X
                xx.Y = xLB.Location.Y + zz
            Case 4
                '向左移
                xx.X = xLB.Location.X - zz
                xx.Y = xLB.Location.Y

        End Select

        'Label1.LocationX = Val(TextBox1.Text)
        'Label1.Location.Y= Val(TextBox2.Text)
        xLB.Location = xx
    End Sub





    Private Sub Form_Spec_CHK2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress


    End Sub

    Private Sub Form_Spec_CHK2_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Dim Xpath As String
        Dim BBB As String
        Dim default_Color_file As String

        Me.Text = Default_Title

        '-------------------
        ' 取得預設的Item設定顏色
        '-------------------
        default_Color_file = Form_Spec.Get_Source_DIR & "\APP\Default_item.txt"

        If (File.Exists(default_Color_file)) Then

            USE_Color_Tool = True
            ColorTool.init()
            ColorTool.set_Select_item(CMB_PreSet_Item)
            ColorTool.set_FColor_item(CMB_Font_Color)
            ColorTool.set_Color_item(CMB_BG_Color)

            ColorTool.set_Src_file(default_Color_file)
            ColorTool.Load_Src_file()
        Else
            Create_New_ColorSetting_file(default_Color_file)

        End If

        Load_SearchPath()

        'Overlap_windows()
        'CMB_Search_Path.Items.Add(Form_Spec.Get_Source_DIR)
        'CMB_Search_Path.SelectedIndex = 0

        'Xpath = Form_Spec.Get_Source_DIR & "\APP\Search_Path.txt"
        'Read_Search_Path(CMB_Search_Path, Xpath, False)

        '----------------
        ' 建立LOG資料夾
        '----------------

        BBB = Form_Spec.Get_filePath_x_Append & "\LOG"

        If Not IO.Directory.Exists(BBB) Then

            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(BBB)
        End If


        '========================
        '   加入提示項目
        '========================
        ToolTip1.IsBalloon = True
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
        ToolTip1.ToolTipTitle = "Note:"
        ToolTip1.UseAnimation = True
        ToolTip1.UseFading = True
        '------------------------
        'ToolTip1.SetToolTip(Button6, "ABC")

        ToolTip1.SetToolTip(Btn_to_Modify, "Copy_to_Modify")      'EEPROM
    End Sub
    Public Sub Load_SearchPath()
        Dim Xpath As String

        CMB_Search_Path.Items.Clear()
        CMB_Search_Path.Items.Add(Form_Spec.Get_Source_DIR)
        CMB_Search_Path.SelectedIndex = 0

        Xpath = Form_Spec.Get_Source_DIR & "\APP\Search_Path.txt"
        Read_Search_Path(CMB_Search_Path, Xpath, False)
    End Sub
    Public Sub Create_New_ColorSetting_file(ByVal AAA As String)
        'Dim filename1 As String
        Dim filenum As Integer
        Dim content As String

        filenum = FreeFile()
        FileOpen(filenum, AAA, OpenMode.Output)


        content = "CHK, 0, 5"
        PrintLine(filenum, content)
        content = "Wait_, 0, 7"
        PrintLine(filenum, content)
        content = "TP_, 0, 3"
        PrintLine(filenum, content)
        content = "VB_, 0, 4"
        PrintLine(filenum, content)
        content = "HAND_, 0, 5"
        PrintLine(filenum, content)
        content = "LA_, 0, 6"
        PrintLine(filenum, content)
        content = "OSC_, 0, 7"
        PrintLine(filenum, content)
        content = "Pass_, 0, 21"
        PrintLine(filenum, content)
        content = "OK_, 0, 2"
        PrintLine(filenum, content)
        content = "NG_, 0, 1"
        PrintLine(filenum, content)
        content = "Panel_, 0, 11"
        PrintLine(filenum, content)
        content = "Code_, 0, 12"
        PrintLine(filenum, content)
        content = "JUMP_, 13, 0"
        PrintLine(filenum, content)
        content = "XAppend, 10, 0"
        PrintLine(filenum, content)
        content = "MT1_, 0, 15"
        PrintLine(filenum, content)
        content = "MT2_, 0, 16"
        PrintLine(filenum, content)
        content = "RM1_, 0, 17"
        PrintLine(filenum, content)
        content = "RM2_, 0, 18"

        PrintLine(filenum, content)


        FileClose(filenum)


    End Sub

    Private Sub Form_Spec_CHK2_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick


    End Sub

    Private Sub Form_Spec_CHK2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles Me.MouseDoubleClick

    End Sub

    Private Sub Form_Spec_CHK2_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

    End Sub

    Private Sub Form_Spec_CHK2_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove





    End Sub

    Private Sub Form_Spec_CHK2_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        'Dim dd1 As Integer
        'Dim dd2 As Integer



    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)


        ' Draw_RectangleTest(TB_XX, TB_YY, TB_WW, TB_HH)
    End Sub

    Private Sub Form_Spec_CHK2_Move(sender As Object, e As EventArgs) Handles Me.Move
        '[kk00]
        'Form_Spec_CHK.Location = Me.Location
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TB_Result_Info.TextChanged

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Create_New_Item()

    End Sub
    Public Sub Create_New_Item()
        Form_Spec.FullName_item_CSV = Form_Spec.Get_CSV_Filename()

        '1.新增
        If (Val(TB_WW.Text) <> 0) And (Val(TB_HH.Text) <> 0) Then
            Save_Spec_CheckItem(Form_Spec.FullName_item_CSV)
        End If

        '2.重新載入
        Load_Page_Label()

        TB_CreateDate.Clear()
        TB_ModifyDate.Clear()
    End Sub
    Private Sub Save_Spec_CheckItem(TextString As String)   '[WR00]
        Dim filename1 As String
        Dim filenum As Integer
        Dim i As Integer
        Dim content As String
        Dim tempx As String
        Dim temp_log As String


        'filename1 = "d:\ex1.txt"                                        '依檔名
        'filename1 = TestResult_Output_Path & LB_GotFileName.Text & ".csv"                                      '依檔名

        filename1 = TextString
        'error 處理   'joe

        ' ResultFileName = filename1                  '開啟結果檔時使用


        '============檢查檔案是否存在==========
        If (File.Exists(filename1)) Then

        Else
            'Create_New_Script_file()
            'If (File.Exists(filename1)) Then
            'Else
            If filename1 = "" Then
                MsgBox("請檢查檔案路徑:" & filename1)
                Return
            End If

            'End If
        End If

        '[20180113] [V067a]
        'wait_ms(10)                     '寫入行資料時, 發生檔案被其它地方使用的情形, 可能是電腦變快了, 前面的程序還沒有完成，所以才加這個delay。



        filenum = FreeFile()
        FileOpen(filenum, filename1, OpenMode.Append)


        i = Val(TextBox8.Text)                                                  '1: Number
        tempx = CheckBox1.Checked                                               '0: Enable
        content = tempx & "," & i

        content = content & "," & TB_Title1.Text & "," & TB_subTitle1.Text          '2:Title1  3:Title1_sub
        content = content & "," & TB_Title2.Text & "," & TB_subTitle2.Text          '4:Title2, 5:Title2_sub
        content = content & "," & TB_Title3.Text & "," & TB_subTitle3.Text          '6:Title3, 7:Title3_sub

        content = content & "," & TB_XX.Text & "," & TB_YY.Text & "," & TB_WW.Text & "," & TB_HH.Text '8,9,10,11 方形的長寛
        temp_log = content
        '------------
        '   分次寫(1)
        '------------
        Print(filenum, content)

        '-------------------------
        '   顏色儲存值的範圍保護
        '-------------------------
        If (CMB_Font_Color.SelectedIndex = -1) Then
            CMB_Font_Color.SelectedIndex = 0
        End If

        If (CMB_BG_Color.SelectedIndex = -1) Then
            CMB_BG_Color.SelectedIndex = 0
        End If

        content = "," & CMB_Font_Color.SelectedIndex & "," & CMB_BG_Color.SelectedIndex & "," & TB_ColorG.Text & "," & TB_ColorB.Text   '12,13,14,15 顏色

        If TB_CreateDate.Text = "" Then
            TB_CreateDate.Text = Get_Now_Format()
        End If

        If TB_ModifyDate.Text = "" Then
            TB_ModifyDate.Text = Get_Now_Format()
        End If

        'content = content & "," & TB_Result_Info.Text & "," & TB_Note.Text & "," & TB_Display.Text            '16:Resolt, 17:Note
        content = content & "," & TB_Result_Info.Text & "," & TB_Note.Text & "," & TB_Display.Text & "," & TB_CreateDate.Text & "," & TB_ModifyDate.Text            '16:Resolt, 17:Note
        '------------
        '   分次寫(2)
        '------------
        PrintLine(filenum, content)
        temp_log = temp_log & content


        FileClose(filenum)

        TextBox8.Text = i + 1
        Page_Items_Now += 1
        LB_Page_Items.Text = Page_Items_Now & " 項"

        Save_LOG(1, temp_log)
    End Sub
    'Private Sub CheckItem_ReNew_and_Save(TextString As String, ByVal ItNo As Integer, ByVal KK As Block_Item)  '20200809
    Private Sub CheckItem_ReNew_and_Save(str_FileName As String)  '20200809   [WR01]
        ' Dim filename1 As String
        Dim filenum As Integer
        Dim i As Integer


        ' filename1 = TextString

        '============檢查檔案是否存在==========
        If (File.Exists(str_FileName)) Then

        Else
            'Create_New_Script_file()
            'If (File.Exists(filename1)) Then
            'Else
            If str_FileName = "" Then
                MsgBox("請檢查檔案路徑:" & str_FileName)
                Return
            End If

            'End If
        End If

        '[20180113] [V067a]
        'wait_ms(10)                     '寫入行資料時, 發生檔案被其它地方使用的情形, 可能是電腦變快了, 前面的程序還沒有完成，所以才加這個delay。



        filenum = FreeFile()

        FileOpen(filenum, str_FileName, OpenMode.Output)

        Dim temp_Be_Used As String

        i = 1


        '=================================
        '   讀取所有的方塊
        '=================================
        For Each obj As Block_Item In block_x


            temp_Be_Used = obj.Draw      '判斷是否刪除

            If temp_Be_Used.ToUpper = "TRUE" Then
                Writh_Line(filenum, obj, i)
                i += 1
            End If

            'Debug.Print("After" & i)
        Next




        FileClose(filenum)


        '總項次
        TextBox8.Text = i

    End Sub
    Public Sub Save_LOG(ByRef tTYPE As Integer, ByRef AAA As String)
        Dim filename1 As String
        Dim content As String
        Dim filenum As Integer

        If Form_Spec.CKB_compare.Checked = False Then
            filename1 = Form_Spec.Get_filePath_x_Append & "\LOG\log.txt"
        Else
            filename1 = Form_Spec.Get_filePath_x_Append & "\LOG\log_Compare.txt"
        End If


        content = Get_Now_Format() & "," & "NewADD" & "," & "PAGE_" & Form_Spec.Get_Page_Now & "," & AAA
        filenum = FreeFile()

        FileOpen(filenum, filename1, OpenMode.Append)




        PrintLine(filenum, content)


        FileClose(filenum)
    End Sub
    Private Sub LOG_Write_Block(ByVal KK As Block_Item, ByVal NN As Integer)   '[WR02]
        Dim filename1 As String
        Dim content As String
        Dim filenum As Integer
        'Dim content As String

        'filename1 = Form_Spec.Get_file_Path_x & "\LOG\log.txt"

        If Form_Spec.CKB_compare.Checked = False Then
            filename1 = Form_Spec.Get_filePath_x_Append & "\LOG\log.txt"
        Else
            filename1 = Form_Spec.Get_filePath_x_Append & "\LOG\log_Compare.txt"
        End If

        'i = ItNo                                                  '1: Number

        content = Get_Now_Format() & ",Modify," & "PAGE_" & Form_Spec.Get_Page_Now & ","
        filenum = FreeFile()

        FileOpen(filenum, filename1, OpenMode.Append)

        'tempx = KK.Draw                                               '0: Enable

        content = content & KK.Draw & "," & NN & "," & KK.Title1 & "," & KK.Title1_sub   '2:Title1    '3:Title1_sub

        'Print(filenum, content)

        content = content & "," & KK.Title2 & "," & KK.Title2_sub         '4:Title2, 5:Title2_sub
        content = content & "," & KK.Title3 & "," & KK.Title3_sub            '6:Title3, 7:Title3_sub

        content = content & "," & KK.XX & "," & KK.YY & "," & KK.Width & "," & KK.Height '8,9,10,11 方形的長寛
        Print(filenum, content)


        content = "," & KK.Font_Color & "," & KK.BG_Color & "," & KK.Code_FileName & "," & KK.Code_LineNumber   '12,13,14,15 顏色

        content = content & "," & KK.Result & "," & KK.Note & "," & KK.Disp & "," & KK.Date_Create & "," & KK.Date_Modify         '16:Resolt, 17:Note

        PrintLine(filenum, content)
        FileClose(filenum)
    End Sub
    Private Sub Writh_Line(ByVal filenum As Integer, ByVal KK As Block_Item, ByVal NN As Integer)   '[WR02]
        Dim content As String

        'i = ItNo                                                  '1: Number

        'tempx = KK.Draw                                               '0: Enable

        content = KK.Draw & "," & NN & "," & KK.Title1 & "," & KK.Title1_sub   '2:Title1    '3:Title1_sub

        'Print(filenum, content)

        content = content & "," & KK.Title2 & "," & KK.Title2_sub         '4:Title2, 5:Title2_sub
        content = content & "," & KK.Title3 & "," & KK.Title3_sub            '6:Title3, 7:Title3_sub

        content = content & "," & KK.XX & "," & KK.YY & "," & KK.Width & "," & KK.Height '8,9,10,11 方形的長寛
        Print(filenum, content)


        content = "," & KK.Font_Color & "," & KK.BG_Color & "," & KK.Code_FileName & "," & KK.Code_LineNumber   '12,13,14,15 顏色

        content = content & "," & KK.Result & "," & KK.Note & "," & KK.Disp & "," & KK.Date_Create & "," & KK.Date_Modify         '16:Resolt, 17:Note

        PrintLine(filenum, content)

    End Sub


    Private Sub Load_Page_Label()
        'Dim Name_String As String
        'Dim str_array() As String

        Form_Spec_CHK_Top.Graph_Clear()




        If Form_Spec_CHK_Top.LB_cnt_Enable = True Then
            Form_Spec_CHK_Top.Release_LB()

            For Each obj As Block_Item In block_x
                obj.Draw = False      '讀取前, 要先清空
            Next

        End If

        ' Label22.Text = Form_Spec_CHK_Top.LB_cnt_Enable


        '-------------------
        '   讀取檔案
        '   空白:預設 (*目前使用)
        '   其它: _Set1, _Set2, _Set3
        '-------------------
        'Check_Pattern_and_Save()

        'If CMB_SaveType.Text = "" Then

        '    CheckItem_Read_and_Draw(Form_Spec.FullName_item_CSV, True)
        'Else
        '    str_array = Form_Spec.FullName_item_CSV.Split(".")

        '    Form_Spec.Pattern_Type = CMB_SaveType.Text
        '    Name_String = str_array(0) & "_" & CMB_SaveType.Text & ".csv"

        '    CheckItem_Read_and_Draw(Name_String, True)
        'End If

        'Get_CSV_Pattern()

        CheckItem_Read_and_Draw(Get_CSV_Pattern(), True)
    End Sub
    Public Sub CheckItem_Read_and_Draw(ByVal SourceFile As String, ByVal Dmsg As Boolean)       ' [RD00]
        Dim Item_Counter As Integer = 0
        Dim x_leng As Integer = 0

        Dim t_X As Integer
        Dim t_Y As Integer
        Dim t_Width As Integer
        Dim t_Height As Integer


        Dim line_counter As Integer

        Dim Parameter_Line As String = ""

        Dim strArr2() As String

        line_counter = 0

        If (File.Exists(SourceFile)) Then


            Dim fs_Parameter As FileStream = New FileStream(SourceFile, FileMode.Open)
            Using sr_Parameter As StreamReader = New StreamReader(fs_Parameter,
                                                        System.Text.Encoding.Default)

                If Item_Counter = 0 Then
                    ReDim block_x(Item_Counter)
                Else
                    ReDim Preserve block_x(Item_Counter)
                End If


                While (sr_Parameter.EndOfStream <> True)

                    'If Item_Counter <> 0 Then
                    ReDim Preserve block_x(Item_Counter)
                    'ReDim block_x(Item_Counter)

                    ' End If
                    '-------------
                    '   讀第1行
                    '------------

                    line_counter += 1

                    Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題

                    strArr2 = Parameter_Line.Split(",")

                    '----------
                    '取得資料量
                    '----------
                    x_leng = UBound(strArr2)
                    'Debug.Print("UBound:" & x_leng)
                    block_x(Item_Counter).Draw = strArr2(0)
                    block_x(Item_Counter).Number = Val(strArr2(1))
                    block_x(Item_Counter).Title1 = strArr2(2)
                    block_x(Item_Counter).Title1_sub = strArr2(3)
                    block_x(Item_Counter).Title2 = strArr2(4)
                    block_x(Item_Counter).Title2_sub = strArr2(5)
                    block_x(Item_Counter).Title3 = strArr2(6)
                    block_x(Item_Counter).Title3_sub = strArr2(7)


                    block_x(Item_Counter).XX = Val(strArr2(8))
                    block_x(Item_Counter).YY = Val(strArr2(9))
                    block_x(Item_Counter).Width = Val(strArr2(10))
                    block_x(Item_Counter).Height = Val(strArr2(11))

                    block_x(Item_Counter).Font_Color = Val(strArr2(12))
                    block_x(Item_Counter).BG_Color = Val(strArr2(13))

                    block_x(Item_Counter).Code_FileName = strArr2(14)
                    block_x(Item_Counter).Code_LineNumber = strArr2(15)

                    block_x(Item_Counter).Result = strArr2(16)
                    block_x(Item_Counter).Note = strArr2(17)

                    ' Debug.Print("x_leng:" & x_leng)

                    '-------------------------------
                    '   擴充顯示項目調整功能 - 新建
                    '-------------------------------
                    If x_leng = 18 Then
                        If strArr2(18) = "" Then
                            block_x(Item_Counter).Disp = 128
                        Else
                            block_x(Item_Counter).Disp = strArr2(18)
                        End If
                    ElseIf x_leng < 18 Then
                        '----------------
                        '   小於18項時, 補上
                        '----------------
                        block_x(Item_Counter).Disp = 128
                    End If

                    '-------------------------------
                    '   擴充顯示項目調整功能 - 新建
                    '-------------------------------
                    If x_leng = 20 Then
                        If strArr2(18) = "" Then
                            block_x(Item_Counter).Disp = 128
                        Else
                            block_x(Item_Counter).Disp = strArr2(18)
                        End If

                        If strArr2(19) = "" Then
                            block_x(Item_Counter).Date_Create = Get_Now_Format()
                        Else
                            block_x(Item_Counter).Date_Create = strArr2(19)
                        End If

                        If strArr2(20) = "" Then
                            block_x(Item_Counter).Date_Modify = Get_Now_Format()
                        Else
                            block_x(Item_Counter).Date_Modify = strArr2(20)
                        End If
                    Else
                        block_x(Item_Counter).Date_Create = Get_Now_Format()
                        block_x(Item_Counter).Date_Modify = Get_Now_Format()
                    End If




                    t_X = Val(strArr2(8))           'X
                    t_Y = Val(strArr2(9))           'Y
                    t_Width = Val(strArr2(10))      'Width
                    t_Height = Val(strArr2(11))      'Height


                    block_x(Item_Counter).rect = New RectangleF(t_X, t_Y, t_Width, t_Height)
                    block_x(Item_Counter).Total_Area = (t_Width) * (t_Height)


                    Item_Counter += 1

                End While


                sr_Parameter.Close()                                                          '關閉檔案


                'Draw_Item_block()

                '累加給下一個新項次使用
                TextBox8.Text = line_counter + 1

                If (line_counter <> 0) Then
                    '----------------------------
                    '   由面積進行排序
                    '----------------------------
                    Dim Compare() As Integer = Nothing
                    ReDim Compare(UBound(block_x))


                    For i = 0 To UBound(block_x)
                        ' Debug.Print(i)
                        Compare(i) = block_x(i).Total_Area
                    Next



                    Array.Sort(Compare, block_x)
                    Array.Reverse(block_x)      '由大至小

                    For i = 0 To UBound(block_x)
                        Draw_Item_block(block_x(i))
                    Next

                End If

            End Using

        Else

            line_counter = 1

            If Dmsg = True Then
                MsgBox("Please check: " & SourceFile & ", 此為Log檔案的路徑")
            Else
                'Debug.Print("Please check: " & SourceFile)
            End If

        End If


        Page_Items_Now = Item_Counter
        LB_Page_Items.Text = Page_Items_Now & "項"

        'Dim fs_Parameter As FileStream = New FileStream("port_setting.csv", FileMode.Open)

    End Sub
    Public Sub Draw_Item_block(ByVal XX As Block_Item)
        Dim t_X As Integer
        Dim t_Y As Integer
        Dim t_Width As Integer
        Dim t_Height As Integer
        Dim Dstr As String
        Dim tFont_Color As Byte
        Dim tBG_Color As Byte

        '-----------------------
        '   是否繪製
        '-----------------------

        If XX.Draw = "TRUE" Then

            t_X = XX.XX       'X
            t_Y = XX.YY       'Y
            t_Width = XX.Width     'Width
            t_Height = XX.Height     'Height


            ' Dstr = strArr2(2) & "_" & strArr2(3) & "_" & strArr2(4) & "_" & strArr2(5) & "_" & strArr2(6) & "_" & strArr2(7)


            '-----------------------
            '   顯示文字項目處理
            '-----------------------

            Dstr = ""

            If ((XX.Disp And 128) = 128) Then
                'Dstr = Dstr & strArr2(2)
                Dstr = Dstr & XX.Title1
            End If

            If ((XX.Disp And 64) = 64) Then
                'Dstr = Dstr & "_" & strArr2(3)
                Dstr = Dstr & "_" & XX.Title1_sub
            End If

            If ((XX.Disp And 32) = 32) Then
                'Dstr = Dstr & "_" & strArr2(4)
                Dstr = Dstr & "_" & XX.Title2
            End If

            If ((XX.Disp And 16) = 16) Then
                'Dstr = Dstr & "_" & strArr2(5)
                Dstr = Dstr & "_" & XX.Title2_sub
            End If

            If ((XX.Disp And 8) = 8) Then
                'Dstr = Dstr & "_" & strArr2(6)
                Dstr = Dstr & "_" & XX.Title3
            End If

            If ((XX.Disp And 4) = 4) Then
                'Dstr = Dstr & "_" & strArr2(7)
                Dstr = Dstr & "_" & XX.Title3_sub
            End If

            If ((XX.Disp And 2) = 2) Then
                'Dstr = Dstr & "_" & strArr2(16)
                Dstr = Dstr & "_" & XX.Result
            End If

            If ((XX.Disp And 1) = 1) Then
                'Dstr = Dstr & "_" & strArr2(17)
                Dstr = Dstr & "_" & XX.Note
            End If







            tFont_Color = XX.Font_Color   '12:字顏色
            tBG_Color = XX.BG_Color     '13:背景色 

            ' If Dcolor = 0 Then
            'Draw_RectangleNoUse(DD1, DD2, DD3, DD4, Dstr)
            'Else
            Form_Spec_CHK_Top.Draw_RectangleNow(t_X, t_Y, t_Width, t_Height, Dstr, tFont_Color, tBG_Color)
            'End If



        End If
    End Sub

    Public Sub Read_Search_Path(ByRef XXX As ComboBox, ByVal SourceFile As String, ByVal dMSG As Boolean)       '20200818 [PH00]


        Dim Parameter_Line As String = ""



        If (File.Exists(SourceFile)) Then


            Dim fs_Parameter As FileStream = New FileStream(SourceFile, FileMode.Open)
            Using sr_Parameter As StreamReader = New StreamReader(fs_Parameter,
                                                        System.Text.Encoding.Default)

                While (sr_Parameter.EndOfStream <> True)


                    Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題

                    If (Mid(Parameter_Line, 1, 1) <> "#") Then

                        XXX.Items.Add(Parameter_Line)
                    End If



                End While


                sr_Parameter.Close()                                                          '關閉檔案




            End Using

        Else


            If dMSG = True Then
                MsgBox("Please check: " & SourceFile & ", 此為Log檔案的路徑")
            End If

        End If


        'Dim fs_Parameter As FileStream = New FileStream("port_setting.csv", FileMode.Open)

    End Sub
    'Public Sub Add_Button_Label(ByVal XX As Integer, ByVal YY As Integer, ByVal SS As String, ByVal CC As Color)


    '    Dim fnt As New Font("Cooper Black", 12, FontStyle.Bold)

    '    Dim xxx As New Point
    '    xxx.X = XX + 1
    '    xxx.Y = YY + 1

    '    LB_Array(LB_cnt) = New Label
    '    LB_Array(LB_cnt).Font = fnt
    '    LB_Array(LB_cnt).Name = "LBA_" & LB_cnt
    '    LB_Array(LB_cnt).Location = xxx
    '    LB_Array(LB_cnt).Text = SS
    '    LB_Array(LB_cnt).AutoSize = True


    '    LB_Array(LB_cnt).BackColor = Color.White

    '    Me.Controls.Add(LB_Array(LB_cnt))

    '    AddHandler LB_Array(LB_cnt).Click, AddressOf Btns_Click

    '    LB_cnt += 1



    '    If LB_cnt <> 0 Then
    '        ReDim Preserve LB_Array(LB_cnt)

    '    End If

    '    LB_cnt_Enable = True

    'End Sub
    '========================================
    '   在物件上按下click '[ST00]
    '========================================
    Public Sub Btns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ' Dim xx_str As String

        'Label1.Text = ""


        'If Label1.Text = "" Then
        '    Label1.Text = sender.Text
        'Else
        '    Label1.Text &= vbCrLf & sender.Text

        'End If

        BTN_Action_Enable = True
        'BTN_SSWITCH_Enable = True

        LB_SelectedItem.Text = sender.name


        Reflash_BlockItemContent_ListBox()




    End Sub
    '[ST01]
    Public Sub Reflash_BlockItemContent_ListBox()
        Dim xx_p As Point

        Dim xx As Integer

        Dim strArr2() As String
        Dim xx_cnt As Integer = 0
        Dim yy_array() As String


        'Debug.Print("Select:" & LB_SelectedItem.Text)

        ListBox_Item_content.Items.Clear()

        If Page_Items_Now = 0 Then
            Exit Sub
        End If

        If LB_SelectedItem.Text = "" Then
            Exit Sub
        End If


        strArr2 = LB_SelectedItem.Text.Split("_")
        'ListBox1.Text = ""

        ' xx = Val(sender.text)
        xx = Val(strArr2(1))
        Select_Label = xx

        'If Label6.Text = "" Then
        '    Label6.Text = sender.Text
        'Else
        '    Label6.Text &= vbCrLf & sender.Text
        'End If
        'ListBox1.

        'ListBox1.Items.Add(sender.text)
        'ListBox1.Items.Add(" ")

        Get_Counter_to_String(xx_cnt)
        'ListBox1.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Draw)

        Get_Counter_to_String(xx_cnt)
        'ListBox1.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Number)
        ListBox_Item_content.Items.Add("")

        Debug.Print("sss:" & UBound(block_x))

        If xx <= UBound(block_x) Then
            'If xx <= block_x.Length Then
            Debug.Print("sss:" & block_x.Length & "," & xx)

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Title1)
            'ListBox1.Items.Add("")



            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Title1_sub)
            ListBox_Item_content.Items.Add("")

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Title2)

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Title2_sub)
            ListBox_Item_content.Items.Add("")

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Title3)

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Title3_sub)
            ListBox_Item_content.Items.Add("")

            Get_Counter_to_String(xx_cnt)
            'ListBox1.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).XX)
            TBx.Text = block_x(xx).XX

            Get_Counter_to_String(xx_cnt)
            'ListBox1.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).YY)
            TBy.Text = block_x(xx).YY

            '20210104

            Dim opD1_plus As Integer = 0
            Dim opD2_plus As Integer = 0

            'If Form_Spec_CHK_Top.S_W_Ratio = 1 Then    '20210104
            If block_x(xx).XX > 765 Then
                If Form_Spec_CHK_Top.S_W_Ratio = 1 Then
                    opD1_plus = 0
                Else
                    opD1_plus = -16
                End If


            Else
                opD1_plus = 0

            End If


            If block_x(xx).XX > 765 Then
                opD2_plus = 16
                'opD1_plus = 0
            Else
                opD2_plus = 0

            End If
            'End If

            Dim x_dpi_base As Single
            If Form_Spec.CKB_Lo_dpi.Checked = True Then

                x_dpi_base = 0.5
            Else
                x_dpi_base = 1
            End If

            xx_p.X = Form_Spec_CHK_Top.x_base + ((opD2_plus + ((block_x(xx).XX - Form_Spec_CHK_Top.x_base - opD2_plus) * Form_Spec_CHK_Top.S_W_Ratio) + opD1_plus) * x_dpi_base)
            xx_p.Y = Form_Spec_CHK_Top.y_base + (((block_x(xx).YY - Form_Spec_CHK_Top.y_base) * Form_Spec_CHK_Top.S_H_Ratio) * x_dpi_base)

            Form_Spec_CHK_Top.PicBox_Arrow.Location = xx_p
            Form_Spec_CHK_Top.PicBox_Arrow.Visible = True
            BTN_Select_flag = True

            Get_Counter_to_String(xx_cnt)
            'ListBox1.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Width)
            TBw.Text = block_x(xx).Width

            Get_Counter_to_String(xx_cnt)
            'ListBox1.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Height)
            TBh.Text = block_x(xx).Height

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Font_Color)

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).BG_Color)
            ListBox_Item_content.Items.Add("")

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Code_FileName)
            TB_KeyWord.Text = block_x(xx).Code_FileName

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Code_LineNumber)
            TB_Code_Line.Text = block_x(xx).Code_LineNumber.Replace("LN_", "")

            ListBox_Item_content.Items.Add("")
            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Result)

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Note)
            Npp_file_name = block_x(xx).Note
            Get_Counter_to_String(xx_cnt)
            'ListBox1.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Disp)
            ListBox_Item_content.Items.Add("")

            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Date_Create)
            ListBox_Item_content.Items.Add(Get_Counter_to_String(xx_cnt) & block_x(xx).Date_Modify)
            '-------------
            '   Label顯示的處理
            '-------------
            Set_to_disp_CheckBox(block_x(xx).Disp)

            '======================
            '  特別行為的處理
            '======================





            If BTN_Action_Enable = True Then

                '--------------------
                '   切換頁面
                '--------------------
                If block_x(xx).Title1.Contains("JUMP_") Then

                    yy_array = block_x(xx).Title1.Split("_")
                    TextBox1.Text = yy_array(1)

                    If CKB_JUMP.Checked = True Then
                        Jump_to_Page(Val(TextBox1.Text))
                        BTN_Action_Enable = False
                        'BTN_SSWITCH_Enable = False
                        Exit Sub
                    End If

                End If

                '--------------------
                '   切換仕樣
                '--------------------
                If CKB_XLINK.Checked = True Then
                    If block_x(xx).Title1.Contains("XLINK") Then

                        Form_Spec.LB_LastSpec.Text = block_x(xx).Title1_sub & "," & block_x(xx).Title2 & "," & block_x(xx).Title2_sub
                        Form_Spec.Switch_temp_Spec()
                        BTN_Action_Enable = False
                        Exit Sub
                    End If
                End If

                '--------------------
                '   開啟附件
                '--------------------
                If CKB_append.Checked = True Then
                    If block_x(xx).Title1.Contains("XAppend") Then

                        Open_Xfile(block_x(xx).Note)
                    End If

                End If

                BTN_Action_Enable = False
                'Exit Sub
            End If






        End If


    End Sub
    Public Sub Set_to_disp_CheckBox(ByVal dValue As Byte)

        Set_CHECKBOX(CKB_Title1, dValue, 128)
        Set_CHECKBOX(CheckBox4, dValue, 64)
        Set_CHECKBOX(CheckBox5, dValue, 32)
        Set_CHECKBOX(CheckBox6, dValue, 16)
        Set_CHECKBOX(CheckBox7, dValue, 8)
        Set_CHECKBOX(CheckBox8, dValue, 4)
        Set_CHECKBOX(CheckBox9, dValue, 2)
        Set_CHECKBOX(CheckBox10, dValue, 1)

    End Sub
    Private Sub Set_CHECKBOX(ByRef AA As CheckBox, ByVal bb As Byte, ByVal CC As Byte)
        If ((bb And CC) = CC) Then
            AA.Checked = True
        Else
            AA.Checked = False
        End If
    End Sub

    Private Function Get_Counter_to_String(ByRef AA As Integer) As String
        Dim Bb As Integer
        Bb = AA
        AA += 1

        Return "[" & Bb & "],"

    End Function

    'Private Sub button_Click(sender As Object, e As EventArgs)

    '    '

    '    Label1.Text = sender.Text

    '    ' textBox4.Text = "索引值: " + ((Button)(sender)).TabIndex.ToString();


    'End Sub






    'End Sub




    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    '    Picture_ReSize_save(100, 200, Form_Spec.Get_Full_Small_Name)
    'End Sub
    Public Sub Picture_ReSize_save(ByVal xx As Integer, ByVal yy As Integer, ByVal zz As String)

        'picturebox => img
        Dim img As Image = Form_Spec_CHK_Top.PicBox_TopMain.Image

        Dim bmp As New Bitmap(xx, yy)

        ' img => bmp
        Dim gfx As Graphics = Graphics.FromImage(bmp)
        gfx.DrawImage(img, 0, 0, xx, yy)

        gfx.Dispose()

        bmp.Save(zz, System.Drawing.Imaging.ImageFormat.Png)


        bmp.Dispose()
    End Sub
    Public Sub Picture_save_New(ByRef AAA As PictureBox, ByVal zz As String)

        'picturebox => img
        Dim img As Image = AAA.Image

        Dim bmp As New Bitmap(AAA.Width, AAA.Height)

        ' img => bmp
        Dim gfx As Graphics = Graphics.FromImage(bmp)
        gfx.DrawImage(img, 0, 0, AAA.Width, AAA.Height)

        gfx.Dispose()

        bmp.Save(zz, System.Drawing.Imaging.ImageFormat.Png)

        bmp.Dispose()
    End Sub

    Public Sub Picture_save_New_Chart(ByRef dChart As DataVisualization.Charting.Chart, ByVal zz As String)


        'dChart.SaveImage(zz, ChartImageFormat.Png)

        dChart.SaveImage(Form_Spec.Get_Source_DIR & "\Png\" & zz & ".png", System.Drawing.Imaging.ImageFormat.Png)



    End Sub




    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click

        Form_Spec.Page_Now()
        Debug.Print("NFS_9")
        Form_Spec.Update_for_Sync()
        ' Form_Spec.Update_pic()
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Form_Spec.Page_Up()
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        Form_Spec.Page_Dn()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CKB_Ori_Graphic.CheckedChanged
        Debug.Print("NFS_10")
        Form_Spec.Update_for_Sync()
    End Sub

    Public Sub Button7_Click_1(sender As Object, e As EventArgs)
        'Form_Spec_CHK_Top.Release_LB()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        '--------------
        '   修改Item內容
        '--------------
        Modify_BlockItem()
    End Sub

    Public Sub Modify_BlockItem() '[MD00]
        'Dim tp_lb As Block_Item
        Dim Modify_Value As String

        Modify_Value = TB_Modify_Value.Text

        Select_Label_Item = CMB_ItemType.SelectedIndex

        Select Case Select_Label_Item
            Case 0
                block_x(Select_Label).Draw = Modify_Value

            Case 1
                block_x(Select_Label).Number = Val(Modify_Value)
            Case 2
                block_x(Select_Label).Title1 = Modify_Value
            Case 3
                block_x(Select_Label).Title1_sub = Modify_Value
            Case 4
                block_x(Select_Label).Title2 = Modify_Value
            Case 5
                block_x(Select_Label).Title2_sub = Modify_Value
            Case 6
                block_x(Select_Label).Title3 = Modify_Value
            Case 7
                block_x(Select_Label).Title3_sub = Modify_Value
            Case 8
                block_x(Select_Label).XX = Val(Modify_Value)
            Case 9
                block_x(Select_Label).YY = Val(Modify_Value)
            Case 10
                block_x(Select_Label).Width = Val(Modify_Value)
            Case 11
                block_x(Select_Label).Height = Val(Modify_Value)
            Case 12
                '-------------
                ' 調整 - 字型顏色
                '-------------
                block_x(Select_Label).Font_Color = Val(Modify_Value)
                CMB_Font_Color.SelectedIndex = Val(Modify_Value)
            Case 13
                '-------------
                ' 調整 - 背景顏色
                '-------------
                block_x(Select_Label).BG_Color = Val(Modify_Value)
                CMB_BG_Color.SelectedIndex = Val(Modify_Value)

            Case 14
                '-------------
                ' 調整 - Code_FileName
                '-------------
                block_x(Select_Label).Code_FileName = Modify_Value
            Case 15
                '-------------
                ' 調整 - Code_LineNumber
                '-------------
                block_x(Select_Label).Code_LineNumber = Modify_Value

            Case 16
                block_x(Select_Label).Result = Modify_Value
            Case 17
                block_x(Select_Label).Note = Modify_Value

            Case 18
                '-----------------
                ' 調整 - 顯示狀態
                '-----------------
                block_x(Select_Label).Disp = Get_Disp_Chkbox()

            Case 19
                'Date_Create
                block_x(Select_Label).Date_Create = Modify_Value
            Case 20
                'Date_Modify
                block_x(Select_Label).Date_Modify = Modify_Value


            Case 21
                '-------------
                ' 調整 - 大小
                '-------------
                block_x(Select_Label).XX = Val(TB_XX.Text)
                block_x(Select_Label).YY = Val(TB_YY.Text)
                block_x(Select_Label).Width = Val(TB_WW.Text)
                block_x(Select_Label).Height = Val(TB_HH.Text)



        End Select




        If CKB_Date_Lock.Checked = True Then

        Else
            block_x(Select_Label).Date_Modify = Get_Now_Format()
        End If


        LOG_Write_Block(block_x(Select_Label), Select_Label)



        Check_Pattern_and_Save()





        Load_Page_Label()

        Reflash_BlockItemContent_ListBox()

    End Sub
    Public Sub Check_Pattern_and_Save()
        'Dim Name_String As String
        'Dim str_array() As String

        'If CMB_SaveType.Text = "" Then


        '    CheckItem_ReNew_and_Save(Form_Spec.FullName_item_CSV)

        'Else
        '    Form_Spec.Pattern_Type = CMB_SaveType.Text

        '    str_array = Form_Spec.FullName_item_CSV.Split(".")

        '    Name_String = str_array(0) & "_" & CMB_SaveType.Text & ".csv"

        '    CheckItem_ReNew_and_Save(Name_String)

        'End If

        CheckItem_ReNew_and_Save(Get_CSV_Pattern())
    End Sub
    Public Function Get_CSV_Pattern() As String
        'Dim str_array() As String
        Dim Rtn_Name As String

        'If CMB_SaveType.Text = "" Then

        '    Rtn_Name = Form_Spec.FullName_item_CSV
        '    'CheckItem_ReNew_and_Save()

        'Else
        '    Form_Spec.Pattern_Type = CMB_SaveType.Text

        '    str_array = Form_Spec.FullName_item_CSV.Split(".")

        '    Rtn_Name = str_array(0) & "_" & CMB_SaveType.Text & ".csv"

        '    'CheckItem_ReNew_and_Save(Name_String)

        'End If

        Rtn_Name = Form_Spec.FullName_item_CSV

        Return Rtn_Name
    End Function

    '

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        CMB_ItemType.SelectedIndex = 0
        TB_Modify_Value.Text = "false"
    End Sub
    Public Sub Delete_Item()
        '===================
        '   刪除方塊
        '===================
        CMB_ItemType.SelectedIndex = 0
        TB_Modify_Value.Text = "false"
        Modify_BlockItem()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_PreSet_Item.SelectedIndexChanged
        TB_Title1.Text = CMB_PreSet_Item.Text

        If USE_Color_Tool = True Then
            ColorTool.Get_Color(CMB_PreSet_Item.SelectedIndex)
        Else
            If CMB_PreSet_Item.SelectedIndex = 12 Then            'JUMP
                CMB_Font_Color.SelectedIndex = 13               '鮭魚色
            Else
                CMB_Font_Color.SelectedIndex = 0
            End If
        End If

    End Sub
    '[ST02]
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox_Item_content.SelectedIndexChanged
        Dim xx As String
        Dim zz As String
        Dim z1 As String
        Dim z2 As String
        Dim xNO As Integer
        Dim yy_array() As String

        Dim Server_array() As String

        xx = ListBox_Item_content.SelectedItem

        If xx <> "" Then

            '-------------------
            '   選擇項目的處理
            '-------------------
            yy_array = xx.Split(",")


            '-------------------
            '   移除前面的括號_處理
            '-------------------
            zz = yy_array(0)
            z1 = zz.Replace("[", "")
            z2 = z1.Replace("]", "")

            xNO = Val(z2)
            'Debug.Print(xNO)

            CMB_ItemType.SelectedIndex = xNO
            CMB_ItemType.Text = CMB_ItemType.SelectedItem
            '-------------------
            '   仕樣書轉跳的處理
            '-------------------

            'If yy_array(1).Contains("XLINK") Then
            '    Debug.Print("Xlink to:")
            '    Form_Spec.LB_LastSpec.Text = Get_Xlink()
            '    Form_Spec.Switch_temp_Spec()

            'End If

            '-------------------
            '   特定字串_處理
            '-------------------
            If yy_array(1).Contains("JUMP_") Then

                Server_array = yy_array(1).Split("_")
                TextBox1.Text = Server_array(1)

                If CKB_JUMP.Checked = True Then
                    Jump_to_Page(Val(TextBox1.Text))
                End If

            End If

            '-------------------
            '   開啟附件的處理
            '-------------------
            'If yy_array(1).Contains("XAppend") Then
            '    If CKB_append.Checked = True Then
            '        Open_Xfile()
            '    End If

            'End If

            '---------------
            ' Redmine_Server 的判斷
            '---------------
            '(1)issue
            If yy_array(1).Contains("RM1_") Then
                Sever_Type = 1
                RadioButton1.Checked = True
                Server_array = yy_array(1).Split("_")
                TextBox17.Text = Server_array(1)

                If CheckBox12.Checked = True Then
                    Open_Redmine()
                End If

            End If

            '(2)Projects
            If yy_array(1).Contains("RM2_") Then
                Sever_Type = 2
                RadioButton2.Checked = True
                Server_array = yy_array(1).Split("_")
                TextBox17.Text = Server_array(1)
                'Debug.Print(UBound(Server_array))
                If UBound(Server_array) >= 2 Then
                    TextBox17.Text = TextBox17.Text & "_" & Server_array(2)
                    If CheckBox12.Checked = True Then
                        Open_Redmine()
                    End If
                End If

            End If

            '---------------
            ' Mantis_Server 的判斷
            '---------------
            If yy_array(1).Contains("MT1_") Then
                Sever_Type = 1
                RadioButton1.Checked = True
                Server_array = yy_array(1).Split("_")
                TextBox21.Text = Server_array(1)
                If CheckBox11.Checked = True Then
                    Open_Mantis()
                End If
            End If

            If yy_array(1).Contains("MT2_") Then
                Sever_Type = 2
                RadioButton2.Checked = True
                Server_array = yy_array(1).Split("_")
                TextBox21.Text = Server_array(1)
                If CheckBox11.Checked = True Then
                    Open_Mantis()
                End If
            End If

            '---------------
            '   行號的判斷
            '---------------
            If yy_array(1).Contains("LN_") Then

                Server_array = yy_array(1).Split("_")
                TB_Code_Line.Text = Server_array(1)

            End If

            If yy_array(1).Contains("Line_") Then

                Server_array = yy_array(1).Split("_")
                TB_Code_Line.Text = Server_array(1)

            End If

            If yy_array(1).Contains("LINE_") Then

                Server_array = yy_array(1).Split("_")
                TB_Code_Line.Text = Server_array(1)

            End If

            If yy_array(1).Contains("line_") Then

                Server_array = yy_array(1).Split("_")
                TB_Code_Line.Text = Server_array(1)

            End If

            TB_KeyWord.Text = yy_array(1)
        End If

    End Sub
    Public Function Get_Xlink() As String
        Dim Out_srt As String = ""
        Dim xx_srt As String
        Dim xx_srt_array() As String

        xx_srt = ListBox_Item_content.Items(2)
        Debug.Print("X1:" & xx_srt)
        If xx_srt <> "" Then
            xx_srt_array = xx_srt.Split(",")
            Out_srt = xx_srt_array(1) & ","
        End If

        xx_srt = ListBox_Item_content.Items(4)
        Debug.Print("X2:" & xx_srt)
        If xx_srt <> "" Then
            xx_srt_array = xx_srt.Split(",")
            Out_srt = Out_srt & xx_srt_array(1) & ","
        End If

        xx_srt = ListBox_Item_content.Items(5)
        Debug.Print("X3:" & xx_srt)
        If xx_srt <> "" Then
            xx_srt_array = xx_srt.Split(",")
            Out_srt = Out_srt & xx_srt_array(1)
        End If

        Return Out_srt
        'ListBox_Item_content.Items(3) & "," & ListBox_Item_content.Items(4) & "," & ListBox_Item_content.Items(5)

    End Function
    Public Sub Open_Xfile(ByVal XX As String)
        'Dim zz_str As String
        'Dim zz_array() As String
        'zz_str = ListBox_Item_content.Items(17)
        'zz_array = zz_str.Split(",")
        'OpenX(zz_array(1))
        OpenX(XX)
    End Sub
    Public Sub Open_Xfile2()
        Dim zz_str As String
        Dim zz_array() As String
        zz_str = ListBox_Item_content.Items(17)
        zz_array = zz_str.Split(",")
        OpenX(zz_array(1))
        'OpenX(XX)
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Btn_Search.Click
        Dim Fpath As String
        Dim tempFpath As String
        Dim Ftrg As String

        Dim Is_ABS_Path As Boolean

        'Dim Disk_id As Char
        Dim Disk_id As String


        Label25.Text = "搜尋中..."

        ListBox_Search_Result.Items.Clear()


        If CKB_All_Path.Checked = False Then


            '路徑
            Fpath = CMB_Search_Path.Text

            'KeyWord  + 副檔名
            If CKB_All_FileType.Checked = True Then
                Ftrg = "*" & TB_KeyWord.Text & "*.*"
            Else
                Ftrg = TB_KeyWord.Text
            End If
            'C:\Joe\work\1_SW_Design


            GetFileList(Fpath, Ftrg, ListBox_Search_Result, True)
        Else
            '----------------------------
            '   檢查所有路徑
            '----------------------------
            For Each DD As String In CMB_Search_Path.Items
                Fpath = DD


                Is_ABS_Path = False

                '------------------------------
                '   絕對路徑 / 相對路徑的處理
                '------------------------------
                Disk_id = Mid(Fpath, 1, 1)

                If Disk_id.ToUpper >= "A" And Disk_id.ToUpper <= "Z" Then

                    If Mid(Fpath, 2, 2) = ":\" Then
                        Is_ABS_Path = True

                    End If

                End If

                '------------------------------
                '   相對路徑的處理
                '------------------------------
                If Is_ABS_Path = False Then
                    tempFpath = Application.StartupPath & "\" & Fpath
                    'Label22.Text = tempFpath

                    Fpath = Application.StartupPath & "\" & Fpath
                End If


                'KeyWord  + 副檔名
                If CKB_All_FileType.Checked = True Then
                    Ftrg = "*" & TB_KeyWord.Text & "*.*"
                Else
                    Ftrg = TB_KeyWord.Text
                End If

                Debug.Print("Search Path:" & Fpath)
                GetFileList(Fpath, Ftrg, ListBox_Search_Result, True)

            Next
        End If

        Label25.Text = "完成"

    End Sub
    Public Sub Do_Search()
        Dim Fpath As String
        Dim Ftrg As String

        Label25.Text = "搜尋中..."
        ListBox_Search_Result.Items.Clear()


        '路徑
        Fpath = CMB_Search_Path.Text

        'KeyWord  + 副檔名
        If CKB_All_FileType.Checked = True Then
            Ftrg = "*" & TB_KeyWord.Text & "*.*"
        Else
            Ftrg = TB_KeyWord.Text
        End If
        'C:\Joe\work\1_SW_Design


        GetFileList(Fpath, Ftrg, ListBox_Search_Result, True)

        Label25.Text = "完成"

    End Sub

    ' 遞迴搜尋子目錄內的符合檔案, 取代 Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories 屬性之 Bug
    ' 需一個 ListBox 物件
    ' 取得目錄搜尋列表
    Public Sub GetFileList(ByRef 路徑 As String, ByRef 檔名 As String, ByRef 接收容器 As ListBox, Optional ByRef 深度搜尋 As Boolean = False)
        Dim dirAttr As Microsoft.VisualBasic.FileAttribute
        Try
            For Each fd In My.Computer.FileSystem.GetFiles(路徑, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, 檔名)
                接收容器.Items.Add(fd.ToString)
            Next

            ' 喘口氣
            My.Application.DoEvents()

            If 深度搜尋 Then
                For Each ff In My.Computer.FileSystem.GetDirectories(路徑)
                    dirAttr = FileSystem.GetAttr(ff.ToString)
                    If (dirAttr And FileAttribute.System) = 0 Then GetFileList(ff.ToString, 檔名, 接收容器, 深度搜尋)
                Next
            End If
        Catch ex As UnauthorizedAccessException
        Catch ex As IOException
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim Full_target_name As String

        Full_target_name = ListBox_Search_Result.SelectedItem

        '===================
        '　開啟Report
        '===================

        If (File.Exists(Full_target_name)) Then
            '        ReportFileName
            If (Full_target_name <> "") Then
                System.Diagnostics.Process.Start(Full_target_name)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else
            MsgBox("找不到該檔案:" & Full_target_name)
        End If
    End Sub
    Public Sub OpenX(ByVal AAA As String)
        ' Dim Full_target_name As String

        ' Full_target_name = ListBox_Search_Result.SelectedItem

        '===================
        '　開啟Report
        '===================

        If (File.Exists(AAA)) Then
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

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TB_WW.TextChanged

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TB_subTitle1.TextChanged

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        TB_subTitle1.Text = ""
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        TB_subTitle2.Text = ""
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TB_subTitle3.Text = ""
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        '====================
        '   產生小圖
        '====================
        Form_Spec.Update_for_Sync_Create_SmallPic()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        TB_CreateDate.Text = Get_Now_Format()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim xxx As String

        '----------------
        '路徑+檔名
        '----------------
        xxx = ListBox_Search_Result.SelectedItem & " -n" & Val(TB_Code_Line.Text)


        'xxx = "D:\PROGRAMS\ElectronicBus_ECU\EV_BusECU\trunk\Sources\06_AC.c -n200"
        'D:\PROGRAMS\ElectronicBus_ECU\EV_BusECU\trunk\Sources\06_AC.c
        'D:\joe\npp_6792_bin
        ' Shell("d:\npp069\notepad++.exe " & xxx, vbNormalFocus)
        'Shell("D:\joe\npp_6792_bin\notepad++.exe " & xxx, vbNormalFocus)
        ' Shell(Form_Spec.Get_Source_DIR & "\APP\npp069\notepad++.exe " & xxx, vbNormalFocus)
        'Shell(Form_Spec.Get_Source_DIR & "\APP\npp069\notepad++.exe " & xxx, vbNormalFocus)

        Shell(Form_Spec.NotePad_path & " " & xxx, vbNormalFocus)

        ' NotePad_path()
    End Sub
    Public Sub Open_by_Npp(ByVal AAA As String)
        Dim xxx As String

        '----------------
        '路徑+檔名
        '----------------
        xxx = Npp_file_name & " -n" & Val(TB_Code_Line.Text)


        'xxx = "D:\PROGRAMS\ElectronicBus_ECU\EV_BusECU\trunk\Sources\06_AC.c -n200"
        'D:\PROGRAMS\ElectronicBus_ECU\EV_BusECU\trunk\Sources\06_AC.c
        'D:\joe\npp_6792_bin
        ' Shell("d:\npp069\notepad++.exe " & xxx, vbNormalFocus)
        'Shell("D:\joe\npp_6792_bin\notepad++.exe " & xxx, vbNormalFocus)
        ' Shell(Form_Spec.Get_Source_DIR & "\APP\npp069\notepad++.exe " & xxx, vbNormalFocus)
        'Shell(Form_Spec.Get_Source_DIR & "\APP\npp069\notepad++.exe " & xxx, vbNormalFocus)
        Debug.Print("NotePad_path:" & Form_Spec.NotePad_path)
        Shell(Form_Spec.NotePad_path & " " & xxx, vbNormalFocus)

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox_Search_Result.SelectedIndexChanged

    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        TB_Modify_Value.Clear()
    End Sub



    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        'Form_Spec.LA_DIR

        Dim temp_path As String
        Dim temp_para As String

        temp_path = Form_Spec.LA_DIR
        temp_para = " "
        'ListBox2.SelectedItem

        Call Shell(temp_path & temp_para & ListBox_Search_Result.SelectedItem)

    End Sub
    Public Sub Open_by_LA(ByVal AAA As String)
        Dim temp_path As String
        Dim temp_para As String
        'Npp_file_name
        temp_path = Form_Spec.LA_DIR
        temp_para = " "
        'ListBox2.SelectedItem

        Call Shell(temp_path & temp_para & AAA)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click



        TB_Display.Text = Get_Disp_Chkbox()


    End Sub
    Private Function Get_Disp_Chkbox()
        Dim xx As Integer
        xx = 0

        If CKB_Title1.Checked = True Then
            xx += 128
        End If


        If CheckBox4.Checked = True Then
            xx += 64
        End If


        If CheckBox5.Checked = True Then
            xx += 32
        End If


        If CheckBox6.Checked = True Then
            xx += 16
        End If


        If CheckBox7.Checked = True Then
            xx += 8
        End If


        If CheckBox8.Checked = True Then
            xx += 4
        End If


        If CheckBox9.Checked = True Then
            xx += 2
        End If


        If CheckBox10.Checked = True Then
            xx += 1
        End If

        Return xx

    End Function
    Public Sub Load_Compare_Offset_File()
        Dim filename1 As String
        filename1 = Form_Spec.Get_filePath_x_Append & "\Index\Compare_Offset.txt"


        If (File.Exists(filename1)) Then
            ' MsgBox("找到該路徑:" & filename1)
            'Get_Spec_path(filename1)
            Compare_Page_List.Items.Clear()



            Dim Parameter_Line As String = ""


            Dim fs As FileStream = New FileStream(filename1, FileMode.Open)
            Using sr As StreamReader = New StreamReader(fs, _
                                                        System.Text.Encoding.Default)

                While (sr.EndOfStream <> True)

                    '送出相關設定
                    Parameter_Line = sr.ReadLine()


                    If Compare_Page_List.Items.IndexOf(Val(Parameter_Line)) = -1 Then
                        '----------------------
                        ' 取不重複的值
                        '----------------------
                        Compare_Page_List.Items.Add(Val(Parameter_Line))
                    End If


                End While


                sr.Close()                                                          '關閉檔案




            End Using

        Else
            'MsgBox("No_File:" & filename1)

            'Create_New_file()
        End If

        'For Each Xval As Integer In Compare_Page_List.Items

        '    Debug.Print("Listbox_Add:" & Xval)
        'Next

    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Btn_Modify_BlockDisplay.Click
        '---------------------------
        '   調整方面的顯示項目
        '---------------------------
        CMB_ItemType.SelectedIndex = 18     '調整顯示項目
        Modify_BlockItem()




    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click

    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click

        Open_Mantis()

    End Sub
    Private Sub Open_Mantis()
        Dim xxx As String
        Dim yyy As String

        If RadioButton1.Checked = True Then
            xxx = Form_Spec.Mantis_Server
        Else
            xxx = Form_Spec.Mantis_Server2
        End If

        'xxx = Form_Spec.Mantis_Server
        yyy = TextBox21.Text
        Process.Start("iexplore.exe", xxx & yyy)
    End Sub
    Private Sub Open_Redmine()
        Dim xxx As String
        Dim yyy As String
        If RadioButton1.Checked = True Then
            xxx = Form_Spec.Redmine_Server
        Else
            xxx = Form_Spec.Redmine_Server2
        End If

        yyy = TextBox17.Text

        Process.Start("iexplore.exe", xxx & yyy)
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        TB_Note.Clear()
    End Sub

    Private Sub Button7_Click_2(sender As Object, e As EventArgs) Handles Button7.Click
        Jump_to_Page(Val(TextBox1.Text))
    End Sub
    Public Sub Jump_to_Page(ByVal tPage)
        Form_Spec.Change_Page(tPage)
    End Sub



    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        TB_ModifyDate.Text = Get_Now_Format()
    End Sub
    Public Function Get_Now_Format()
        Return Now.ToString("yyyy/MM/dd_HH:mm:ss")
    End Function

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        TB_Title1.Text = ""
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        TB_Title2.Text = ""
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        TB_Title3.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim log_filename As String
        If Form_Spec.CKB_compare.Checked = False Then
            log_filename = Form_Spec.Get_filePath_x_Append & "\LOG\log.txt"
        Else
            log_filename = Form_Spec.Get_filePath_x_Append & "\LOG\log_Compare.txt"
        End If

        If (File.Exists(log_filename)) = True Then
            '        ReportFileName
            ' If (filename_html <> "") Then
            System.Diagnostics.Process.Start(log_filename)
            'Else
        Else

            MsgBox("No File Name:" & log_filename)
            ' End If

        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TB_ColorB.Text = ""
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        TB_ColorG.Text = ""
    End Sub

    Private Sub TB_LastPage_DoubleClick(sender As Object, e As EventArgs) Handles TB_LastPage.DoubleClick
        Jump_to_Page(Val(TB_LastPage.Text))
    End Sub

    Private Sub TB_LastPage_TextChanged(sender As Object, e As EventArgs) Handles TB_LastPage.TextChanged

    End Sub

    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click
        Jump_to_Page(Val(TB_LastPage.Text))
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        Dim filename1 As String
        Dim content As String
        Dim filenum As Integer

        filename1 = Form_Spec.Get_filePath_x_Append & "\Index\Compare_Offset.txt"
        filenum = FreeFile()
        FileOpen(filenum, filename1, OpenMode.Append)


        content = Form_Spec.Get_Page_Now

        PrintLine(filenum, content)


        FileClose(filenum)

        Load_Compare_Offset_File()

    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        ' Form_Spec.FullName_item_CSV
        If (File.Exists(Form_Spec.FullName_item_CSV)) = True Then
            '        ReportFileName
            If (Form_Spec.FullName_item_CSV <> "") Then
                System.Diagnostics.Process.Start(Form_Spec.FullName_item_CSV)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else

            MsgBox("找不到該檔案:" & Form_Spec.FullName_item_CSV)
        End If

    End Sub

    Private Sub CMB_SaveType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_SaveType.SelectedIndexChanged
        '-------------------
        '   工作區類別改變
        '-------------------

        Form_Spec.SetX_Str = CMB_SaveType.Text

        If Setup_first = False Then
            Setup_first = True
        Else

            Change_WorkArea_SetX()

        End If


    End Sub
    Public Sub Change_WorkArea_SetX()
        Dim location_XY As Point
        Dim Form_Size_W As Integer
        Dim Form_Size_H As Integer

        'Dim xxx As String
        Dim small_open As Boolean = False

        '---------------------
        '   如果縮圖區有開啟時， 暫時先關閉
        '---------------------
        If Form_Spec_All_Pc.Created = True Then
            Form_Size_W = Form_Spec_All_Pc.Width
            Form_Size_H = Form_Spec_All_Pc.Height

            location_XY = Form_Spec_All_Pc.Location
            Form_Spec_All_Pc.Close()
            small_open = True
        End If

        Form_Spec.Pattern_Type = CMB_SaveType.Text


        'Debug.Print("From Change_WorkArea_SetX")
        ReNew_Working_SetX_Path(True)





        'xxx = Form_Spec.Get_CSV_Filename()



        '---------------------
        '   如果縮圖區有本來是開啟的， 再打開
        '---------------------
        If small_open = True Then

            Form_Spec_All_Pc.Show()
            small_open = False
            Form_Spec_All_Pc.Location = location_XY
            Form_Spec_All_Pc.Width = Form_Size_W
            Form_Spec_All_Pc.Height = Form_Size_H
        End If



        Form_Spec_CHK_Top.Graph_Clear()
        Debug.Print("NFS_11")
        Form_Spec.Update_for_Sync()

        '------------------
        '   回到第一頁
        '------------------
        Form_Spec.TrackBar1.Value = 1
        Form_Spec.Change_Page(Form_Spec.TrackBar1.Value)



        Form_Spec_CHK_Top.Focus()

        '------------------
        '   讀取Title
        '------------------
        Form_Note.Read_FormTitle()
        'Form_Spec_CHK_Top.Timer1.Enabled = True    '20210112

        '--------------------------------
        '   讀取 補差頁的Page_Offset設定
        '--------------------------------
        Load_Compare_Offset_File()
        'Check_Pattern_and_Save()
    End Sub
    '========================================
    '   tCreat_DIR=true => 建立資料夾
    '========================================
    Public Sub ReNew_Working_SetX_Path(ByVal tCreat_DIR As Boolean)

        'Debug.Print("ReNew_Working_SetX_Path:" & Now)


        '-------------------
        '   依設定的工作檔組別
        '-------------------

        ' If CMB_SaveType.Text = "" Then
        '-------------------
        '   以後改為Set_0
        '-------------------
        'Form_Spec.Get_filePath_x_Append = Form_Spec.Get_file_Path_x
        ' Else
        '-------------------
        '   以後改為Set_1 ~ Set_xx
        '-------------------
        If CMB_SaveType.Text <> "" Then
            Form_Spec.Set_filePath_x_Append(Form_Spec.Get_file_Path_x & "\" & CMB_SaveType.Text)



            Debug.Print(Form_Spec.Get_filePath_x_Append & "\LOG")

            CHK_DIR_if_None_then_Create(Form_Spec.Get_filePath_x_Append & "\LOG")


            Form_Spec.FullName_Index_Path = Form_Spec.Get_filePath_x_Append & "\Index"
            Form_Spec.FullName_item_Path = Form_Spec.Get_filePath_x_Append & "\DATA_CSV"

            Form_Spec.Set_Full_PathSmall_Name(Form_Spec.Get_WorkPath_x_Append & "\sp")

            Form_Spec.Set_CHK_PicPath(Form_Spec.Get_WorkPath_x_Append & "\Normal")

            Form_Spec.FullName_ALL = Form_Spec.Get_filePath_x_Append & "\ALL"

            'temp_index_dir = Form_Spec.Get_filePath_x_Append & "\Index"

            If tCreat_DIR = True Then
                CHK_DIR_if_None_then_Create(Form_Spec.FullName_item_Path)
                CHK_DIR_if_None_then_Create(Form_Spec.Get_Full_PathSmall_Name)
                CHK_DIR_if_None_then_Create(Form_Spec.Get_CHK_PicPath)
                CHK_DIR_if_None_then_Create(Form_Spec.FullName_ALL)
                CHK_DIR_if_None_then_Create(Form_Spec.FullName_Index_Path)
            End If


            If Form_Spec.CKB_compare.Checked = True Then

                Form_Spec.FullName_Index_Path &= "\compare"
                Form_Spec.FullName_item_Path &= "\compare"


                Form_Spec.FullName_ALL &= "\compare"
            End If


            If tCreat_DIR = True Then
                Debug.Print("Idx_path:" & Form_Spec.FullName_Index_Path)
                CHK_DIR_if_None_then_Create(Form_Spec.FullName_Index_Path)
                CHK_DIR_if_None_then_Create(Form_Spec.FullName_item_Path)
                CHK_DIR_if_None_then_Create(Form_Spec.Get_Full_PathSmall_Name)
                CHK_DIR_if_None_then_Create(Form_Spec.Get_CHK_PicPath)
                CHK_DIR_if_None_then_Create(Form_Spec.FullName_ALL)

            End If
        End If
    End Sub
    Public Sub CHK_DIR_if_None_then_Create(ByVal AAA As String)
        If Not IO.Directory.Exists(AAA) Then

            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(AAA)
        End If
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        'If RadioButton2.Checked = True Then
        'OpenFileDialog1.InitialDirectory = TB_User_Path.Text
        OpenFileDialog1.InitialDirectory = Form_Spec.Get_Source_DIR
        'End If


        'OpenFileDialog1.DefaultExt = "csv"
        'OpenFileDialog1.Filter = Enabled
        OpenFileDialog1.ShowDialog()
        ListBox_Search_Result.Items.Clear()
        ListBox_Search_Result.Items.Add(OpenFileDialog1.FileName)
        'LB_FullPath.Text = TB_NewScriptName.Text
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Btn_to_Modify.Click
        'TB_Modify_Value.Text = ListBox_Search_Result.Items(0)
        TB_Modify_Value.Text = ListBox_Search_Result.SelectedItem
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        Copy_to_word_Area()

    End Sub
    Public Sub Copy_to_word_Area()
        TB_Title1.Text = Get_List_Item_Text(1)
        TB_subTitle1.Text = Get_List_Item_Text(2)
        TB_Title2.Text = Get_List_Item_Text(4)
        TB_subTitle2.Text = Get_List_Item_Text(5)
        TB_Title3.Text = Get_List_Item_Text(7)
        TB_subTitle3.Text = Get_List_Item_Text(8)

        CMB_Font_Color.SelectedIndex = Val(Get_List_Item_Text(10))
        CMB_BG_Color.SelectedIndex = Val(Get_List_Item_Text(11))

        TB_ColorG.Text = Get_List_Item_Text(13)
        TB_ColorB.Text = Get_List_Item_Text(14)

        TB_Result_Info.Text = Get_List_Item_Text(16)
        TB_Note.Text = Get_List_Item_Text(17)

        TB_Display.Text = Get_Disp_Chkbox()

        TB_XX.Text = TBx.Text
        TB_YY.Text = TBy.Text
        TB_WW.Text = TBw.Text
        TB_HH.Text = TBh.Text
    End Sub
    Public Function Get_List_Item_Text(ByVal AAA As Integer) As String
        Dim ori_str As String
        Dim ori_str_array() As String

        Try
            ori_str = ListBox_Item_content.Items(AAA)

            If ori_str <> "" Then
                ori_str_array = ori_str.Split(",")
                Return ori_str_array(1)
            Else
                Return ""
            End If

        Catch ex As Exception
            Return ""
        Finally

        End Try

    End Function

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        TB_Title1.Text = ""
        TB_subTitle1.Text = ""
        TB_Title2.Text = ""
        TB_subTitle2.Text = ""
        TB_Title3.Text = ""
        TB_subTitle3.Text = ""

        CMB_Font_Color.SelectedIndex = 0
        CMB_BG_Color.SelectedIndex = 0

        TB_ColorG.Text = ""
        TB_ColorB.Text = ""

        TB_Result_Info.Text = ""
        TB_Note.Text = ""

        TB_Display.Text = 0
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        Form_SetX_List.Show()
        'Form_Note.Show()
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Clipboard.SetText(TB_KeyWord.Text)
    End Sub

    'Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
    '    Form_SetX_List.Show()

    'End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles BTN_Get_Jump_Page.Click
        'TB_Title1.Clear()

        'CMB_PreSet_Item.SelectedIndex = 12                'jump
        CMB_PreSet_Item.SelectedIndex = 15               'jump
        CMB_Font_Color.SelectedIndex = 13           '13:鮭魚

        TB_Title1.Text = "JUMP_" & TB_LastPage.Text
    End Sub

    Private Sub CMB_Font_Color_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Font_Color.SelectedIndexChanged

    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        Dim Full_target_name As String

        'default_Color_file = Form_Spec.Get_Source_DIR & "\APP\Default_item.txt"
        Full_target_name = Form_Spec.Get_Source_DIR & "\APP\Default_item.txt"

        '===================
        '　開啟Report
        '===================

        If (File.Exists(Full_target_name)) Then
            '        ReportFileName
            If (Full_target_name <> "") Then
                System.Diagnostics.Process.Start(Full_target_name)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else
            MsgBox("找不到該檔案:" & Full_target_name)
        End If
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        ColorTool.Load_Src_file()
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        Dim Full_target_name As String

        'default_Color_file = Form_Spec.Get_Source_DIR & "\APP\Default_item.txt"
        Full_target_name = Form_Spec.Get_Source_DIR & "\APP\Search_Path.txt"

        '===================
        '　開啟Report
        '===================

        If (File.Exists(Full_target_name)) Then
            '        ReportFileName
            If (Full_target_name <> "") Then
                System.Diagnostics.Process.Start(Full_target_name)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else
            MsgBox("找不到該檔案:" & Full_target_name)
        End If

    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        Load_SearchPath()
    End Sub

    Private Sub Button24_Click_1(sender As Object, e As EventArgs) Handles Button24.Click
        CMB_ItemType.SelectedIndex = 17
        TB_Modify_Value.Text = ListBox_Search_Result.SelectedItem
        Modify_BlockItem()

    End Sub
End Class