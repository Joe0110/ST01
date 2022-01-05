<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Spec_CHK_Top
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_Spec_CHK_Top))
        Me.PicBox_TopMain = New System.Windows.Forms.PictureBox()
        Me.BTN_Overlap = New System.Windows.Forms.Button()
        Me.PicBox_Arrow = New System.Windows.Forms.PictureBox()
        Me.PicBox_Drag = New System.Windows.Forms.PictureBox()
        Me.PicBox_TopSub = New System.Windows.Forms.PictureBox()
        Me.Timer_OnTOP = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_Resize = New System.Windows.Forms.Timer(Me.components)
        CType(Me.PicBox_TopMain, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBox_Arrow, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBox_Drag, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBox_TopSub, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBox_TopMain
        '
        Me.PicBox_TopMain.Location = New System.Drawing.Point(29, 21)
        Me.PicBox_TopMain.Name = "PicBox_TopMain"
        Me.PicBox_TopMain.Size = New System.Drawing.Size(720, 960)
        Me.PicBox_TopMain.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicBox_TopMain.TabIndex = 62
        Me.PicBox_TopMain.TabStop = False
        '
        'BTN_Overlap
        '
        Me.BTN_Overlap.Font = New System.Drawing.Font("PMingLiU", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.BTN_Overlap.Location = New System.Drawing.Point(787, 61)
        Me.BTN_Overlap.Name = "BTN_Overlap"
        Me.BTN_Overlap.Size = New System.Drawing.Size(67, 23)
        Me.BTN_Overlap.TabIndex = 63
        Me.BTN_Overlap.Text = "重疊"
        Me.BTN_Overlap.UseVisualStyleBackColor = True
        '
        'PicBox_Arrow
        '
        Me.PicBox_Arrow.Image = CType(resources.GetObject("PicBox_Arrow.Image"), System.Drawing.Image)
        Me.PicBox_Arrow.Location = New System.Drawing.Point(787, 118)
        Me.PicBox_Arrow.Name = "PicBox_Arrow"
        Me.PicBox_Arrow.Size = New System.Drawing.Size(14, 17)
        Me.PicBox_Arrow.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicBox_Arrow.TabIndex = 107
        Me.PicBox_Arrow.TabStop = False
        Me.PicBox_Arrow.Visible = False
        '
        'PicBox_Drag
        '
        Me.PicBox_Drag.BackColor = System.Drawing.Color.Red
        Me.PicBox_Drag.Location = New System.Drawing.Point(813, 216)
        Me.PicBox_Drag.Name = "PicBox_Drag"
        Me.PicBox_Drag.Size = New System.Drawing.Size(31, 27)
        Me.PicBox_Drag.TabIndex = 108
        Me.PicBox_Drag.TabStop = False
        Me.PicBox_Drag.Visible = False
        '
        'PicBox_TopSub
        '
        Me.PicBox_TopSub.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.PicBox_TopSub.Location = New System.Drawing.Point(765, 21)
        Me.PicBox_TopSub.Name = "PicBox_TopSub"
        Me.PicBox_TopSub.Size = New System.Drawing.Size(51, 34)
        Me.PicBox_TopSub.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicBox_TopSub.TabIndex = 109
        Me.PicBox_TopSub.TabStop = False
        '
        'Timer_OnTOP
        '
        Me.Timer_OnTOP.Interval = 500
        '
        'Timer_Resize
        '
        Me.Timer_Resize.Interval = 300
        '
        'Form_Spec_CHK_Top
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(903, 1012)
        Me.Controls.Add(Me.PicBox_TopSub)
        Me.Controls.Add(Me.PicBox_Drag)
        Me.Controls.Add(Me.PicBox_Arrow)
        Me.Controls.Add(Me.BTN_Overlap)
        Me.Controls.Add(Me.PicBox_TopMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "Form_Spec_CHK_Top"
        Me.Text = "Form_Spec_CHK_Top"
        CType(Me.PicBox_TopMain, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBox_Arrow, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBox_Drag, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBox_TopSub, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PicBox_TopMain As System.Windows.Forms.PictureBox
    Friend WithEvents BTN_Overlap As System.Windows.Forms.Button
    Friend WithEvents PicBox_Arrow As System.Windows.Forms.PictureBox
    Friend WithEvents PicBox_Drag As System.Windows.Forms.PictureBox
    Friend WithEvents PicBox_TopSub As System.Windows.Forms.PictureBox
    Friend WithEvents Timer_OnTOP As System.Windows.Forms.Timer
    Friend WithEvents Timer_Resize As System.Windows.Forms.Timer
End Class
