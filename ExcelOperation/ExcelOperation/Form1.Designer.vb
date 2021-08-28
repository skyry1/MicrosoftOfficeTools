<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DirPath_TextBox = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ActiveCell_TextBox = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SheetZoom_TextBox = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Filter_CheckBox = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(25, 423)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(130, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "実行"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DirPath_TextBox
        '
        Me.DirPath_TextBox.Location = New System.Drawing.Point(25, 59)
        Me.DirPath_TextBox.Name = "DirPath_TextBox"
        Me.DirPath_TextBox.Size = New System.Drawing.Size(692, 35)
        Me.DirPath_TextBox.TabIndex = 1
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(587, 100)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(130, 40)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "参照"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(143, 30)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "実行ディレクトリ"
        '
        'ActiveCell_TextBox
        '
        Me.ActiveCell_TextBox.Location = New System.Drawing.Point(25, 224)
        Me.ActiveCell_TextBox.Name = "ActiveCell_TextBox"
        Me.ActiveCell_TextBox.Size = New System.Drawing.Size(175, 35)
        Me.ActiveCell_TextBox.TabIndex = 4
        Me.ActiveCell_TextBox.Text = "A1"
        Me.ActiveCell_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(25, 191)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 30)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "アクティブセル"
        '
        'SheetZoom_TextBox
        '
        Me.SheetZoom_TextBox.Location = New System.Drawing.Point(25, 312)
        Me.SheetZoom_TextBox.Name = "SheetZoom_TextBox"
        Me.SheetZoom_TextBox.Size = New System.Drawing.Size(175, 35)
        Me.SheetZoom_TextBox.TabIndex = 6
        Me.SheetZoom_TextBox.Text = "100"
        Me.SheetZoom_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(25, 279)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(97, 30)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "表示倍率"
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(12, 684)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(705, 40)
        Me.ProgressBar.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 651)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(108, 30)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "プロセスバー"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(25, 383)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(0, 30)
        Me.Label5.TabIndex = 10
        '
        'Filter_CheckBox
        '
        Me.Filter_CheckBox.AutoSize = True
        Me.Filter_CheckBox.Location = New System.Drawing.Point(25, 383)
        Me.Filter_CheckBox.Name = "Filter_CheckBox"
        Me.Filter_CheckBox.Size = New System.Drawing.Size(154, 34)
        Me.Filter_CheckBox.TabIndex = 11
        Me.Filter_CheckBox.Text = "フィルター解除"
        Me.Filter_CheckBox.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 30.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(731, 736)
        Me.Controls.Add(Me.Filter_CheckBox)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ProgressBar)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.SheetZoom_TextBox)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ActiveCell_TextBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.DirPath_TextBox)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Form1"
        Me.Text = "Excel操作"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents DirPath_TextBox As TextBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents ActiveCell_TextBox As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents SheetZoom_TextBox As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents ProgressBar As ProgressBar
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Filter_CheckBox As CheckBox
End Class
