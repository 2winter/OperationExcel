<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ex = New System.Windows.Forms.GroupBox()
        Me.ShowResult = New System.Windows.Forms.TextBox()
        Me.Save_Page = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.OpenSheetResult1 = New System.Windows.Forms.TextBox()
        Me.OpenSheet1 = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Sheet_Index = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SheetNum = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.VersionText = New System.Windows.Forms.TextBox()
        Me.PageText = New System.Windows.Forms.TextBox()
        Me.ChangeOneByOne = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.LoadDstPath = New System.Windows.Forms.Button()
        Me.Save_DstPath = New System.Windows.Forms.TextBox()
        Me.ex.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(32, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(89, 23)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "加载总表路径"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ex
        '
        Me.ex.Controls.Add(Me.ShowResult)
        Me.ex.Controls.Add(Me.Save_Page)
        Me.ex.Controls.Add(Me.Button3)
        Me.ex.Controls.Add(Me.OpenSheetResult1)
        Me.ex.Controls.Add(Me.OpenSheet1)
        Me.ex.Controls.Add(Me.Label11)
        Me.ex.Controls.Add(Me.Sheet_Index)
        Me.ex.Controls.Add(Me.Label5)
        Me.ex.Controls.Add(Me.SheetNum)
        Me.ex.Controls.Add(Me.Label1)
        Me.ex.Controls.Add(Me.Label2)
        Me.ex.Controls.Add(Me.TextBox4)
        Me.ex.Controls.Add(Me.TextBox5)
        Me.ex.Location = New System.Drawing.Point(32, 86)
        Me.ex.Name = "ex"
        Me.ex.Size = New System.Drawing.Size(326, 297)
        Me.ex.TabIndex = 22
        Me.ex.TabStop = False
        Me.ex.Text = "总表"
        '
        'ShowResult
        '
        Me.ShowResult.Location = New System.Drawing.Point(141, 206)
        Me.ShowResult.Name = "ShowResult"
        Me.ShowResult.ReadOnly = True
        Me.ShowResult.Size = New System.Drawing.Size(100, 21)
        Me.ShowResult.TabIndex = 30
        '
        'Save_Page
        '
        Me.Save_Page.Location = New System.Drawing.Point(16, 206)
        Me.Save_Page.Name = "Save_Page"
        Me.Save_Page.Size = New System.Drawing.Size(94, 23)
        Me.Save_Page.TabIndex = 28
        Me.Save_Page.Text = "页码版本缓存"
        Me.Save_Page.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(244, 268)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "CloseExcel"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'OpenSheetResult1
        '
        Me.OpenSheetResult1.Location = New System.Drawing.Point(272, 87)
        Me.OpenSheetResult1.Name = "OpenSheetResult1"
        Me.OpenSheetResult1.ReadOnly = True
        Me.OpenSheetResult1.Size = New System.Drawing.Size(47, 21)
        Me.OpenSheetResult1.TabIndex = 27
        '
        'OpenSheet1
        '
        Me.OpenSheet1.Location = New System.Drawing.Point(184, 87)
        Me.OpenSheet1.Name = "OpenSheet1"
        Me.OpenSheet1.Size = New System.Drawing.Size(75, 23)
        Me.OpenSheet1.TabIndex = 26
        Me.OpenSheet1.Text = "打开Sheet"
        Me.OpenSheet1.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(12, 90)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(77, 12)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "SheetIndex："
        '
        'Sheet_Index
        '
        Me.Sheet_Index.Location = New System.Drawing.Point(95, 87)
        Me.Sheet_Index.Name = "Sheet_Index"
        Me.Sheet_Index.Size = New System.Drawing.Size(62, 21)
        Me.Sheet_Index.TabIndex = 24
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(39, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(71, 12)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Sheet表数："
        '
        'SheetNum
        '
        Me.SheetNum.Location = New System.Drawing.Point(141, 38)
        Me.SheetNum.Name = "SheetNum"
        Me.SheetNum.ReadOnly = True
        Me.SheetNum.Size = New System.Drawing.Size(67, 21)
        Me.SheetNum.TabIndex = 11
        Me.SheetNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 145)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 12)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "行数："
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(153, 145)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 12)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "列数："
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(61, 142)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.ReadOnly = True
        Me.TextBox4.Size = New System.Drawing.Size(59, 21)
        Me.TextBox4.TabIndex = 9
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(200, 142)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.ReadOnly = True
        Me.TextBox5.Size = New System.Drawing.Size(59, 21)
        Me.TextBox5.TabIndex = 10
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.VersionText)
        Me.GroupBox2.Controls.Add(Me.PageText)
        Me.GroupBox2.Controls.Add(Me.ChangeOneByOne)
        Me.GroupBox2.Controls.Add(Me.ListBox1)
        Me.GroupBox2.Controls.Add(Me.Button4)
        Me.GroupBox2.Location = New System.Drawing.Point(372, 86)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(308, 297)
        Me.GroupBox2.TabIndex = 23
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "目标Excel"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(135, 262)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 12)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "版本:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(135, 237)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 12)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "页码:"
        '
        'VersionText
        '
        Me.VersionText.Location = New System.Drawing.Point(170, 259)
        Me.VersionText.Name = "VersionText"
        Me.VersionText.ReadOnly = True
        Me.VersionText.Size = New System.Drawing.Size(100, 21)
        Me.VersionText.TabIndex = 19
        '
        'PageText
        '
        Me.PageText.Location = New System.Drawing.Point(170, 232)
        Me.PageText.Name = "PageText"
        Me.PageText.ReadOnly = True
        Me.PageText.Size = New System.Drawing.Size(100, 21)
        Me.PageText.TabIndex = 18
        '
        'ChangeOneByOne
        '
        Me.ChangeOneByOne.Location = New System.Drawing.Point(28, 232)
        Me.ChangeOneByOne.Name = "ChangeOneByOne"
        Me.ChangeOneByOne.Size = New System.Drawing.Size(75, 23)
        Me.ChangeOneByOne.TabIndex = 17
        Me.ChangeOneByOne.Text = "挨个替换"
        Me.ChangeOneByOne.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 12
        Me.ListBox1.Location = New System.Drawing.Point(28, 54)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(242, 172)
        Me.ListBox1.TabIndex = 16
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(28, 25)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(85, 23)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "替换版本"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(127, 24)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(358, 21)
        Me.TextBox1.TabIndex = 20
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'LoadDstPath
        '
        Me.LoadDstPath.Location = New System.Drawing.Point(32, 57)
        Me.LoadDstPath.Name = "LoadDstPath"
        Me.LoadDstPath.Size = New System.Drawing.Size(89, 23)
        Me.LoadDstPath.TabIndex = 24
        Me.LoadDstPath.Text = "保存目标路径"
        Me.LoadDstPath.UseVisualStyleBackColor = True
        '
        'Save_DstPath
        '
        Me.Save_DstPath.Location = New System.Drawing.Point(127, 57)
        Me.Save_DstPath.Name = "Save_DstPath"
        Me.Save_DstPath.Size = New System.Drawing.Size(358, 21)
        Me.Save_DstPath.TabIndex = 21
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(718, 413)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ex)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.LoadDstPath)
        Me.Controls.Add(Me.Save_DstPath)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ex.ResumeLayout(False)
        Me.ex.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ex As System.Windows.Forms.GroupBox
    Friend WithEvents Save_Page As System.Windows.Forms.Button
    Friend WithEvents OpenSheetResult1 As System.Windows.Forms.TextBox
    Friend WithEvents OpenSheet1 As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Sheet_Index As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents SheetNum As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents LoadDstPath As System.Windows.Forms.Button
    Friend WithEvents Save_DstPath As System.Windows.Forms.TextBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents ShowResult As System.Windows.Forms.TextBox
    Friend WithEvents ChangeOneByOne As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents VersionText As System.Windows.Forms.TextBox
    Friend WithEvents PageText As System.Windows.Forms.TextBox

End Class
