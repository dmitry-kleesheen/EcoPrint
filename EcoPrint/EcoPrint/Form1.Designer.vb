<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ButtonClose = New System.Windows.Forms.Button()
        Me.ButtonPrint = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.LinkLabelFile = New System.Windows.Forms.LinkLabel()
        Me.TextBoxFile = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.ButtonUP = New System.Windows.Forms.Button()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ButtonDOWN = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ButtonOdnodelete = New System.Windows.Forms.Button()
        Me.ButtonOdnoadd = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.ListBoxLog = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CmbPrinter = New System.Windows.Forms.ComboBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ButtonHelp = New System.Windows.Forms.Button()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonClose
        '
        Me.ButtonClose.Location = New System.Drawing.Point(457, 651)
        Me.ButtonClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(100, 28)
        Me.ButtonClose.TabIndex = 0
        Me.ButtonClose.Tag = "f"
        Me.ButtonClose.Text = "Close"
        Me.ButtonClose.UseVisualStyleBackColor = True
        '
        'ButtonPrint
        '
        Me.ButtonPrint.Location = New System.Drawing.Point(349, 651)
        Me.ButtonPrint.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonPrint.Name = "ButtonPrint"
        Me.ButtonPrint.Size = New System.Drawing.Size(100, 28)
        Me.ButtonPrint.TabIndex = 1
        Me.ButtonPrint.Tag = "f"
        Me.ButtonPrint.Text = "Print"
        Me.ButtonPrint.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LinkLabelFile)
        Me.GroupBox1.Controls.Add(Me.TextBoxFile)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(21, 42)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(561, 106)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Tag = "f"
        Me.GroupBox1.Text = "Document Word"
        '
        'LinkLabelFile
        '
        Me.LinkLabelFile.AutoSize = True
        Me.LinkLabelFile.Location = New System.Drawing.Point(485, 48)
        Me.LinkLabelFile.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LinkLabelFile.Name = "LinkLabelFile"
        Me.LinkLabelFile.Size = New System.Drawing.Size(54, 17)
        Me.LinkLabelFile.TabIndex = 2
        Me.LinkLabelFile.TabStop = True
        Me.LinkLabelFile.Tag = "f"
        Me.LinkLabelFile.Text = "Browse"
        Me.ToolTip1.SetToolTip(Me.LinkLabelFile, "Select a Word file using the ""Browse"" link")
        '
        'TextBoxFile
        '
        Me.TextBoxFile.AllowDrop = True
        Me.TextBoxFile.Location = New System.Drawing.Point(57, 44)
        Me.TextBoxFile.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBoxFile.Name = "TextBoxFile"
        Me.TextBoxFile.ReadOnly = True
        Me.TextBoxFile.Size = New System.Drawing.Size(419, 22)
        Me.TextBoxFile.TabIndex = 1
        Me.TextBoxFile.Tag = "f"
        Me.ToolTip1.SetToolTip(Me.TextBoxFile, "Выберите файл Word с помощью комманды ""Обзор"" - либо перетащите файл в это поле.." &
        ".")
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 48)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "File:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ButtonUP)
        Me.GroupBox2.Controls.Add(Me.ButtonDOWN)
        Me.GroupBox2.Controls.Add(Me.DataGridView1)
        Me.GroupBox2.Controls.Add(Me.ButtonOdnodelete)
        Me.GroupBox2.Controls.Add(Me.ButtonOdnoadd)
        Me.GroupBox2.Location = New System.Drawing.Point(17, 210)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(561, 302)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Tag = "f"
        Me.GroupBox2.Text = "Print Jobs"
        '
        'ButtonUP
        '
        Me.ButtonUP.ImageIndex = 1
        Me.ButtonUP.ImageList = Me.ImageList1
        Me.ButtonUP.Location = New System.Drawing.Point(441, 160)
        Me.ButtonUP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonUP.Name = "ButtonUP"
        Me.ButtonUP.Size = New System.Drawing.Size(100, 34)
        Me.ButtonUP.TabIndex = 4
        Me.ButtonUP.Text = "Up"
        Me.ButtonUP.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonUP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.ButtonUP.UseVisualStyleBackColor = True
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "DownloadLog_16x.png")
        Me.ImageList1.Images.SetKeyName(1, "UploadFile_16x.png")
        '
        'ButtonDOWN
        '
        Me.ButtonDOWN.ImageIndex = 0
        Me.ButtonDOWN.ImageList = Me.ImageList1
        Me.ButtonDOWN.Location = New System.Drawing.Point(441, 202)
        Me.ButtonDOWN.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonDOWN.Name = "ButtonDOWN"
        Me.ButtonDOWN.Size = New System.Drawing.Size(100, 34)
        Me.ButtonDOWN.TabIndex = 4
        Me.ButtonDOWN.Text = "Down"
        Me.ButtonDOWN.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonDOWN.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.ButtonDOWN.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowDrop = True
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2})
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView1.Location = New System.Drawing.Point(16, 23)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(413, 271)
        Me.DataGridView1.TabIndex = 3
        Me.DataGridView1.Tag = "f"
        '
        'ButtonOdnodelete
        '
        Me.ButtonOdnodelete.Location = New System.Drawing.Point(440, 65)
        Me.ButtonOdnodelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonOdnodelete.Name = "ButtonOdnodelete"
        Me.ButtonOdnodelete.Size = New System.Drawing.Size(100, 34)
        Me.ButtonOdnodelete.TabIndex = 1
        Me.ButtonOdnodelete.Tag = "f"
        Me.ButtonOdnodelete.Text = "Delete"
        Me.ButtonOdnodelete.UseVisualStyleBackColor = True
        '
        'ButtonOdnoadd
        '
        Me.ButtonOdnoadd.Location = New System.Drawing.Point(441, 23)
        Me.ButtonOdnoadd.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonOdnoadd.Name = "ButtonOdnoadd"
        Me.ButtonOdnoadd.Size = New System.Drawing.Size(100, 34)
        Me.ButtonOdnoadd.TabIndex = 1
        Me.ButtonOdnoadd.Tag = "f"
        Me.ButtonOdnoadd.Text = "Add"
        Me.ButtonOdnoadd.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.ListBoxLog)
        Me.GroupBox4.Location = New System.Drawing.Point(17, 519)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox4.Size = New System.Drawing.Size(561, 113)
        Me.GroupBox4.TabIndex = 4
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Log"
        '
        'ListBoxLog
        '
        Me.ListBoxLog.FormattingEnabled = True
        Me.ListBoxLog.ItemHeight = 16
        Me.ListBoxLog.Location = New System.Drawing.Point(9, 21)
        Me.ListBoxLog.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ListBoxLog.Name = "ListBoxLog"
        Me.ListBoxLog.Size = New System.Drawing.Size(543, 84)
        Me.ListBoxLog.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(23, 170)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(108, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Select a printer:"
        '
        'CmbPrinter
        '
        Me.CmbPrinter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbPrinter.FormattingEnabled = True
        Me.CmbPrinter.Location = New System.Drawing.Point(132, 166)
        Me.CmbPrinter.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CmbPrinter.Name = "CmbPrinter"
        Me.CmbPrinter.Size = New System.Drawing.Size(444, 24)
        Me.CmbPrinter.TabIndex = 6
        Me.CmbPrinter.Tag = "f"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ButtonHelp
        '
        Me.ButtonHelp.Location = New System.Drawing.Point(21, 651)
        Me.ButtonHelp.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonHelp.Name = "ButtonHelp"
        Me.ButtonHelp.Size = New System.Drawing.Size(100, 28)
        Me.ButtonHelp.TabIndex = 7
        Me.ButtonHelp.Tag = "f"
        Me.ButtonHelp.Text = "About ..."
        Me.ButtonHelp.UseVisualStyleBackColor = True
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'Column1
        '
        Me.Column1.HeaderText = "Pages #"
        Me.Column1.MinimumWidth = 6
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column1.Width = 130
        '
        'Column2
        '
        Me.Column2.HeaderText = "Print type"
        Me.Column2.MinimumWidth = 6
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column2.Width = 135
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(593, 704)
        Me.Controls.Add(Me.ButtonHelp)
        Me.Controls.Add(Me.CmbPrinter)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ButtonPrint)
        Me.Controls.Add(Me.ButtonClose)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ButtonClose As Button
    Friend WithEvents ButtonPrint As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents LinkLabelFile As LinkLabel
    Friend WithEvents TextBoxFile As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents ButtonOdnodelete As Button
    Friend WithEvents ButtonOdnoadd As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents CmbPrinter As ComboBox
    Friend WithEvents ListBoxLog As ListBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents ButtonHelp As Button
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents PageSetupDialog1 As PageSetupDialog
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ButtonDOWN As Button
    Friend WithEvents ButtonUP As Button
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
End Class
