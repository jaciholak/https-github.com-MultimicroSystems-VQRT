<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmShowHideGrid
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmShowHideGrid))
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.tgShow = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.btnShowSaveExit = New System.Windows.Forms.Button()
        Me.cmdShowAllOn = New System.Windows.Forms.Button()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cmdShowAllOff = New System.Windows.Forms.Button()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboSavePrintOption = New System.Windows.Forms.ComboBox()
        Me.StatusStrip.SuspendLayout()
        CType(Me.tgShow, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 446)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(294, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(39, 17)
        Me.ToolStripStatusLabel.Text = "Status"
        '
        'tgShow
        '
        Me.tgShow.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.tgShow.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tgShow.EmptyRows = True
        Me.tgShow.GroupByCaption = "Drag a column header here to group by that column"
        Me.tgShow.Images.Add(CType(resources.GetObject("tgShow.Images"), System.Drawing.Image))
        Me.tgShow.Location = New System.Drawing.Point(0, 0)
        Me.tgShow.Name = "tgShow"
        Me.tgShow.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tgShow.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tgShow.PreviewInfo.ZoomFactor = 75.0R
        Me.tgShow.PrintInfo.PageSettings = CType(resources.GetObject("tgShow.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tgShow.Size = New System.Drawing.Size(294, 411)
        Me.tgShow.TabIndex = 9
        Me.tgShow.Text = "C1TrueDBGrid1"
        Me.tgShow.UseCompatibleTextRendering = False
        Me.tgShow.PropBag = resources.GetString("tgShow.PropBag")
        '
        'btnShowSaveExit
        '
        Me.btnShowSaveExit.Location = New System.Drawing.Point(9, 5)
        Me.btnShowSaveExit.Name = "btnShowSaveExit"
        Me.btnShowSaveExit.Size = New System.Drawing.Size(92, 23)
        Me.btnShowSaveExit.TabIndex = 11
        Me.btnShowSaveExit.Text = "Save - Exit"
        Me.btnShowSaveExit.UseVisualStyleBackColor = True
        '
        'cmdShowAllOn
        '
        Me.cmdShowAllOn.Location = New System.Drawing.Point(107, 5)
        Me.cmdShowAllOn.Name = "cmdShowAllOn"
        Me.cmdShowAllOn.Size = New System.Drawing.Size(58, 23)
        Me.cmdShowAllOn.TabIndex = 12
        Me.cmdShowAllOn.Text = "All On"
        Me.cmdShowAllOn.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnShowSaveExit)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cmdShowAllOff)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cmdShowAllOn)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(294, 446)
        Me.SplitContainer1.SplitterDistance = 31
        Me.SplitContainer1.TabIndex = 14
        '
        'cmdShowAllOff
        '
        Me.cmdShowAllOff.Location = New System.Drawing.Point(171, 5)
        Me.cmdShowAllOff.Name = "cmdShowAllOff"
        Me.cmdShowAllOff.Size = New System.Drawing.Size(58, 23)
        Me.cmdShowAllOff.TabIndex = 12
        Me.cmdShowAllOff.Text = "All Off"
        Me.cmdShowAllOff.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.tgShow)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.cmdDelete)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboSavePrintOption)
        Me.SplitContainer2.Panel2Collapsed = True
        Me.SplitContainer2.Size = New System.Drawing.Size(294, 411)
        Me.SplitContainer2.SplitterDistance = 342
        Me.SplitContainer2.TabIndex = 10
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(226, 9)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(58, 23)
        Me.cmdDelete.TabIndex = 17
        Me.cmdDelete.Text = "Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(250, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Click on the dropdown to Select/Create a Template"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Template Name:"
        '
        'cboSavePrintOption
        '
        Me.cboSavePrintOption.FormattingEnabled = True
        Me.cboSavePrintOption.Location = New System.Drawing.Point(102, 9)
        Me.cboSavePrintOption.Name = "cboSavePrintOption"
        Me.cboSavePrintOption.Size = New System.Drawing.Size(121, 21)
        Me.cboSavePrintOption.TabIndex = 13
        Me.cboSavePrintOption.Visible = False
        '
        'frmShowHideGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(294, 468)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.StatusStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmShowHideGrid"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Show Hide Grid Columns"
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        CType(Me.tgShow, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents tgShow As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents btnShowSaveExit As System.Windows.Forms.Button
    Friend WithEvents cmdShowAllOn As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cboSavePrintOption As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdShowAllOff As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button

End Class
