﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormProductRequest
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormProductRequest))
        Me.ToolStripContainer1 = New System.Windows.Forms.ToolStripContainer()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.UcProductRequest1 = New ProductRequestApp.UCProductRequest()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButtonCommit = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonSubmit = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonReSubmit = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonValidate = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonStsCancelled = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonReject = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonComplete = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripContainer1.BottomToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.ContentPanel.SuspendLayout()
        Me.ToolStripContainer1.TopToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripContainer1
        '
        '
        'ToolStripContainer1.BottomToolStripPanel
        '
        Me.ToolStripContainer1.BottomToolStripPanel.Controls.Add(Me.StatusStrip1)
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.TabControl1)
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(1196, 548)
        Me.ToolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStripContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripContainer1.Name = "ToolStripContainer1"
        Me.ToolStripContainer1.Size = New System.Drawing.Size(1196, 595)
        Me.ToolStripContainer1.TabIndex = 0
        Me.ToolStripContainer1.Text = "ToolStripContainer1"
        '
        'ToolStripContainer1.TopToolStripPanel
        '
        Me.ToolStripContainer1.TopToolStripPanel.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripContainer1.TopToolStripPanel.Controls.Add(Me.ToolStrip1)
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1196, 22)
        Me.StatusStrip1.TabIndex = 3
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        Me.ToolStripStatusLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(1048, 17)
        Me.ToolStripStatusLabel2.Spring = True
        Me.ToolStripStatusLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(0, 3)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1196, 542)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.UcProductRequest1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1188, 516)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Product Request"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'UcProductRequest1
        '
        Me.UcProductRequest1.Location = New System.Drawing.Point(6, 9)
        Me.UcProductRequest1.Name = "UcProductRequest1"
        Me.UcProductRequest1.Size = New System.Drawing.Size(1180, 504)
        Me.UcProductRequest1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DataGridView1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1188, 516)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Approval Status"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.DataGridView1.ColumnHeadersHeight = 35
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4})
        Me.DataGridView1.Location = New System.Drawing.Point(25, 20)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(996, 270)
        Me.DataGridView1.TabIndex = 0
        '
        'Column1
        '
        Me.Column1.DataPropertyName = "statusname"
        Me.Column1.HeaderText = "Status"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 150
        '
        'Column2
        '
        Me.Column2.DataPropertyName = "modifiedby"
        Me.Column2.HeaderText = "Modified by"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 150
        '
        'Column3
        '
        Me.Column3.DataPropertyName = "latestupdate"
        DataGridViewCellStyle1.Format = "dd-MMM-yyyy"
        Me.Column3.DefaultCellStyle = DataGridViewCellStyle1
        Me.Column3.HeaderText = "Latest Update"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        '
        'Column4
        '
        Me.Column4.DataPropertyName = "remarks"
        Me.Column4.HeaderText = "Remarks"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        Me.Column4.Width = 200
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripSeparator1, Me.ToolStripButtonCommit, Me.ToolStripButtonSubmit, Me.ToolStripButtonReSubmit, Me.ToolStripButtonValidate, Me.ToolStripButtonStsCancelled, Me.ToolStripButtonReject, Me.ToolStripButtonComplete})
        Me.ToolStrip1.Location = New System.Drawing.Point(3, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(604, 25)
        Me.ToolStrip1.TabIndex = 0
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(157, 22)
        Me.ToolStripButton1.Text = "Print Product Request Form"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButtonCommit
        '
        Me.ToolStripButtonCommit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonCommit.Image = CType(resources.GetObject("ToolStripButtonCommit.Image"), System.Drawing.Image)
        Me.ToolStripButtonCommit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonCommit.Name = "ToolStripButtonCommit"
        Me.ToolStripButtonCommit.Size = New System.Drawing.Size(55, 22)
        Me.ToolStripButtonCommit.Text = "Commit"
        '
        'ToolStripButtonSubmit
        '
        Me.ToolStripButtonSubmit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonSubmit.Image = CType(resources.GetObject("ToolStripButtonSubmit.Image"), System.Drawing.Image)
        Me.ToolStripButtonSubmit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonSubmit.Name = "ToolStripButtonSubmit"
        Me.ToolStripButtonSubmit.Size = New System.Drawing.Size(49, 22)
        Me.ToolStripButtonSubmit.Text = "Submit"
        '
        'ToolStripButtonReSubmit
        '
        Me.ToolStripButtonReSubmit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonReSubmit.Image = CType(resources.GetObject("ToolStripButtonReSubmit.Image"), System.Drawing.Image)
        Me.ToolStripButtonReSubmit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonReSubmit.Name = "ToolStripButtonReSubmit"
        Me.ToolStripButtonReSubmit.Size = New System.Drawing.Size(66, 22)
        Me.ToolStripButtonReSubmit.Text = "Re-submit"
        '
        'ToolStripButtonValidate
        '
        Me.ToolStripButtonValidate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonValidate.Image = CType(resources.GetObject("ToolStripButtonValidate.Image"), System.Drawing.Image)
        Me.ToolStripButtonValidate.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonValidate.Name = "ToolStripButtonValidate"
        Me.ToolStripButtonValidate.Size = New System.Drawing.Size(53, 22)
        Me.ToolStripButtonValidate.Text = "Validate"
        '
        'ToolStripButtonStsCancelled
        '
        Me.ToolStripButtonStsCancelled.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonStsCancelled.Image = CType(resources.GetObject("ToolStripButtonStsCancelled.Image"), System.Drawing.Image)
        Me.ToolStripButtonStsCancelled.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonStsCancelled.Name = "ToolStripButtonStsCancelled"
        Me.ToolStripButtonStsCancelled.Size = New System.Drawing.Size(100, 22)
        Me.ToolStripButtonStsCancelled.Text = "Status-Cancelled"
        '
        'ToolStripButtonReject
        '
        Me.ToolStripButtonReject.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonReject.Image = CType(resources.GetObject("ToolStripButtonReject.Image"), System.Drawing.Image)
        Me.ToolStripButtonReject.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonReject.Name = "ToolStripButtonReject"
        Me.ToolStripButtonReject.Size = New System.Drawing.Size(43, 22)
        Me.ToolStripButtonReject.Text = "Reject"
        '
        'ToolStripButtonComplete
        '
        Me.ToolStripButtonComplete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonComplete.Image = CType(resources.GetObject("ToolStripButtonComplete.Image"), System.Drawing.Image)
        Me.ToolStripButtonComplete.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButtonComplete.Name = "ToolStripButtonComplete"
        Me.ToolStripButtonComplete.Size = New System.Drawing.Size(63, 22)
        Me.ToolStripButtonComplete.Text = "Complete"
        '
        'FormProductRequest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1196, 595)
        Me.Controls.Add(Me.ToolStripContainer1)
        Me.Name = "FormProductRequest"
        Me.Text = "FormProductRequest"
        Me.ToolStripContainer1.BottomToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.BottomToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ContentPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolStripContainer1 As System.Windows.Forms.ToolStripContainer
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Public WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButtonCommit As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonSubmit As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonReSubmit As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonValidate As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonStsCancelled As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonReject As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonComplete As System.Windows.Forms.ToolStripButton
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UcProductRequest1 As ProductRequestApp.UCProductRequest
End Class
