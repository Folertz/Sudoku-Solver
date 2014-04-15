<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SudokuSolver
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SudokuSolver))
        Me.ButtonClear = New System.Windows.Forms.Button()
        Me.ButtonSolve = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.StatusError = New System.Windows.Forms.ToolStripStatusLabel()
        Me.OwnerLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.EventLog = New System.Windows.Forms.RichTextBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ClearLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonClear
        '
        Me.ButtonClear.Location = New System.Drawing.Point(214, 184)
        Me.ButtonClear.Name = "ButtonClear"
        Me.ButtonClear.Size = New System.Drawing.Size(90, 23)
        Me.ButtonClear.TabIndex = 0
        Me.ButtonClear.Text = "Clear Grid"
        Me.ButtonClear.UseVisualStyleBackColor = True
        '
        'ButtonSolve
        '
        Me.ButtonSolve.Location = New System.Drawing.Point(310, 184)
        Me.ButtonSolve.Name = "ButtonSolve"
        Me.ButtonSolve.Size = New System.Drawing.Size(88, 23)
        Me.ButtonSolve.TabIndex = 1
        Me.ButtonSolve.Text = "Solve"
        Me.ButtonSolve.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusError, Me.OwnerLabel})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 214)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(507, 22)
        Me.StatusStrip1.Stretch = False
        Me.StatusStrip1.TabIndex = 2
        '
        'StatusError
        '
        Me.StatusError.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusError.Name = "StatusError"
        Me.StatusError.Size = New System.Drawing.Size(341, 17)
        Me.StatusError.Spring = True
        Me.StatusError.Text = "Can be solved.."
        Me.StatusError.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OwnerLabel
        '
        Me.OwnerLabel.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OwnerLabel.Name = "OwnerLabel"
        Me.OwnerLabel.Size = New System.Drawing.Size(151, 17)
        Me.OwnerLabel.Text = "© 2013 Xertz Productions"
        '
        'EventLog
        '
        Me.EventLog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.EventLog.ContextMenuStrip = Me.ContextMenuStrip1
        Me.EventLog.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EventLog.Location = New System.Drawing.Point(214, 12)
        Me.EventLog.Name = "EventLog"
        Me.EventLog.ReadOnly = True
        Me.EventLog.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical
        Me.EventLog.Size = New System.Drawing.Size(279, 166)
        Me.EventLog.TabIndex = 3
        Me.EventLog.Text = ""
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClearLogToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(125, 26)
        '
        'ClearLogToolStripMenuItem
        '
        Me.ClearLogToolStripMenuItem.Name = "ClearLogToolStripMenuItem"
        Me.ClearLogToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
        Me.ClearLogToolStripMenuItem.Text = "Clear Log"
        '
        'SudokuSolver
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(507, 236)
        Me.Controls.Add(Me.EventLog)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.ButtonSolve)
        Me.Controls.Add(Me.ButtonClear)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(523, 275)
        Me.MinimumSize = New System.Drawing.Size(523, 275)
        Me.Name = "SudokuSolver"
        Me.Text = "Sudoku Solver"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButtonClear As System.Windows.Forms.Button
    Friend WithEvents ButtonSolve As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents EventLog As System.Windows.Forms.RichTextBox
    Friend WithEvents StatusError As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ClearLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OwnerLabel As System.Windows.Forms.ToolStripStatusLabel

End Class
