<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ControlDataRowEdit
  Inherits System.Windows.Forms.UserControl

  'UserControl overrides dispose to clean up the component list.
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
    Me.ListView = New System.Windows.Forms.ListView()
    Me.TextBox = New System.Windows.Forms.TextBox()
    Me.verticalScrollBar = New System.Windows.Forms.VScrollBar()
    Me.SuspendLayout()
    '
    'ListView
    '
    Me.ListView.Dock = System.Windows.Forms.DockStyle.Fill
    Me.ListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
    Me.ListView.LabelWrap = False
    Me.ListView.Location = New System.Drawing.Point(0, 0)
    Me.ListView.MultiSelect = False
    Me.ListView.Name = "ListView"
    Me.ListView.Size = New System.Drawing.Size(331, 307)
    Me.ListView.TabIndex = 0
    Me.ListView.UseCompatibleStateImageBehavior = False
    '
    'TextBox
    '
    Me.TextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.TextBox.Location = New System.Drawing.Point(0, 0)
    Me.TextBox.Multiline = True
    Me.TextBox.Name = "TextBox"
    Me.TextBox.Size = New System.Drawing.Size(100, 20)
    Me.TextBox.TabIndex = 1
    Me.TextBox.Text = "In place edit"
    '
    'verticalScrollBar
    '
    Me.verticalScrollBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.verticalScrollBar.Location = New System.Drawing.Point(312, 2)
    Me.verticalScrollBar.Name = "verticalScrollBar"
    Me.verticalScrollBar.Size = New System.Drawing.Size(17, 303)
    Me.verticalScrollBar.TabIndex = 2
    Me.verticalScrollBar.Visible = False
    '
    'ControlDataRowEdit
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.verticalScrollBar)
    Me.Controls.Add(Me.TextBox)
    Me.Controls.Add(Me.ListView)
    Me.DoubleBuffered = True
    Me.Name = "ControlDataRowEdit"
    Me.Size = New System.Drawing.Size(331, 307)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Public WithEvents ListView As System.Windows.Forms.ListView
  Public WithEvents TextBox As System.Windows.Forms.TextBox
  Private WithEvents verticalScrollBar As System.Windows.Forms.VScrollBar

End Class
