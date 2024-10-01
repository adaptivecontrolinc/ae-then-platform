<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormGallons
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
    Me.buttonOK = New System.Windows.Forms.Button()
    Me.buttonCancel = New System.Windows.Forms.Button()
    Me.textBoxGallons = New System.Windows.Forms.TextBox()
    Me.labelGallons = New System.Windows.Forms.Label()
    Me.SuspendLayout()
    '
    'buttonOK
    '
    Me.buttonOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.buttonOK.Location = New System.Drawing.Point(66, 72)
    Me.buttonOK.Name = "buttonOK"
    Me.buttonOK.Size = New System.Drawing.Size(100, 28)
    Me.buttonOK.TabIndex = 0
    Me.buttonOK.Text = "OK"
    Me.buttonOK.UseVisualStyleBackColor = True
    '
    'buttonCancel
    '
    Me.buttonCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.buttonCancel.Location = New System.Drawing.Point(172, 72)
    Me.buttonCancel.Name = "buttonCancel"
    Me.buttonCancel.Size = New System.Drawing.Size(100, 28)
    Me.buttonCancel.TabIndex = 1
    Me.buttonCancel.Text = "Cancel"
    Me.buttonCancel.UseVisualStyleBackColor = True
    '
    'textBoxGallons
    '
    Me.textBoxGallons.Location = New System.Drawing.Point(19, 39)
    Me.textBoxGallons.Name = "textBoxGallons"
    Me.textBoxGallons.Size = New System.Drawing.Size(253, 20)
    Me.textBoxGallons.TabIndex = 2
    '
    'labelGallons
    '
    Me.labelGallons.Location = New System.Drawing.Point(16, 16)
    Me.labelGallons.Name = "labelGallons"
    Me.labelGallons.Size = New System.Drawing.Size(100, 20)
    Me.labelGallons.TabIndex = 3
    Me.labelGallons.Text = "Gallons"
    Me.labelGallons.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'FormMimicGallons
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(284, 112)
    Me.ControlBox = False
    Me.Controls.Add(Me.labelGallons)
    Me.Controls.Add(Me.textBoxGallons)
    Me.Controls.Add(Me.buttonCancel)
    Me.Controls.Add(Me.buttonOK)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "FormMimicGallons"
    Me.ShowIcon = False
    Me.ShowInTaskbar = False
    Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
    Me.Text = "Form Gallons"
    Me.TopMost = True
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Private WithEvents buttonOK As System.Windows.Forms.Button
  Private WithEvents buttonCancel As System.Windows.Forms.Button
  Private WithEvents textBoxGallons As System.Windows.Forms.TextBox
  Private WithEvents labelGallons As System.Windows.Forms.Label
End Class
