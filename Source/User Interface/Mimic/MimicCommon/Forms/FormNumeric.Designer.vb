<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormNumeric
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
    Me.buttonCancel = New System.Windows.Forms.Button()
    Me.buttonOK = New System.Windows.Forms.Button()
    Me.textBoxValue = New System.Windows.Forms.TextBox()
    Me.keysNumericMain = New KeysNumeric()
    Me.SuspendLayout()
    '
    'buttonCancel
    '
    Me.buttonCancel.Location = New System.Drawing.Point(116, 254)
    Me.buttonCancel.Name = "buttonCancel"
    Me.buttonCancel.Size = New System.Drawing.Size(80, 28)
    Me.buttonCancel.TabIndex = 3
    Me.buttonCancel.Text = "Cancel"
    Me.buttonCancel.UseVisualStyleBackColor = True
    '
    'buttonOK
    '
    Me.buttonOK.Location = New System.Drawing.Point(12, 254)
    Me.buttonOK.Name = "buttonOK"
    Me.buttonOK.Size = New System.Drawing.Size(80, 28)
    Me.buttonOK.TabIndex = 4
    Me.buttonOK.Text = "OK"
    Me.buttonOK.UseVisualStyleBackColor = True
    '
    'textBoxValue
    '
    Me.textBoxValue.Location = New System.Drawing.Point(12, 12)
    Me.textBoxValue.Name = "textBoxValue"
    Me.textBoxValue.Size = New System.Drawing.Size(186, 21)
    Me.textBoxValue.TabIndex = 5
    '
    'keysNumericMain
    '
    Me.keysNumericMain.KeySpacing = 2
    Me.keysNumericMain.Location = New System.Drawing.Point(12, 40)
    Me.keysNumericMain.Name = "keysNumericMain"
    Me.keysNumericMain.Size = New System.Drawing.Size(184, 186)
    Me.keysNumericMain.TabIndex = 6
    '
    'FormNumeric
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(208, 292)
    Me.ControlBox = False
    Me.Controls.Add(Me.keysNumericMain)
    Me.Controls.Add(Me.textBoxValue)
    Me.Controls.Add(Me.buttonOK)
    Me.Controls.Add(Me.buttonCancel)
    Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "FormNumeric"
    Me.ShowIcon = False
    Me.ShowInTaskbar = False
    Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Form Numeric"
    Me.TopMost = True
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Private WithEvents buttonCancel As System.Windows.Forms.Button
  Private WithEvents buttonOK As System.Windows.Forms.Button
  Private WithEvents textBoxValue As System.Windows.Forms.TextBox
  Private WithEvents keysNumericMain As KeysNumeric
End Class
