<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormNumberPad
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
    Me.controlNumberPad = New ControlNumberPad()
    Me.SuspendLayout()
    '
    'buttonCancel
    '
    Me.buttonCancel.Location = New System.Drawing.Point(112, 247)
    Me.buttonCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
    Me.buttonCancel.Name = "buttonCancel"
    Me.buttonCancel.Size = New System.Drawing.Size(90, 49)
    Me.buttonCancel.TabIndex = 3
    Me.buttonCancel.Text = "Cancel"
    Me.buttonCancel.UseVisualStyleBackColor = True
    '
    'buttonOK
    '
    Me.buttonOK.Location = New System.Drawing.Point(9, 247)
    Me.buttonOK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
    Me.buttonOK.Name = "buttonOK"
    Me.buttonOK.Size = New System.Drawing.Size(90, 49)
    Me.buttonOK.TabIndex = 4
    Me.buttonOK.Text = "OK"
    Me.buttonOK.UseVisualStyleBackColor = True
    '
    'controlNumberPad
    '
    Me.controlNumberPad.Amount = "0"
    Me.controlNumberPad.AmountDouble = 0.0R
    Me.controlNumberPad.AmountInteger = 0
    Me.controlNumberPad.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.controlNumberPad.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.controlNumberPad.KeySize = New System.Drawing.Size(64, 48)
    Me.controlNumberPad.Location = New System.Drawing.Point(9, 10)
    Me.controlNumberPad.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
    Me.controlNumberPad.Name = "controlNumberPad"
    Me.controlNumberPad.ShowAmount = True
    Me.controlNumberPad.Size = New System.Drawing.Size(193, 229)
    Me.controlNumberPad.TabIndex = 6
    '
    'FormNumberPad
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackColor = System.Drawing.Color.Gainsboro
    Me.ClientSize = New System.Drawing.Size(209, 304)
    Me.ControlBox = False
    Me.Controls.Add(Me.controlNumberPad)
    Me.Controls.Add(Me.buttonOK)
    Me.Controls.Add(Me.buttonCancel)
    Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "FormNumberPad"
    Me.ShowIcon = False
    Me.ShowInTaskbar = False
    Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Form Number Pad"
    Me.TopMost = True
    Me.ResumeLayout(False)

  End Sub
  Private WithEvents buttonCancel As System.Windows.Forms.Button
  Private WithEvents buttonOK As System.Windows.Forms.Button
  Private WithEvents controlNumberPad As ControlNumberPad
End Class
