<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserControlKeypad
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
    Me.TextBoxValue = New System.Windows.Forms.TextBox
    Me.ButtonDecimal = New System.Windows.Forms.Button
    Me.Button0 = New System.Windows.Forms.Button
    Me.ButtonClear = New System.Windows.Forms.Button
    Me.Button9 = New System.Windows.Forms.Button
    Me.Button8 = New System.Windows.Forms.Button
    Me.Button7 = New System.Windows.Forms.Button
    Me.Button6 = New System.Windows.Forms.Button
    Me.Button5 = New System.Windows.Forms.Button
    Me.Button4 = New System.Windows.Forms.Button
    Me.Button3 = New System.Windows.Forms.Button
    Me.Button2 = New System.Windows.Forms.Button
    Me.Button1 = New System.Windows.Forms.Button
    Me.SuspendLayout()
    '
    'TextBoxValue
    '
    Me.TextBoxValue.Location = New System.Drawing.Point(0, 4)
    Me.TextBoxValue.Name = "TextBoxValue"
    Me.TextBoxValue.Size = New System.Drawing.Size(168, 21)
    Me.TextBoxValue.TabIndex = 25
    Me.TextBoxValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '
    'ButtonDecimal
    '
    Me.ButtonDecimal.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.ButtonDecimal.Location = New System.Drawing.Point(112, 148)
    Me.ButtonDecimal.Name = "ButtonDecimal"
    Me.ButtonDecimal.Size = New System.Drawing.Size(56, 40)
    Me.ButtonDecimal.TabIndex = 24
    Me.ButtonDecimal.Text = "."
    '
    'Button0
    '
    Me.Button0.Location = New System.Drawing.Point(56, 148)
    Me.Button0.Name = "Button0"
    Me.Button0.Size = New System.Drawing.Size(56, 40)
    Me.Button0.TabIndex = 23
    Me.Button0.Text = "0"
    '
    'ButtonClear
    '
    Me.ButtonClear.Location = New System.Drawing.Point(0, 148)
    Me.ButtonClear.Name = "ButtonClear"
    Me.ButtonClear.Size = New System.Drawing.Size(56, 40)
    Me.ButtonClear.TabIndex = 22
    Me.ButtonClear.Text = "Clear"
    '
    'Button9
    '
    Me.Button9.Location = New System.Drawing.Point(112, 108)
    Me.Button9.Name = "Button9"
    Me.Button9.Size = New System.Drawing.Size(56, 40)
    Me.Button9.TabIndex = 21
    Me.Button9.Text = "9"
    '
    'Button8
    '
    Me.Button8.Location = New System.Drawing.Point(56, 108)
    Me.Button8.Name = "Button8"
    Me.Button8.Size = New System.Drawing.Size(56, 40)
    Me.Button8.TabIndex = 20
    Me.Button8.Text = "8"
    '
    'Button7
    '
    Me.Button7.Location = New System.Drawing.Point(0, 108)
    Me.Button7.Name = "Button7"
    Me.Button7.Size = New System.Drawing.Size(56, 40)
    Me.Button7.TabIndex = 19
    Me.Button7.Text = "7"
    '
    'Button6
    '
    Me.Button6.Location = New System.Drawing.Point(112, 68)
    Me.Button6.Name = "Button6"
    Me.Button6.Size = New System.Drawing.Size(56, 40)
    Me.Button6.TabIndex = 18
    Me.Button6.Text = "6"
    '
    'Button5
    '
    Me.Button5.Location = New System.Drawing.Point(56, 68)
    Me.Button5.Name = "Button5"
    Me.Button5.Size = New System.Drawing.Size(56, 40)
    Me.Button5.TabIndex = 17
    Me.Button5.Text = "5"
    '
    'Button4
    '
    Me.Button4.Location = New System.Drawing.Point(0, 68)
    Me.Button4.Name = "Button4"
    Me.Button4.Size = New System.Drawing.Size(56, 40)
    Me.Button4.TabIndex = 16
    Me.Button4.Text = "4"
    '
    'Button3
    '
    Me.Button3.Location = New System.Drawing.Point(112, 28)
    Me.Button3.Name = "Button3"
    Me.Button3.Size = New System.Drawing.Size(56, 40)
    Me.Button3.TabIndex = 15
    Me.Button3.Text = "3"
    '
    'Button2
    '
    Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Button2.Location = New System.Drawing.Point(56, 28)
    Me.Button2.Name = "Button2"
    Me.Button2.Size = New System.Drawing.Size(56, 40)
    Me.Button2.TabIndex = 14
    Me.Button2.Text = "2"
    '
    'Button1
    '
    Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Button1.Location = New System.Drawing.Point(0, 28)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(56, 40)
    Me.Button1.TabIndex = 13
    Me.Button1.Text = "1"
    '
    'UserControlKeypad
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.TextBoxValue)
    Me.Controls.Add(Me.ButtonDecimal)
    Me.Controls.Add(Me.Button0)
    Me.Controls.Add(Me.ButtonClear)
    Me.Controls.Add(Me.Button9)
    Me.Controls.Add(Me.Button8)
    Me.Controls.Add(Me.Button7)
    Me.Controls.Add(Me.Button6)
    Me.Controls.Add(Me.Button5)
    Me.Controls.Add(Me.Button4)
    Me.Controls.Add(Me.Button3)
    Me.Controls.Add(Me.Button2)
    Me.Controls.Add(Me.Button1)
    Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Name = "UserControlKeypad"
    Me.Size = New System.Drawing.Size(168, 192)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TextBoxValue As System.Windows.Forms.TextBox
  Friend WithEvents ButtonDecimal As System.Windows.Forms.Button
  Friend WithEvents Button0 As System.Windows.Forms.Button
  Friend WithEvents ButtonClear As System.Windows.Forms.Button
  Friend WithEvents Button9 As System.Windows.Forms.Button
  Friend WithEvents Button8 As System.Windows.Forms.Button
  Friend WithEvents Button7 As System.Windows.Forms.Button
  Friend WithEvents Button6 As System.Windows.Forms.Button
  Friend WithEvents Button5 As System.Windows.Forms.Button
  Friend WithEvents Button4 As System.Windows.Forms.Button
  Friend WithEvents Button3 As System.Windows.Forms.Button
  Friend WithEvents Button2 As System.Windows.Forms.Button
  Friend WithEvents Button1 As System.Windows.Forms.Button

End Class
