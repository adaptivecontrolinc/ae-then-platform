<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPassword
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
    Me.TextBoxValue = New System.Windows.Forms.TextBox
    Me.Button0 = New System.Windows.Forms.Button
    Me.ButtonCancel = New System.Windows.Forms.Button
    Me.Button9 = New System.Windows.Forms.Button
    Me.Button8 = New System.Windows.Forms.Button
    Me.Button7 = New System.Windows.Forms.Button
    Me.Button6 = New System.Windows.Forms.Button
    Me.Button5 = New System.Windows.Forms.Button
    Me.Button4 = New System.Windows.Forms.Button
    Me.Button3 = New System.Windows.Forms.Button
    Me.Button2 = New System.Windows.Forms.Button
    Me.Button1 = New System.Windows.Forms.Button
    Me.ButtonAccept = New System.Windows.Forms.Button
    Me.SuspendLayout()
    '
    'TextBoxValue
    '
    Me.TextBoxValue.Location = New System.Drawing.Point(6, 12)
    Me.TextBoxValue.Name = "TextBoxValue"
    Me.TextBoxValue.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
    Me.TextBoxValue.Size = New System.Drawing.Size(168, 20)
    Me.TextBoxValue.TabIndex = 0
    Me.TextBoxValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    Me.TextBoxValue.UseSystemPasswordChar = True
    '
    'Button0
    '
    Me.Button0.Location = New System.Drawing.Point(6, 176)
    Me.Button0.Name = "Button0"
    Me.Button0.Size = New System.Drawing.Size(56, 40)
    Me.Button0.TabIndex = 10
    Me.Button0.Text = "0"
    '
    'ButtonCancel
    '
    Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.ButtonCancel.Location = New System.Drawing.Point(62, 176)
    Me.ButtonCancel.Name = "ButtonCancel"
    Me.ButtonCancel.Size = New System.Drawing.Size(56, 40)
    Me.ButtonCancel.TabIndex = 11
    Me.ButtonCancel.Text = "Cancel"
    '
    'Button9
    '
    Me.Button9.Location = New System.Drawing.Point(118, 130)
    Me.Button9.Name = "Button9"
    Me.Button9.Size = New System.Drawing.Size(56, 40)
    Me.Button9.TabIndex = 9
    Me.Button9.Text = "9"
    '
    'Button8
    '
    Me.Button8.Location = New System.Drawing.Point(62, 130)
    Me.Button8.Name = "Button8"
    Me.Button8.Size = New System.Drawing.Size(56, 40)
    Me.Button8.TabIndex = 8
    Me.Button8.Text = "8"
    '
    'Button7
    '
    Me.Button7.Location = New System.Drawing.Point(6, 130)
    Me.Button7.Name = "Button7"
    Me.Button7.Size = New System.Drawing.Size(56, 40)
    Me.Button7.TabIndex = 7
    Me.Button7.Text = "7"
    '
    'Button6
    '
    Me.Button6.Location = New System.Drawing.Point(118, 84)
    Me.Button6.Name = "Button6"
    Me.Button6.Size = New System.Drawing.Size(56, 40)
    Me.Button6.TabIndex = 6
    Me.Button6.Text = "6"
    '
    'Button5
    '
    Me.Button5.Location = New System.Drawing.Point(62, 84)
    Me.Button5.Name = "Button5"
    Me.Button5.Size = New System.Drawing.Size(56, 40)
    Me.Button5.TabIndex = 5
    Me.Button5.Text = "5"
    '
    'Button4
    '
    Me.Button4.Location = New System.Drawing.Point(6, 84)
    Me.Button4.Name = "Button4"
    Me.Button4.Size = New System.Drawing.Size(56, 40)
    Me.Button4.TabIndex = 4
    Me.Button4.Text = "4"
    '
    'Button3
    '
    Me.Button3.Location = New System.Drawing.Point(118, 38)
    Me.Button3.Name = "Button3"
    Me.Button3.Size = New System.Drawing.Size(56, 40)
    Me.Button3.TabIndex = 3
    Me.Button3.Text = "3"
    '
    'Button2
    '
    Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Button2.Location = New System.Drawing.Point(62, 38)
    Me.Button2.Name = "Button2"
    Me.Button2.Size = New System.Drawing.Size(56, 40)
    Me.Button2.TabIndex = 2
    Me.Button2.Text = "2"
    '
    'Button1
    '
    Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Button1.Location = New System.Drawing.Point(6, 38)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(56, 40)
    Me.Button1.TabIndex = 1
    Me.Button1.Text = "1"
    '
    'ButtonAccept
    '
    Me.ButtonAccept.Location = New System.Drawing.Point(118, 176)
    Me.ButtonAccept.Name = "ButtonAccept"
    Me.ButtonAccept.Size = New System.Drawing.Size(56, 40)
    Me.ButtonAccept.TabIndex = 12
    Me.ButtonAccept.Text = "Accept"
    '
    'FormPassword
    '
    Me.AcceptButton = Me.ButtonAccept
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.ButtonCancel
    Me.ClientSize = New System.Drawing.Size(182, 218)
    Me.ControlBox = False
    Me.Controls.Add(Me.ButtonAccept)
    Me.Controls.Add(Me.TextBoxValue)
    Me.Controls.Add(Me.Button0)
    Me.Controls.Add(Me.ButtonCancel)
    Me.Controls.Add(Me.Button9)
    Me.Controls.Add(Me.Button8)
    Me.Controls.Add(Me.Button7)
    Me.Controls.Add(Me.Button6)
    Me.Controls.Add(Me.Button5)
    Me.Controls.Add(Me.Button4)
    Me.Controls.Add(Me.Button3)
    Me.Controls.Add(Me.Button2)
    Me.Controls.Add(Me.Button1)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "FormPassword"
    Me.Text = "Enter Password"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TextBoxValue As System.Windows.Forms.TextBox
  Friend WithEvents Button0 As System.Windows.Forms.Button
  Friend WithEvents ButtonCancel As System.Windows.Forms.Button
  Friend WithEvents Button9 As System.Windows.Forms.Button
  Friend WithEvents Button8 As System.Windows.Forms.Button
  Friend WithEvents Button7 As System.Windows.Forms.Button
  Friend WithEvents Button6 As System.Windows.Forms.Button
  Friend WithEvents Button5 As System.Windows.Forms.Button
  Friend WithEvents Button4 As System.Windows.Forms.Button
  Friend WithEvents Button3 As System.Windows.Forms.Button
  Friend WithEvents Button2 As System.Windows.Forms.Button
  Friend WithEvents Button1 As System.Windows.Forms.Button
  Friend WithEvents ButtonAccept As System.Windows.Forms.Button
End Class
