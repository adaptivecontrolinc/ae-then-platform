Namespace Utilities.Color

  Public Module Color

    Public Function ConvertRgbToColor(ByVal rgb As Integer) As System.Drawing.Color

      Dim red As Integer = rgb And 255
      Dim green As Integer = (rgb \ 256) And 255
      Dim blue As Integer = (rgb \ 65536) And 255

      Return System.Drawing.Color.FromArgb(red, green, blue)
    End Function

    Public Function ConvertColorToRgb(ByVal color As System.Drawing.Color) As Integer

      Dim red As Integer = color.R                 ' 
      Dim green As Integer = color.G * 256         ' 2^8
      Dim blue As Integer = color.B * 65536        ' 2^16

      Return red + green + blue
    End Function

    Function ConvertColorToStatusColor(color As System.Drawing.Color) As Integer
      Return color.ToArgb And 16777215
    End Function

  End Module

End Namespace