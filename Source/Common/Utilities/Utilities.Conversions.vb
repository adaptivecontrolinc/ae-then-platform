Namespace Utilities.Conversions

  Public Module Conversions

    ' Conversions we are using 
    '   1 Kilogram = 2.204623 pounds
    '   1 US Gallon = 3.785412 liters
    '   1 UK Gallon = 4.54609 liters


    Public ReadOnly Property CelsiusToFarenheit(celsius As Double) As Double
      Get
        Return 32 + (celsius * 9 / 5)
      End Get
    End Property

    Public ReadOnly Property FarenheitToCelsius(farenheit As Double) As Double
      Get
        Return ((farenheit - 32) / 9) * 5
      End Get
    End Property

    Public ReadOnly Property KilogramsToPounds() As Double
      Get
        Return 2.204623
      End Get
    End Property

    Public ReadOnly Property PoundsToKilograms() As Double
      Get
        Return 1 / KilogramsToPounds
      End Get
    End Property

    Public ReadOnly Property GramsToPounds() As Double
      Get
        Return KilogramsToPounds / 1000
      End Get
    End Property

    Public ReadOnly Property PoundsToGrams() As Double
      Get
        Return PoundsToKilograms * 1000
      End Get
    End Property

    Public ReadOnly Property OuncesToGrams() As Double
      Get
        Return PoundsToGrams / 16
      End Get
    End Property

    Public ReadOnly Property LitersToGrams As Double
      Get
        Return 1000
      End Get
    End Property

    Public ReadOnly Property LitersToKilograms As Double
      Get
        Return 1
      End Get
    End Property

    Public ReadOnly Property GallonsUSToLiters() As Double
      Get
        Return 3.785412
      End Get
    End Property

    Public ReadOnly Property LitersToGallonsUS() As Double
      Get
        Return 1 / GallonsUSToLiters
      End Get
    End Property

    Public ReadOnly Property GallonsUKToLiters() As Double
      Get
        Return 4.54609
      End Get
    End Property

    Public ReadOnly Property LitersToGallonsUK() As Double
      Get
        Return 1 / GallonsUKToLiters
      End Get
    End Property

    Public ReadOnly Property GallonsUKToGallonsUS() As Double
      Get
        Return 1.20095
      End Get
    End Property

    Public ReadOnly Property GallonsUSToGallonsUK() As Double
      Get
        Return 1 / GallonsUKToGallonsUS
      End Get
    End Property

    Public ReadOnly Property GallonsUSToPounds() As Double
      Get
        Return 8.3454
      End Get
    End Property

    Public ReadOnly Property PoundsToGallonsUS As Double
      Get
        Return 1 / GallonsUSToPounds
      End Get
    End Property

    Public ReadOnly Property GallonsUKToPounds() As Double
      Get
        Return 10.0224
      End Get
    End Property

    Public ReadOnly Property PoundsToGallonsUK As Double
      Get
        Return 1 / GallonsUKToPounds
      End Get
    End Property

    Public ReadOnly Property MetersToYards() As Double
      Get
        Return 1.0936
      End Get
    End Property

    Public ReadOnly Property YardsToMeters() As Double
      Get
        Return 1 / MetersToYards
      End Get
    End Property


    Public ReadOnly Property BarToPsi() As Double
      Get
        Return 14.5037738
      End Get
    End Property
    Public ReadOnly Property PsiToBar() As Double
      Get
        Return 1 / BarToPsi
      End Get
    End Property


    Public Function ColorToOleRgb(ByVal argb As Integer) As Integer
      ' In OleRGB red is the least significant 8-bits, in .Net blue is the least significant bit (BGR versus RGB)

      Dim color = Drawing.Color.FromArgb(argb)
      Return ColorToOleRgb(color)
    End Function

    Public Function ColorToOleRgb(ByVal color As System.Drawing.Color) As Integer
      ' In OleRGB red is the least significant 8-bits, in .Net blue is the least significant bit (BGR versus RGB)

      Dim red As Integer = color.R                 ' 
      Dim green As Integer = color.G * 256         ' 2^8
      Dim blue As Integer = color.B * 65536        ' 2^16

      Return red + green + blue
    End Function

    Public Function OleRgbToColor(ByVal rgb As Integer) As System.Drawing.Color
      ' In OleRGB red is the least significant 8-bits, in .Net blue is the least significant bit (BGR versus RGB)
      Dim red As Integer = rgb And 255
      Dim green As Integer = (rgb \ 256) And 255
      Dim blue As Integer = (rgb \ 65536) And 255

      Return Drawing.Color.FromArgb(red, green, blue)
    End Function

    Public Function OleRgbOrArgbToColor(ByVal colorInt As Integer) As System.Drawing.Color
      If colorInt < 0 OrElse colorInt > &HFFFFFF Then
        Return Drawing.Color.FromArgb(colorInt)
      Else
        Return OleRgbToColor(colorInt)
      End If
    End Function

  End Module

End Namespace
