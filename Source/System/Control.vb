#If ADAPTIVECONTROLNAMESPACE Or TARGET <> "library" Then
Namespace AdaptiveControl
#End If

Public Enum Unit
  Seconds
  Minutes
  Temperature
  TemperatureTenths
End Enum

''' <summary>Marks an item as not being logged into histories.</summary>
<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)>
Public NotInheritable Class NoLogHistoryAttribute : Inherits Attribute
End Class

''' <summary>Specifies units for displaying the item's value.</summary>
<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)>
Public NotInheritable Class UnitAttribute : Inherits Attribute
  ReadOnly unit_ As Unit
  Sub New(value As Unit)
    unit_ = value
  End Sub
  ReadOnly Property Unit() As Unit
    Get
      Return unit_
    End Get
  End Property
End Class

''' <summary>Specifies a format for displaying the item's value.</summary>
<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)>
Public NotInheritable Class ValueFormatAttribute : Inherits Attribute
  ReadOnly format_ As String
  Sub New(format As String)
    format_ = format
  End Sub
  ReadOnly Property Format() As String
    Get
      Return format_
    End Get
  End Property
End Class

''' <summary>Specifies a translation of the item's name into a particular language.</summary>
<AttributeUsage(AttributeTargets.All, AllowMultiple:=True)>
Public NotInheritable Class TranslateAttribute : Inherits Attribute
  ReadOnly language_, translation_ As String
  ''' <summary>Initializes a new instance of the <see cref="TranslateAttribute"/> class with a language and a translation.</summary>
  ''' <param name="language">The language.</param>
  ''' <param name="translation">The translation.</param>
  Sub New(language As String, translation As String)
    language_ = language : translation_ = translation
  End Sub
  ''' <summary>Gets the language.</summary>
  ''' <returns>the language.</returns>
  ReadOnly Property Language() As String
    Get
      Return language_
    End Get
  End Property
  ''' <summary>Gets the translation.</summary>
  ''' <returns>the translation.</returns>
  ReadOnly Property Translation() As String
    Get
      Return translation_
    End Get
  End Property
#If Not NETFX_CORE Then
  Overrides ReadOnly Property TypeId As Object
    Get
      Return Me
    End Get
  End Property
#End If
End Class

''' <summary>Specifies a translation of the item's description into a particular language.</summary>
<AttributeUsage(AttributeTargets.All, AllowMultiple:=True)>
Public NotInheritable Class TranslateDescriptionAttribute : Inherits Attribute
  ReadOnly language_, description_ As String
  ''' <summary>Initializes a new instance of the <see cref="TranslateDescriptionAttribute"/> class with a language and a description.</summary>
  ''' <param name="language">The language.</param>
  ''' <param name="description">The translation of the description.</param>
  Sub New(language As String, description As String)
    language_ = language : description_ = description
  End Sub
  ''' <summary>Gets the language.</summary>
  ''' <returns>the language.</returns>
  ReadOnly Property Language() As String
    Get
      Return language_
    End Get
  End Property
  ''' <summary>Gets the translated description.</summary>
  ''' <returns>the translated description.</returns>
  ReadOnly Property Description() As String
    Get
      Return description_
    End Get
  End Property
#If Not NETFX_CORE Then
  Overrides ReadOnly Property TypeId As Object
    Get
      Return Me
    End Get
  End Property
#End If
End Class

''' <summary>Specifies a translation of the item's category into a particular language.</summary>
<AttributeUsage(AttributeTargets.All, AllowMultiple:=True)>
Public NotInheritable Class TranslateCategoryAttribute : Inherits Attribute
  ReadOnly language_, category_ As String
  ''' <summary>Initializes a new instance of the <see cref="TranslateCategoryAttribute"/> class with a language and a category.</summary>
  ''' <param name="language">The language.</param>
  ''' <param name="category">The translation of the category.</param>
  Sub New(language As String, category As String)
    language_ = language : category_ = category
  End Sub
  ''' <summary>Gets the language.</summary>
  ''' <returns>the language.</returns>
  ReadOnly Property Language() As String
    Get
      Return language_
    End Get
  End Property
  ''' <summary>Gets the translated category.</summary>
  ''' <returns>the translated category.</returns>
  ReadOnly Property Category() As String
    Get
      Return category_
    End Get
  End Property
#If Not NETFX_CORE Then
  Overrides ReadOnly Property TypeId As Object
    Get
      Return Me
    End Get
  End Property
#End If
End Class

''' <summary>Specifies a translation of a screen button's text into a particular language.</summary>
<AttributeUsage(AttributeTargets.Method, AllowMultiple:=True)>
Public NotInheritable Class TranslateScreenButtonAttribute : Inherits Attribute
  ReadOnly order_ As Integer, language_, text_ As String

  ''' <summary>Initializes a new instance of the <see cref="TranslateScreenButtonAttribute"/> class with a language and a category.</summary>
  ''' <param name="order">The order of the button.</param>
  ''' <param name="language">The language.</param>
  ''' <param name="text">The translation of the button text.</param>
  Sub New(order As Integer, language As String, text As String)
    order_ = order : language_ = language : text_ = text
  End Sub
  ''' <summary>Gets the order.</summary>
  ''' <returns>The order.</returns>
  ReadOnly Property Order() As Integer
    Get
      Return order_
    End Get
  End Property
  ''' <summary>Gets the language.</summary>
  ''' <returns>The language.</returns>
  ReadOnly Property Language() As String
    Get
      Return language_
    End Get
  End Property
  ''' <summary>Gets the translated button text.</summary>
  ''' <returns>The translated button text.</returns>
  ReadOnly Property Text() As String
    Get
      Return text_
    End Get
  End Property
#If Not NETFX_CORE Then
  Overrides ReadOnly Property TypeId As Object
    Get
      Return Me
    End Get
  End Property
#End If
End Class

''' <summary>Specifies a translation of a command into a particular language.</summary>
<AttributeUsage(AttributeTargets.Class, AllowMultiple:=True)>
Public NotInheritable Class TranslateCommandAttribute : Inherits Attribute
  ReadOnly language_, longName_, parameters_ As String
  ''' <summary>Initializes a new instance of the <see cref="TranslateCommandAttribute"/> class with a language, long name and parameters.</summary>
  ''' <param name="language">The language.</param>
  ''' <param name="longName">The translation of the long name.</param>
  ''' <param name="parameters">The translation of the parameters.</param>
  Sub New(language As String, longName As String, Optional parameters As String = "")
    language_ = language : longName_ = longName : parameters_ = parameters
  End Sub
  ''' <summary>Gets the language.</summary>
  ''' <returns>The language.</returns>
  ReadOnly Property Language() As String
    Get
      Return language_
    End Get
  End Property
  ''' <summary>Gets the translated long name.</summary>
  ''' <returns>The translated long name.</returns>
  ReadOnly Property LongName() As String
    Get
      Return longName_
    End Get
  End Property
  ''' <summary>Gets the translated parameters.</summary>
  ''' <returns>the translated parameters.</returns>
  ReadOnly Property Parameters() As String
    Get
      Return parameters_
    End Get
  End Property

#If Not NETFX_CORE Then
  Overrides ReadOnly Property TypeId As Object
    Get
      Return Me
    End Get
  End Property
#End If
End Class


''' <summary>Specifies that the item's value should be persisted while the control system is not running.</summary>
Public NotInheritable Class PersistAttribute : Inherits Attribute
End Class

Public Enum Running
  NotRunning
  RunningButPaused
  RunningNow
End Enum

''' <summary>Specifies the valid range of values for a parameter variable.</summary>
<AttributeUsage(AttributeTargets.Field)>
Public NotInheritable Class ParameterAttribute : Inherits Attribute
  ReadOnly minimumValue_ As Integer, maximumValue_ As Integer

  ''' <summary>Initializes a new instance of the <see cref="ParameterAttribute"/> class.</summary>
  ''' <param name="minimumValue">The minimum value allowed for this parameter.</param>
  ''' <param name="maximumValue">The maximum value allowed for this parameter.</param>
  Sub New(minimumValue As Integer, maximumValue As Integer)
    minimumValue_ = minimumValue : maximumValue_ = maximumValue
  End Sub

  ''' <summary>Gets the minimum value.</summary>
  ''' <returns>the minimum value.</returns>
  ReadOnly Property MinimumValue() As Integer
    Get
      Return minimumValue_
    End Get
  End Property

  ''' <summary>Gets the maximum value.</summary>
  ''' <returns>the maximum value.</returns>
  ReadOnly Property MaximumValue() As Integer
    Get
      Return maximumValue_
    End Get
  End Property
End Class

''' <summary>Specifies the index of the parameter within a set.</summary>
<AttributeUsage(AttributeTargets.Field)>
Public NotInheritable Class ParameterIndexAttribute : Inherits Attribute
  ReadOnly index_ As Integer

  ''' <summary>Initializes a new instance of the <see cref="ParameterIndexAttribute"/> class.</summary>
  ''' <param name="index">The index of this parameter within a set.</param>
  Sub New(index As Integer)
    index_ = index
  End Sub

  ''' <summary>Gets the index.</summary>
  ''' <returns>the index.</returns>
  ReadOnly Property Index() As Integer
    Get
      Return index_
    End Get
  End Property
End Class

Public Enum IOType
  Normal
  Timer
  Dinp
  Dout
  Aninp
  Anout
  Temp
  Pid
  Counter
  Unknown = 255
End Enum

Public Enum Override
  Prevent
  Allow
End Enum

''' <summary>Specifies I/O information for a variable.</summary>
<AttributeUsage(AttributeTargets.Field)>
Public NotInheritable Class IOAttribute : Inherits Attribute
  Implements IComparable(Of IOAttribute)

  <DebuggerBrowsable(DebuggerBrowsableState.Never)>
  ReadOnly ioType_ As IOType, allowOverride_ As Override, device_, format_ As String
  <DebuggerBrowsable(DebuggerBrowsableState.Never)>
  Private channel_ As Integer

  ''' <summary>Initializes a new instance of the <see cref="IOAttribute"/> class.</summary>
  ''' <param name="ioType">The type of the input or output.</param>
  ''' <param name="channel">The channel number to show - has no additional meaning.</param>
  ''' <param name="allowOverride">For an output, whether manually overriding the output value is allowed.</param>
  ''' <param name="device">An optional device type - like Motor.</param>
  ''' <param name="format">An optional format for displaying the value.</param>
  Sub New(ioType As IOType, channel As Integer,
          Optional allowOverride As Override = Override.Allow,
          Optional device As String = "", Optional format As String = "")
    ioType_ = ioType : channel_ = channel : allowOverride_ = allowOverride
    device_ = device : format_ = format
  End Sub

  ''' <summary>Gets the i/o type.</summary>
  ''' <returns>the i/o type.</returns>
  ReadOnly Property IOType() As IOType
    Get
      Return ioType_
    End Get
  End Property

  ''' <summary>Gets or sets the channel number.</summary>
  ''' <returns>the channel number.</returns>
  Property Channel() As Integer
    Get
      Return channel_
    End Get
    Set(value As Integer)
      channel_ = value
    End Set
  End Property

  ''' <summary>Gets whether manually overriding the output value is allowed.</summary>
  ''' <returns>true if manually overriding the output value is allowed; otherwise, false.</returns>
  ReadOnly Property AllowOverride() As Override
    Get
      Return allowOverride_
    End Get
  End Property

  ''' <summary>Gets the device.</summary>
  ''' <returns>the device.</returns>
  ReadOnly Property Device() As String
    Get
      Return device_
    End Get
  End Property

  ''' <summary>Gets the format for displaying the value.</summary>
  ''' <returns>the format for displaying the value.</returns>
  ReadOnly Property Format() As String
    Get
      Return format_
    End Get
  End Property

  Function CompareTo(other As IOAttribute) As Integer Implements IComparable(Of IOAttribute).CompareTo
    ' First compare type, then channel
    Dim ret = ioType_ - other.ioType_
    Return If(ret <> 0, ret, channel_ - other.channel_)
  End Function
End Class

''' <summary>Specifies that this control system can only run one program at a time, and so for example UI's should 
''' only associate a single program number with a batch.</summary>
<AttributeUsage(AttributeTargets.Assembly)>
Public NotInheritable Class SingleProgramAttribute : Inherits Attribute
  ReadOnly value_ As Boolean
  Sub New(value As Boolean)
    value_ = value
  End Sub

  ' Required to create from a serialized string
  Sub New(str As String)
    value_ = Boolean.Parse(str)
  End Sub

  ReadOnly Property Value() As Boolean
    Get
      Return value_
    End Get
  End Property
  Overrides Function ToString() As String
    Return value_.ToString
  End Function
End Class

''' <summary>Specifies that this assembly should allow IP functionality in its program editor.</summary>
<AttributeUsage(AttributeTargets.Assembly)>
Public NotInheritable Class IncludeIPAttribute : Inherits Attribute
  ReadOnly value_ As Boolean
  Sub New(value As Boolean)
    value_ = value
  End Sub

  ' Required to create from a serialized string
  Sub New(str As String)
    value_ = Boolean.Parse(str)
  End Sub

  ReadOnly Property Value() As Boolean
    Get
      Return value_
    End Get
  End Property
  Overrides Function ToString() As String
    Return value_.ToString
  End Function
End Class

''' <summary>Specifies the maximum program number allowed for programs.</summary>
<AttributeUsage(AttributeTargets.Assembly)>
Public NotInheritable Class MaximumProgramNumberAttribute : Inherits Attribute
  ReadOnly value_ As Integer
  Sub New(value As Integer)
    value_ = value
  End Sub

  ' Required to create from a serialized string
  Sub New(str As String)
    value_ = Integer.Parse(str, Globalization.NumberStyles.None, InvariantCulture)
  End Sub

  ReadOnly Property Value() As Integer
    Get
      Return value_
    End Get
  End Property

  Overrides Function ToString() As String
    Return value_.ToString(System.Globalization.CultureInfo.InvariantCulture)
  End Function
End Class


''' <summary>Adds extra profile information to a command, for a name like "pH".</summary>
<AttributeUsage(AttributeTargets.Class Or AttributeTargets.Field)>
Public Class ProfileAttribute : Inherits Attribute
  ReadOnly name_, gradient_, target_, minutes_ As String

  ''' <summary>Initializes a new instance of the <see cref="ProfileAttribute"/> class.</summary>
  ''' <param name="name">The name for example "pH"</param>
  ''' <param name="gradient">The gradient.</param>
  ''' <param name="target">The target.</param>
  ''' <param name="minutes">The minutes.</param>
  Sub New(name As String, gradient As String, target As String, minutes As String)
    name_ = name
    gradient_ = gradient : target_ = target : minutes_ = minutes
  End Sub

  ''' <summary>Gets the name.</summary>
  ''' <returns>the name.</returns>
  ReadOnly Property Name() As String
    Get
      Return name_
    End Get
  End Property
  ''' <summary>Gets the gradient.</summary>
  ''' <returns>the gradient.</returns>
  ReadOnly Property Gradient() As String
    Get
      Return gradient_
    End Get
  End Property
  ''' <summary>Gets the target.</summary>
  ''' <returns>the target.</returns>
  ReadOnly Property Target() As String
    Get
      Return target_
    End Get
  End Property
  ''' <summary>Gets the minutes.</summary>
  ''' <returns>the minutes.</returns>
  ReadOnly Property Minutes() As String
    Get
      Return minutes_
    End Get
  End Property
End Class


<Flags()>
Public Enum CommandType
  Standard = 0
  BatchParameter = 1
  ParallelCommand = 2
End Enum

''' <summary>Specifies that this class is a command.</summary>
<AttributeUsage(AttributeTargets.Class Or AttributeTargets.Field)>
Public Class CommandAttribute : Inherits Attribute
  ReadOnly longName_, parameters_, gradient_, target_, minutes_ As String, commandType_ As CommandType

  ''' <summary>Initializes a new instance of the <see cref="CommandAttribute"/> class with a long name and other optional parameters.</summary>
  ''' <param name="longName">The long name.</param>
  ''' <param name="parameters">The parameters.</param>
  ''' <param name="gradient">The gradient.</param>
  ''' <param name="target">The target.</param>
  ''' <param name="minutes">The minutes.</param>
  ''' <param name="commandType">The command type.</param>
  Sub New(longName As String, Optional parameters As String = "",
          Optional gradient As String = "", Optional target As String = "",
          Optional minutes As String = "", Optional commandType As CommandType = CommandType.Standard)
    longName_ = longName : parameters_ = parameters
    gradient_ = gradient : target_ = target : minutes_ = minutes
    commandType_ = commandType
  End Sub

  ''' <summary>Gets the long name.</summary>
  ''' <returns>the long name.</returns>
  ReadOnly Property LongName() As String
    Get
      Return longName_
    End Get
  End Property
  ''' <summary>Gets the parameters.</summary>
  ''' <returns>the parameters.</returns>
  ReadOnly Property Parameters() As String
    Get
      Return parameters_
    End Get
  End Property
  ''' <summary>Gets the gradient.</summary>
  ''' <returns>the gradient.</returns>
  ReadOnly Property Gradient() As String
    Get
      Return gradient_
    End Get
  End Property
  ''' <summary>Gets the target.</summary>
  ''' <returns>the target.</returns>
  ReadOnly Property Target() As String
    Get
      Return target_
    End Get
  End Property
  ''' <summary>Gets the minutes.</summary>
  ''' <returns>the minutes.</returns>
  ReadOnly Property Minutes() As String
    Get
      Return minutes_
    End Get
  End Property
  ''' <summary>Gets the command type.</summary>
  ''' <returns>the command type.</returns>
  ReadOnly Property CommandType() As CommandType
    Get
      Return commandType_
    End Get
  End Property

  ReadOnly Property IsBatchParameter() As Boolean
    Get
      Return (commandType_ And CommandType.BatchParameter) <> 0
    End Get
  End Property
  ReadOnly Property IsParallelCommand() As Boolean
    Get
      Return (commandType_ And CommandType.ParallelCommand) <> 0
    End Get
  End Property
End Class

Public Enum ButtonImage
  Vessel = 1
  Dial
  Thermometer
  FillBucket
  SideVessel
  Beam
  Information
End Enum

''' <summary>Defines a button for displaying screen information.</summary>
<AttributeUsage(AttributeTargets.Method, AllowMultiple:=True)>
Public NotInheritable Class ScreenButtonAttribute : Inherits Attribute
  ReadOnly text_ As String, order_ As Integer, image_ As ButtonImage

  ''' <summary>Initializes a new instance of the <see cref="ScreenButtonAttribute"/> class.</summary>
  ''' <param name="text">The text to go on the button.</param>
  ''' <param name="order">The order of this button - use 1, 2, 3, etc for your buttons.</param>
  ''' <param name="image">The image to display on the button.</param>
  Sub New(text As String, order As Integer, image As ButtonImage)
    text_ = text : order_ = order : image_ = image
  End Sub

  ''' <summary>Gets the button text.</summary>
  ''' <returns>the button text.</returns>
  ReadOnly Property Text() As String
    Get
      Return text_
    End Get
  End Property

  ''' <summary>Gets the order.</summary>
  ''' <returns>the order.</returns>
  ReadOnly Property Order() As Integer
    Get
      Return order_
    End Get
  End Property

  ''' <summary>Gets the button image.</summary>
  ''' <returns>the button image.</returns>
  ReadOnly Property Image() As ButtonImage
    Get
      Return image_
    End Get
  End Property
#If Not NETFX_CORE Then
  Overrides ReadOnly Property TypeId As Object
    Get
      Return Me
    End Get
  End Property
#End If
End Class

<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)>
Public NotInheritable Class GraphSetAttribute : Inherits Attribute
  ReadOnly sets_ As String()
  Sub New(ParamArray sets As String())
    sets_ = sets
  End Sub
  ReadOnly Property Sets As String()
    Get
      Return sets_
    End Get
  End Property
End Class

''' <summary>Specifies how this item should be shown as a trace on a history graph.</summary>
<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)>
Public NotInheritable Class GraphTraceAttribute : Inherits Attribute
  ReadOnly minimumValue_ As Integer, maximumValue_ As Integer, minimumY_ As Integer, maximumY_ As Integer,
           colorName_, format_ As String

  ''' <summary>Initializes a new instance of the <see cref="GraphTraceAttribute"/> class with minimum and maximum values, and other optional parameters.</summary>
  ''' <param name="minimumValue">The minimum value of the variable which will shown on the graph.  Any smaller value is not shown.</param>
  ''' <param name="maximumValue">The maximum value of the variable which will shown on the graph.  Any larger value is not shown.</param>
  ''' <param name="minimumY">The bottom Y position on the graph corresponding to the <paramref name="minimumValue"/>.  Use 0 for the bottom of the graph.</param>
  ''' <param name="maximumY">The top Y position on the graph corresponding to the <paramref name="maximumValue"/>.  Use 10000 for the top of the graph.</param>
  ''' <param name="colorName">The color to display this graph trace.  If not specified, then the colors Olive, Red, Magenta, Blue, Navy, Maroon, and Green are used in rotation.</param>
  ''' <param name="format">The format in which to display the value.</param>
  Sub New(minimumValue As Integer, maximumValue As Integer,
          Optional minimumY As Integer = 0, Optional maximumY As Integer = 10000,
          Optional colorName As String = "", Optional format As String = "")
    minimumValue_ = minimumValue : maximumValue_ = maximumValue
    minimumY_ = minimumY : maximumY_ = maximumY
    colorName_ = colorName : format_ = format
  End Sub

  ''' <summary>Gets the minimum value.</summary>
  ''' <returns>the minimum value.</returns>
  ReadOnly Property MinimumValue() As Integer
    Get
      Return minimumValue_
    End Get
  End Property

  ''' <summary>Gets the maximum value.</summary>
  ''' <returns>the maximum value.</returns>
  ReadOnly Property MaximumValue() As Integer
    Get
      Return maximumValue_
    End Get
  End Property

  ''' <summary>Gets the minimum Y.</summary>
  ''' <returns>the command type.</returns>
  <DefaultValue(0)>
  ReadOnly Property MinimumY() As Integer
    Get
      Return minimumY_
    End Get
  End Property

  ''' <summary>Gets the maximum Y.</summary>
  ''' <returns>the command type.</returns>
  <DefaultValue(10000)>
  ReadOnly Property MaximumY() As Integer
    Get
      Return maximumY_
    End Get
  End Property

  ''' <summary>Gets the color name.</summary>
  ''' <returns>the color name.</returns>
  <DefaultValue("")>
  ReadOnly Property ColorName() As String
    Get
      Return colorName_
    End Get
  End Property

  ''' <summary>Gets the format for displaying values.</summary>
  ''' <returns>the format for displaying values.</returns>
  <DefaultValue("")>
  ReadOnly Property Format() As String
    Get
      Return format_
    End Get
  End Property
End Class

''' <summary>Specifies a label for display on a history graph.</summary>
<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property, AllowMultiple:=True)>
Public NotInheritable Class GraphLabelAttribute : Inherits Attribute
  ReadOnly label_ As String, value_ As Integer

  ''' <summary>Initializes a new instance of the <see cref="GraphLabelAttribute"/> class.</summary>
  ''' <param name="label">The text of the label.</param>
  ''' <param name="value">The variable value corresponding to this label's position.</param>
  Sub New(label As String, value As Integer)
    label_ = label : value_ = value
  End Sub

  ''' <summary>Gets the text of the label.</summary>
  ''' <returns>the text of the label.</returns>
  ReadOnly Property Label() As String
    Get
      Return label_
    End Get
  End Property

  ''' <summary>Gets the variable value corresponding to this label's position.</summary>
  ''' <returns>the variable value corresponding to this label's position.</returns>
  ReadOnly Property Value() As Integer
    Get
      Return value_
    End Get
  End Property

#If Not NETFX_CORE Then
  Overrides ReadOnly Property TypeId As Object
    Get
      Return Me
    End Get
  End Property
#End If
End Class

''' <summary>Specifies by how much the channel number increases for an array of I/O items.</summary>
<AttributeUsage(AttributeTargets.Field)>
Public NotInheritable Class IOChannelIncrementAttribute : Inherits Attribute
  ReadOnly increment_ As Integer

  ''' <summary>Initializes a new instance of the <see cref="IOChannelIncrementAttribute"/> class.</summary>
  ''' <param name="increment">The increment in channel number, which defaults to 1 if this <see cref="IOChannelIncrementAttribute"/> is not specified.</param>
  Sub New(increment As Integer)
    increment_ = increment
  End Sub

  ''' <summary>Gets the increment.</summary>
  ''' <returns>the increment.</returns>
  ReadOnly Property Increment() As Integer
    Get
      Return increment_
    End Get
  End Property
End Class

Public Enum AckStateValue
  Off
  UnackMessage
  AckMessage
  QQ
  MessageIsDone
  QQIsDoneAndAnswerIsYes
  QQIsDoneAndAnswerIsNo
End Enum
Public Enum Mode
  Run
  Debug
  Test
  Override
End Enum


#If TARGET = "library" Then
''' <summary>Defines methods that are used to run a command.</summary>
Public Interface ACCommand
  ''' <summary>Called when a command is started.</summary>
  ''' <param name="param">The parameter values.</param>
  ''' <returns>True if the command should step on, False if not.</returns>
  Function Start(ParamArray param() As Integer) As Boolean
  ''' <summary>Called to run the command.  Called all the time, whether or not the command is active.</summary>
  ''' <returns>true if the command should step on, false if not.</returns>
  Function Run() As Boolean
  ''' <summary>Call to cancel the command, after which <see cref="IsOn"/> should return False.</summary>
  Sub Cancel()
  ''' <summary>Gets whether the command is currently active.</summary>
  ''' <returns>true if the command is active; otherwise, false.</returns>
  ReadOnly Property IsOn() As Boolean
  ''' <summary>Called when the values of parameter values have been changed.</summary>
  ''' <param name="param">The new parameter values.</param>
  Sub ParametersChanged(ParamArray param() As Integer)
End Interface
#End If


#If Not NOADAPTIVECONTROLPARENT Then
Public Interface ACControlCodeSubSet
  ''' <summary>Gets whether a job is running right now.</summary>
  ''' <returns>true if a job is running right now; otherwise, false.</returns>
  ReadOnly Property IsProgramRunning() As Boolean

  ''' <summary>Called each time a job (program) stops.</summary>
  Sub ProgramStop(history As Object, historyBytes As Byte())
End Interface
#End If

''' <summary>A formula for the maximum gradient for temperature profiles, for example 'Parameters_MaxGrad
''' This attribute should be applied to the ControlCode class.</summary>
<AttributeUsage(AttributeTargets.Class)>
Public NotInheritable Class MaximumGradientAttribute : Inherits Attribute
  ReadOnly gradient_ As String

  Sub New(gradient As String)
    gradient_ = gradient
  End Sub

  ReadOnly Property Gradient() As String
    Get
      Return gradient_
    End Get
  End Property
End Class


#If Not NOADAPTIVECONTROLPARENT Then
''' <summary>Defines methods that are used to run the control code.</summary>
Public Interface ACControlCode
  ''' <summary>Called just once when the control system starts up for the first time.</summary>
  ''' <remarks>This will be shortly after the instance ctor for the class has been called.  All other control systems will have been created at this time, for example.</remarks>
  Sub StartUp()

  ''' <summary>Called just once when the control system is being shut down.  At this time, there will be approximately 3 seconds remaining before <see cref="ReadInputs"/>, <see cref="Run"/> and <see cref="WriteOutputs"/> stop being called.</summary>
  Sub ShutDown()

  ''' <summary>Called to read input values into the given arrays.</summary>
  ''' <param name="dinp">An array into which digital input values should be stored.</param>
  ''' <param name="aninp">An array into which analog input values should be stored.</param>
  ''' <param name="temp">An array into which temperature values should be stored.</param>
  ''' <returns>true if a cycle of I/O reads has completed without errors.</returns>
  Function ReadInputs(dinp() As Boolean, aninp() As Short, temp() As Short) As Boolean

  ''' <summary>Called all the time to run the control.</summary>
  Sub Run()

  ''' <summary>Called to write output values found in the given arrays.</summary>
  ''' <param name="dout">An array in which digital output values are found.</param>
  ''' <param name="anout">An array in which analog output values are found.</param>
  Sub WriteOutputs(dout() As Boolean, anout() As Short)

  ''' <summary>Called to get text information for display on a screen.</summary>
  ''' <param name="screen">The screen number.</param>
  ''' <param name="row">An array of 100 rows in which text can be stored. <paramref name="row"/>[0] is not used.</param>
  Sub DrawScreen(screen As Integer, row() As String)

  ''' <summary>Called each time a job (program) starts.</summary>
  Sub ProgramStart()

  ''' <summary>Called each time a job (program) stops.</summary>
  Sub ProgramStop()
End Interface

Public Enum ButtonPosition
  [Operator]
  Expert
End Enum
Public Enum LogEventType
  [Error] = 2
  Warning = 4
  Information = 8
End Enum



Public Enum StandardButton
  WorkList
  Program
  Graph
  Mimic
  History
  IO
  Variables
  Parameters
  Programs
  Clean
  Sleep
  Ladder
  Camera
End Enum

''' <summary>Provides methods for accessing the environment in which the control code is running.</summary>
Public Interface ACParent
  ''' <summary>Gets the acknowledgment state.</summary>
  ''' <returns>the acknowledgment state.</returns>
  ReadOnly Property AckState() As AckStateValue

  ''' <summary>Gets the text currently showing on the button used for messages.</summary>
  ''' <returns>the text currently showing on the button used for messages.</returns>
  ReadOnly Property ButtonText() As String

  ''' <summary>Gets or sets a message for the operator.</summary>
  ''' <returns>a message for the operator.</returns>
  ''' <remarks>Becomes an empty string again when the operator acknowledges the message.</remarks>
  Property Signal() As String

  ''' <summary>Gets a list of unacknowledged alarms, separated by commas.</summary>
  ''' <returns>a list of unacknowledged alarms, separated by commas.</returns>
  ReadOnly Property UnacknowledgedAlarms() As String

  ''' <summary>Simulates pressing the Run, Pause, Stop, Yes and No buttons.</summary>
  ''' <remarks>Only when the values go from false to true are they acted on.</remarks>
  Sub PressButtons(run As Boolean, pause As Boolean, [stop] As Boolean, yes As Boolean, no As Boolean)

  ''' <summary>Simulates pressing the Forward and Backward buttons.</summary>
  ''' <remarks>Only when the values go from false to true are they acted on.</remarks>
  Sub PressForwardBackward(forward As Boolean, backward As Boolean)

  ''' <summary>Gets whether a job is running right now. You can also get this by checking the <see cref="Running"/> property.</summary>
  ''' <returns>true if a job is running right now; otherwise, false.</returns>
  ReadOnly Property IsProgramRunning() As Boolean

  ''' <summary>Gets whether a job is running right now, but has been paused. You can also get this by checking the <see cref="Running"/> property.</summary>
  ''' <returns>true if a job is running right now, but has been paused; otherwise, false.</returns>
  ReadOnly Property IsPaused() As Boolean

  ''' <summary>Gets whether there is an unacknowledged signal.</summary>
  ''' <returns>true if there is an unacknowledged signal; otherwise, false.</returns>
  ReadOnly Property IsSignalUnacknowledged() As Boolean

  ''' <summary>Gets whether there is an alarm.</summary>
  ''' <returns>true if there is an alarm; otherwise, false.</returns>
  ReadOnly Property IsAlarmActive() As Boolean

  ''' <summary>Gets whether there is an unacknowledged alarm.</summary>
  ''' <returns>true if there is an unacknowledged alarm; otherwise, false.</returns>
  ReadOnly Property IsAlarmUnacknowledged() As Boolean

  ''' <summary>Gets a message.</summary>
  ''' <param name="messageNumber">The message number, 1 to 99</param>
  ''' <returns>the message of given number.</returns>
  ReadOnly Property Message(messageNumber As Integer) As String

  ''' <summary>Gets or sets the running mode.</summary>
  ''' <returns>the running mode.</returns>
  Property Mode() As Mode

  ''' <summary>Gets the number of the program that the sequencer has reached.</summary>
  ''' <returns>the number of the program that the sequencer has reached.</returns>
  ''' <remarks>This will change if the sequencer is paused and the step is being changed.  
  ''' Watch out - if program numbers can be strings, then this will have to return 0.</remarks>
  ReadOnly Property ProgramNumber() As Integer


  ''' <summary>Gets the number of the program that the sequencer has reached.</summary>
  ''' <returns>the number of the program that the sequencer has reached.</returns>
  ''' <remarks>This will change if the sequencer is paused and the step is being changed.</remarks>
  ReadOnly Property ProgramNumberAsString() As String

  ''' <summary>Gets the name of the program that the sequencer has reached.</summary>
  ''' <returns>the name of the program that the sequencer has reached.</returns>
  ''' <remarks>This will change if the sequencer is paused and the step is being changed.</remarks>
  ReadOnly Property ProgramName() As String

  ''' <summary>Gets the number of the step that the sequencer has reached.</summary>
  ''' <returns>the number of the step that the sequencer has reached.</returns>
  ''' <remarks>This will change if the sequencer is paused and the step is being changed.</remarks>
  ReadOnly Property StepNumber() As Integer

  ''' <summary>Gets the text of the step that the sequencer has reached.</summary>
  ''' <returns>the text of the step that the sequencer has reached.</returns>
  ''' <remarks>This will change if the sequencer is paused and the step is being changed.</remarks>
  ReadOnly Property StepText() As String

  ''' <summary>Gets the time that the sequencer has spent in the current step versus how long it is supposed to spend there in total.</summary>
  ''' <returns>the time that the sequencer has spent in the current step versus how long it is supposed to spend there in total.</returns>
  ''' <remarks>For example, "10 / 15" means 10 minutes out of an expected total of 15 minutes.</remarks>
  ReadOnly Property TimeInStep() As String

  ''' <summary>Gets the current step number, relative to all steps in all programs.</summary>
  ''' <returns>the current step number, relative to all steps in all programs.</returns>
  ReadOnly Property CurrentStep() As Integer

  ''' <summary>Gets the current changing step number, relative to all steps in all programs.</summary>
  ''' <returns>the current changing step number, relative to all steps in all programs.</returns>
  ReadOnly Property ChangingStep() As Integer

  ''' <summary>Gets the name of the job being run.</summary>
  ''' <returns>the name of the job being run.</returns>
  ReadOnly Property Job() As String

  ''' <summary>Gets information for all steps, prefixed by extra information.</summary>
  ''' <returns>information for all steps, prefixed by extra information.</returns>
  ReadOnly Property PrefixedSteps() As String

  ''' <summary>Gets information for all messages.</summary>
  ''' <returns>information for all messages.</returns>
  ReadOnly Property Messages() As String

  ''' <summary>Gets information for all programs.</summary>
  ''' <returns>information for all programs.</returns>
  ReadOnly Property Programs() As String

  ''' <summary>Gets information for time in steps.</summary>
  ''' <returns>information for time in steps.</returns>
  ReadOnly Property TimeInSteps() As String

  ''' <summary>Gets the time elapsed since the job started.</summary>
  ''' <returns>the time elapsed since the job started.</returns>
  ReadOnly Property ElapsedTime() As TimeSpan

  ''' <summary>Gets the time that should have elapsed since the job started.</summary>
  ''' <returns>the time that should have elapsed since the job started.</returns>
  ReadOnly Property ElapsedTimeExpected() As TimeSpan

  ''' <summary>Gets the time remaining until the job ends.</summary>
  ''' <returns>the time remaining until the job ends.</returns>
  ReadOnly Property RemainingTime() As TimeSpan

  ''' <summary>Gets the numbers of all programs that were specified when the current job was started.</summary>
  ''' <returns>the numbers of all programs that were specified when the current job was started.</returns>
  ReadOnly Property ProgramNumbers() As System.Collections.ObjectModel.ReadOnlyCollection(Of Integer)

  ''' <summary>Gets the names of all active alarms, separated by commas.</summary>
  ''' <returns>the names of all active alarms, separated by commas.</returns>
  ReadOnly Property ActiveAlarms() As String

  ''' <summary>Gets whether the system is sleeping.</summary>
  ''' <returns>true if the system is sleeping; otherwise, false.</returns>
  ReadOnly Property IsSleeping() As Boolean

  ''' <summary>Gets the name of our own control system.</summary>
  ''' <returns>the name of our own control system.</returns>
  ReadOnly Property ControlSystemName() As String

  ''' <summary>Gets the number of our own control system, relative to others, starting from 1.</summary>
  ''' <returns>the number of our own control system, relative to others, starting from 1.</returns>
  ReadOnly Property ControlSystemNumber() As Integer

  ''' <summary>Gets the 'ControlCode' instance of another control system.</summary>
  ''' <param name="index">The index of the control system, starting from 1.</param>
  ''' <returns>the 'ControlCode' instance of another control system.</returns>
  ReadOnly Property ControlSystem(index As Integer) As Object

  ''' <summary>Gets the total number of control systems.</summary>
  ''' <returns>the total number of control systems.</returns>
  ReadOnly Property ControlSystemCount() As Integer

  ''' <summary>Gets the 'ControlCode' instance of the control system that is the master of coupling.</summary>
  ''' <returns>the 'ControlCode' instance of the control system that is the master of coupling.</returns>
  ''' <remarks>Returns null if our control system is not coupled.  Also returns null if our control system is coupled, and we are ourself the master.</remarks>
  ReadOnly Property CouplingMaster() As Object

  ''' <summary>Gets whether we are running in coupled mode.</summary>
  ''' <returns>true if running in coupled mode; otherwise, false.</returns>
  ReadOnly Property Coupled() As Boolean

  ''' <summary>Gets the coupling combination number, starting from 1, or 0 if not coupled.</summary>
  ''' <returns>the coupling combination number, starting from 1, or 0 if not coupled.</returns>
  ReadOnly Property CouplingCombination() As Integer

  ''' <summary>Sets whether the current coupled function is ready.</summary>
  Property CouplingReady() As Boolean

  ''' <summary>Gets whether the current coupled function is ready for all coupled control systems.</summary>
  ''' <returns>true if the current coupled function is ready for all coupled control systems; otherwise, false.</returns>
  ReadOnly Property CouplingAllReady() As Boolean

  Property CouplingReady2() As Boolean
  ReadOnly Property CouplingAllReady2() As Boolean
  Property CouplingReady3() As Boolean
  ReadOnly Property CouplingAllReady3() As Boolean

  ''' <summary>Sets whether the current coupled function is safe.</summary>
  Property CouplingSafe() As Boolean

  ''' <summary>Gets whether the current coupled function is safe for all coupled control systems.</summary>
  ''' <returns>true if the current coupled function is safe for all coupled control systems; otherwise, false.</returns>
  ReadOnly Property CouplingAllSafe() As Boolean

  Property CouplingSafe2() As Boolean
  ReadOnly Property CouplingAllSafe2() As Boolean
  Property CouplingSafe3() As Boolean
  ReadOnly Property CouplingAllSafe3() As Boolean

  ''' <summary>Adds a button to one of the main toolbars.</summary>
  ''' <param name="button">The button to add, a ToolStripItem, usually a ToolStripButton.</param>
  ''' <param name="position">The toolbar on which to add the button.</param>
  ''' <param name="view">The view to be shown when the button is pressed.</param>
  Sub AddButton(button As Object, position As ButtonPosition, Optional view As Object = Nothing)

  ''' <summary>Adds a standard button (like Variables) to one of the main toolbars.</summary>
  ''' <param name="button">The button to add.</param>
  ''' <param name="position">The toolbar on which to add the button.</param>
  ''' <param name="text">Optionally, the text to show on the button, if standard text is not wanted.</param>
  ''' <param name="options">Further options for the standard button.</param>
  ''' <remarks>If used with <paramref name="button"/>=<see cref="StandardButton.Program"/>, then the following <paramref name="options"/> can be used for the program window:  Pause=1, Run=2, ChangeStep=4,Stop=8, YesOrNo=16, Halt=32.</remarks>
  Sub AddStandardButton(button As StandardButton, position As ButtonPosition, Optional text As String = Nothing,
                        Optional options As Integer = 0)

  ''' <summary>Presses a button on one of the main toolbars.</summary>
  ''' <param name="position">The toolbar.</param>
  ''' <param name="index">The zero-based index of the button on the toolbar.</param>
  Sub PressButton(position As ButtonPosition, index As Integer)

  ''' <summary>Makes a button on one of the main toolbars visible or not visible.</summary>
  ''' <param name="position">The toolbar.</param>
  ''' <param name="index">The zero-based index of the burron on the toolbar.</param>
  ''' <param name="visible">Whether the button is visible.</param>
  Sub SetButtonVisible(position As ButtonPosition, index As Integer, visible As Boolean)

  ''' <summary>Sets the text for a button on one of the main toolbars.</summary>
  ''' <param name="position">The toolbar.</param>
  ''' <param name="index">The zero-based index of the burron on the toolbar.</param>
  ''' <param name="text">The new text for the button.</param>
  Sub SetButtonText(position As ButtonPosition, index As Integer, text As String)

  ''' <summary>Makes the expert button visible or not visible.</summary>
  ''' <param name="visible">Whether the expert button is visible.</param>
  Sub SetExpertButtonVisible(visible As Boolean)

  ''' <summary>Set a non-standard status view to show at the top of each standard window.</summary>
  ''' <param name="value">The control to use as a status view.</param>
  ''' <remarks>Only works if there is a single control system.</remarks>
  Sub SetStatusView(value As Object)

  ''' <summary>Executes an SQL statement against the control system database and returns data in a <see cref="DataTable"/>.</summary>
  ''' <param name="sql">The SQL text.</param>
  ''' <returns>the data.</returns>
  Function DbGetDataTable(sql As String) As DataTable

  ''' <summary>Executes an SQL statement against the control system database.</summary>
  ''' <param name="sql">The SQL text.</param>
  ''' <param name="parameterNamesAndValues">Optional parameter names and values for the SQL statement.</param>
  ''' <returns>the number of rows affected.</returns>
  Function DbExecute(sql As String, ParamArray parameterNamesAndValues() As Object) As Integer

  ''' <summary>Makes a <see cref="String"/> suitable for use in an SQL statement from the given value.</summary>
  ''' <param name="value">The value to convert to a <see cref="String"/>.</param>
  ''' <returns>a <see cref="String"/> suitable for use in an SQL statement.</returns>
  Function DbSqlString(value As Object) As String

  ''' <summary>Updates the schema of the control system database with the given SQL schema text.</summary>
  ''' <param name="sql">The SQL schema text.</param>
  Sub DbUpdateSchema(sql As String)

  ''' <summary>Starts a job running.</summary>
  ''' <param name="jobName">The job.</param>
  ''' <param name="programNumbers">The program numbers.</param>
  ''' <param name="substituteSteps">The steps to substitute.</param>
  Sub StartJob(jobName As String, dataId As Object, programNumbers As IEnumerable(Of Integer),
               Optional substituteSteps As IEnumerable(Of String) = Nothing)

  ''' <summary>Starts a job running.</summary>
  ''' <param name="jobName">The job.</param>
  ''' <param name="steps">The steps to run.</param>
  Sub StartJob(jobName As String, dataId As Object, ParamArray steps() As String)

  ''' <summary>Stops a job.</summary>
  Sub StopJob()

  ''' <summary>Prevents standard database schema activities.</summary>
  Sub NonStandardDb()

  ''' <summary>Creates a standard view.</summary>
  ''' <param name="typeName">One of "History", "Variables", "IO" or "Parameters"</param>
  ''' <returns>the created view.</returns>
  Function CreateView(typeName As String) As Object

  ''' <summary>Creates a history from serialized data.</summary>
  ''' <param name="bytes">The bytes from which to create the history</param>
  ''' <returns>an instance of the History object.</returns>
  Function CreateHistory(bytes() As Byte) As Object

  ''' <summary>Creates a proxy to a object of type 'ControlCode' found at the given data source.</summary>
  ''' <param name="dataSource">The data source.</param>
  ''' <returns>the proxy object.</returns>
  Function CreateControlCodeProxy(dataSource As Object) As Object

  ''' <summary>Creates a proxy to a object of given type found at the given data source.</summary>
  ''' <param name="classToProxy">The type to proxy.</param>
  ''' <param name="dataSource">The data source.</param>
  ''' <returns>the proxy object.</returns>
  Function CreateProxy(classToProxy As Type, dataSource As Object) As Object

  ''' <summary>Gets the live history.</summary>
  ''' <returns>the live history.</returns>
  ReadOnly Property History() As Object

  ''' <summary>Creates an instance of <see cref="RemoteValues"/> that can read and write the variables in a control system.</summary>
  ''' <param name="connect">The connect path to the control system.</param>
  ''' <param name="autoRefreshInterval">How often to read the property values, measured in ms.</param>
  ''' <param name="properties">The names of properties to automatically read from the control system.</param>
  ''' <returns>an instance of <see cref="RemoteValues"/>.</returns>
  Function CreateRemoteValues(connect As String, autoRefreshInterval As Integer, ParamArray properties() As String) As RemoteValues

  ''' <summary>Logs an event to the control system log file.</summary>
  ''' <param name="eventType">The type of the event.</param>
  ''' <param name="id">The id of the event, usually an <see cref="[Enum]"/> value.</param>
  ''' <param name="message">The message.</param>
  Sub LogEvent(eventType As LogEventType, id As Object, message As String)

  ''' <summary>Logs an exception to the control system log file.</summary>
  ''' <param name="ex">The exception.</param>
  Sub LogException(ex As Exception)

  ''' <summary>Gets the value of a setting (from the control system setup file).</summary>
  ''' <param name="name">The name of the setting.</param>
  ''' <returns>the value of the setting, or an empty string if the <paramref name="name"/> is not found.</returns>
  ReadOnly Property Setting(name As String) As String

  ''' <summary>Sets the channel number of an IO variable to a different value.  Use 0 to hide the variable.</summary>
  ''' <param name="name">The name of the IO variable.</param>
  ''' <param name="channel">The new channel number, or 0 to hide the variable.</param>
  Sub SetIOChannel(name As String, channel As Integer)


  ''' <summary>Gets the number of io items.</summary>
  Function GetIOCount() As Integer

  ''' <summary>Gets information about a particular io item.</summary>
  ''' <param name="index">The index of the item, from 0 to GetIOCount() - 1.</param>
  ''' <param name="name">Returns the name.</param>
  ''' <param name="ioType">Returns the IO type.</param>
  ''' <param name="channel">Returns the channel number.</param>
  ''' <param name="allowOverride">Returns True if the io can be overridden.</param>
  ''' <param name="device">Returns the device.</param>
  ''' <param name="format">Returns the display format.</param>
  Sub GetIOItem(index As Integer, ByRef name As String, ByRef ioType As IOType, ByRef channel As Integer, ByRef allowOverride As Override, ByRef device As String, ByRef format As String)

  ''' <summary>Gets the name of the culture (language) for which the control system is displaying text.</summary>
  ''' <returns>the name of the culture (language) for which the control system is displaying text.</returns>
  ReadOnly Property CultureName() As String

  ''' <summary>Sets the name of the culture (language) for which the control system is displaying text.</summary>
  Sub SetCultureName(value As String)

  ''' <summary>Gets the data found in the attached dongle, if any.</summary>
  ''' <returns>the data found in the attached dongle, or null if no dongle is in use.</returns>
  ReadOnly Property DongleBytes() As Byte()

  ''' <summary>Sets the visibility of the operator toolbar at the bottom.</summary>
  Sub SetOperatorToolbarVisible(value As Boolean)

  Property Forces As String

  ''' <summary>Sets a cancel handler to perhaps set a non-standard action when a ToolStripButton is clicked.</summary>
  ''' <param name="handler">The handler routine.</param>
  Sub SetToolStripButtonClickHandler(handler As CancelEventHandler)

  Function AddSequencerToNetwork(name As String, component As Object) As Object

  Sub SetBlockAll(value As Boolean)


  ''' <summary>Returns whether some UI functions are locked out at the request of Plant Explorer.</summary>
  ReadOnly Property Locked As Boolean

  ReadOnly Property ProgramStepCount As Integer
  ReadOnly Property ProgramStep(ofs As Integer) As ProgramStep

  Sub SendToFlowBook(tags As String, subject As String, body As String, priority As Net.Mail.MailPriority)
  ReadOnly Property JobAsHtml As String

  ''' <summary>Gets the sequencer for this control system.</summary>
  ''' <returns>the sequencer for this control system.</returns>
  ReadOnly Property ControlSequencer() As Object

  ''' <summary>If the value of any of the machine parameters is changed, then returns True one. 
  ''' Also returns True once at startup.</summary>
  ReadOnly Property ParametersChanged As Boolean

  ''' <summary>If Yes is pressed, returns True once.</summary>
  ReadOnly Property YesPressed As Boolean
  ''' <summary>If No is pressed, returns True once.</summary>
  ReadOnly Property NoPressed As Boolean

  ''' <summary>Sets whether values for a variable of the given name will be logged in histories.</summary>
  Sub SetLogHistory(name As String, log As Boolean)

  ''' <summary>Sets how often data is logged to a history - use TimeSpan.MinValue for continuous logging.</summary>
  Sub SetHistoryLogInterval(value As TimeSpan)

  ReadOnly Property UnableToStart As Boolean

  ''' <summary>Gets the current temperature of the CPU in tenths of a degree C, so 354 is 35.4C</summary>
  ReadOnly Property CpuTemperature() As Integer

  ''' <summary>Gets whether the expert menu is visible.</summary>
  ReadOnly Property ExpertMode As Boolean

  ''' <summary>Makes the expert menu visible or not visible.</summary>
  ''' <param name="value">Whether the expert menu is visible.</param>
  Sub SetExpertMode(value As Boolean)

  ''' <summary>Sets whether the system is sleeping.</summary>
  Sub SetSleep(value As Boolean)

End Interface

Public Interface ProgramStep
  ReadOnly Property Command As String
  Property Notes As String
  ReadOnly Property ParameterCount As Integer
  Property Parameter(ofs As Integer) As String
  ReadOnly Property TotalTime As TimeSpan
  ReadOnly Property Elapsed As TimeSpan
  ReadOnly Property ExtraTime As TimeSpan
  ReadOnly Property TemperatureAfter As Integer
  ReadOnly Property IpLevel As Integer
  ReadOnly Property ProgramNumber As String
  ReadOnly Property StepNumber As Integer
  ReadOnly Property TimeInStep() As TimeSpan
End Interface


Public Interface RemoteValues
  ''' <summary>Sets values for properties of given names.</summary>
  ''' <example>
  ''' <code>SetValues("Var1", 46, "Var2", "Abc")</code>
  ''' </example>
  ''' <param name="namesAndValues">Pairs of names and values, separated by commas.</param>
  ''' <remarks>Performed asynchronously, so it is not possible to tell whether the method succeeded.</remarks>
  Sub SetValues(ParamArray namesAndValues() As Object)

  ''' <summary>Gets the latest values of a property.</summary>
  ''' <param name="index">The zero-based index of the property, based on the order the properties were named in the call to <see cref="ACParent.CreateRemoteValues"></see>.</param>
  ''' <returns>the latest values of a property.</returns>
  Default ReadOnly Property Value(index As Integer) As Object

  ''' <summary>Gets whether the last refresh of the values went ok.</summary>
  ''' <returns>true if the last refresh of the values went ok; otherwise, false.</returns>
  ''' <remarks>As long as the values have been refreshed successfully at least once, then it is ok to call <see cref="Value">.</see></remarks>
  ReadOnly Property LastOK() As Boolean
End Interface
#End If

#If ADAPTIVECONTROLNAMESPACE Or TARGET <> "library" Then
End Namespace
#End If

