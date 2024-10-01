Public Class PhoenixTrio

  Public SerialPort As String
  Public Comms As Ports.Modbus
  Public CommsTimer As New Timer

  Private ComsFaulted As Boolean
  Private ComsFaultTimer As New Timer
  '  Public FaultCode As YaskawaFault
  Public FaultRead, FaultWrite As Integer
  Public WriteOk As Integer

  ' Register 0020H   
  Public StatusWord As Short
  Public StatusOK As Boolean           'bit 0
  Public StatusAlarm As Boolean        'bit 1
  Public StatusBatMode As Boolean      'bit 2
  Public StatusReady As Boolean        'bit 3
  Public StatusRemote As Boolean       'bit 4
  Public StatusBatStart As Boolean     'bit 5
  Public StatusGreenLed As Boolean     'bit 12
  Public StatusYelloLed As Boolean     'bit 13
  Public StatusRedLed As Boolean       'bit 14


  'Public OutputVoltage As Short         '10mV
  'Public OutputCurrent As Short         '10mA
  'Public BatteryVoltage As Short        '10mV
  'Public ChargeCurrent As Short         '10mA
  'Public DeviceTemp As Short            'K


  Public Sub New(ByVal port As String)
    'Setup Communication Ports 
    '>> Use this method so that if the xml file doesn't declare the port, no "is nothing" exceptions will be thrown <<
    If Not String.IsNullOrEmpty(port) Then
      Try
        Me.SerialPort = port

        ' Old way, single read then divide out elements
        '  Comms = New Ports.Modbus(New Ports.ModbusTcp(address, 502))  'TCP/IP
        Comms = New Ports.Modbus(New Ports.SerialPort(port, 115200, System.IO.Ports.Parity.Even, 8, System.IO.Ports.StopBits.One).BaseStream)

        'Reset Communication Port Alarm Timeout
        CommsTimer.Seconds = 10

      Catch ex As Exception
      End Try
    End If
  End Sub

  Function ReadTrio() As Boolean
    Dim readSuccess As Boolean

    'Test that Comms has been succesfully created (address was set)
    If Comms IsNot Nothing Then

      ' 16 words starting at 0020H = 32 Dec 
      Dim valueSet0020H(1) As Short
      Select Case Comms.Read(192, 30001 + 8194, valueSet0020H)
        Case Ports.Modbus.Result.OK
          StatusWord = valueSet0020H(1)               ' 0020H = Drive Status 1

          ' DECODE STATUS WORD
          Dim bitArrayWord1 As New BitArray(BitConverter.GetBytes(StatusWord))
          StatusOK = bitArrayWord1(0)
          StatusAlarm = bitArrayWord1(1)
          StatusBatMode = bitArrayWord1(2)
          StatusReady = bitArrayWord1(3)
          StatusRemote = bitArrayWord1(4)
          StatusBatStart = bitArrayWord1(5)
          StatusGreenLed = bitArrayWord1(12)
          StatusYelloLed = bitArrayWord1(13)
          StatusRedLed = bitArrayWord1(14)

          ComsFaultTimer.Seconds = 5

        Case Ports.Modbus.Result.Busy
        Case Ports.Modbus.Result.HwFault
        Case Ports.Modbus.Result.Fault

      End Select

#If 0 Then

      ' 0x2006 = 8198
      Static valueSetAnalogValues(6) As Short
      Select Case Comms.Read(192, 40001 + 8198, valueSetAnalogValues)
        Case Ports.Modbus.Result.OK
          ' Live Values
          OutputVoltage = valueSetAnalogValues(1)       '10mV
          OutputCurrent = valueSetAnalogValues(2)         '10mA
          BatteryVoltage = valueSetAnalogValues(3)       '10mV
          ChargeCurrent = valueSetAnalogValues(4)        '10mA
          DeviceTemp = valueSetAnalogValues(5)        'K

        Case Ports.Modbus.Result.Busy
        Case Ports.Modbus.Result.HwFault
        Case Ports.Modbus.Result.Fault
      End Select


#End If

      ' Timeout
      ComsFaulted = ComsFaultTimer.Finished

    End If
    Return readSuccess
  End Function

  ' TODO - Add Function for TRIO (Below from Yaskawa)
  'Function GetAlarmMessage(code As Integer) As String

  '  Dim resultHex = code.ToString("X")
  '  Select Case resultHex
  '    Case "" : Return Nothing

  '    Case "1" : Return "Undervoltage (Uv)"
  '    Case "2" : Return "Drive Overheat Warning (ov)"

  '    Case "3B" : Return "Safe Disable Input (HbbF)"
  '    Case "3C" : Return "Safe Disable Input (Hbb)"

  '  End Select
  '  Return Nothing
  'End Function

End Class
