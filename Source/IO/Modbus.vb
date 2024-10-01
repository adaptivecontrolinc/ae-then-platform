'  Batch Control Up To 3.2.155
'  [2013-04-10] - Update with David Tindale for working with Single arrays on read agains Automation-DIrect DL206 REAL v-memory
'  [2012-11-07]

' Modbus protocol
Namespace Ports
  Public Class Modbus
    Private Enum State
      Idle
      TxRxBusy
      TxRxComplete
      Tx
      Rx
      Complete
    End Enum

    Public Enum Result
      Busy
      OK
      Fault
      HwFault
    End Enum

    Private ReadOnly stm_ As Stream
    Private rx_() As Byte, state_ As State, result_ As Result, waitFailTime_ As Date
    Private red_ As Integer
    Private asyncResult_ As IAsyncResult
    Private oks_, faults_, hwFaults_ As Integer
    Private Structure WorkData : Dim FirstRegister, SlaveAddress As Integer, IsRead As Boolean, Func As Func : End Structure
    Private work_ As WorkData
    Private ReadOnly writeOptimisation_ As New WriteOptimisation

    Private Enum Func As Byte
      ReadOutputTable = 1       ' read DOUTs  00001 
      ReadInputTable            ' read DINPs  10001
      ReadHoldingRegisters      ' read V's
      ReadInputRegisters        ' read ANINPs 30001
      ForceSingleOutput
      PresetSingleRegister
      ForceMultipleOutputs = 15
      PresetMultipleRegisters
    End Enum

    Public Sub New(ByVal stm As System.IO.Stream)
      stm.ReadTimeout = 200  ' should allow for greatest possible time to receive data, for example 100 bytes @ 19200 -> 50ms plus some delay
      stm_ = New Stream(stm)
    End Sub

    Friend ReadOnly Property BaseStream() As Stream
      Get
        Return stm_
      End Get
    End Property

    Friend ReadOnly Property OKs() As Integer
      Get
        Return oks_
      End Get
    End Property
    Public ReadOnly Property Faults() As Integer
      Get
        Return faults_
      End Get
    End Property
    Public ReadOnly Property HwFaults() As Integer
      Get
        Return hwFaults_
      End Get
    End Property
    Private Sub SetResult(ByVal value As Result)
      result_ = value
      Select Case value
        Case Result.OK : oks_ += 1
        Case Result.Fault : faults_ += 1
        Case Else : hwFaults_ += 1
      End Select
    End Sub

    Private Shared Function Hash(ByVal array() As Byte, ByVal ibStart As Integer, ByVal cbSize As Integer) As Byte()
      Dim crc As Integer = &HFFFF  ' do it in 32 bits instead of 16 bit unsigned
      Do While cbSize > 0
        crc = crc Xor array(ibStart)
        For shiftCount As Integer = 0 To 7
          If (crc And 1) <> 0 Then
            crc = (crc >> 1) Xor &HA001
          Else
            crc >>= 1
          End If
        Next shiftCount
        ibStart += 1 : cbSize -= 1
      Loop
      Dim all4Bytes() As Byte = System.BitConverter.GetBytes(crc)
      Return New Byte() {all4Bytes(0), all4Bytes(1)}
    End Function

    Protected Sub BeginWriteAndRead(ByVal tx() As Byte, ByVal rxCount As Integer)
      Dim txCount As Integer = tx.Length

      ' The CRC should go in the last 2 bytes
      Dim crcBytes() As Byte = Hash(tx, 0, txCount - 2)
      tx(txCount - 2) = crcBytes(0)
      tx(txCount - 1) = crcBytes(1)

      ' Make a suitable rx_
      rx_ = New Byte(rxCount - 1) {}

      ' Begin the write of the bytes
      If stm_.Stream.CanSeek Then  ' CanSeek for us means callback async's work
        state_ = State.TxRxBusy : waitFailTime_ = Date.UtcNow.AddSeconds(5)
        stm_.Flush()
        stm_.BeginWrite(tx, 0, tx.Length, TxComplete_, Nothing)
      Else
        state_ = State.Tx
        stm_.Flush()
        asyncResult_ = stm_.BeginWrite(tx, 0, tx.Length, Nothing, Nothing)
      End If
    End Sub

    Private ReadOnly TxComplete_ As New AsyncCallback(AddressOf TxComplete)
    Private Sub TxComplete(ByVal ar As IAsyncResult)
      stm_.EndWrite(ar)  ' tidy up
      ' Start reading, feeding in the rx timeouts each time
      stm_.BeginRead(rx_, 0, rx_.Length, RxComplete_, Nothing)
    End Sub

    Private ReadOnly RxComplete_ As New AsyncCallback(AddressOf RxComplete)
    Private Sub RxComplete(ByVal ar As IAsyncResult)
      red_ = stm_.EndRead(ar)
      state_ = State.TxRxComplete
    End Sub

    Protected Sub RunStateMachine(ByVal slaveAddress As Integer, ByVal firstRegister As Integer, ByVal isRead As Boolean)
      Dim red As Integer
      Select Case state_
        Case State.TxRxBusy
          If Date.UtcNow >= waitFailTime_ Then
            stm_.Stream.Close()
            red = -1
            GoTo txRxComplete
          End If

        Case State.TxRxComplete
          red = red_
          GoTo txRxComplete

        Case State.Tx
          ' Wait for the tx to complete
          If Not asyncResult_.IsCompleted Then Exit Sub
          stm_.EndWrite(asyncResult_)  ' tidy up

          ' Start reading, feeding in the rx timeouts each time
          asyncResult_ = stm_.BeginRead(rx_, 0, rx_.Length, Nothing, Nothing)
          state_ = State.Rx : GoTo stateRx ' go straight on to the next state

        Case State.Rx
stateRx:
          If Not asyncResult_.IsCompleted Then Exit Sub ' it'll be completed soon
          red = stm_.EndRead(asyncResult_) : asyncResult_ = Nothing

txRxComplete:
          Dim rxCount As Integer = rx_.Length
          If red = -1 Then
            SetResult(Result.HwFault)
          ElseIf red <> rxCount Then        ' not enough bytes ?
            SetResult(Result.Fault)
          Else
            ' Check the incoming CRC as well - unless there are two zeroes in the data.  This is
            ' just a lazy exception for modbus tcp/ip
            Dim crc0 As Byte = rx_(rxCount - 2), crc1 As Byte = rx_(rxCount - 1)
            If crc0 = 0 AndAlso crc1 = 0 Then
              SetResult(Result.OK)
            Else
              Dim crcBytes() As Byte = Hash(rx_, 0, rxCount - 2)
              If crcBytes(0) <> crc0 OrElse crcBytes(1) <> crc1 Then
                SetResult(Result.Fault)
              Else
                SetResult(Result.OK)
              End If
            End If
          End If
          state_ = State.Complete : GoTo stateComplete ' go straight on to the next state

          ' We need these states because we may have to sit here in these states
          ' waiting for the correct person to pick up the data.
        Case State.Complete
stateComplete:
          ' Maybe it's someone else's job - TODO: timeout if no-one comes for it for a long time
          If work_.SlaveAddress <> slaveAddress OrElse work_.FirstRegister <> firstRegister _
              OrElse work_.IsRead <> isRead Then Exit Sub
          state_ = State.Idle ' the end of this job
      End Select
    End Sub

    Private Shared Function GetBitCount(ByVal typ As Type) As Integer
      If typ Is GetType(Boolean) Then Return 1
      If typ Is GetType(Int16) OrElse typ Is GetType(UInt16) Then Return 16
      If typ Is GetType(Int32) OrElse typ Is GetType(UInt32) OrElse typ Is GetType(Single) Then Return 32
    End Function

#If 0 Then
    ' TODO: don't always prep up for a Read or Write - this could save a lot of time !
    ' But this is not quite it
    ' The saving will be mostly before calls to Write()
    Public ReadOnly Property Idle As Boolean
      Get
        Return state_ = State.Idle
      End Get
    End Property
#End If

    ' We always ignore element 0 of the 'values' array - this is to make it easier for the engineer writing
    ' the control-code who calls this.
    Public Function Read(ByVal slaveAddress As Integer, ByVal firstRegister As Integer, ByVal values As Array) As Result
      ' In modbus, the high bytes comes first

      ' One less because we ignore element 0
      Dim bitCount As Integer = GetBitCount(values.GetType.GetElementType),
          count As Integer = values.Length - 1,
          totalBits As Integer = count * bitCount,
          shortCount As Integer = totalBits \ 16 : If totalBits Mod 16 <> 0 Then shortCount += 1

      ' Start a completely new task
      If state_ = State.Idle Then
        work_.SlaveAddress = slaveAddress : work_.FirstRegister = firstRegister : work_.IsRead = True
        ' ReadOutputTable
        If firstRegister <= 9999 Then
          If bitCount <> 1 Then Throw New ArgumentOutOfRangeException
          Dim tx(8 - 1) As Byte : tx(0) = CType(slaveAddress, Byte)
          work_.Func = Func.ReadOutputTable : tx(1) = work_.Func
          BitConverterLittleEndian.GetBytes(CType(firstRegister - 1, Short)).CopyTo(tx, 2)
          BitConverterLittleEndian.GetBytes(CType(count, Short)).CopyTo(tx, 4)
          Dim byteCount As Integer = (count + 7) \ 8
          BeginWriteAndRead(tx, 5 + byteCount)

          ' ReadInputTable
        ElseIf firstRegister >= 10001 AndAlso firstRegister <= 19999 Then
          If bitCount <> 1 Then Throw New ArgumentOutOfRangeException
          Dim tx(8 - 1) As Byte : tx(0) = CType(slaveAddress, Byte)
          work_.Func = Func.ReadInputTable : tx(1) = work_.Func
          BitConverterLittleEndian.GetBytes(CType(firstRegister - 10001, Short)).CopyTo(tx, 2)
          BitConverterLittleEndian.GetBytes(CType(count, Short)).CopyTo(tx, 4)
          Dim byteCount As Integer = (count + 7) \ 8
          BeginWriteAndRead(tx, 5 + byteCount)

          ' ReadInputRegisters
        ElseIf firstRegister >= 30001 AndAlso firstRegister <= 39999 Then
          Dim tx(8 - 1) As Byte : tx(0) = CType(slaveAddress, Byte)
          work_.Func = Func.ReadInputRegisters : tx(1) = work_.Func
          BitConverterLittleEndian.GetBytes(CType(firstRegister - 30001, Short)).CopyTo(tx, 2)
          BitConverterLittleEndian.GetBytes(CType(shortCount, Short)).CopyTo(tx, 4)
          BeginWriteAndRead(tx, 5 + 2 * shortCount)

          ' ReadHoldingRegisters
        ElseIf firstRegister >= 40001 Then  ' read holding registers
          Dim tx(8 - 1) As Byte : tx(0) = CType(slaveAddress, Byte)
          work_.Func = Func.ReadHoldingRegisters : tx(1) = work_.Func
          BitConverterLittleEndian.GetBytes(CType(firstRegister - 40001, Short)).CopyTo(tx, 2)
          BitConverterLittleEndian.GetBytes(CType(shortCount, Short)).CopyTo(tx, 4)
          BeginWriteAndRead(tx, 5 + 2 * shortCount)
        Else
          Throw New ArgumentOutOfRangeException("firstRegister")
        End If
      End If


      ' See if we're finished
      RunStateMachine(slaveAddress, firstRegister, True)
      If state_ <> State.Idle Then Return Result.Busy ' not yet 

      If result_ = Result.OK Then
        ' Ok, store these new values - we may have to unscramble the hi-lo order
        Select Case work_.Func
          ' Data here is in single bytes
          Case Func.ReadOutputTable, Func.ReadInputTable
            For i As Integer = 0 To count - 1
              values.SetValue((rx_(3 + i \ 8) And (1 << (i And 7))) <> 0, i + 1)
            Next i

            ' Data is returned in exactly the same way for both of these
          Case Func.ReadInputRegisters, Func.ReadHoldingRegisters
            ' Convert the words back round to low byte first
            Dim uShortArray() As UShort = TryCast(values, UShort())
            If uShortArray IsNot Nothing Then
              For i As Integer = 0 To count - 1
                uShortArray(i + 1) = BitConverterLittleEndian.ToUInt16(rx_, 3 + 2 * i)
              Next i
            Else
              Dim singleArray() As Single = TryCast(values, Single())
              If singleArray IsNot Nothing Then
                For i As Integer = 0 To count - 1
                  singleArray(i + 1) = BitConverterLittleEndian.ToSingle(rx_, 3 + 4 * i)
                Next i
              Else
                Dim data(shortCount - 1) As Short
                For i As Integer = 0 To shortCount - 1
                  data(i) = BitConverterLittleEndian.ToInt16(rx_, 3 + 2 * i)
                Next i

                Dim booleanArray() As Boolean = TryCast(values, Boolean())
                If booleanArray IsNot Nothing Then
                  For i As Integer = 0 To count - 1
                    booleanArray(i + 1) = (data(i \ 16) And (1 << (i And 15))) <> 0
                  Next i
                Else
                  Dim shortArray() As Short = DirectCast(values, Short())
                  For i As Integer = 0 To count - 1
                    shortArray(i + 1) = data(i)
                  Next i
                End If
              End If
            End If
        End Select
      End If
      Return result_
    End Function

    Public Function Write(ByVal slaveAddress As Integer, ByVal firstRegister As Integer, ByVal values As Array, ByVal writeMode As WriteMode) As Result
      If state_ = State.Idle Then
        ' Start a completely new task

        ' Optionally, do write-optimisation, meaning we usually do not write the same values to the same
        ' registers in the same slave.
        If writeMode = Ports.WriteMode.Optimised AndAlso
           writeOptimisation_.RecentlyWritten(values, slaveAddress, firstRegister) Then Return Result.OK

        work_.SlaveAddress = slaveAddress : work_.FirstRegister = firstRegister : work_.IsRead = False
        Dim bitCount As Integer = GetBitCount(values.GetType.GetElementType),
            count As Integer = values.Length - 1,
            totalBits As Integer = count * bitCount,
            shortCount As Integer = totalBits \ 16 : If totalBits Mod 16 <> 0 Then shortCount += 1

        ' ForceMultipleOutputs
        If firstRegister >= 10001 AndAlso firstRegister <= 19999 Then
          If bitCount <> 1 Then Throw New NotSupportedException
          Dim byteCount As Integer = count \ 8 : If count Mod 8 <> 0 Then byteCount += 1
          Dim tx(9 + byteCount - 1) As Byte
          tx(0) = CType(slaveAddress, Byte)
          work_.Func = Func.ForceMultipleOutputs : tx(1) = work_.Func
          Dim startAddressBytes() As Byte = BitConverterLittleEndian.GetBytes(CType(firstRegister - 10001, Short)),
              pointsBytes() As Byte = BitConverterLittleEndian.GetBytes(CType(count, Short))
          tx(2) = startAddressBytes(0) : tx(3) = startAddressBytes(1)
          tx(4) = pointsBytes(0) : tx(5) = pointsBytes(1)
          tx(6) = CType(byteCount, Byte)

          Dim boolArray() As Boolean = DirectCast(values, Boolean())
          For i As Integer = 0 To count - 1
            If boolArray(i + 1) Then
              tx(7 + (i \ 8)) = CType(tx(7 + (i \ 8)) Or (1 << (i And 7)), Byte)
            End If
          Next i
          BeginWriteAndRead(tx, 8)

          ' PresetSingleRegister or PresetMultipleRegisters
        ElseIf firstRegister >= 40001 Then
          ' Make the request
          Dim tx() As Byte
          Dim fn As Func
          If shortCount = 1 Then
            fn = Func.PresetSingleRegister
            tx = New Byte(8 - 1) {}
          Else
            fn = Func.PresetMultipleRegisters
            tx = New Byte(7 + 2 * shortCount + 2 - 1) {}
          End If

          tx(0) = CType(slaveAddress, Byte)
          work_.Func = fn : tx(1) = work_.Func
          Dim startAddressBytes() As Byte = BitConverterLittleEndian.GetBytes(CType(firstRegister - 40001, Short))
          tx(2) = startAddressBytes(0) : tx(3) = startAddressBytes(1)

          If shortCount > 1 Then
            Dim pointsBytes() As Byte = BitConverterLittleEndian.GetBytes(CType(shortCount, Short))
            tx(4) = pointsBytes(0) : tx(5) = pointsBytes(1)
            tx(6) = CType(2 * shortCount, Byte)
          End If

          Dim data(shortCount * 2 - 1) As Byte
          Dim uShortArray() As UShort = TryCast(values, UShort())
          If uShortArray IsNot Nothing Then
            For i As Integer = 0 To count - 1
              BitConverterLittleEndian.GetBytes(uShortArray(i + 1)).CopyTo(data, i * 2)
            Next i
          Else
            Dim singleArray() As Single = TryCast(values, Single())
            If singleArray IsNot Nothing Then
              For i As Integer = 0 To count - 1
                BitConverterLittleEndian.GetBytes(singleArray(i + 1)).CopyTo(data, i * 4)
              Next i
            Else
              Dim booleanArray() As Boolean = TryCast(values, Boolean())
              If booleanArray IsNot Nothing Then
                Dim i As Integer
                Do While i * 16 < count
                  Dim startIndex As Integer = i * 16,
                      encodeBitCount As Integer = Math.Min(count - startIndex, 16)
                  BitConverterLittleEndian.GetBytes(BitEncoding.ToInt16(booleanArray, startIndex + 1, encodeBitCount)).CopyTo(data, i * 2)
                  i += 1
                Loop
              Else
                Dim shortArray() As Short = DirectCast(values, Short())
                For i As Integer = 0 To count - 1
                  BitConverterLittleEndian.GetBytes(shortArray(i + 1)).CopyTo(data, i * 2)
                Next i
              End If
            End If
          End If

          ' Put in the data and begin the write
          If shortCount = 1 Then
            data.CopyTo(tx, 4)
          Else
            data.CopyTo(tx, 7)
          End If
          BeginWriteAndRead(tx, 8)
        Else
          Throw New ArgumentOutOfRangeException("firstRegister")
        End If
      End If

      ' See if we're finished
      RunStateMachine(slaveAddress, firstRegister, False)
      If state_ <> State.Idle Then Return Result.Busy ' not yet 

      If result_ = Result.OK AndAlso writeMode = Ports.WriteMode.Optimised Then
        writeOptimisation_.SuccessfulWrite(values, slaveAddress, firstRegister)
      End If
      Return result_
    End Function

    Private NotInheritable Class BitConverterLittleEndian
      Private Sub New()
      End Sub
      Public Shared Function GetBytes(ByVal value As Int16) As Byte()
        Dim ret() As Byte = System.BitConverter.GetBytes(value)
        Dim b As Byte = ret(0) : ret(0) = ret(1) : ret(1) = b
        Return ret
      End Function
      Public Shared Function GetBytes(ByVal value As UInt16) As Byte()
        Dim ret() As Byte = System.BitConverter.GetBytes(value)
        Dim b As Byte = ret(0) : ret(0) = ret(1) : ret(1) = b
        Return ret
      End Function
      Public Shared Function GetBytes(ByVal value As Int32) As Byte()
        Dim ret() As Byte = System.BitConverter.GetBytes(value)
        Return New Byte() {ret(3), ret(2), ret(1), ret(0)}
      End Function
      Public Shared Function GetBytes(ByVal value As Single) As Byte()
        Dim ret() As Byte = System.BitConverter.GetBytes(value)
        Return New Byte() {ret(3), ret(2), ret(1), ret(0)}
      End Function

      Public Shared Function ToInt16(ByVal value() As Byte, ByVal startIndex As Integer) As Short
        Return System.BitConverter.ToInt16(New Byte() {value(startIndex + 1), value(startIndex)}, 0)
      End Function
      Public Shared Function ToUInt16(ByVal value() As Byte, ByVal startIndex As Integer) As UShort
        Return System.BitConverter.ToUInt16(New Byte() {value(startIndex + 1), value(startIndex)}, 0)
      End Function
      Public Shared Function ToInt32(ByVal value() As Byte, ByVal startIndex As Integer) As Integer
        Return System.BitConverter.ToInt32(New Byte() {value(startIndex + 3), value(startIndex + 2), value(startIndex + 1), value(startIndex)}, 0)
      End Function
      Public Shared Function ToSingle(ByVal value() As Byte, ByVal startIndex As Integer) As Single
        ' [2013-04-10] updated with David Tindale
        ' Some old, un-identified plc had the bytes this way round
        '    Return System.BitConverter.ToSingle(New Byte() {value(startIndex + 3), value(startIndex + 2), value(startIndex + 1), value(startIndex)}, 0)
        ' but for Automation-Direct DL205 PLC, reading V-memory as REAL value (2 16-bit words, 32-bit floating value)
        ' we need this instead...
        Return System.BitConverter.ToSingle(New Byte() {value(startIndex + 1), value(startIndex + 0), value(startIndex + 3), value(startIndex + 2)}, 0)
      End Function
    End Class

    Private NotInheritable Class BitEncoding
      Private Sub New()
      End Sub

      Public Shared Function ToByte(ByVal value() As Boolean, ByVal startIndex As Integer, ByVal length As Integer) As Byte
        Dim ret As Integer
        For i As Integer = length - 1 To 0 Step -1
          ret *= 2
          If value(startIndex + i) Then ret += 1
        Next i
        Return CType(ret, Byte)
      End Function

      Public Shared Function ToInt16(ByVal value() As Boolean, ByVal startIndex As Integer, ByVal length As Integer) As Short
        Dim ret As Integer
        For i As Integer = length - 1 To 0 Step -1
          ret *= 2
          If value(startIndex + i) Then ret += 1
        Next i
        If ret < 32768 Then Return CType(ret, Short)
        Return CType(ret - 65536, Short)
      End Function
    End Class
  End Class

  ' -----------------------------------------------------------------
  ' Allow for the MBAP header, etc, when communicating using Modbus Tcp/Ip
  Public Class ModbusTcp : Inherits NetworkPort
    Private transactionId_ As Short, rxBuffer_(), rxOriginalBuffer_() As Byte, rxOriginalOffset_ As Integer

    Public Sub New(ByVal computer As String, ByVal port As Integer)
      MyBase.New(computer, port)
    End Sub

    Public Overrides Function BeginWrite(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, ByVal callback As AsyncCallback, ByVal state As Object) As IAsyncResult
      ' Add a MBAP header at the front and strip the CRC-16 off the end
      transactionId_ += 1S  ' will overflow back to 0

      Dim tx(6 + count - 2 - 1) As Byte
      Dim by() As Byte = System.BitConverter.GetBytes(transactionId_) : tx(0) = by(1) : tx(1) = by(0)
      by = System.BitConverter.GetBytes(count - 2) : tx(4) = by(1) : tx(5) = by(0)
      Array.Copy(buffer, offset, tx, 6, count - 2)
      Return MyBase.BeginWrite(tx, 0, tx.Length, callback, state)
    End Function

    Public Overrides Function BeginRead(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, _
                                        ByVal callback As AsyncCallback, ByVal state As Object) As IAsyncResult
      ' Assign a temporary bigger buffer to allow for the MBAP header
      rxBuffer_ = New Byte(6 + count - 2 - 1) {}
      rxOriginalBuffer_ = buffer : rxOriginalOffset_ = offset
      Return MyBase.BeginRead(rxBuffer_, 0, rxBuffer_.Length, callback, state)
    End Function

    Public Overrides Function EndRead(ByVal asyncResult As IAsyncResult) As Integer
      Dim ret As Integer = MyBase.EndRead(asyncResult)
      If ret = 0 Then Return 0
      Array.Copy(rxBuffer_, 6, rxOriginalBuffer_, rxOriginalOffset_, rxBuffer_.Length - 6)
      Return ret - 6 + 2
    End Function
  End Class
End Namespace
