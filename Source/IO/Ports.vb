'  Batch Control Up To 3.2.193
'  [2021-06-04]

Namespace Ports
#Const ASYNCREADANDWRITE = Not CF
  Public Class SerialPort : Implements IDisposable
    Private ReadOnly portName_ As String, baudRate_ As Integer, parity_ As System.IO.Ports.Parity, _
                     dataBits_ As Integer, stopBits_ As System.IO.Ports.StopBits, options_ As OptionsValue
    Private hCom_ As IntPtr = INVALID_HANDLE_VALUE
    Private Shared INVALID_HANDLE_VALUE As New IntPtr(-1) ' TODO: is this really const ?
    Private ReadOnly stream_ As New SerialStream(Me)
    Private readTotalTimeoutConstant_, readIntervalTimeout_, readTotalTimeoutMultiplier_ As Integer

    <Flags()> _
    Public Enum OptionsValue
      None = 0
      Sync = 1
    End Enum
    Public Sub New(ByVal portName As String, ByVal baudRate As Integer, ByVal parity As System.IO.Ports.Parity, _
                   ByVal dataBits As Integer, ByVal stopBits As System.IO.Ports.StopBits, Optional ByVal options As OptionsValue = OptionsValue.None)
#If CF Then
    If Not portName.EndsWith(":") Then portName &= ":" ' need a trailing colon in CE
#End If

      portName_ = portName : baudRate_ = baudRate : parity_ = parity : dataBits_ = dataBits : stopBits_ = stopBits : options_ = options
    End Sub

    Private Function TryToOpen() As Boolean
      If IsOpen Then Return True ' already open

      ' Open the port, etc
      Const GENERIC_READ As Integer = &H80000000, GENERIC_WRITE As Integer = &H40000000, _
            OPEN_EXISTING As Integer = 3

#If ASYNCREADANDWRITE Then
      ' TODO: actually, it seems that we can do overlapped calls to ReadFile and WriteFile even
      ' if we leave this as 0.  In this case, the I/O operations are serialized (i.e. one has
      ' to complete before the other can start, but that's fine for us anyway).
      ' This means we can also call ReadFile non-overlapped which makes our 
      ' implementation of Read() and Write() cleaner.
      Dim openFlags As Integer
      If (options_ And OptionsValue.Sync) = 0 Then openFlags = &H40000000 ' FILE_FLAG_OVERLAPPED
#Else
      Const openFlags = 0
#End If
      hCom_ = NativeMethods.CreateFile(PortName, GENERIC_READ Or GENERIC_WRITE, 0, IntPtr.Zero, OPEN_EXISTING, openFlags, IntPtr.Zero)
      If Not IsOpen Then Return False ' no exceptions: instead, we will try to pen whenever the port is attempted

      ' Set the timeouts
      SetTimeouts()

      ' Also set all the baud-rate stuff via a DCB
      ' The packing alignment needed is not currently possible with the CF, so 
      ' we have to resort to the BitConverter to get the job done
      Dim dcb(100 - 1) As Byte : NativeMethods.GetCommState(hCom_, dcb)
      System.BitConverter.GetBytes(BaudRate).CopyTo(dcb, 4)
      dcb(18) = CType(DataBits, Byte)
      dcb(19) = CType(Parity, Byte)
      Dim bStopBit As Byte
      Select Case StopBits
        Case System.IO.Ports.StopBits.Two : bStopBit = 2
        Case System.IO.Ports.StopBits.OnePointFive : bStopBit = 1
      End Select
      dcb(20) = bStopBit
      NativeMethods.SetCommState(hCom_, dcb)
      Return True
    End Function

    Private ReadOnly Property IsOpen() As Boolean
      Get
        Return Not hCom_.Equals(INVALID_HANDLE_VALUE)
      End Get
    End Property

    Private Sub Close()
      If IsOpen Then
        NativeMethods.CloseHandle(hCom_)
        hCom_ = INVALID_HANDLE_VALUE
      End If
    End Sub

    ' CF does not implement Component in the way we expect, so we do the stuff ourself here
    Protected Overrides Sub Finalize()
      Try
        Dispose(False)
      Finally
        MyBase.Finalize()
      End Try
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If disposing Then Close()
    End Sub

    Public Overrides Function ToString() As String
      Return portName_ & "," & baudRate_.ToString(Globalization.CultureInfo.InvariantCulture) & ","c & parity_.ToString.Substring(0, 1) & ","c & _
             dataBits_.ToString(Globalization.CultureInfo.InvariantCulture) & ","c & CType(stopBits_, Integer).ToString(Globalization.CultureInfo.InvariantCulture)
    End Function

    ''' <summary>Gets the port for communications, including but not limited to all available COM ports.</summary>
    Public ReadOnly Property PortName() As String
      Get
        Return portName_
      End Get
    End Property
    ''' <summary>Gets the serial baud rate.</summary>
    Public ReadOnly Property BaudRate() As Integer
      Get
        Return baudRate_
      End Get
    End Property
    ''' <summary>Gets the parity-checking protocol.</summary>
    Public ReadOnly Property Parity() As System.IO.Ports.Parity
      Get
        Return parity_
      End Get
    End Property
    ''' <summary>Gets the standard length of data bits per byte.</summary>
    Public ReadOnly Property DataBits() As Integer
      Get
        Return dataBits_
      End Get
    End Property
    ''' <summary>Gets the standard number of stopbits per byte.</summary>
    Public ReadOnly Property StopBits() As System.IO.Ports.StopBits
      Get
        Return stopBits_
      End Get
    End Property
    ''' <summary>Gets the underlying <see cref="System.IO.Stream"/>.</summary>
    Public ReadOnly Property BaseStream() As System.IO.Stream
      Get
        Return stream_
      End Get
    End Property

    ' This is handy so we can use a SerialPort anywhere where we need a Stream
    Public Shared Widening Operator CType(ByVal port As SerialPort) As System.IO.Stream
      Return port.stream_
    End Operator

    ''' <summary>Gets or sets a constant used to calculate the total time-out period for read operations, in milliseconds.</summary>
    ''' <remarks> For each read operation, this value is added to the product of the ReadTotalTimeoutMultiplier member and the requested number of bytes.</remarks>
    Public Property ReadTotalTimeoutConstant() As Integer
      Get
        Return readTotalTimeoutConstant_
      End Get
      Set(ByVal value As Integer)
        readTotalTimeoutConstant_ = value
        SetTimeouts()
      End Set
    End Property

    ''' <summary>Gets or sets the maximum time allowed to elapse between the arrival of two bytes on the communications line, in milliseconds.</summary>
    ''' <remarks> During a Read operation, the time period begins when the first byte is received.
    '''  If the interval between the arrival of any two bytes exceeds this amount, the Read operation is completed and any buffered data is returned.
    '''  A value of zero indicates that interval time-outs are not used.</remarks>
    Public Property ReadIntervalTimeout() As Integer
      Get
        Return readIntervalTimeout_
      End Get
      Set(ByVal value As Integer)
        ' A small value like 2 can give unexpected results, so make this a minimum of 10ms
        ' In our modbus driver, for example, we do not use this at all, as we are content to 
        ' rely on the big overall timeout since we always know what the return packet
        ' length should be.  For other protocols, like Beacon, we do not have this luxury,
        ' and would need to use this again.
        If value <> 0 AndAlso value < 10 Then value = 10
        readIntervalTimeout_ = value
        SetTimeouts()
      End Set
    End Property

    ''' <summary>Gets or sets the multiplier used to calculate the total time-out period for read operations, in milliseconds.</summary>
    ''' <remarks>For each read operation, this value is multiplied by the requested number of bytes to be read.</remarks>
    Public Property ReadTotalTimeoutMultiplier() As Integer
      Get
        Return readTotalTimeoutMultiplier_
      End Get
      Set(ByVal value As Integer)
        readTotalTimeoutMultiplier_ = value
        SetTimeouts()
      End Set
    End Property

    Private Sub SetTimeouts()
      If Not IsOpen Then Exit Sub
      Dim ct As New COMMTIMEOUTS
      ct.ReadIntervalTimeout = readIntervalTimeout_
      ct.ReadTotalTimeoutMultiplier = readTotalTimeoutMultiplier_
      ct.ReadTotalTimeoutConstant = readTotalTimeoutConstant_
      NativeMethods.SetCommTimeouts(hCom_, ct)
    End Sub

    <Flags()> _
    Public Enum PurgeActions
      TXAbort = 1   ' Kill the pending/current writes to the comm port.
      RXAbort = 2   ' Kill the pending/current reads to the comm port.
      TXClear = 4   ' Kill the transmit queue if there.
      RXClear = 8   ' Kill the typeahead buffer if there.
      All = TXAbort Or RXAbort Or TXClear Or RXClear
    End Enum

    ''' <summary>Purges data from buffers.</summary>
    Public Sub Purge(ByVal action As PurgeActions)
      NativeMethods.PurgeComm(hCom_, action)
    End Sub

    ''' <summary>Reads a number of bytes from the System.IO.Ports.SerialPort input buffer and writes those bytes into a byte array at the specified offset.</summary>
    Public Function Read(ByVal buffer As Byte(), ByVal offset As Integer, ByVal count As Integer) As Integer
      ' The implementation is actually inside the stream class
      Return stream_.Read(buffer, offset, count)
    End Function

    ''' <summary>Writes a specified number of bytes to the serial port using data from a buffer.</summary>
    Public Sub Write(ByVal buffer As Byte(), ByVal offset As Integer, ByVal count As Integer)
      stream_.Write(buffer, offset, count)
    End Sub

    ''' <summary>Sets a new parity.</summary>
    Public Sub OverrideParity(ByVal value As System.IO.Ports.Parity)
      Dim dcb(100 - 1) As Byte : NativeMethods.GetCommState(hCom_, dcb)
      dcb(19) = CType(value, Byte)
      NativeMethods.SetCommState(hCom_, dcb)
    End Sub

    ''' <summary>Writes the given bytes with mark parity, and then resets to space parity.</summary>
    Public Sub WriteWithMarkParity(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
      OverrideParity(System.IO.Ports.Parity.Mark)
      ' Send data
      stream_.Write(buffer, offset, count)
      OverrideParity(System.IO.Ports.Parity.Space)
    End Sub

    Public Function GetStatistics() As SerialPortStatistics
      Dim ret As New SerialPortStatistics
      ret.BytesIn = stream_.TotalIn : ret.BytesOut = stream_.TotalOut
      Return ret
    End Function

    ' -----------------------------------------------------------
    Public Class SerialStream : Inherits System.IO.Stream
      Private ReadOnly port_ As SerialPort
      Private totalIn_, totalOut_ As Integer

      Public Sub New(ByVal port As SerialPort)
        port_ = port
      End Sub

      Public ReadOnly Property Port() As SerialPort
        Get
          Return port_
        End Get
      End Property
      Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then port_.Dispose()
        MyBase.Dispose(disposing)
      End Sub

      Public ReadOnly Property TotalIn() As Integer
        Get
          Return totalIn_
        End Get
      End Property
      Public ReadOnly Property TotalOut() As Integer
        Get
          Return totalOut_
        End Get
      End Property

      Public Overrides Function ToString() As String
        Return port_.ToString
      End Function

      Public Overrides Property ReadTimeout() As Integer
        Get
          Return port_.ReadTotalTimeoutConstant
        End Get
        Set(ByVal value As Integer)
          port_.ReadTotalTimeoutConstant = value
        End Set
      End Property

      Public Overrides Sub Flush()
        port_.Purge(PurgeActions.All)
      End Sub

      Public Overrides ReadOnly Property CanRead() As Boolean
        Get
          Return True
        End Get
      End Property
      Public Overrides ReadOnly Property CanWrite() As Boolean
        Get
          Return True
        End Get
      End Property
      Public Overrides ReadOnly Property CanSeek() As Boolean
        Get
          Return False
        End Get
      End Property
      Public Overrides ReadOnly Property Length() As Long
        Get
          Return 0
        End Get
      End Property

      Public Overrides Property Position() As Long
        Get
          Return 0
        End Get
        Set(ByVal value As Long)
        End Set
      End Property
      Public Overrides Function Seek(ByVal offset As Long, ByVal origin As System.IO.SeekOrigin) As Long
        Return 0
      End Function
      Public Overrides Sub SetLength(ByVal value As Long)
      End Sub

#If Not ASYNCREADANDWRITE Then
      ' The sync versions are straightforward
      Public Overrides Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
        Dim red As Integer
        If offset = 0 Then
          NativeMethods.ReadFile(port_.hCom_, buffer, count, red, Nothing)
        Else
          ' This is a nuisance, but we can do it
          Dim buf2(count - 1) As Byte
          NativeMethods.ReadFile(port_.hCom_, buf2, count, red, Nothing)
          Array.Copy(buf2, 0, buffer, offset, red)
        End If
        if red > 0 then totalIn_ += red
        Return red
      End Function
      Public Overrides Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
        Dim wrote As Integer
        If offset = 0 Then
          NativeMethods.WriteFile(port_.hCom_, buffer, count, wrote, Nothing)
        Else
          ' This is a nuisance, but we can do it
          Dim buf2(count - 1) As Byte : Array.Copy(buffer, offset, buf2, 0, count)
          NativeMethods.WriteFile(port_.hCom_, buf2, count, wrote, Nothing)
        End If
        totalOut_ += wrote
      End Sub
#Else
      ' On the desktop, but not on CF, we can provide our own versions of Begin/End/Read/Write, rather than
      ' relying on the delegate versions that IO.Stream will otherwise synthesise.
      ' Our versions are better because we can wait directly on the Win32 file implementation.
      Private ReadOnly asyncSerialPort_ As New AsyncSerialPort

      Public Overrides Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
        Dim red As Integer
        If (port_.options_ And OptionsValue.Sync) = 0 Then  ' async
          red = EndRead(BeginRead(buffer, offset, count, Nothing, Nothing))
        Else
          If Not port_.TryToOpen Then Return -1
          If offset = 0 Then
            If NativeMethods.ReadFile(port_.hCom_, buffer, count, red, Nothing) = 0 Then
              Throw New System.ComponentModel.Win32Exception
            End If
          Else
            Dim buf2(count - 1) As Byte  ' a bit annoying...
            If NativeMethods.ReadFile(port_.hCom_, buf2, count, red, Nothing) = 0 Then
              Throw New System.ComponentModel.Win32Exception
            End If
            Array.Copy(buf2, 0, buffer, offset, red)
          End If
        End If
        If red > 0 Then totalIn_ += red
        Return red
      End Function

      Public Overrides Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
        If (port_.options_ And OptionsValue.Sync) = 0 Then
          EndWrite(BeginWrite(buffer, offset, count, Nothing, Nothing))
          totalOut_ += count
        Else
          If offset <> 0 Then Throw New NotSupportedException
          '  If Not port_.TryToOpen Then Return -1
          Dim wrote As Integer
          If NativeMethods.WriteFile(port_.hCom_, buffer, count, wrote, Nothing) = 0 Then
            Throw New System.ComponentModel.Win32Exception
          End If
          totalOut_ += wrote
        End If
      End Sub

      Public Overrides Function BeginRead(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, _
                                          ByVal callback As AsyncCallback, ByVal state As Object) As IAsyncResult
        If Not port_.TryToOpen Then Return New QuietAsyncResult(callback, state, Date.MinValue) ' don't even bother trying to open if not already open

        If (port_.options_ And OptionsValue.Sync) <> 0 Then Throw New NotImplementedException
        If callback IsNot Nothing Then Throw New NotSupportedException
        asyncSerialPort_.State = state
        Return asyncSerialPort_.BeginRead(port_.hCom_, buffer, offset, count)
      End Function
      Public Overrides Function EndRead(ByVal asyncResult As IAsyncResult) As Integer
        Dim ret As Integer = asyncSerialPort_.EndRead(asyncResult)
        ' If there has been a bad timeout problem, close the port
        If ret = -1 Then port_.Close()
        Return ret
      End Function
      Public Overrides Function BeginWrite(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, _
                                           ByVal callback As AsyncCallback, ByVal state As Object) As IAsyncResult
        ' If the port is not open, and can not be opened, then just return something useless
        If Not port_.TryToOpen() Then Return New QuietAsyncResult(Nothing, state, Date.UtcNow + TimeSpan.FromSeconds(1))

        If (port_.options_ And OptionsValue.Sync) <> 0 Then Throw New NotImplementedException
        If callback IsNot Nothing Then Throw New NotSupportedException
        asyncSerialPort_.State = state
        Return asyncSerialPort_.BeginWrite(port_.hCom_, buffer, offset, count)
      End Function
      Public Overrides Sub EndWrite(ByVal asyncResult As IAsyncResult)
        If TypeOf asyncResult Is QuietAsyncResult Then Exit Sub ' quietly do nothing

        ' If there has been a bad timeout problem, close the port
        If asyncSerialPort_.EndWrite(asyncResult) = -1 Then port_.Close()
      End Sub

      ' --------------------------------------------
      Private Class AsyncSerialPort
        Implements IAsyncResult

        Private overlapped_ As NativeMethods.OVERLAPPED, overlappedHandle_ As Runtime.InteropServices.GCHandle
        Private hCom_ As IntPtr, bufferHandle_ As Runtime.InteropServices.GCHandle, bufferIsPinned_ As Boolean
        Private state_ As Object, timeoutTime_ As Date, timedOut_ As Boolean

        ' Interop will pin buffers only during the function calls, so we have to pin the overlapped_ memory ourself
        Public Sub New()
          overlapped_ = New NativeMethods.OVERLAPPED
          overlappedHandle_ = Runtime.InteropServices.GCHandle.Alloc(overlapped_, Runtime.InteropServices.GCHandleType.Pinned)
        End Sub

        Private Const ERROR_IO_PENDING As Integer = 997  ' not an error, just a confirmation that the work is pending

        Public Function BeginRead(ByVal hCom As IntPtr, ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As IAsyncResult
          If offset <> 0 Then Throw New NotSupportedException
          PinBuffer(buffer)
          Dim red As Integer  ' always an overlapped read, so this will not be set
          Static count2 As Integer
          count2 = count
          If NativeMethods.ReadFile(hCom, buffer, count2, red, overlapped_) = 0 AndAlso _
               Runtime.InteropServices.Marshal.GetLastWin32Error <> ERROR_IO_PENDING Then
            UnpinBuffer()
            Return New QuietAsyncResult(Nothing, state_, Date.MinValue)
          End If
          timeoutTime_ = Date.UtcNow + TimeSpan.FromSeconds(10) : hCom_ = hCom
          Return Me
        End Function

        Public Function EndRead(ByVal asyncResult As IAsyncResult) As Integer
          If TypeOf asyncResult Is QuietAsyncResult Then Return -1
          Dim bytes As Integer
          If timedOut_ Then
            bytes = -1
          Else
            NativeMethods.GetOverlappedResult(hCom_, overlapped_, bytes, True)  ' only returns when complete
          End If
          UnpinBuffer()
          Return bytes
        End Function

        Public Function BeginWrite(ByVal hCom As IntPtr, ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As IAsyncResult
          If offset <> 0 Then Throw New NotSupportedException

          ' Make the native, overlapped call
          PinBuffer(buffer)
          If NativeMethods.WriteFile(hCom, buffer, count, 0, overlapped_) = 0 AndAlso _
               Runtime.InteropServices.Marshal.GetLastWin32Error <> ERROR_IO_PENDING Then
            UnpinBuffer()
            Return New QuietAsyncResult(Nothing, state_, Date.MinValue)
          End If
          timeoutTime_ = Date.UtcNow + TimeSpan.FromSeconds(2)  ' for safety, an overall timeout of our own, in case the Win serial driver fails us
          hCom_ = hCom
          Return Me
        End Function

        Public Function EndWrite(ByVal asyncResult As IAsyncResult) As Integer
          If TypeOf asyncResult Is QuietAsyncResult Then Return -1
          Dim bytes As Integer
          If timedOut_ Then
            bytes = -1
          Else
            NativeMethods.GetOverlappedResult(hCom_, overlapped_, bytes, True) ' only returns when complete
          End If
          UnpinBuffer()
          Return bytes
        End Function

        Private Sub PinBuffer(ByVal buffer() As Byte)
          bufferHandle_ = Runtime.InteropServices.GCHandle.Alloc(buffer, Runtime.InteropServices.GCHandleType.Pinned)
          bufferIsPinned_ = True
        End Sub
        Private Sub UnpinBuffer()
          If bufferIsPinned_ Then
            bufferHandle_.Free()
            bufferIsPinned_ = False
          End If
        End Sub
        Private ReadOnly Property IsCompleted() As Boolean Implements IAsyncResult.IsCompleted
          Get
            If overlapped_.HasIoCompleted Then
              timedOut_ = False : Return True
            End If
            If Date.UtcNow >= timeoutTime_ Then
              timedOut_ = True : Return True
            End If
            Return False
          End Get
        End Property

        ' The rest are not much interest
        Private ReadOnly Property AsyncState() As Object Implements IAsyncResult.AsyncState
          Get
            Return state_
          End Get
        End Property
        Public Property State() As Object
          Get
            Return state_
          End Get
          Set(ByVal value As Object)
            state_ = value
          End Set
        End Property
        Private ReadOnly Property AsyncWaitHandle() As System.Threading.WaitHandle Implements IAsyncResult.AsyncWaitHandle
          Get
            Throw New NotImplementedException
            '            Return New system.Threading.WaitHandle(overlapped_.hEvent)  ' TODO: almost
            Return Nothing
          End Get
        End Property
        Private ReadOnly Property CompletedSynchronously() As Boolean Implements IAsyncResult.CompletedSynchronously
          Get
            Return False
          End Get
        End Property
      End Class
#End If
    End Class

    ' ---------------------------------------------------------
    Private NotInheritable Class NativeMethods
      Private Sub New()
      End Sub
#If CF Then
      Public Declare Function CloseHandle Lib "coredll" (ByVal handle As IntPtr) As <Runtime.InteropServices.MarshalAs( Runtime.InteropServices.UnmanagedType.Bool)>  Boolean
      Public Declare Function CreateEvent Lib "coredll" (ByVal eventAttributes As IntPtr, <Runtime.InteropServices.MarshalAs( Runtime.InteropServices.UnmanagedType.Bool)>  ByVal manualReset As Boolean, <Runtime.InteropServices.MarshalAs( Runtime.InteropServices.UnmanagedType.Bool)>  ByVal initialState As Boolean, ByVal name As String) As IntPtr
      Public Declare Function CreateFile Lib "coredll" ( _
          ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, _
          ByVal lpSecurityAttributes As IntPtr, ByVal dwCreationDisposition As Integer, _
          ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As IntPtr) As IntPtr
      Public Declare Function GetCommState Lib "coredll" (ByVal hCom As IntPtr, ByVal dcb() As Byte) As <Runtime.InteropServices.MarshalAs( Runtime.InteropServices.UnmanagedType.Bool)>  Boolean
      Public Declare Function PurgeComm Lib "coredll" (ByVal hCom As IntPtr, ByVal flags As Integer) As <Runtime.InteropServices.MarshalAs( Runtime.InteropServices.UnmanagedType.Bool)>  Boolean
      Public Declare Function ReadFile Lib "coredll" (ByVal hFile As IntPtr, ByVal buffer As Byte(), ByVal numberOfBytesToRead As Integer, ByRef numberOfBytesRead As Integer, ByVal unsupported As IntPtr) As Integer
      Public Declare Function SetCommState Lib "coredll" (ByVal hCom As IntPtr, ByVal dcb() As Byte) As <Runtime.InteropServices.MarshalAs( Runtime.InteropServices.UnmanagedType.Bool)>  Boolean
      Public Declare Function SetCommTimeouts Lib "coredll" (ByVal hCom As IntPtr, ByVal ct As COMMTIMEOUTS) As  <Runtime.InteropServices.MarshalAs( Runtime.InteropServices.UnmanagedType.Bool)>  Boolean
      Public Declare Function WriteFile Lib "coredll" (ByVal hFile As IntPtr, ByVal buffer As Byte(), ByVal numberOfBytesToWrite As Integer, ByRef numberOfBytesWritten As Integer, ByVal unsupported As IntPtr) As Integer
#Else
      Public Declare Function CloseHandle Lib "kernel32" (ByVal handle As IntPtr) As <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> Boolean
      Public Declare Auto Function CreateFile Lib "kernel32" ( _
          ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, _
          ByVal lpSecurityAttributes As IntPtr, ByVal dwCreationDisposition As Integer, _
          ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As IntPtr) As IntPtr
      Public Declare Function GetCommState Lib "kernel32" (ByVal hCom As IntPtr, ByVal dcb() As Byte) As <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> Boolean
      Public Declare Function PurgeComm Lib "kernel32" (ByVal hCom As IntPtr, ByVal flags As Integer) As <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> Boolean
      Public Declare Function SetCommState Lib "kernel32" (ByVal hCom As IntPtr, ByVal dcb() As Byte) As <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> Boolean
      Public Declare Function SetCommTimeouts Lib "kernel32" (ByVal hCom As IntPtr, ByVal ct As COMMTIMEOUTS) As <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> Boolean
      Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As IntPtr, ByVal buffer As Byte(), ByVal numberOfBytesToRead As Integer, ByRef numberOfBytesRead As Integer, ByVal ol As OVERLAPPED) As Integer
      Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As IntPtr, ByVal buffer As Byte(), ByVal numberOfBytesToWrite As Integer, ByRef numberOfBytesWritten As Integer, ByVal ol As OVERLAPPED) As Integer
      Public Declare Auto Function CreateEvent Lib "kernel32" (ByVal eventAttributes As IntPtr, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> ByVal manualReset As Boolean, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> ByVal initialState As Boolean, ByVal name As String) As IntPtr
      Public Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As IntPtr, ByVal overlapped As OVERLAPPED, ByRef numberOfBytesTransferred As Integer, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> ByVal wait As Boolean) As <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.Bool)> Boolean

      ' ----------------------------------------------------------------
      <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> _
      Public Class OVERLAPPED
        Implements IDisposable

        Public Internal As Integer
        Public InternalHigh As Integer
        Public Offset As Integer
        Public OffsetHigh As Integer
        Public hEvent As IntPtr


        Public Sub New()
          hEvent = NativeMethods.CreateEvent(IntPtr.Zero, True, False, Nothing)  ' must be a manual reset event
        End Sub

        Public Function HasIoCompleted() As Boolean
          Const STATUS_PENDING As Integer = &H103
          Return Internal <> STATUS_PENDING
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
          Dispose(True)
          GC.SuppressFinalize(Me)
        End Sub

        Private Sub Dispose(ByVal disposing As Boolean)
          If Not hEvent.Equals(IntPtr.Zero) Then
            NativeMethods.CloseHandle(hEvent)
            hEvent = IntPtr.Zero
          End If
        End Sub
        Protected Overrides Sub finalize()
          Dispose(False)
        End Sub
      End Class
#End If
    End Class

    ' -----------------------------------------
    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> _
    Private Class COMMTIMEOUTS
      Public ReadIntervalTimeout As Integer
      Public ReadTotalTimeoutMultiplier As Integer
      Public ReadTotalTimeoutConstant As Integer
      Public WriteTotalTimeoutMultiplier As Integer
      Public WriteTotalTimeoutConstant As Integer
    End Class
  End Class


  ' -----------------------------------------------------------------------
  ' A network, sockets, stream that keeps re-connecting, and doesn't throw exceptions.
  ' Instead of exceptions, it just returns 0 bytes from Read's and quietly skips Write's
  ' Only works on ipv4 - probably not a problem
  Public Class NetworkPort : Inherits System.IO.Stream
    Private ReadOnly computer_ As String, port_ As Integer
    Private socket_ As Net.Sockets.Socket, receiveTimeout_ As Integer

    Public Sub New(ByVal computer As String, ByVal port As Integer)
      computer_ = computer : port_ = port
    End Sub

    Public Overrides Function ToString() As String
      Return computer_ & ":" & port_.ToString(Globalization.CultureInfo.InvariantCulture)
    End Function

    Public Overrides Sub Close()
      If socket_ IsNot Nothing Then socket_.Close() : socket_ = Nothing
    End Sub

    Public Overrides Property ReadTimeout() As Integer
      Get
        Return receiveTimeout_
      End Get
      Set(ByVal value As Integer)
        If receiveTimeout_ = value Then Exit Property
        receiveTimeout_ = value
        If socket_ IsNot Nothing Then socket_.ReceiveTimeout = value
      End Set
    End Property


    Public Overrides Function BeginWrite(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, ByVal callback As AsyncCallback, ByVal state As Object) As IAsyncResult
      ' If the socket is not currently connected, then ensure that is attempted first
      If socket_ Is Nothing Then Return New ConnectThenSend(Me, buffer, offset, count, callback, state)
      Try
        Return socket_.BeginSend(buffer, offset, count, Net.Sockets.SocketFlags.None, callback, state)
      Catch
        Close()
        Return New QuietAsyncResult(callback, state, Date.MinValue)
      End Try
    End Function

    ' Nothing worth worrying about
    Public Overrides Sub EndWrite(ByVal asyncResult As IAsyncResult)
      If TypeOf asyncResult Is QuietAsyncResult Then Exit Sub ' quietly do nothing
      If Not TypeOf asyncResult Is ConnectThenSend Then
        Try
          socket_.EndSend(asyncResult)
        Catch
          Close()
        End Try
      End If
    End Sub

    Public Overrides Function BeginRead(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, _
                                        ByVal callback As AsyncCallback, ByVal state As Object) As IAsyncResult
      ' We assume a previous Write will have done its best to get the socket connected
      If socket_ Is Nothing Then Return New QuietAsyncResult(callback, state, Date.MinValue)

      Try
        Return socket_.BeginReceive(buffer, offset, count, Net.Sockets.SocketFlags.None, callback, state)
      Catch
        Close()  ' some socket exception
        Return New QuietAsyncResult(callback, state, Date.MinValue)  ' handly it quietly
      End Try
    End Function

    Public Overrides Function EndRead(ByVal asyncResult As IAsyncResult) As Integer
      If TypeOf asyncResult Is QuietAsyncResult Then Return 0 ' no bytes received
      Try
        Return socket_.EndReceive(asyncResult)
      Catch
        Close()
        Return 0   ' sanitize it
      End Try
    End Function

    ' We define simple not async code as well, as this may be all that our callers need
    Public Overrides Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
      If socket_ Is Nothing Then Return 0 ' should have been opened by a preceding Write
      Try
        Return socket_.Receive(buffer, offset, count, Net.Sockets.SocketFlags.None)
      Catch
        Close()
        Return 0
      End Try
    End Function

    Public Overrides Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
      If socket_ Is Nothing Then
        socket_ = New Net.Sockets.Socket(Net.Sockets.AddressFamily.InterNetwork, Net.Sockets.SocketType.Stream, Net.Sockets.ProtocolType.Tcp)
        socket_.ReceiveTimeout = receiveTimeout_
        Try
          socket_.Connect(computer_, port_)
        Catch
          Close()
        End Try
      End If
      Try
        socket_.Send(buffer, offset, count, Net.Sockets.SocketFlags.None)
      Catch
        Close()
      End Try
    End Sub

    ' These are necessary, but un-interesting
    Public Overrides ReadOnly Property CanRead() As Boolean
      Get
        Return True
      End Get
    End Property
    Public Overrides ReadOnly Property CanSeek() As Boolean
      Get
        Return True ' means async call-backs are available
      End Get
    End Property
    Public Overrides ReadOnly Property CanWrite() As Boolean
      Get
        Return True
      End Get
    End Property
    Public Overrides Sub Flush()
    End Sub
    Public Overrides ReadOnly Property Length() As Long
      Get
        Throw New NotSupportedException
      End Get
    End Property
    Public Overrides Property Position() As Long
      Get
        Throw New NotSupportedException
      End Get
      Set(ByVal value As Long)
        Throw New NotSupportedException
      End Set
    End Property
    Public Overrides Function Seek(ByVal offset As Long, ByVal origin As System.IO.SeekOrigin) As Long
      Throw New NotSupportedException
    End Function
    Public Overrides Sub SetLength(ByVal value As Long)
      Throw New NotSupportedException
    End Sub

    ' ------------------------------------------------------------------
    Private Class ConnectThenSend : Implements IAsyncResult
      Private ReadOnly owner_ As NetworkPort, socket_ As Net.Sockets.Socket, _
                       buffer_() As Byte, offset_, count_ As Integer, _
                       callback_ As AsyncCallback, state_ As Object
      Private completed_ As Boolean

      Public Sub New(ByVal owner As NetworkPort, ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, ByVal callback As AsyncCallback, ByVal state As Object)
        owner_ = owner : buffer_ = buffer : offset_ = offset : count_ = count
        callback_ = callback : state_ = state
        socket_ = New Net.Sockets.Socket(Net.Sockets.AddressFamily.InterNetwork, Net.Sockets.SocketType.Stream, Net.Sockets.ProtocolType.Tcp)
        socket_.ReceiveTimeout = owner.receiveTimeout_
        owner_.socket_ = socket_
        Try
          ' This BeginConnect may sometimes (but not always) throw a SocketException
          socket_.BeginConnect(owner.computer_, owner.port_, AddressOf OnConnected, Nothing)
        Catch
          owner_.Close()
          completed_ = True  ' give up
          If callback_ IsNot Nothing Then callback_(Me)
        End Try
      End Sub

      Private Sub OnConnected(ByVal ar As IAsyncResult)
        Try
          socket_.EndConnect(ar)
          socket_.BeginSend(buffer_, offset_, count_, Net.Sockets.SocketFlags.None, AddressOf OnSent, Nothing)
        Catch ex As Exception
          completed_ = True  ' give up
          If socket_ Is owner_.socket_ Then
          owner_.Close()
          If callback_ IsNot Nothing Then callback_(Me)
          Else
            socket_.Close()
          End If
        End Try
      End Sub

      Private Sub OnSent(ByVal ar As IAsyncResult)
        Try
          socket_.EndSend(ar)
        Catch
          If socket_ Is owner_.socket_ Then
          owner_.Close()
          Else
            socket_.Close()
          End If
        End Try
        completed_ = True
        If callback_ IsNot Nothing Then callback_(Me)
      End Sub

      Private ReadOnly Property AsyncState() As Object Implements IAsyncResult.AsyncState
        Get
          Return state_
        End Get
      End Property
      Private ReadOnly Property AsyncWaitHandle() As System.Threading.WaitHandle Implements IAsyncResult.AsyncWaitHandle
        Get
          Throw New NotImplementedException
        End Get
      End Property
      Private ReadOnly Property CompletedSynchronously() As Boolean Implements IAsyncResult.CompletedSynchronously
        Get
          Return False
        End Get
      End Property
      Private ReadOnly Property IsCompleted() As Boolean Implements IAsyncResult.IsCompleted
        Get
          Return completed_
        End Get
      End Property
    End Class
  End Class

  ' ------------------------------------------------------------------
  Friend Class QuietAsyncResult : Implements IAsyncResult
    Private ReadOnly state_ As Object, completeTime_ As Date
    Public Sub New(ByVal callback As AsyncCallback, ByVal state As Object, ByVal completeTime As Date)
      state_ = state : completeTime_ = completeTime
      If callback IsNot Nothing Then callback(Me)
    End Sub
    Private ReadOnly Property AsyncState() As Object Implements IAsyncResult.AsyncState
      Get
        Return state_
      End Get
    End Property
    Private ReadOnly Property AsyncWaitHandle() As System.Threading.WaitHandle Implements IAsyncResult.AsyncWaitHandle
      Get
        Throw New NotImplementedException
      End Get
    End Property
    Private ReadOnly Property CompletedSynchronously() As Boolean Implements IAsyncResult.CompletedSynchronously
      Get
        Return False
      End Get
    End Property
    Private ReadOnly Property IsCompleted() As Boolean Implements IAsyncResult.IsCompleted
      Get
        Return completeTime_ = Date.MinValue OrElse Date.UtcNow >= completeTime_
      End Get
    End Property
  End Class

  Public Class StreamLogger : Inherits System.IO.Stream
    Private ReadOnly stm_ As System.IO.Stream, writeFile_ As System.IO.FileStream

    Public Sub New(ByVal stm As System.IO.Stream, ByVal path As String)
      stm_ = stm
      writeFile_ = New System.IO.FileStream(path, System.IO.FileMode.Create)
    End Sub

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
      If disposing Then stm_.Dispose() : writeFile_.Close()
      MyBase.Dispose(disposing)
    End Sub

    Public Overrides Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
      Dim ret As Integer = stm_.Read(buffer, offset, count)
      writeFile_.Write(buffer, offset, ret)  ' anything read we copy to our lof file
      Return ret
    End Function
    Public Overrides Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
      stm_.Write(buffer, offset, count)
    End Sub

    Public Overrides ReadOnly Property CanRead() As Boolean
      Get
        Return stm_.CanSeek
      End Get
    End Property
    Public Overrides ReadOnly Property CanSeek() As Boolean
      Get
        Return stm_.CanSeek
      End Get
    End Property

    Public Overrides ReadOnly Property CanWrite() As Boolean
      Get
        Return stm_.CanWrite
      End Get
    End Property

    Public Overrides Sub Flush()
      stm_.Flush()
    End Sub

    Public Overrides ReadOnly Property Length() As Long
      Get
        Return stm_.Length
      End Get
    End Property

    Public Overrides Property Position() As Long
      Get
        Return stm_.Position
      End Get
      Set(ByVal value As Long)
        stm_.Position = value
      End Set
    End Property


    Public Overrides Function Seek(ByVal offset As Long, ByVal origin As System.IO.SeekOrigin) As Long
      Return stm_.Seek(offset, origin)
    End Function

    Public Overrides Sub SetLength(ByVal value As Long)
      stm_.SetLength(value)
    End Sub
  End Class

  ' -----------------------------------------------------------------------------------
  ' Encapsulates a stream, adding built-in diagnostics that are useful for the control system
  Public Class Stream : Implements IDisposable
    Private ReadOnly stm_ As System.IO.Stream
    Private logComEvents_ As Boolean, logFileName_ As String, writeFile_ As System.IO.TextWriter

    ' Used for diagnostics
    Private latestRx_() As Byte : Private ReadOnly comEvents_ As New ComEventCollection

    Public Sub New(ByVal stm As System.IO.Stream)
      stm_ = stm
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
      CloseLogFile()
    End Sub

    Private Sub CloseLogFile()
      If writeFile_ IsNot Nothing Then writeFile_.Close() : writeFile_ = Nothing
    End Sub

    Public ReadOnly Property Stream() As System.IO.Stream
      Get
        Return stm_
      End Get
    End Property

    Public Overrides Function ToString() As String
      Return stm_.ToString
    End Function

    Public Property LogComEvents() As Boolean
      Get
        Return logComEvents_
      End Get
      Set(ByVal value As Boolean)
        If logComEvents_ = value Then Exit Property
        logComEvents_ = value
#If 0 Then
        LogFileName = If(value, "c:\ComsDiags.txt", Nothing)
#End If
      End Set
    End Property

    Public Property LogFileName As String
      Get
        Return logFileName_
      End Get
      Set(ByVal value As String)
        If logFileName_ = value Then Exit Property
        CloseLogFile()
        logFileName_ = value
        If Not String.IsNullOrEmpty(value) Then
          Dim sw As New System.IO.StreamWriter(value) : sw.AutoFlush = True
          writeFile_ = sw
        End If
      End Set
    End Property

    Public Sub Flush()
      stm_.Flush()
    End Sub

    Public Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
      If logComEvents_ Then AddEvents(ComEvent.EventTypeValue.Tx, buffer, offset, count)
      stm_.Write(buffer, offset, count)
    End Sub
    Public Function BeginWrite(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, _
                               ByVal callback As AsyncCallback, ByVal state As Object) As IAsyncResult
      If logComEvents_ Then AddEvents(ComEvent.EventTypeValue.Tx, buffer, offset, count)
      Return stm_.BeginWrite(buffer, offset, count, callback, state)
    End Function
    Public Sub EndWrite(ByVal asyncResult As IAsyncResult)
      stm_.EndWrite(asyncResult)
    End Sub

    Private Sub AddEvents(ByVal eventType As ComEvent.EventTypeValue, ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
      If count <= 0 Then Exit Sub ' -1 means hardware fault, for instance
      ' Keep a note for diagnostics - unfortunately, each byte will have the same time on it
      Dim time As Date = Date.UtcNow
      Dim events(count - 1) As ComEvent
      For i As Integer = 0 To events.Length - 1
        Dim ev As New ComEvent(time, eventType, buffer(offset + i))
        events(i) = ev
        If writeFile_ IsNot Nothing Then writeFile_.WriteLine(ev.ToString)
      Next i
      SyncLock comEvents_
        comEvents_.AddRange(events)
      End SyncLock
    End Sub

    Public Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
      Dim ret As Integer = stm_.Read(buffer, offset, count)
      If logComEvents_ Then AddEvents(ComEvent.EventTypeValue.Rx, buffer, offset, ret)
      Return ret
    End Function

    Public Function BeginRead(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, _
                              ByVal callback As AsyncCallback, ByVal state As Object) As IAsyncResult
      latestRx_ = buffer
      Return stm_.BeginRead(buffer, offset, count, callback, state)
    End Function

    Public Function EndRead(ByVal asyncResult As IAsyncResult) As Integer
      Dim ret As Integer = stm_.EndRead(asyncResult)
      If logComEvents_ Then AddEvents(ComEvent.EventTypeValue.Rx, latestRx_, 0, ret)
      Return ret
    End Function

    Friend ReadOnly Property ComEvents() As ComEvent()
      Get
        ' Returns a thread-safe snapshot
        Dim ret() As ComEvent
        SyncLock comEvents_
          Dim count As Integer = comEvents_.Count
          ret = New ComEvent(count - 1) {}
          For i As Integer = 0 To count - 1
            ret(i) = comEvents_(i)
          Next i
        End SyncLock
        Return ret
      End Get
    End Property

    ' ----------------------------
    Public Class ComEvent
      Private ReadOnly time_ As Date, eventType_ As EventTypeValue, ch_ As Byte

      Public Sub New(ByVal time As Date, ByVal eventType As EventTypeValue, ByVal ch As Byte)
        time_ = time : eventType_ = eventType : ch_ = ch
      End Sub
      Public ReadOnly Property Time() As Date
        Get
          Return time_
        End Get
      End Property
      Public ReadOnly Property EventType() As EventTypeValue
        Get
          Return eventType_
        End Get
      End Property
      Public ReadOnly Property Ch() As Byte
        Get
          Return ch_
        End Get
      End Property

      Public Overrides Function ToString() As String
        Return time_.ToString("HH:mm:ss.fff", System.Globalization.CultureInfo.InvariantCulture) & ", " & eventType_.ToString & ", " & ch_.ToString("X2", System.Globalization.CultureInfo.InvariantCulture)
      End Function

      Public Enum EventTypeValue
        Rx
        Tx
      End Enum
    End Class
    ' ----------------------------
    Private Class ComEventCollection : Inherits LimitedCollection(Of ComEvent)
      Public Sub New()
        MyBase.New(1000)  ' limiting capacity
      End Sub
#If 0 Then
      Public Overloads Sub Add(ByVal time As Date, ByVal eventType As ComEvent.EventTypeValue, ByVal ch As Byte)
        Add(New ComEvent(time, eventType, ch))
      End Sub
#End If
    End Class

    Private Class LimitedCollection(Of T) : Implements ICollection(Of T)
      Private ReadOnly coll_() As T, capacity_ As Integer
      Private count_ As Integer, firstOfs_ As Integer

      Public Sub New(ByVal capacity As Integer)
        capacity_ = capacity : coll_ = New T(capacity - 1) {}
      End Sub

      Public Sub Add(ByVal value As T) Implements ICollection(Of T).Add
        If count_ < capacity_ Then
          coll_(count_) = value : count_ += 1
        Else
          coll_(firstOfs_) = value
          firstOfs_ += 1 : If firstOfs_ = capacity_ Then firstOfs_ = 0
        End If
      End Sub

      ' TODO: probably could optimise this a bit
      Public Sub AddRange(ByVal value() As T)
        For Each x As T In value : Add(x) : Next x
      End Sub

      Default Public ReadOnly Property Item(ByVal index As Integer) As T
        Get
          index += firstOfs_ : If index >= capacity_ Then index -= capacity_
          Return coll_(index)
        End Get
      End Property

      Public ReadOnly Property Count() As Integer Implements ICollection(Of T).Count
        Get
          Return count_
        End Get
      End Property

      Public Sub Clear() Implements ICollection(Of T).Clear
        count_ = 0 : firstOfs_ = 0
      End Sub

      Public Function Contains(ByVal item As T) As Boolean Implements ICollection(Of T).Contains
        Dim ofs As Integer = firstOfs_
        For i As Integer = 0 To count_ - 1
          If coll_(ofs).Equals(item) Then Return True
          ofs += 1 : If ofs = capacity_ Then ofs = 0
        Next i
        Return False
      End Function

      Private Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
        Return Nothing
      End Function


      Private ReadOnly Property IsReadOnly() As Boolean Implements ICollection(Of T).IsReadOnly
        Get
          Return False
        End Get
      End Property
      Private Function Remove(ByVal item As T) As Boolean Implements ICollection(Of T).Remove
        Throw New NotSupportedException
      End Function
      Private Sub CopyTo(ByVal array() As T, ByVal arrayIndex As Integer) Implements ICollection(Of T).CopyTo
        Throw New NotSupportedException
      End Sub
      Private Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
      End Function
    End Class
  End Class

  ''' <summary>Statistics about activity on a serial port.</summary>
  Public Class SerialPortStatistics
    Public BytesIn As Integer
    Public BytesOut As Integer
    Public FramingErrors As Integer
    Public ParityErrors As Integer
    Public HardwareOverrunErrors As Integer
    Public BufferOverrunErrors As Integer
  End Class

  ''' <summary>Statistics about activity on a socket.</summary>
  Public Class SocketStatistics
    Public BytesIn As Integer
    Public BytesOut As Integer
  End Class

  Public Enum WriteMode
    Optimised
    Always
  End Enum

  ' ------------------------------------------------------
  Friend Class WriteOptimisation
    Private ReadOnly recentWrites_ As New List(Of RecentWrite)

    Const writeDelay As Integer = 5

    ''' <summary>Returns true if exactly the given values were written recently with the same params.</summary>
    Public Function RecentlyWritten(ByVal values As Array, ByVal ParamArray params() As Object) As Boolean
      ' Make invalid anything that's got too old
      Dim expiry As Date = Date.UtcNow - TimeSpan.FromSeconds(writeDelay)

      ' See if we have something the same recently
      For i As Integer = 0 To recentWrites_.Count - 1
        Dim rw As RecentWrite = recentWrites_.Item(i)
        If Not rw.OlderThan(expiry) AndAlso rw.EqualParams(params) AndAlso rw.EqualsValues(values) Then Return True
      Next i
      Return False
    End Function


    ''' <summary>Returns true if exactly the given values were written recently with the same params.</summary>
    Public Sub SuccessfulWrite(ByVal values As Array, ByVal ParamArray params() As Object)
      ' Make invalid anything that's got too old
      Dim now As Date = Date.UtcNow
      Dim expiry As Date = now - TimeSpan.FromSeconds(writeDelay)

      ' See if we have something the same recently
      Dim freeRw As RecentWrite
      For i As Integer = 0 To recentWrites_.Count - 1
        Dim rw As RecentWrite = recentWrites_.Item(i)
        If rw.OlderThan(expiry) Then
          If freeRw Is Nothing Then freeRw = rw
        Else
          If rw.EqualParams(params) Then
            If rw.EqualsValues(values) Then Exit Sub
            ' If done recently with the same params, but with different values, then re-use the slot
            rw.ReplaceValues(values, now)
            Exit Sub
          End If
        End If
      Next i
      ' Need a new (or expired) slot
      If freeRw Is Nothing Then freeRw = New RecentWrite : recentWrites_.Add(freeRw)

      ' Store the values
      freeRw.SetNewValues(values, params, now)
    End Sub


#If 0 Then
  'Older version of WriteOptimisation class, not using the "SuccessfulWrite Sub
  ' ------------------------------------------------------
  Friend Class WriteOptimisation
    Private ReadOnly recentWrites_ As New List(Of RecentWrite)

    Const writeDelay As Integer = 5

    ''' <summary>Returns true if exactly the given values were written recently with the same params.</summary>
    Public Function RecentlyWrote(ByVal values As Array, ByVal ParamArray params() As Object) As Boolean
      ' Make invalid anything that's got too old
      Dim now As Date = Date.UtcNow
      Dim expiry As Date = now - TimeSpan.FromSeconds(writeDelay)

      ' See if we have something the same recently
      Dim freeRw As RecentWrite
      For i As Integer = 0 To recentWrites_.Count - 1
        Dim rw As RecentWrite = recentWrites_.Item(i)
        If rw.OlderThan(expiry) Then
          If freeRw Is Nothing Then freeRw = rw
        Else
          If rw.EqualParams(params) Then
            If rw.EqualsValues(values) Then Return True
                ' If done recently with the same params, but with different values, then re-use the slot
            rw.ReplaceValues(values, now)
                Return False ' got to do a write
              End If
            End If
      Next i
      ' Need a new (or expired) slot
      If freeRw Is Nothing Then freeRw = New RecentWrite : recentWrites_.Add(freeRw)

      ' Store the values
      freeRw.SetNewValues(values, params, now)
      Return False
    End Function


#End If
    ' ---------------------------------------
    Private Class RecentWrite
      Private time_ As Date, values_ As Array, params_() As Object, _
              valuesBoolean_() As Boolean, valuesShort_() As Short

      Public Sub ExpireIfOlder(ByVal value As Date)
        If time_ < value Then time_ = Date.MinValue : values_ = Nothing : params_ = Nothing
      End Sub

      Public Function OlderThan(ByVal value As Date) As Boolean
        Return time_ < value
      End Function
      Public Sub ReplaceValues(ByVal values As Array, ByVal time As Date)
        Array.Copy(values, values_, values_.Length)  ' assumes that target array is already the correct length and type
        time_ = time
      End Sub
      Public Sub SetNewValues(ByVal values As Array, ByVal params() As Object, ByVal time As Date)
        values_ = DirectCast(values.Clone, Array)
        valuesBoolean_ = TryCast(values_, Boolean()) : valuesShort_ = TryCast(values_, Short())
        params_ = params ' don't bother cloning this one
        time_ = time
      End Sub
      Public Function EqualParams(ByVal other() As Object) As Boolean
        If params_.Length <> other.Length Then Return False
        Dim same As Boolean = True
        For i As Integer = 0 To params_.Length - 1
          If Not params_(i).Equals(other(i)) Then Return False
        Next i
        Return True
      End Function

      Public Function EqualsValues(ByVal other As Array) As Boolean
        If values_.Length <> other.Length Then Return False
        If valuesBoolean_ IsNot Nothing Then
          Dim oth() As Boolean = DirectCast(other, Boolean())
          For i As Integer = 0 To valuesBoolean_.Length - 1
            If valuesBoolean_(i) <> oth(i) Then Return False
          Next i
        ElseIf valuesShort_ IsNot Nothing Then
          Dim oth() As Short = DirectCast(other, Short())
          For i As Integer = 0 To valuesShort_.Length - 1
            If valuesShort_(i) <> oth(i) Then Return False
          Next i
        Else
          For i As Integer = 0 To values_.Length - 1
            If Not values_.GetValue(i).Equals(other.GetValue(i)) Then Return False
          Next i
        End If
        Return True
      End Function
    End Class
  End Class

  ' ------------------------------------------------------------
  Public NotInheritable Class BitConverter
    Private Sub New()
    End Sub
    Public Shared Function GetBooleans(ByVal value() As Short, ByVal startIndex As Integer, ByVal count As Integer) As Boolean()
      Dim ret(count * 16) As Boolean  ' for convenience, element 0 is not used
      For i As Integer = 0 To count * 16 - 1
        ret(i + 1) = (value(i \ 16 + startIndex) And (1 << (i And 15))) <> 0
      Next i
      Return ret
    End Function
    Public Shared Function GetBooleans(ByVal value() As Integer, ByVal startIndex As Integer, ByVal count As Integer) As Boolean()
      Dim ret(count * 32) As Boolean  ' for convenience, element 0 is not used
      For i As Integer = 0 To count * 32 - 1
        ret(i + 1) = (value(i \ 32 + startIndex) And (1 << (i And 31))) <> 0
      Next i
      Return ret
    End Function
    Public Shared Function GetInt16s(ByVal value() As Boolean, ByVal startIndex As Integer, ByVal count As Integer) As Short()
      Dim outCount As Integer = count \ 16
      If (count Mod 16) <> 0 Then outCount += 1
      Dim ret(outCount) As Short  ' for convenience, element 0 is not used
      For i As Integer = 0 To count - 1
        If value(startIndex + i) Then
          Dim ofs As Integer = (i \ 16) + 1
          ret(ofs) = ret(ofs) Or (1S << (i And 15))
        End If
      Next i
      Return ret
    End Function
    Public Shared Function GetInt32s(ByVal value() As Boolean, ByVal startIndex As Integer, ByVal count As Integer) As Integer()
      Dim outCount As Integer = count \ 32
      If (count Mod 32) <> 0 Then outCount += 1
      Dim ret(outCount) As Integer  ' for convenience, element 0 is not used
      For i As Integer = 0 To count - 1
        If value(startIndex + i) Then
          Dim ofs As Integer = (i \ 32) + 1
          ret(ofs) = ret(ofs) Or (1 << (i And 31))
        End If
      Next i
      Return ret
    End Function
  End Class
End Namespace