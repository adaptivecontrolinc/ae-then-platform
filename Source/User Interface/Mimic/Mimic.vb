Imports Utilities.Translations

Public Class Mimic
  Public ControlCode As ControlCode
  Public Remoted As Boolean

  Public Sub New()
    DoubleBuffered = True  ' no flicker 
    InitializeComponent()

    ' NOTE: Always have this set as none for a mimic
    Me.AutoScaleMode = Windows.Forms.AutoScaleMode.None
  End Sub

  Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles Timer1.Tick
    'There's an odd bug related to boxing the integer values that creates probelms with .ToString
    '  to get round this feature we divide the values by 1 which seems to force the values to be boxed correctly
    Try
      'Make sure we have a connection
      If ControlCode Is Nothing Then Exit Sub

      'Check to see if we are remoted (Plant Explorer / History)
      Remoted = Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode)
      If Remoted Then buttonManualControl.Visible = False

      With ControlCode

        HeatCoolOutput.Value = .IO.HeatCoolOutput
        If Not .IO.CoolSelect AndAlso Not .IO.HeatSelect Then
          HeatCoolOutput.TemperatureChange = TemperatureChange.None
        End If


        LblTank1Status.Text = ("Tank 1") & (": ") & .Tank1Status
        LblTank1Level.Text = ("Tank 1") & (": ") & (.Tank1Level / 10).ToString("#0.0") & "%"

        levelBarMachine.Value = .MachineLevel / 10
        levelBarReserve.Value = .ReserveLevel / 10
        levelBarAdd.Value = .AddLevel / 10
        levelBarTank1.Value = .Tank1Level / 10


        If .PumpControl.IsRunning Then
          buttonPumpRequest.Text = Translate("Pump Stop")
        Else : buttonPumpRequest.Text = Translate("Pump Start")
        End If
        labelPumpStatus.Text = .PumpControl.StateString
        labelPumpReversalStatus.Text = .PumpControl.StateFlowString

        labelMachineLevel.Text = Translate("Level") & (": ") & (.MachineLevel / 10).ToString("#0.0") & "%"

        labelMachinePressure.Text = Translate("Machine Pressure") & (": ") & .MachinePressureDisplay
        labelPackageDP.Text = .PackageDpStr

        '********************************************************************************************
        '******   PANEL FORM - MANUAL CONTROL
        '********************************************************************************************

        ' Add Tank GroupBox: 
        GroupBoxAddTank.Text = Translate("Add Tank")
        labelAddStatusPanelButtons.Text = .AddStatus
        labelAddLevelPanelButtons.Text = Translate("Level") & (": ") & (.AddLevel / 10).ToString("#0.0") & "%"
        labelAddStatusMimic.Text = .AddStatus
        labelAddLevelMimic.Text = Translate("Level") & (": ") & (.AddLevel / 10).ToString("#0.0") & "%"

        ' Add Ready Pushbuttons
        buttonAddReady.Text = Translate("Add Ready")
        buttonAddReady1.Text = Translate("Add Ready")
        If .AddReady Then
          buttonAddReady.BackColor = Color.Green
          buttonAddReady.ForeColor = Color.White
          buttonAddReady1.BackColor = Color.Green
          buttonAddReady1.ForeColor = Color.White
        Else
          buttonAddReady.BackColor = Color.WhiteSmoke
          buttonAddReady.ForeColor = Color.Black
          buttonAddReady1.BackColor = Color.WhiteSmoke
          buttonAddReady1.ForeColor = Color.Black
        End If

        ' Add Fill
        buttonAddFill.Text = Translate("Fill")
        If .AddControl.FillIsActive Then
          buttonAddFill.BackColor = Color.Green
        Else
          buttonAddFill.BackColor = Color.Transparent
        End If

        ' Add Drain
        buttonAddDrain.Text = Translate("Drain")
        If .AddControl.DrainIsActive Then
          buttonAddDrain.BackColor = Color.Green
        Else
          buttonAddDrain.BackColor = Color.Transparent
        End If

        ' Add Cancel
        buttonAddCancel.Text = Translate("Cancel")
        If .AddControl.CancelActive Then
          buttonAddCancel.BackColor = Color.Green
        Else
          buttonAddCancel.BackColor = Color.Transparent
        End If

        ' Reserve Tank GroupBox:
        GroupBoxReserveTank.Text = Translate("Reserve Tank")
        labelReserveStatusPanelButtons.Text = .ReserveStatus
        labelReserveLevelPanelButtons.Text = Translate("Level") & (": ") & (.ReserveLevel / 10).ToString("#0.0") & "%"

        labelReserveLevelMimic.Text = Translate("Level") & (": ") & (.ReserveLevel / 10).ToString("#0.0") & "%"
        LabelReserveStatusMimic.Text = .ReserveStatus

        ' Reserve Ready Button
        If .ReserveReady Then
          buttonReserveReady.BackColor = Color.Green
          buttonReserveReady1.BackColor = Color.Green
        Else
          buttonReserveReady.BackColor = Color.WhiteSmoke
          buttonReserveReady1.BackColor = Color.WhiteSmoke
        End If

        ' Reserve Fill
        buttonReserveFill.Text = Translate("Fill")
        If .ReserveControl.IsFillActive Then
          buttonReserveFill.BackColor = Color.Green
        Else
          buttonReserveFill.BackColor = Color.Transparent
        End If

        ' Reserve Heat
        buttonReserveHeat.Text = Translate("Heat")
        If .ReserveControl.IsHeatActive Then
          buttonReserveHeat.BackColor = Color.Green
        Else
          buttonReserveHeat.BackColor = Color.Transparent
        End If

        ' Reserve Drain
        buttonReserveDrain.Text = Translate("Drain")
        If .ReserveControl.IsDrainActive Then
          buttonReserveDrain.BackColor = Color.Green
        Else
          buttonReserveDrain.BackColor = Color.Transparent
        End If

        ' Reserve Cancel
        buttonReserveCancel.Text = Translate("Cancel")

        ' Main Pump GroupBox:
        labelPumpStatusPanelButtons.Text = .PumpControl.StateString
        labelPumpOutputPanelButtons.Text = Translate("Pump Speed") & (": ") & (.IO.PumpSpeedOutput / 10).ToString("#0.0") & "%"

      End With

    Catch ex As Exception
      Dim message As String = ex.ToString
    End Try
  End Sub

#Region " PANEL MIMIC INTERFACE "

  Private Sub buttonManualControl_Click(sender As Object, e As EventArgs) Handles buttonManualControl.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      ' Assign the default position (upper left corner)
      PanelButtons.Location = New Point(5, 5)
      ' Display the button panel
      Me.PanelButtons.Visible = True
      Timer2.Interval = 30000
      Timer2.Start()
    End With
  End Sub

  Private Sub buttonClose_Click(sender As Object, e As EventArgs) Handles buttonClose.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      Me.PanelButtons.Visible = False
    End With
  End Sub

  Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
    'There's an odd bug related to boxing the integer values that creates probelms with .ToString
    '  to get round this feature we divide the values by 1 which seems to force the values to be boxed correctly
    Try
      'Make sure we have a connection
      If ControlCode Is Nothing Then Exit Sub

      If ControlCode.AddControl.IsActive Then Exit Sub
      If ControlCode.ReserveControl.IsActive Then Exit Sub

      ' Timer2 timed out - Hide the mimic panel
      If PanelButtons.Visible Then PanelButtons.Visible = False

    Catch ex As Exception
    End Try
  End Sub

  '==============================================  ADD PANEL BUTTONS/DISPLAY ==========================================

  Private Sub buttonAddReady_Click(sender As Object, e As EventArgs) Handles buttonAddReady.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .AddReady = Not .AddReady
    End With
  End Sub

  Private Sub buttonAddReady1_Click(sender As Object, e As EventArgs) Handles buttonAddReady1.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .AddReady = Not .AddReady
    End With
  End Sub

  Private Sub buttonAddFill_Click(sender As Object, e As EventArgs) Handles buttonAddFill.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      Using newForm As New FormNumberPad
        newForm.Text = Translate("Set Desired Add Fill Level") & (" (0-100)...")
        newForm.Amount = ""
        If newForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
          .AddControl.FillManual(EFillType.Cold, newForm.AmountInteger * 10)
        End If
      End Using
    End With
  End Sub

  Private Sub buttonAddDrain_Click(sender As Object, e As EventArgs) Handles buttonAddDrain.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .AddControl.DrainManual()
    End With
  End Sub

  Private Sub buttonAddCancel_Click(sender As Object, e As EventArgs) Handles buttonAddCancel.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .AddControl.Cancel()
    End With
  End Sub


  '==============================================  RESERVE PANEL BUTTONS/DISPLAY ==========================================

  Private Sub buttonReserveReady_Click(sender As Object, e As EventArgs) Handles buttonReserveReady.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .ReserveReady = Not .ReserveReady
    End With
  End Sub

  Private Sub buttonReserveReady1_Click(sender As Object, e As EventArgs) Handles buttonReserveReady1.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .ReserveReady = Not .ReserveReady
    End With
  End Sub

  Private Sub buttonReserveFill_Click(sender As Object, e As EventArgs) Handles buttonReserveFill.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      Using newForm As New FormNumberPad
        newForm.Text = Translate("Set Desired Reserve Fill Level") & (" (0-100)...")
        newForm.Amount = ""
        If newForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
          .ReserveControl.ManualFill(EFillType.Cold, newForm.AmountInteger * 10)
        End If
      End Using
    End With
  End Sub

  Private Sub buttonReserveHeat_Click(sender As Object, e As EventArgs) Handles buttonReserveHeat.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      Using newForm As New FormNumberPad
        newForm.Text = Translate("Set Desired Reserve Temp... ")
        newForm.Amount = ""
        If newForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
          Dim desiredTemp As Integer = newForm.AmountInteger * 10
          If desiredTemp > 1350 Then
            MessageBox.Show(Translate("Desired Temp is too high. Please choose a lower temp."), Translate("Reserve Heat"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
          Else
            If .ReserveLevel < 350 Then
              MessageBox.Show(Translate("Level too low to heat.  Please fill Reserve Tank"), Translate("Reserve Heat"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
              ' OK to begin heating
              .ReserveControl.ManualHeat(desiredTemp)
            End If
          End If
        End If
      End Using
    End With
  End Sub

  ' [2016-08-16] Removed this function as there are too many things that can go wrong at this time
  Private Sub buttonReserveTransfer_Click(sender As Object, e As EventArgs)
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .ReserveControl.ManualTransfer()
    End With
  End Sub

  Private Sub buttonReserveDrain_Click(sender As Object, e As EventArgs) Handles buttonReserveDrain.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .ReserveControl.ManualDrain()
    End With
  End Sub

  Private Sub buttonReserveCancel_Click(sender As Object, e As EventArgs) Handles buttonReserveCancel.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .ReserveControl.Cancel()
    End With
  End Sub


  Private Sub ButtonTank1Cancel_Click(sender As Object, e As EventArgs) Handles ButtonTank1Cancel.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    ' Temporary button to cancel kitchen activity - Delete once KP / AT issue is sorted
    With ControlCode
      .CK.Cancel()
      .KA.Cancel()
      .KP.Cancel()
      .KR.Cancel()
      .WK.Cancel()
      .LA.KP1.Cancel()
      .LA.Cancel()
    End With

  End Sub




  '================================================  MAIN PUMP ============================================

  Private Sub buttonPumpRequest_Click(sender As Object, e As EventArgs) Handles buttonPumpRequest.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      ' Check Pump Stopped or Running, If stopped - make sure it's ok to start
      If .PumpControl.IsStopped Then
        If Not .Parent.IsProgramRunning Then
          MessageBox.Show(Translate("Pump Disabled - Program Not Running"), Translate("Pump Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
          Exit Sub
        End If
        If .EStop Then
          MessageBox.Show(Translate("Pump Disabled - EStop Active"), Translate("Pump Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
          Exit Sub
        End If
        If Not .MachineClosed Then
          MessageBox.Show(Translate("Pump Disabled - Lids Not Closed"), Translate("Pump Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
          Exit Sub
        End If
        If Not .PumpControl.PumpEnable Then
          MessageBox.Show(Translate("Pump Disabled - Pump Not Enabled"), Translate("Pump Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
          Exit Sub
        End If

        ' Made it this far 
        .PumpControl.StartManual()
      Else
        ' Stop the pump
        .PumpControl.StopMainPump()
      End If

    End With
  End Sub

  Private Sub buttonPumpStart_Click(sender As Object, e As EventArgs) Handles buttonPumpStart.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      If Not .Parent.IsProgramRunning Then
        MessageBox.Show(Translate("Pump Disabled - Program Not Running"), Translate("Pump Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If .EStop Then
        MessageBox.Show(Translate("Pump Disabled - EStop Active"), Translate("Pump Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If Not .MachineClosed Then
        MessageBox.Show(Translate("Pump Disabled - Lids Not Closed"), Translate("Pump Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If Not .PumpControl.PumpEnable Then
        MessageBox.Show(Translate("Pump Disabled - Pump Not Enabled"), Translate("Pump Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If

      ' Made it this far, start the pump
      If Not .IO.MainPump Then .PumpControl.StartManual()
    End With
  End Sub

  Private Sub buttonPumpStop_Click(sender As Object, e As EventArgs) Handles buttonPumpStop.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      .PumpControl.StopMainPump()
    End With
  End Sub

  Private Sub ButtonLevelFlush_Click(sender As Object, e As EventArgs) Handles ButtonLevelFlush.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      If .Parent.IsProgramRunning Then
        MessageBox.Show(Translate("Level Flush Disabled- Program Running"), Translate("Level Flush Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If .EStop Then
        MessageBox.Show(Translate("Level Flush  Disabled - EStop Active"), Translate("Level Flush Start Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If

      ' Made it this far, start or stop
      If .LevelFlush.IsActive Then
        .LevelFlush.Cancel()
      Else : .LevelFlush.LevelFlushStartManual()
      End If

    End With
  End Sub


  '================================================  HALF-LOAD ============================================
  ' 2022-02-03 - Mt. Holly Then Platorm does not have Half Load Functionality:
#If 0 Then

  Private Sub buttonHalfLoad_Click(sender As Object, e As EventArgs) Handles buttonHalfLoad.Click
    'Make sure we have a connection
    If ControlCode Is Nothing Then Exit Sub

    'Check to see if we are remoted (Plant Explorer / History)
    If Runtime.Remoting.RemotingServices.IsTransparentProxy(ControlCode) Then Exit Sub

    With ControlCode
      If Not .Parent.IsProgramRunning Then
        MessageBox.Show(Translate("Machine Idle"), Translate("Isolate Request"), MessageBoxButtons.OK, MessageBoxIcon.Hand)
        Exit Sub
      End If
      If .Parent.IsProgramRunning AndAlso (Not .LD.IsConfirmPortSwitch) Then
        MessageBox.Show(Translate("Must be on Load Command"), Translate("Isolate Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If Not .MachineSafe Then
        MessageBox.Show(Translate("Machine not safe") & Environment.NewLine & .SafetyControl.StateString, Translate("Isolate Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If .MachineLevel > 100 Then
        MessageBox.Show(Translate("Machine Not Empty"), Translate("Isolate Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If .EStop Then
        MessageBox.Show(Translate("ESTOP"), Translate("Isolate Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If .IO.PumpRunning OrElse .IO.MainPump Then
        MessageBox.Show(Translate("Main Pump running"), Translate("Isolate Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If Not .IO.KierLidClosed Then
        MessageBox.Show(Translate("Machine Lids Not Closed"), Translate("Isolate Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If
      If Not .IO.SampleLidClosed Then
        MessageBox.Show(Translate("Sample Lid Not Closed"), Translate("Isolate Request"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Exit Sub
      End If

      ' Made it this far, switch the KierIsolate state
      .KierBLoaded = Not .KierBLoaded
      ' .HL.KierBLoaded = .KierBLoaded

    End With
  End Sub

#End If
#End Region

End Class
