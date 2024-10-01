<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Mimic
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
    Me.components = New System.ComponentModel.Container()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.buttonPumpRequest = New System.Windows.Forms.Button()
    Me.buttonManualControl = New System.Windows.Forms.Button()
    Me.buttonAddReady = New System.Windows.Forms.Button()
    Me.buttonReserveReady = New System.Windows.Forms.Button()
    Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
    Me.LabelReserveStatusMimic = New System.Windows.Forms.Label()
    Me.ButtonLevelFlush = New System.Windows.Forms.Button()
    Me.labelMachineLevel = New MimicControls.Label()
    Me.IO_ReserveRunback = New MimicDeviceValve()
    Me.IO_HoldDown = New MimicControls.Output()
    Me.IO_FlowReverse = New MimicControls.Output()
    Me.IO_FillHot = New MimicDeviceValve()
    Me.IO_MachineDrainHot = New MimicDeviceValve()
    Me.levelBarTank1 = New MimicControls.LevelBar()
    Me.levelBarAdd = New MimicControls.LevelBar()
    Me.levelBarReserve = New MimicControls.LevelBar()
    Me.PanelButtons = New MimicControls.Panel()
    Me.CheckBox1 = New System.Windows.Forms.CheckBox()
    Me.ButtonTank1Cancel = New System.Windows.Forms.Button()
    Me.buttonHalfLoad = New System.Windows.Forms.Button()
    Me.GroupBoxPump = New System.Windows.Forms.GroupBox()
    Me.labelPumpOutputPanelButtons = New MimicControls.Label()
    Me.buttonPumpStop = New System.Windows.Forms.Button()
    Me.buttonPumpStart = New System.Windows.Forms.Button()
    Me.labelPumpStatusPanelButtons = New MimicControls.Label()
    Me.GroupBoxReserveTank = New System.Windows.Forms.GroupBox()
    Me.buttonReserveCancel = New System.Windows.Forms.Button()
    Me.buttonReserveHeat = New System.Windows.Forms.Button()
    Me.labelReserveLevelPanelButtons = New MimicControls.Label()
    Me.buttonReserveDrain = New System.Windows.Forms.Button()
    Me.buttonReserveFill = New System.Windows.Forms.Button()
    Me.buttonReserveReady1 = New System.Windows.Forms.Button()
    Me.labelReserveStatusPanelButtons = New MimicControls.Label()
    Me.buttonClose = New System.Windows.Forms.Button()
    Me.GroupBoxAddTank = New System.Windows.Forms.GroupBox()
    Me.buttonAddCancel = New System.Windows.Forms.Button()
    Me.labelAddLevelPanelButtons = New MimicControls.Label()
    Me.buttonAddDrain = New System.Windows.Forms.Button()
    Me.buttonAddFill = New System.Windows.Forms.Button()
    Me.buttonAddReady1 = New System.Windows.Forms.Button()
    Me.labelAddStatusPanelButtons = New MimicControls.Label()
    Me.labelPumpReversalStatus = New MimicControls.Label()
    Me.labelPumpStatus = New MimicControls.Label()
    Me.labelMachinePressure = New MimicControls.Label()
    Me.labelPackageDP = New MimicControls.Label()
    Me.levelBarMachine = New MimicControls.LevelBar()
    Me.labelAddLevelMimic = New MimicControls.Label()
    Me.labelReserveLevelMimic = New MimicControls.Label()
    Me.IO_AddRunback = New MimicDeviceValve()
    Me.IO_Tank1FillCold = New MimicDeviceValve()
    Me.IO_Tank1Transfer = New MimicDeviceValve()
    Me.IO_ReserveHeat = New MimicDeviceValve()
    Me.IO_ReserveFillCold = New MimicDeviceValve()
    Me.IO_ReserveDrain = New MimicDeviceValve()
    Me.IO_ReserveTransfer = New MimicDeviceValve()
    Me.IO_AddTransfer = New MimicDeviceValve()
    Me.IO_AddMixing = New MimicDeviceValve()
    Me.IO_AddDrain = New MimicDeviceValve()
    Me.IO_AddFillCold = New MimicDeviceValve()
    Me.IO_Tank1TransferToReserve = New MimicDeviceValve()
    Me.IO_Tank1TransferToDrain = New MimicDeviceValve()
    Me.IO_Tank1TransferToAdd = New MimicDeviceValve()
    Me.IO_SystemBlock = New MimicDeviceValve()
    Me.IO_MachineDrain = New MimicDeviceValve()
    Me.IO_FillCold = New MimicDeviceValve()
    Me.IO_Condensate = New MimicDeviceValve()
    Me.IO_WaterDump = New MimicDeviceValve()
    Me.IO_CoolSelect = New MimicDeviceValve()
    Me.IO_HeatSelect = New MimicDeviceValve()
    Me.IO_Airpad = New MimicDeviceValve()
    Me.IO_Vent = New MimicDeviceValve()
    Me.IO_TopWash = New MimicDeviceValve()
    Me.IO_KierLidClosed = New MimicControls.Input()
    Me.IO_AddPumpRunning = New MimicDevicePump()
    Me.IO_PumpRunning = New MimicDevicePump()
    Me.IO_TempInterlock = New MimicControls.Input()
    Me.LblTank1Status = New MimicControls.Label()
    Me.LblTank1Level = New MimicControls.Label()
    Me.HeatCoolOutput = New MimicControls.HeatExchanger()
    Me.labelAddStatusMimic = New MimicControls.Label()
    Me.PanelButtons.SuspendLayout()
    Me.GroupBoxPump.SuspendLayout()
    Me.GroupBoxReserveTank.SuspendLayout()
    Me.GroupBoxAddTank.SuspendLayout()
    Me.SuspendLayout()
    '
    'Timer1
    '
    Me.Timer1.Enabled = True
    Me.Timer1.Interval = 1000
    '
    'buttonPumpRequest
    '
    Me.buttonPumpRequest.Location = New System.Drawing.Point(5, 598)
    Me.buttonPumpRequest.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonPumpRequest.Name = "buttonPumpRequest"
    Me.buttonPumpRequest.Size = New System.Drawing.Size(80, 49)
    Me.buttonPumpRequest.TabIndex = 181
    Me.buttonPumpRequest.Text = "Pump Start"
    Me.buttonPumpRequest.UseVisualStyleBackColor = True
    '
    'buttonManualControl
    '
    Me.buttonManualControl.Location = New System.Drawing.Point(7, 6)
    Me.buttonManualControl.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonManualControl.Name = "buttonManualControl"
    Me.buttonManualControl.Size = New System.Drawing.Size(80, 49)
    Me.buttonManualControl.TabIndex = 235
    Me.buttonManualControl.Text = "Manual Control"
    Me.buttonManualControl.UseVisualStyleBackColor = True
    '
    'buttonAddReady
    '
    Me.buttonAddReady.Location = New System.Drawing.Point(563, 132)
    Me.buttonAddReady.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonAddReady.Name = "buttonAddReady"
    Me.buttonAddReady.Size = New System.Drawing.Size(80, 49)
    Me.buttonAddReady.TabIndex = 182
    Me.buttonAddReady.Text = "Add Ready"
    Me.buttonAddReady.UseVisualStyleBackColor = True
    '
    'buttonReserveReady
    '
    Me.buttonReserveReady.Location = New System.Drawing.Point(563, 76)
    Me.buttonReserveReady.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonReserveReady.Name = "buttonReserveReady"
    Me.buttonReserveReady.Size = New System.Drawing.Size(80, 49)
    Me.buttonReserveReady.TabIndex = 183
    Me.buttonReserveReady.Text = "Reserve Ready"
    Me.buttonReserveReady.UseVisualStyleBackColor = True
    '
    'Timer2
    '
    Me.Timer2.Interval = 30000
    '
    'LabelReserveStatusMimic
    '
    Me.LabelReserveStatusMimic.AutoSize = True
    Me.LabelReserveStatusMimic.BackColor = System.Drawing.Color.Transparent
    Me.LabelReserveStatusMimic.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.LabelReserveStatusMimic.Location = New System.Drawing.Point(648, 79)
    Me.LabelReserveStatusMimic.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
    Me.LabelReserveStatusMimic.Name = "LabelReserveStatusMimic"
    Me.LabelReserveStatusMimic.Size = New System.Drawing.Size(124, 20)
    Me.LabelReserveStatusMimic.TabIndex = 239
    Me.LabelReserveStatusMimic.Text = "Reserve Status"
    Me.LabelReserveStatusMimic.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'ButtonLevelFlush
    '
    Me.ButtonLevelFlush.BackColor = System.Drawing.Color.WhiteSmoke
    Me.ButtonLevelFlush.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
    Me.ButtonLevelFlush.Location = New System.Drawing.Point(7, 325)
    Me.ButtonLevelFlush.Margin = New System.Windows.Forms.Padding(4)
    Me.ButtonLevelFlush.Name = "ButtonLevelFlush"
    Me.ButtonLevelFlush.Size = New System.Drawing.Size(80, 49)
    Me.ButtonLevelFlush.TabIndex = 385
    Me.ButtonLevelFlush.Text = "Level Flush"
    Me.ButtonLevelFlush.UseVisualStyleBackColor = False
    '
    'labelMachineLevel
    '
    Me.labelMachineLevel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.labelMachineLevel.ForeColor = System.Drawing.Color.Black
    Me.labelMachineLevel.Location = New System.Drawing.Point(280, 354)
    Me.labelMachineLevel.Margin = New System.Windows.Forms.Padding(4)
    Me.labelMachineLevel.Name = "labelMachineLevel"
    Me.labelMachineLevel.Size = New System.Drawing.Size(51, 20)
    Me.labelMachineLevel.TabIndex = 296
    Me.labelMachineLevel.Text = "000%"
    Me.labelMachineLevel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'IO_ReserveRunback
    '
    Me.IO_ReserveRunback.Location = New System.Drawing.Point(927, 463)
    Me.IO_ReserveRunback.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_ReserveRunback.Name = "IO_ReserveRunback"
    Me.IO_ReserveRunback.NormallyOn = False
    Me.IO_ReserveRunback.OffFeedback = False
    Me.IO_ReserveRunback.OffFeedbackEnabled = False
    Me.IO_ReserveRunback.OnFeedback = False
    Me.IO_ReserveRunback.OnFeedbackEnabled = False
    Me.IO_ReserveRunback.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_ReserveRunback.Overridden = False
    Me.IO_ReserveRunback.Size = New System.Drawing.Size(13, 25)
    Me.IO_ReserveRunback.TabIndex = 295
    Me.IO_ReserveRunback.Text = "MimicDeviceValve1"
    Me.IO_ReserveRunback.UIEnabled = False
    '
    'IO_HoldDown
    '
    Me.IO_HoldDown.Location = New System.Drawing.Point(199, 138)
    Me.IO_HoldDown.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_HoldDown.Name = "IO_HoldDown"
    Me.IO_HoldDown.Size = New System.Drawing.Size(21, 20)
    Me.IO_HoldDown.TabIndex = 294
    '
    'IO_FlowReverse
    '
    Me.IO_FlowReverse.Location = New System.Drawing.Point(161, 400)
    Me.IO_FlowReverse.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_FlowReverse.Name = "IO_FlowReverse"
    Me.IO_FlowReverse.Size = New System.Drawing.Size(21, 20)
    Me.IO_FlowReverse.TabIndex = 293
    '
    'IO_FillHot
    '
    Me.IO_FillHot.Location = New System.Drawing.Point(373, 618)
    Me.IO_FillHot.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_FillHot.Name = "IO_FillHot"
    Me.IO_FillHot.NormallyOn = False
    Me.IO_FillHot.OffFeedback = False
    Me.IO_FillHot.OffFeedbackEnabled = False
    Me.IO_FillHot.OnFeedback = False
    Me.IO_FillHot.OnFeedbackEnabled = False
    Me.IO_FillHot.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_FillHot.Overridden = False
    Me.IO_FillHot.Size = New System.Drawing.Size(25, 13)
    Me.IO_FillHot.TabIndex = 292
    Me.IO_FillHot.Text = "MimicDeviceValve1"
    Me.IO_FillHot.UIEnabled = False
    '
    'IO_MachineDrainHot
    '
    Me.IO_MachineDrainHot.Location = New System.Drawing.Point(693, 578)
    Me.IO_MachineDrainHot.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_MachineDrainHot.Name = "IO_MachineDrainHot"
    Me.IO_MachineDrainHot.NormallyOn = False
    Me.IO_MachineDrainHot.OffFeedback = False
    Me.IO_MachineDrainHot.OffFeedbackEnabled = False
    Me.IO_MachineDrainHot.OnFeedback = False
    Me.IO_MachineDrainHot.OnFeedbackEnabled = False
    Me.IO_MachineDrainHot.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_MachineDrainHot.Overridden = False
    Me.IO_MachineDrainHot.Size = New System.Drawing.Size(25, 13)
    Me.IO_MachineDrainHot.TabIndex = 291
    Me.IO_MachineDrainHot.Text = "MimicDeviceValve1"
    Me.IO_MachineDrainHot.UIEnabled = False
    '
    'levelBarTank1
    '
    Me.levelBarTank1.EndColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarTank1.ForeColor = System.Drawing.Color.Black
    Me.levelBarTank1.Format = Nothing
    Me.levelBarTank1.Location = New System.Drawing.Point(983, 55)
    Me.levelBarTank1.Margin = New System.Windows.Forms.Padding(4)
    Me.levelBarTank1.MiddleColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarTank1.Name = "levelBarTank1"
    Me.levelBarTank1.Size = New System.Drawing.Size(35, 46)
    Me.levelBarTank1.StartColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarTank1.TabIndex = 288
    '
    'levelBarAdd
    '
    Me.levelBarAdd.EndColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarAdd.ForeColor = System.Drawing.Color.Black
    Me.levelBarAdd.Format = Nothing
    Me.levelBarAdd.Location = New System.Drawing.Point(847, 283)
    Me.levelBarAdd.Margin = New System.Windows.Forms.Padding(4)
    Me.levelBarAdd.MiddleColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarAdd.Name = "levelBarAdd"
    Me.levelBarAdd.Size = New System.Drawing.Size(13, 33)
    Me.levelBarAdd.StartColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarAdd.TabIndex = 287
    '
    'levelBarReserve
    '
    Me.levelBarReserve.EndColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarReserve.ForeColor = System.Drawing.Color.Black
    Me.levelBarReserve.Format = Nothing
    Me.levelBarReserve.Location = New System.Drawing.Point(923, 258)
    Me.levelBarReserve.Margin = New System.Windows.Forms.Padding(4)
    Me.levelBarReserve.MiddleColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarReserve.Name = "levelBarReserve"
    Me.levelBarReserve.Size = New System.Drawing.Size(25, 74)
    Me.levelBarReserve.StartColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarReserve.TabIndex = 286
    '
    'PanelButtons
    '
    Me.PanelButtons.BackColor = System.Drawing.SystemColors.ControlDark
    Me.PanelButtons.Border.InnerBorder = True
    Me.PanelButtons.Border.Style = MimicControls.BorderStyle.DoubleRaised
    Me.PanelButtons.Controls.Add(Me.CheckBox1)
    Me.PanelButtons.Controls.Add(Me.ButtonTank1Cancel)
    Me.PanelButtons.Controls.Add(Me.buttonHalfLoad)
    Me.PanelButtons.Controls.Add(Me.GroupBoxPump)
    Me.PanelButtons.Controls.Add(Me.GroupBoxReserveTank)
    Me.PanelButtons.Controls.Add(Me.buttonClose)
    Me.PanelButtons.Controls.Add(Me.GroupBoxAddTank)
    Me.PanelButtons.Location = New System.Drawing.Point(93, 102)
    Me.PanelButtons.Margin = New System.Windows.Forms.Padding(4)
    Me.PanelButtons.Name = "PanelButtons"
    Me.PanelButtons.Size = New System.Drawing.Size(733, 468)
    Me.PanelButtons.TabIndex = 236
    Me.PanelButtons.Visible = False
    '
    'CheckBox1
    '
    Me.CheckBox1.AutoSize = True
    Me.CheckBox1.Location = New System.Drawing.Point(497, 256)
    Me.CheckBox1.Name = "CheckBox1"
    Me.CheckBox1.Size = New System.Drawing.Size(100, 21)
    Me.CheckBox1.TabIndex = 241
    Me.CheckBox1.Text = "CheckBox1"
    Me.CheckBox1.UseVisualStyleBackColor = True
    '
    'ButtonTank1Cancel
    '
    Me.ButtonTank1Cancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.ButtonTank1Cancel.Location = New System.Drawing.Point(497, 314)
    Me.ButtonTank1Cancel.Margin = New System.Windows.Forms.Padding(4)
    Me.ButtonTank1Cancel.Name = "ButtonTank1Cancel"
    Me.ButtonTank1Cancel.Size = New System.Drawing.Size(215, 42)
    Me.ButtonTank1Cancel.TabIndex = 240
    Me.ButtonTank1Cancel.Text = "Kitchen Cancel"
    Me.ButtonTank1Cancel.UseVisualStyleBackColor = True
    '
    'buttonHalfLoad
    '
    Me.buttonHalfLoad.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonHalfLoad.Location = New System.Drawing.Point(497, 375)
    Me.buttonHalfLoad.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonHalfLoad.Name = "buttonHalfLoad"
    Me.buttonHalfLoad.Size = New System.Drawing.Size(215, 42)
    Me.buttonHalfLoad.TabIndex = 239
    Me.buttonHalfLoad.Text = "Half-Load"
    Me.buttonHalfLoad.UseVisualStyleBackColor = True
    '
    'GroupBoxPump
    '
    Me.GroupBoxPump.Controls.Add(Me.labelPumpOutputPanelButtons)
    Me.GroupBoxPump.Controls.Add(Me.buttonPumpStop)
    Me.GroupBoxPump.Controls.Add(Me.buttonPumpStart)
    Me.GroupBoxPump.Controls.Add(Me.labelPumpStatusPanelButtons)
    Me.GroupBoxPump.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.GroupBoxPump.Location = New System.Drawing.Point(487, 10)
    Me.GroupBoxPump.Margin = New System.Windows.Forms.Padding(4)
    Me.GroupBoxPump.Name = "GroupBoxPump"
    Me.GroupBoxPump.Padding = New System.Windows.Forms.Padding(4)
    Me.GroupBoxPump.Size = New System.Drawing.Size(233, 207)
    Me.GroupBoxPump.TabIndex = 238
    Me.GroupBoxPump.TabStop = False
    Me.GroupBoxPump.Text = "Main Pump"
    '
    'labelPumpOutputPanelButtons
    '
    Me.labelPumpOutputPanelButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.labelPumpOutputPanelButtons.ForeColor = System.Drawing.Color.Black
    Me.labelPumpOutputPanelButtons.Location = New System.Drawing.Point(8, 62)
    Me.labelPumpOutputPanelButtons.Margin = New System.Windows.Forms.Padding(4)
    Me.labelPumpOutputPanelButtons.Name = "labelPumpOutputPanelButtons"
    Me.labelPumpOutputPanelButtons.Size = New System.Drawing.Size(106, 20)
    Me.labelPumpOutputPanelButtons.TabIndex = 241
    Me.labelPumpOutputPanelButtons.Text = "Output 000%"
    Me.labelPumpOutputPanelButtons.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'buttonPumpStop
    '
    Me.buttonPumpStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonPumpStop.Location = New System.Drawing.Point(8, 143)
    Me.buttonPumpStop.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonPumpStop.Name = "buttonPumpStop"
    Me.buttonPumpStop.Size = New System.Drawing.Size(217, 49)
    Me.buttonPumpStop.TabIndex = 238
    Me.buttonPumpStop.Text = "Stop"
    Me.buttonPumpStop.UseVisualStyleBackColor = True
    '
    'buttonPumpStart
    '
    Me.buttonPumpStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonPumpStart.Location = New System.Drawing.Point(8, 92)
    Me.buttonPumpStart.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonPumpStart.Name = "buttonPumpStart"
    Me.buttonPumpStart.Size = New System.Drawing.Size(217, 49)
    Me.buttonPumpStart.TabIndex = 237
    Me.buttonPumpStart.Text = "Start"
    Me.buttonPumpStart.UseVisualStyleBackColor = True
    '
    'labelPumpStatusPanelButtons
    '
    Me.labelPumpStatusPanelButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.labelPumpStatusPanelButtons.ForeColor = System.Drawing.Color.Black
    Me.labelPumpStatusPanelButtons.Location = New System.Drawing.Point(8, 31)
    Me.labelPumpStatusPanelButtons.Margin = New System.Windows.Forms.Padding(4)
    Me.labelPumpStatusPanelButtons.Name = "labelPumpStatusPanelButtons"
    Me.labelPumpStatusPanelButtons.Size = New System.Drawing.Size(57, 20)
    Me.labelPumpStatusPanelButtons.TabIndex = 234
    Me.labelPumpStatusPanelButtons.Text = "Status"
    Me.labelPumpStatusPanelButtons.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'GroupBoxReserveTank
    '
    Me.GroupBoxReserveTank.Controls.Add(Me.buttonReserveCancel)
    Me.GroupBoxReserveTank.Controls.Add(Me.buttonReserveHeat)
    Me.GroupBoxReserveTank.Controls.Add(Me.labelReserveLevelPanelButtons)
    Me.GroupBoxReserveTank.Controls.Add(Me.buttonReserveDrain)
    Me.GroupBoxReserveTank.Controls.Add(Me.buttonReserveFill)
    Me.GroupBoxReserveTank.Controls.Add(Me.buttonReserveReady1)
    Me.GroupBoxReserveTank.Controls.Add(Me.labelReserveStatusPanelButtons)
    Me.GroupBoxReserveTank.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.GroupBoxReserveTank.Location = New System.Drawing.Point(245, 10)
    Me.GroupBoxReserveTank.Margin = New System.Windows.Forms.Padding(4)
    Me.GroupBoxReserveTank.Name = "GroupBoxReserveTank"
    Me.GroupBoxReserveTank.Padding = New System.Windows.Forms.Padding(4)
    Me.GroupBoxReserveTank.Size = New System.Drawing.Size(233, 407)
    Me.GroupBoxReserveTank.TabIndex = 237
    Me.GroupBoxReserveTank.TabStop = False
    Me.GroupBoxReserveTank.Text = "Reserve Tank"
    '
    'buttonReserveCancel
    '
    Me.buttonReserveCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonReserveCancel.Location = New System.Drawing.Point(8, 348)
    Me.buttonReserveCancel.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonReserveCancel.Name = "buttonReserveCancel"
    Me.buttonReserveCancel.Size = New System.Drawing.Size(217, 49)
    Me.buttonReserveCancel.TabIndex = 243
    Me.buttonReserveCancel.Text = "Cancel"
    Me.buttonReserveCancel.UseVisualStyleBackColor = True
    '
    'buttonReserveHeat
    '
    Me.buttonReserveHeat.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonReserveHeat.Location = New System.Drawing.Point(8, 246)
    Me.buttonReserveHeat.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonReserveHeat.Name = "buttonReserveHeat"
    Me.buttonReserveHeat.Size = New System.Drawing.Size(217, 49)
    Me.buttonReserveHeat.TabIndex = 242
    Me.buttonReserveHeat.Text = "Heat"
    Me.buttonReserveHeat.UseVisualStyleBackColor = True
    '
    'labelReserveLevelPanelButtons
    '
    Me.labelReserveLevelPanelButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.labelReserveLevelPanelButtons.ForeColor = System.Drawing.Color.Black
    Me.labelReserveLevelPanelButtons.Location = New System.Drawing.Point(8, 62)
    Me.labelReserveLevelPanelButtons.Margin = New System.Windows.Forms.Padding(4)
    Me.labelReserveLevelPanelButtons.Name = "labelReserveLevelPanelButtons"
    Me.labelReserveLevelPanelButtons.Size = New System.Drawing.Size(49, 20)
    Me.labelReserveLevelPanelButtons.TabIndex = 241
    Me.labelReserveLevelPanelButtons.Text = "Level"
    Me.labelReserveLevelPanelButtons.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'buttonReserveDrain
    '
    Me.buttonReserveDrain.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonReserveDrain.Location = New System.Drawing.Point(8, 194)
    Me.buttonReserveDrain.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonReserveDrain.Name = "buttonReserveDrain"
    Me.buttonReserveDrain.Size = New System.Drawing.Size(217, 49)
    Me.buttonReserveDrain.TabIndex = 239
    Me.buttonReserveDrain.Text = "Drain"
    Me.buttonReserveDrain.UseVisualStyleBackColor = True
    '
    'buttonReserveFill
    '
    Me.buttonReserveFill.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonReserveFill.Location = New System.Drawing.Point(8, 143)
    Me.buttonReserveFill.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonReserveFill.Name = "buttonReserveFill"
    Me.buttonReserveFill.Size = New System.Drawing.Size(217, 49)
    Me.buttonReserveFill.TabIndex = 238
    Me.buttonReserveFill.Text = "Fill"
    Me.buttonReserveFill.UseVisualStyleBackColor = True
    '
    'buttonReserveReady1
    '
    Me.buttonReserveReady1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonReserveReady1.Location = New System.Drawing.Point(8, 92)
    Me.buttonReserveReady1.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonReserveReady1.Name = "buttonReserveReady1"
    Me.buttonReserveReady1.Size = New System.Drawing.Size(217, 49)
    Me.buttonReserveReady1.TabIndex = 237
    Me.buttonReserveReady1.Text = "Ready"
    Me.buttonReserveReady1.UseVisualStyleBackColor = True
    '
    'labelReserveStatusPanelButtons
    '
    Me.labelReserveStatusPanelButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.labelReserveStatusPanelButtons.ForeColor = System.Drawing.Color.Black
    Me.labelReserveStatusPanelButtons.Location = New System.Drawing.Point(8, 31)
    Me.labelReserveStatusPanelButtons.Margin = New System.Windows.Forms.Padding(4)
    Me.labelReserveStatusPanelButtons.Name = "labelReserveStatusPanelButtons"
    Me.labelReserveStatusPanelButtons.Size = New System.Drawing.Size(57, 20)
    Me.labelReserveStatusPanelButtons.TabIndex = 234
    Me.labelReserveStatusPanelButtons.Text = "Status"
    Me.labelReserveStatusPanelButtons.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'buttonClose
    '
    Me.buttonClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.buttonClose.Location = New System.Drawing.Point(12, 422)
    Me.buttonClose.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonClose.Name = "buttonClose"
    Me.buttonClose.Size = New System.Drawing.Size(713, 39)
    Me.buttonClose.TabIndex = 236
    Me.buttonClose.Text = "Close"
    Me.buttonClose.UseVisualStyleBackColor = True
    '
    'GroupBoxAddTank
    '
    Me.GroupBoxAddTank.Controls.Add(Me.buttonAddCancel)
    Me.GroupBoxAddTank.Controls.Add(Me.labelAddLevelPanelButtons)
    Me.GroupBoxAddTank.Controls.Add(Me.buttonAddDrain)
    Me.GroupBoxAddTank.Controls.Add(Me.buttonAddFill)
    Me.GroupBoxAddTank.Controls.Add(Me.buttonAddReady1)
    Me.GroupBoxAddTank.Controls.Add(Me.labelAddStatusPanelButtons)
    Me.GroupBoxAddTank.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.GroupBoxAddTank.Location = New System.Drawing.Point(4, 10)
    Me.GroupBoxAddTank.Margin = New System.Windows.Forms.Padding(4)
    Me.GroupBoxAddTank.Name = "GroupBoxAddTank"
    Me.GroupBoxAddTank.Padding = New System.Windows.Forms.Padding(4)
    Me.GroupBoxAddTank.Size = New System.Drawing.Size(233, 407)
    Me.GroupBoxAddTank.TabIndex = 0
    Me.GroupBoxAddTank.TabStop = False
    Me.GroupBoxAddTank.Text = "Add Tank"
    '
    'buttonAddCancel
    '
    Me.buttonAddCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonAddCancel.Location = New System.Drawing.Point(8, 351)
    Me.buttonAddCancel.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonAddCancel.Name = "buttonAddCancel"
    Me.buttonAddCancel.Size = New System.Drawing.Size(217, 49)
    Me.buttonAddCancel.TabIndex = 244
    Me.buttonAddCancel.Text = "Cancel"
    Me.buttonAddCancel.UseVisualStyleBackColor = True
    '
    'labelAddLevelPanelButtons
    '
    Me.labelAddLevelPanelButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.labelAddLevelPanelButtons.ForeColor = System.Drawing.Color.Black
    Me.labelAddLevelPanelButtons.Location = New System.Drawing.Point(8, 62)
    Me.labelAddLevelPanelButtons.Margin = New System.Windows.Forms.Padding(4)
    Me.labelAddLevelPanelButtons.Name = "labelAddLevelPanelButtons"
    Me.labelAddLevelPanelButtons.Size = New System.Drawing.Size(49, 20)
    Me.labelAddLevelPanelButtons.TabIndex = 241
    Me.labelAddLevelPanelButtons.Text = "Level"
    Me.labelAddLevelPanelButtons.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'buttonAddDrain
    '
    Me.buttonAddDrain.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonAddDrain.Location = New System.Drawing.Point(8, 194)
    Me.buttonAddDrain.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonAddDrain.Name = "buttonAddDrain"
    Me.buttonAddDrain.Size = New System.Drawing.Size(217, 49)
    Me.buttonAddDrain.TabIndex = 239
    Me.buttonAddDrain.Text = "Drain"
    Me.buttonAddDrain.UseVisualStyleBackColor = True
    '
    'buttonAddFill
    '
    Me.buttonAddFill.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonAddFill.Location = New System.Drawing.Point(8, 143)
    Me.buttonAddFill.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonAddFill.Name = "buttonAddFill"
    Me.buttonAddFill.Size = New System.Drawing.Size(217, 49)
    Me.buttonAddFill.TabIndex = 238
    Me.buttonAddFill.Text = "Fill"
    Me.buttonAddFill.UseVisualStyleBackColor = True
    '
    'buttonAddReady1
    '
    Me.buttonAddReady1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.buttonAddReady1.Location = New System.Drawing.Point(8, 92)
    Me.buttonAddReady1.Margin = New System.Windows.Forms.Padding(4)
    Me.buttonAddReady1.Name = "buttonAddReady1"
    Me.buttonAddReady1.Size = New System.Drawing.Size(217, 49)
    Me.buttonAddReady1.TabIndex = 237
    Me.buttonAddReady1.Text = "Ready"
    Me.buttonAddReady1.UseVisualStyleBackColor = True
    '
    'labelAddStatusPanelButtons
    '
    Me.labelAddStatusPanelButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.labelAddStatusPanelButtons.ForeColor = System.Drawing.Color.Black
    Me.labelAddStatusPanelButtons.Location = New System.Drawing.Point(8, 31)
    Me.labelAddStatusPanelButtons.Margin = New System.Windows.Forms.Padding(4)
    Me.labelAddStatusPanelButtons.Name = "labelAddStatusPanelButtons"
    Me.labelAddStatusPanelButtons.Size = New System.Drawing.Size(57, 20)
    Me.labelAddStatusPanelButtons.TabIndex = 234
    Me.labelAddStatusPanelButtons.Text = "Status"
    Me.labelAddStatusPanelButtons.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'labelPumpReversalStatus
    '
    Me.labelPumpReversalStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.labelPumpReversalStatus.ForeColor = System.Drawing.Color.Black
    Me.labelPumpReversalStatus.Location = New System.Drawing.Point(163, 454)
    Me.labelPumpReversalStatus.Margin = New System.Windows.Forms.Padding(4)
    Me.labelPumpReversalStatus.Name = "labelPumpReversalStatus"
    Me.labelPumpReversalStatus.Size = New System.Drawing.Size(128, 20)
    Me.labelPumpReversalStatus.TabIndex = 233
    Me.labelPumpReversalStatus.Text = "Reversal Status"
    Me.labelPumpReversalStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'labelPumpStatus
    '
    Me.labelPumpStatus.ForeColor = System.Drawing.Color.Black
    Me.labelPumpStatus.Location = New System.Drawing.Point(93, 625)
    Me.labelPumpStatus.Margin = New System.Windows.Forms.Padding(4)
    Me.labelPumpStatus.Name = "labelPumpStatus"
    Me.labelPumpStatus.Size = New System.Drawing.Size(88, 17)
    Me.labelPumpStatus.TabIndex = 232
    Me.labelPumpStatus.Text = "Pump Status"
    Me.labelPumpStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'labelMachinePressure
    '
    Me.labelMachinePressure.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.labelMachinePressure.ForeColor = System.Drawing.Color.Black
    Me.labelMachinePressure.Location = New System.Drawing.Point(103, 17)
    Me.labelMachinePressure.Margin = New System.Windows.Forms.Padding(4)
    Me.labelMachinePressure.Name = "labelMachinePressure"
    Me.labelMachinePressure.Size = New System.Drawing.Size(206, 20)
    Me.labelMachinePressure.TabIndex = 231
    Me.labelMachinePressure.Text = "Machine Pressure: 000bar"
    Me.labelMachinePressure.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'labelPackageDP
    '
    Me.labelPackageDP.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.labelPackageDP.ForeColor = System.Drawing.Color.Black
    Me.labelPackageDP.Location = New System.Drawing.Point(161, 427)
    Me.labelPackageDP.Margin = New System.Windows.Forms.Padding(4)
    Me.labelPackageDP.Name = "labelPackageDP"
    Me.labelPackageDP.Size = New System.Drawing.Size(93, 20)
    Me.labelPackageDP.TabIndex = 230
    Me.labelPackageDP.Text = "DP: 000mb"
    Me.labelPackageDP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'levelBarMachine
    '
    Me.levelBarMachine.EndColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarMachine.ForeColor = System.Drawing.Color.Black
    Me.levelBarMachine.Format = Nothing
    Me.levelBarMachine.Location = New System.Drawing.Point(259, 238)
    Me.levelBarMachine.Margin = New System.Windows.Forms.Padding(4)
    Me.levelBarMachine.MiddleColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarMachine.Name = "levelBarMachine"
    Me.levelBarMachine.Size = New System.Drawing.Size(13, 137)
    Me.levelBarMachine.StartColor = System.Drawing.Color.DeepSkyBlue
    Me.levelBarMachine.TabIndex = 228
    '
    'labelAddLevelMimic
    '
    Me.labelAddLevelMimic.ForeColor = System.Drawing.Color.Black
    Me.labelAddLevelMimic.Location = New System.Drawing.Point(648, 159)
    Me.labelAddLevelMimic.Margin = New System.Windows.Forms.Padding(4)
    Me.labelAddLevelMimic.Name = "labelAddLevelMimic"
    Me.labelAddLevelMimic.Size = New System.Drawing.Size(44, 17)
    Me.labelAddLevelMimic.TabIndex = 227
    Me.labelAddLevelMimic.Text = "000%"
    Me.labelAddLevelMimic.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'labelReserveLevelMimic
    '
    Me.labelReserveLevelMimic.ForeColor = System.Drawing.Color.Black
    Me.labelReserveLevelMimic.Location = New System.Drawing.Point(648, 103)
    Me.labelReserveLevelMimic.Margin = New System.Windows.Forms.Padding(4)
    Me.labelReserveLevelMimic.Name = "labelReserveLevelMimic"
    Me.labelReserveLevelMimic.Size = New System.Drawing.Size(44, 17)
    Me.labelReserveLevelMimic.TabIndex = 226
    Me.labelReserveLevelMimic.Text = "000%"
    Me.labelReserveLevelMimic.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'IO_AddRunback
    '
    Me.IO_AddRunback.Location = New System.Drawing.Point(693, 303)
    Me.IO_AddRunback.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_AddRunback.Name = "IO_AddRunback"
    Me.IO_AddRunback.NormallyOn = False
    Me.IO_AddRunback.OffFeedback = False
    Me.IO_AddRunback.OffFeedbackEnabled = False
    Me.IO_AddRunback.OnFeedback = False
    Me.IO_AddRunback.OnFeedbackEnabled = False
    Me.IO_AddRunback.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_AddRunback.Overridden = False
    Me.IO_AddRunback.Size = New System.Drawing.Size(25, 13)
    Me.IO_AddRunback.TabIndex = 225
    Me.IO_AddRunback.Text = "MimicDeviceValve1"
    Me.IO_AddRunback.UIEnabled = False
    '
    'IO_Tank1FillCold
    '
    Me.IO_Tank1FillCold.Location = New System.Drawing.Point(1008, 6)
    Me.IO_Tank1FillCold.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_Tank1FillCold.Name = "IO_Tank1FillCold"
    Me.IO_Tank1FillCold.NormallyOn = False
    Me.IO_Tank1FillCold.OffFeedback = False
    Me.IO_Tank1FillCold.OffFeedbackEnabled = False
    Me.IO_Tank1FillCold.OnFeedback = False
    Me.IO_Tank1FillCold.OnFeedbackEnabled = False
    Me.IO_Tank1FillCold.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_Tank1FillCold.Overridden = False
    Me.IO_Tank1FillCold.Size = New System.Drawing.Size(13, 25)
    Me.IO_Tank1FillCold.TabIndex = 224
    Me.IO_Tank1FillCold.Text = "MimicDeviceValve1"
    Me.IO_Tank1FillCold.UIEnabled = False
    '
    'IO_Tank1Transfer
    '
    Me.IO_Tank1Transfer.Location = New System.Drawing.Point(991, 121)
    Me.IO_Tank1Transfer.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_Tank1Transfer.Name = "IO_Tank1Transfer"
    Me.IO_Tank1Transfer.NormallyOn = False
    Me.IO_Tank1Transfer.OffFeedback = False
    Me.IO_Tank1Transfer.OffFeedbackEnabled = False
    Me.IO_Tank1Transfer.OnFeedback = False
    Me.IO_Tank1Transfer.OnFeedbackEnabled = False
    Me.IO_Tank1Transfer.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_Tank1Transfer.Overridden = False
    Me.IO_Tank1Transfer.Size = New System.Drawing.Size(13, 25)
    Me.IO_Tank1Transfer.TabIndex = 223
    Me.IO_Tank1Transfer.Text = "MimicDeviceValve1"
    Me.IO_Tank1Transfer.UIEnabled = False
    '
    'IO_ReserveHeat
    '
    Me.IO_ReserveHeat.Location = New System.Drawing.Point(1016, 343)
    Me.IO_ReserveHeat.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_ReserveHeat.Name = "IO_ReserveHeat"
    Me.IO_ReserveHeat.NormallyOn = False
    Me.IO_ReserveHeat.OffFeedback = False
    Me.IO_ReserveHeat.OffFeedbackEnabled = False
    Me.IO_ReserveHeat.OnFeedback = False
    Me.IO_ReserveHeat.OnFeedbackEnabled = False
    Me.IO_ReserveHeat.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_ReserveHeat.Overridden = False
    Me.IO_ReserveHeat.Size = New System.Drawing.Size(13, 25)
    Me.IO_ReserveHeat.TabIndex = 222
    Me.IO_ReserveHeat.Text = "MimicDeviceValve1"
    Me.IO_ReserveHeat.UIEnabled = False
    '
    'IO_ReserveFillCold
    '
    Me.IO_ReserveFillCold.Location = New System.Drawing.Point(987, 209)
    Me.IO_ReserveFillCold.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_ReserveFillCold.Name = "IO_ReserveFillCold"
    Me.IO_ReserveFillCold.NormallyOn = False
    Me.IO_ReserveFillCold.OffFeedback = False
    Me.IO_ReserveFillCold.OffFeedbackEnabled = False
    Me.IO_ReserveFillCold.OnFeedback = False
    Me.IO_ReserveFillCold.OnFeedbackEnabled = False
    Me.IO_ReserveFillCold.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_ReserveFillCold.Overridden = False
    Me.IO_ReserveFillCold.Size = New System.Drawing.Size(13, 25)
    Me.IO_ReserveFillCold.TabIndex = 221
    Me.IO_ReserveFillCold.Text = "MimicDeviceValve1"
    Me.IO_ReserveFillCold.UIEnabled = False
    '
    'IO_ReserveDrain
    '
    Me.IO_ReserveDrain.Location = New System.Drawing.Point(992, 416)
    Me.IO_ReserveDrain.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_ReserveDrain.Name = "IO_ReserveDrain"
    Me.IO_ReserveDrain.NormallyOn = False
    Me.IO_ReserveDrain.OffFeedback = False
    Me.IO_ReserveDrain.OffFeedbackEnabled = False
    Me.IO_ReserveDrain.OnFeedback = False
    Me.IO_ReserveDrain.OnFeedbackEnabled = False
    Me.IO_ReserveDrain.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_ReserveDrain.Overridden = False
    Me.IO_ReserveDrain.Size = New System.Drawing.Size(13, 25)
    Me.IO_ReserveDrain.TabIndex = 220
    Me.IO_ReserveDrain.Text = "MimicDeviceValve1"
    Me.IO_ReserveDrain.UIEnabled = False
    '
    'IO_ReserveTransfer
    '
    Me.IO_ReserveTransfer.Location = New System.Drawing.Point(952, 463)
    Me.IO_ReserveTransfer.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_ReserveTransfer.Name = "IO_ReserveTransfer"
    Me.IO_ReserveTransfer.NormallyOn = False
    Me.IO_ReserveTransfer.OffFeedback = False
    Me.IO_ReserveTransfer.OffFeedbackEnabled = False
    Me.IO_ReserveTransfer.OnFeedback = False
    Me.IO_ReserveTransfer.OnFeedbackEnabled = False
    Me.IO_ReserveTransfer.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_ReserveTransfer.Overridden = False
    Me.IO_ReserveTransfer.Size = New System.Drawing.Size(13, 25)
    Me.IO_ReserveTransfer.TabIndex = 219
    Me.IO_ReserveTransfer.Text = "MimicDeviceValve1"
    Me.IO_ReserveTransfer.UIEnabled = False
    '
    'IO_AddTransfer
    '
    Me.IO_AddTransfer.Location = New System.Drawing.Point(791, 468)
    Me.IO_AddTransfer.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_AddTransfer.Name = "IO_AddTransfer"
    Me.IO_AddTransfer.NormallyOn = False
    Me.IO_AddTransfer.OffFeedback = False
    Me.IO_AddTransfer.OffFeedbackEnabled = False
    Me.IO_AddTransfer.OnFeedback = False
    Me.IO_AddTransfer.OnFeedbackEnabled = False
    Me.IO_AddTransfer.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_AddTransfer.Overridden = False
    Me.IO_AddTransfer.Size = New System.Drawing.Size(13, 25)
    Me.IO_AddTransfer.TabIndex = 217
    Me.IO_AddTransfer.Text = "MimicDeviceValve1"
    Me.IO_AddTransfer.UIEnabled = False
    '
    'IO_AddMixing
    '
    Me.IO_AddMixing.Location = New System.Drawing.Point(791, 343)
    Me.IO_AddMixing.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_AddMixing.Name = "IO_AddMixing"
    Me.IO_AddMixing.NormallyOn = False
    Me.IO_AddMixing.OffFeedback = False
    Me.IO_AddMixing.OffFeedbackEnabled = False
    Me.IO_AddMixing.OnFeedback = False
    Me.IO_AddMixing.OnFeedbackEnabled = False
    Me.IO_AddMixing.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_AddMixing.Overridden = False
    Me.IO_AddMixing.Size = New System.Drawing.Size(13, 25)
    Me.IO_AddMixing.TabIndex = 216
    Me.IO_AddMixing.Text = "MimicDeviceValve1"
    Me.IO_AddMixing.UIEnabled = False
    '
    'IO_AddDrain
    '
    Me.IO_AddDrain.Location = New System.Drawing.Point(879, 389)
    Me.IO_AddDrain.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_AddDrain.Name = "IO_AddDrain"
    Me.IO_AddDrain.NormallyOn = False
    Me.IO_AddDrain.OffFeedback = False
    Me.IO_AddDrain.OffFeedbackEnabled = False
    Me.IO_AddDrain.OnFeedback = False
    Me.IO_AddDrain.OnFeedbackEnabled = False
    Me.IO_AddDrain.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_AddDrain.Overridden = False
    Me.IO_AddDrain.Size = New System.Drawing.Size(13, 25)
    Me.IO_AddDrain.TabIndex = 215
    Me.IO_AddDrain.Text = "MimicDeviceValve1"
    Me.IO_AddDrain.UIEnabled = False
    '
    'IO_AddFillCold
    '
    Me.IO_AddFillCold.Location = New System.Drawing.Point(873, 209)
    Me.IO_AddFillCold.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_AddFillCold.Name = "IO_AddFillCold"
    Me.IO_AddFillCold.NormallyOn = False
    Me.IO_AddFillCold.OffFeedback = False
    Me.IO_AddFillCold.OffFeedbackEnabled = False
    Me.IO_AddFillCold.OnFeedback = False
    Me.IO_AddFillCold.OnFeedbackEnabled = False
    Me.IO_AddFillCold.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_AddFillCold.Overridden = False
    Me.IO_AddFillCold.Size = New System.Drawing.Size(13, 25)
    Me.IO_AddFillCold.TabIndex = 214
    Me.IO_AddFillCold.Text = "MimicDeviceValve1"
    Me.IO_AddFillCold.UIEnabled = False
    '
    'IO_Tank1TransferToReserve
    '
    Me.IO_Tank1TransferToReserve.Location = New System.Drawing.Point(935, 209)
    Me.IO_Tank1TransferToReserve.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_Tank1TransferToReserve.Name = "IO_Tank1TransferToReserve"
    Me.IO_Tank1TransferToReserve.NormallyOn = False
    Me.IO_Tank1TransferToReserve.OffFeedback = False
    Me.IO_Tank1TransferToReserve.OffFeedbackEnabled = False
    Me.IO_Tank1TransferToReserve.OnFeedback = False
    Me.IO_Tank1TransferToReserve.OnFeedbackEnabled = False
    Me.IO_Tank1TransferToReserve.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_Tank1TransferToReserve.Overridden = False
    Me.IO_Tank1TransferToReserve.Size = New System.Drawing.Size(13, 25)
    Me.IO_Tank1TransferToReserve.TabIndex = 213
    Me.IO_Tank1TransferToReserve.Text = "MimicDeviceValve1"
    Me.IO_Tank1TransferToReserve.UIEnabled = False
    '
    'IO_Tank1TransferToDrain
    '
    Me.IO_Tank1TransferToDrain.Location = New System.Drawing.Point(903, 209)
    Me.IO_Tank1TransferToDrain.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_Tank1TransferToDrain.Name = "IO_Tank1TransferToDrain"
    Me.IO_Tank1TransferToDrain.NormallyOn = False
    Me.IO_Tank1TransferToDrain.OffFeedback = False
    Me.IO_Tank1TransferToDrain.OffFeedbackEnabled = False
    Me.IO_Tank1TransferToDrain.OnFeedback = False
    Me.IO_Tank1TransferToDrain.OnFeedbackEnabled = False
    Me.IO_Tank1TransferToDrain.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_Tank1TransferToDrain.Overridden = False
    Me.IO_Tank1TransferToDrain.Size = New System.Drawing.Size(13, 25)
    Me.IO_Tank1TransferToDrain.TabIndex = 212
    Me.IO_Tank1TransferToDrain.Text = "MimicDeviceValve1"
    Me.IO_Tank1TransferToDrain.UIEnabled = False
    '
    'IO_Tank1TransferToAdd
    '
    Me.IO_Tank1TransferToAdd.Location = New System.Drawing.Point(839, 209)
    Me.IO_Tank1TransferToAdd.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_Tank1TransferToAdd.Name = "IO_Tank1TransferToAdd"
    Me.IO_Tank1TransferToAdd.NormallyOn = False
    Me.IO_Tank1TransferToAdd.OffFeedback = False
    Me.IO_Tank1TransferToAdd.OffFeedbackEnabled = False
    Me.IO_Tank1TransferToAdd.OnFeedback = False
    Me.IO_Tank1TransferToAdd.OnFeedbackEnabled = False
    Me.IO_Tank1TransferToAdd.Orientation = MimicDevice.EOrientation.Vertical
    Me.IO_Tank1TransferToAdd.Overridden = False
    Me.IO_Tank1TransferToAdd.Size = New System.Drawing.Size(13, 25)
    Me.IO_Tank1TransferToAdd.TabIndex = 211
    Me.IO_Tank1TransferToAdd.Text = "MimicDeviceValve1"
    Me.IO_Tank1TransferToAdd.UIEnabled = False
    '
    'IO_SystemBlock
    '
    Me.IO_SystemBlock.Location = New System.Drawing.Point(275, 382)
    Me.IO_SystemBlock.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_SystemBlock.Name = "IO_SystemBlock"
    Me.IO_SystemBlock.NormallyOn = False
    Me.IO_SystemBlock.OffFeedback = False
    Me.IO_SystemBlock.OffFeedbackEnabled = False
    Me.IO_SystemBlock.OnFeedback = False
    Me.IO_SystemBlock.OnFeedbackEnabled = False
    Me.IO_SystemBlock.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_SystemBlock.Overridden = False
    Me.IO_SystemBlock.Size = New System.Drawing.Size(25, 13)
    Me.IO_SystemBlock.TabIndex = 206
    Me.IO_SystemBlock.Text = "MimicDeviceValve1"
    Me.IO_SystemBlock.UIEnabled = False
    '
    'IO_MachineDrain
    '
    Me.IO_MachineDrain.Location = New System.Drawing.Point(693, 559)
    Me.IO_MachineDrain.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_MachineDrain.Name = "IO_MachineDrain"
    Me.IO_MachineDrain.NormallyOn = False
    Me.IO_MachineDrain.OffFeedback = False
    Me.IO_MachineDrain.OffFeedbackEnabled = False
    Me.IO_MachineDrain.OnFeedback = False
    Me.IO_MachineDrain.OnFeedbackEnabled = False
    Me.IO_MachineDrain.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_MachineDrain.Overridden = False
    Me.IO_MachineDrain.Size = New System.Drawing.Size(25, 13)
    Me.IO_MachineDrain.TabIndex = 202
    Me.IO_MachineDrain.Text = "MimicDeviceValve1"
    Me.IO_MachineDrain.UIEnabled = False
    '
    'IO_FillCold
    '
    Me.IO_FillCold.Location = New System.Drawing.Point(373, 587)
    Me.IO_FillCold.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_FillCold.Name = "IO_FillCold"
    Me.IO_FillCold.NormallyOn = False
    Me.IO_FillCold.OffFeedback = False
    Me.IO_FillCold.OffFeedbackEnabled = False
    Me.IO_FillCold.OnFeedback = False
    Me.IO_FillCold.OnFeedbackEnabled = False
    Me.IO_FillCold.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_FillCold.Overridden = False
    Me.IO_FillCold.Size = New System.Drawing.Size(25, 13)
    Me.IO_FillCold.TabIndex = 201
    Me.IO_FillCold.Text = "MimicDeviceValve1"
    Me.IO_FillCold.UIEnabled = False
    '
    'IO_Condensate
    '
    Me.IO_Condensate.Location = New System.Drawing.Point(693, 484)
    Me.IO_Condensate.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_Condensate.Name = "IO_Condensate"
    Me.IO_Condensate.NormallyOn = False
    Me.IO_Condensate.OffFeedback = False
    Me.IO_Condensate.OffFeedbackEnabled = False
    Me.IO_Condensate.OnFeedback = False
    Me.IO_Condensate.OnFeedbackEnabled = False
    Me.IO_Condensate.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_Condensate.Overridden = False
    Me.IO_Condensate.Size = New System.Drawing.Size(25, 13)
    Me.IO_Condensate.TabIndex = 200
    Me.IO_Condensate.Text = "MimicDeviceValve1"
    Me.IO_Condensate.UIEnabled = False
    '
    'IO_WaterDump
    '
    Me.IO_WaterDump.Location = New System.Drawing.Point(693, 507)
    Me.IO_WaterDump.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_WaterDump.Name = "IO_WaterDump"
    Me.IO_WaterDump.NormallyOn = False
    Me.IO_WaterDump.OffFeedback = False
    Me.IO_WaterDump.OffFeedbackEnabled = False
    Me.IO_WaterDump.OnFeedback = False
    Me.IO_WaterDump.OnFeedbackEnabled = False
    Me.IO_WaterDump.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_WaterDump.Overridden = False
    Me.IO_WaterDump.Size = New System.Drawing.Size(25, 13)
    Me.IO_WaterDump.TabIndex = 199
    Me.IO_WaterDump.Text = "MimicDeviceValve1"
    Me.IO_WaterDump.UIEnabled = False
    '
    'IO_CoolSelect
    '
    Me.IO_CoolSelect.Location = New System.Drawing.Point(693, 454)
    Me.IO_CoolSelect.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_CoolSelect.Name = "IO_CoolSelect"
    Me.IO_CoolSelect.NormallyOn = False
    Me.IO_CoolSelect.OffFeedback = False
    Me.IO_CoolSelect.OffFeedbackEnabled = False
    Me.IO_CoolSelect.OnFeedback = False
    Me.IO_CoolSelect.OnFeedbackEnabled = False
    Me.IO_CoolSelect.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_CoolSelect.Overridden = False
    Me.IO_CoolSelect.Size = New System.Drawing.Size(25, 13)
    Me.IO_CoolSelect.TabIndex = 198
    Me.IO_CoolSelect.Text = "MimicDeviceValve1"
    Me.IO_CoolSelect.UIEnabled = False
    '
    'IO_HeatSelect
    '
    Me.IO_HeatSelect.Location = New System.Drawing.Point(693, 431)
    Me.IO_HeatSelect.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_HeatSelect.Name = "IO_HeatSelect"
    Me.IO_HeatSelect.NormallyOn = False
    Me.IO_HeatSelect.OffFeedback = False
    Me.IO_HeatSelect.OffFeedbackEnabled = False
    Me.IO_HeatSelect.OnFeedback = False
    Me.IO_HeatSelect.OnFeedbackEnabled = False
    Me.IO_HeatSelect.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_HeatSelect.Overridden = False
    Me.IO_HeatSelect.Size = New System.Drawing.Size(25, 13)
    Me.IO_HeatSelect.TabIndex = 197
    Me.IO_HeatSelect.Text = "MimicDeviceValve1"
    Me.IO_HeatSelect.UIEnabled = False
    '
    'IO_Airpad
    '
    Me.IO_Airpad.Location = New System.Drawing.Point(116, 116)
    Me.IO_Airpad.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_Airpad.Name = "IO_Airpad"
    Me.IO_Airpad.NormallyOn = False
    Me.IO_Airpad.OffFeedback = False
    Me.IO_Airpad.OffFeedbackEnabled = False
    Me.IO_Airpad.OnFeedback = False
    Me.IO_Airpad.OnFeedbackEnabled = False
    Me.IO_Airpad.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_Airpad.Overridden = False
    Me.IO_Airpad.Size = New System.Drawing.Size(25, 13)
    Me.IO_Airpad.TabIndex = 196
    Me.IO_Airpad.Text = "MimicDeviceValve1"
    Me.IO_Airpad.UIEnabled = False
    '
    'IO_Vent
    '
    Me.IO_Vent.Location = New System.Drawing.Point(115, 145)
    Me.IO_Vent.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_Vent.Name = "IO_Vent"
    Me.IO_Vent.NormallyOn = False
    Me.IO_Vent.OffFeedback = False
    Me.IO_Vent.OffFeedbackEnabled = False
    Me.IO_Vent.OnFeedback = False
    Me.IO_Vent.OnFeedbackEnabled = False
    Me.IO_Vent.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_Vent.Overridden = False
    Me.IO_Vent.Size = New System.Drawing.Size(25, 13)
    Me.IO_Vent.TabIndex = 195
    Me.IO_Vent.Text = "MimicDeviceValve1"
    Me.IO_Vent.UIEnabled = False
    '
    'IO_TopWash
    '
    Me.IO_TopWash.Location = New System.Drawing.Point(115, 175)
    Me.IO_TopWash.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_TopWash.Name = "IO_TopWash"
    Me.IO_TopWash.NormallyOn = False
    Me.IO_TopWash.OffFeedback = False
    Me.IO_TopWash.OffFeedbackEnabled = False
    Me.IO_TopWash.OnFeedback = False
    Me.IO_TopWash.OnFeedbackEnabled = False
    Me.IO_TopWash.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_TopWash.Overridden = False
    Me.IO_TopWash.Size = New System.Drawing.Size(25, 13)
    Me.IO_TopWash.TabIndex = 194
    Me.IO_TopWash.Text = "MimicDeviceValve1"
    Me.IO_TopWash.UIEnabled = False
    '
    'IO_KierLidClosed
    '
    Me.IO_KierLidClosed.Location = New System.Drawing.Point(115, 238)
    Me.IO_KierLidClosed.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_KierLidClosed.Name = "IO_KierLidClosed"
    Me.IO_KierLidClosed.Size = New System.Drawing.Size(16, 15)
    Me.IO_KierLidClosed.TabIndex = 191
    Me.IO_KierLidClosed.UIEnabled = False
    '
    'IO_AddPumpRunning
    '
    Me.IO_AddPumpRunning.Location = New System.Drawing.Point(809, 402)
    Me.IO_AddPumpRunning.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_AddPumpRunning.Name = "IO_AddPumpRunning"
    Me.IO_AddPumpRunning.NormallyOn = False
    Me.IO_AddPumpRunning.OffFeedback = False
    Me.IO_AddPumpRunning.OffFeedbackEnabled = False
    Me.IO_AddPumpRunning.OnFeedback = False
    Me.IO_AddPumpRunning.OnFeedbackEnabled = False
    Me.IO_AddPumpRunning.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_AddPumpRunning.Overridden = False
    Me.IO_AddPumpRunning.Size = New System.Drawing.Size(31, 28)
    Me.IO_AddPumpRunning.TabIndex = 190
    Me.IO_AddPumpRunning.Text = "MimicDevicePump1"
    Me.IO_AddPumpRunning.UIEnabled = False
    '
    'IO_PumpRunning
    '
    Me.IO_PumpRunning.Location = New System.Drawing.Point(555, 551)
    Me.IO_PumpRunning.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_PumpRunning.Name = "IO_PumpRunning"
    Me.IO_PumpRunning.NormallyOn = False
    Me.IO_PumpRunning.OffFeedback = False
    Me.IO_PumpRunning.OffFeedbackEnabled = False
    Me.IO_PumpRunning.OnFeedback = False
    Me.IO_PumpRunning.OnFeedbackEnabled = False
    Me.IO_PumpRunning.Orientation = MimicDevice.EOrientation.Horizontal
    Me.IO_PumpRunning.Overridden = False
    Me.IO_PumpRunning.Size = New System.Drawing.Size(31, 28)
    Me.IO_PumpRunning.TabIndex = 185
    Me.IO_PumpRunning.Text = "MimicDevicePump1"
    Me.IO_PumpRunning.UIEnabled = False
    '
    'IO_TempInterlock
    '
    Me.IO_TempInterlock.Location = New System.Drawing.Point(648, 409)
    Me.IO_TempInterlock.Margin = New System.Windows.Forms.Padding(4)
    Me.IO_TempInterlock.Name = "IO_TempInterlock"
    Me.IO_TempInterlock.Size = New System.Drawing.Size(16, 15)
    Me.IO_TempInterlock.TabIndex = 177
    Me.IO_TempInterlock.UIEnabled = False
    '
    'LblTank1Status
    '
    Me.LblTank1Status.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.LblTank1Status.ForeColor = System.Drawing.Color.Black
    Me.LblTank1Status.Location = New System.Drawing.Point(648, 17)
    Me.LblTank1Status.Margin = New System.Windows.Forms.Padding(4)
    Me.LblTank1Status.Name = "LblTank1Status"
    Me.LblTank1Status.Size = New System.Drawing.Size(102, 20)
    Me.LblTank1Status.TabIndex = 176
    Me.LblTank1Status.Text = "Tank1Status"
    Me.LblTank1Status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'LblTank1Level
    '
    Me.LblTank1Level.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.LblTank1Level.ForeColor = System.Drawing.Color.Black
    Me.LblTank1Level.Location = New System.Drawing.Point(648, 46)
    Me.LblTank1Level.Margin = New System.Windows.Forms.Padding(4)
    Me.LblTank1Level.Name = "LblTank1Level"
    Me.LblTank1Level.Size = New System.Drawing.Size(94, 20)
    Me.LblTank1Level.TabIndex = 173
    Me.LblTank1Level.Text = "Tank1Level"
    Me.LblTank1Level.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'HeatCoolOutput
    '
    Me.HeatCoolOutput.ForeColor = System.Drawing.Color.Black
    Me.HeatCoolOutput.Format = "0.0\%"
    Me.HeatCoolOutput.Location = New System.Drawing.Point(608, 420)
    Me.HeatCoolOutput.Margin = New System.Windows.Forms.Padding(4)
    Me.HeatCoolOutput.Name = "HeatCoolOutput"
    Me.HeatCoolOutput.NumberScale = 10
    Me.HeatCoolOutput.Orientation = System.Windows.Forms.Orientation.Vertical
    Me.HeatCoolOutput.Size = New System.Drawing.Size(56, 117)
    Me.HeatCoolOutput.TabIndex = 151
    '
    'labelAddStatusMimic
    '
    Me.labelAddStatusMimic.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.labelAddStatusMimic.ForeColor = System.Drawing.Color.Black
    Me.labelAddStatusMimic.Location = New System.Drawing.Point(648, 132)
    Me.labelAddStatusMimic.Margin = New System.Windows.Forms.Padding(4)
    Me.labelAddStatusMimic.Name = "labelAddStatusMimic"
    Me.labelAddStatusMimic.Size = New System.Drawing.Size(91, 20)
    Me.labelAddStatusMimic.TabIndex = 237
    Me.labelAddStatusMimic.Text = "Add Status"
    Me.labelAddStatusMimic.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'Mimic
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackColor = System.Drawing.Color.WhiteSmoke
    Me.BackgroundImage = Global.My.Resources.Resources.ThenPlatform_AEMXN
    Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
    Me.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.Controls.Add(Me.PanelButtons)
    Me.Controls.Add(Me.ButtonLevelFlush)
    Me.Controls.Add(Me.labelMachineLevel)
    Me.Controls.Add(Me.IO_ReserveRunback)
    Me.Controls.Add(Me.IO_HoldDown)
    Me.Controls.Add(Me.IO_FlowReverse)
    Me.Controls.Add(Me.IO_FillHot)
    Me.Controls.Add(Me.IO_MachineDrainHot)
    Me.Controls.Add(Me.levelBarTank1)
    Me.Controls.Add(Me.levelBarAdd)
    Me.Controls.Add(Me.levelBarReserve)
    Me.Controls.Add(Me.LabelReserveStatusMimic)
    Me.Controls.Add(Me.buttonManualControl)
    Me.Controls.Add(Me.labelPumpReversalStatus)
    Me.Controls.Add(Me.labelPumpStatus)
    Me.Controls.Add(Me.labelMachinePressure)
    Me.Controls.Add(Me.labelPackageDP)
    Me.Controls.Add(Me.levelBarMachine)
    Me.Controls.Add(Me.labelAddLevelMimic)
    Me.Controls.Add(Me.labelReserveLevelMimic)
    Me.Controls.Add(Me.IO_AddRunback)
    Me.Controls.Add(Me.IO_Tank1FillCold)
    Me.Controls.Add(Me.IO_Tank1Transfer)
    Me.Controls.Add(Me.IO_ReserveHeat)
    Me.Controls.Add(Me.IO_ReserveFillCold)
    Me.Controls.Add(Me.IO_ReserveDrain)
    Me.Controls.Add(Me.IO_ReserveTransfer)
    Me.Controls.Add(Me.IO_AddTransfer)
    Me.Controls.Add(Me.IO_AddMixing)
    Me.Controls.Add(Me.IO_AddDrain)
    Me.Controls.Add(Me.IO_AddFillCold)
    Me.Controls.Add(Me.IO_Tank1TransferToReserve)
    Me.Controls.Add(Me.IO_Tank1TransferToDrain)
    Me.Controls.Add(Me.IO_Tank1TransferToAdd)
    Me.Controls.Add(Me.IO_SystemBlock)
    Me.Controls.Add(Me.IO_MachineDrain)
    Me.Controls.Add(Me.IO_FillCold)
    Me.Controls.Add(Me.IO_Condensate)
    Me.Controls.Add(Me.IO_WaterDump)
    Me.Controls.Add(Me.IO_CoolSelect)
    Me.Controls.Add(Me.IO_HeatSelect)
    Me.Controls.Add(Me.IO_Airpad)
    Me.Controls.Add(Me.IO_Vent)
    Me.Controls.Add(Me.IO_TopWash)
    Me.Controls.Add(Me.IO_KierLidClosed)
    Me.Controls.Add(Me.IO_AddPumpRunning)
    Me.Controls.Add(Me.IO_PumpRunning)
    Me.Controls.Add(Me.buttonReserveReady)
    Me.Controls.Add(Me.buttonAddReady)
    Me.Controls.Add(Me.buttonPumpRequest)
    Me.Controls.Add(Me.IO_TempInterlock)
    Me.Controls.Add(Me.LblTank1Status)
    Me.Controls.Add(Me.LblTank1Level)
    Me.Controls.Add(Me.HeatCoolOutput)
    Me.Controls.Add(Me.labelAddStatusMimic)
    Me.ForeColor = System.Drawing.Color.Black
    Me.Margin = New System.Windows.Forms.Padding(4)
    Me.Name = "Mimic"
    Me.Size = New System.Drawing.Size(1064, 651)
    Me.PanelButtons.ResumeLayout(False)
    Me.PanelButtons.PerformLayout()
    Me.GroupBoxPump.ResumeLayout(False)
    Me.GroupBoxPump.PerformLayout()
    Me.GroupBoxReserveTank.ResumeLayout(False)
    Me.GroupBoxReserveTank.PerformLayout()
    Me.GroupBoxAddTank.ResumeLayout(False)
    Me.GroupBoxAddTank.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
  Friend WithEvents LblTank1Level As MimicControls.Label
  Friend WithEvents LblTank1Status As MimicControls.Label
  Friend WithEvents buttonPumpRequest As System.Windows.Forms.Button
  Friend WithEvents IO_ReserveTransfer As MimicDeviceValve
  Friend WithEvents IO_ReserveDrain As MimicDeviceValve
  Friend WithEvents IO_ReserveFillCold As MimicDeviceValve
  Friend WithEvents IO_ReserveHeat As MimicDeviceValve
  Friend WithEvents IO_Tank1Transfer As MimicDeviceValve
  Friend WithEvents IO_Tank1FillCold As MimicDeviceValve
  Friend WithEvents labelReserveLevelMimic As MimicControls.Label
  Friend WithEvents labelPumpStatus As MimicControls.Label
  Friend WithEvents labelPumpReversalStatus As MimicControls.Label
  Friend WithEvents buttonManualControl As System.Windows.Forms.Button
  Friend WithEvents HeatCoolOutput As MimicControls.HeatExchanger
  Friend WithEvents IO_TempInterlock As MimicControls.Input
  Friend WithEvents buttonAddReady As System.Windows.Forms.Button
  Friend WithEvents buttonReserveReady As System.Windows.Forms.Button
  Friend WithEvents IO_PumpRunning As MimicDevicePump
  Friend WithEvents IO_AddPumpRunning As MimicDevicePump
  Friend WithEvents IO_KierLidClosed As MimicControls.Input
  Friend WithEvents IO_TopWash As MimicDeviceValve
  Friend WithEvents IO_Vent As MimicDeviceValve
  Friend WithEvents IO_Airpad As MimicDeviceValve
  Friend WithEvents IO_HeatSelect As MimicDeviceValve
  Friend WithEvents IO_CoolSelect As MimicDeviceValve
  Friend WithEvents IO_WaterDump As MimicDeviceValve
  Friend WithEvents IO_Condensate As MimicDeviceValve
  Friend WithEvents IO_FillCold As MimicDeviceValve
  Friend WithEvents IO_MachineDrain As MimicDeviceValve
  Friend WithEvents IO_SystemBlock As MimicDeviceValve
  Friend WithEvents IO_Tank1TransferToAdd As MimicDeviceValve
  Friend WithEvents IO_Tank1TransferToDrain As MimicDeviceValve
  Friend WithEvents IO_Tank1TransferToReserve As MimicDeviceValve
  Friend WithEvents IO_AddFillCold As MimicDeviceValve
  Friend WithEvents IO_AddDrain As MimicDeviceValve
  Friend WithEvents IO_AddMixing As MimicDeviceValve
  Friend WithEvents IO_AddTransfer As MimicDeviceValve
  Friend WithEvents IO_AddRunback As MimicDeviceValve
  Friend WithEvents labelAddLevelMimic As MimicControls.Label
  Friend WithEvents levelBarMachine As MimicControls.LevelBar
  Friend WithEvents labelPackageDP As MimicControls.Label
  Friend WithEvents labelMachinePressure As MimicControls.Label
  Friend WithEvents PanelButtons As MimicControls.Panel
  Friend WithEvents buttonClose As System.Windows.Forms.Button
  Friend WithEvents GroupBoxAddTank As System.Windows.Forms.GroupBox
  Friend WithEvents buttonAddReady1 As System.Windows.Forms.Button
  Friend WithEvents labelAddStatusPanelButtons As MimicControls.Label
  Friend WithEvents buttonAddFill As System.Windows.Forms.Button
  Friend WithEvents buttonAddDrain As System.Windows.Forms.Button
  Friend WithEvents Timer2 As System.Windows.Forms.Timer
  Friend WithEvents labelAddLevelPanelButtons As MimicControls.Label
  Friend WithEvents GroupBoxReserveTank As System.Windows.Forms.GroupBox
  Friend WithEvents labelReserveLevelPanelButtons As MimicControls.Label
  Friend WithEvents buttonReserveDrain As System.Windows.Forms.Button
  Friend WithEvents buttonReserveFill As System.Windows.Forms.Button
  Friend WithEvents buttonReserveReady1 As System.Windows.Forms.Button
  Friend WithEvents labelReserveStatusPanelButtons As MimicControls.Label
  Friend WithEvents buttonReserveHeat As System.Windows.Forms.Button
  Friend WithEvents buttonReserveCancel As System.Windows.Forms.Button
  Friend WithEvents buttonAddCancel As System.Windows.Forms.Button
  Friend WithEvents labelAddStatusMimic As MimicControls.Label
  Friend WithEvents GroupBoxPump As System.Windows.Forms.GroupBox
  Friend WithEvents labelPumpOutputPanelButtons As MimicControls.Label
  Friend WithEvents buttonPumpStop As System.Windows.Forms.Button
  Friend WithEvents buttonPumpStart As System.Windows.Forms.Button
  Friend WithEvents labelPumpStatusPanelButtons As MimicControls.Label
  Friend WithEvents LabelReserveStatusMimic As System.Windows.Forms.Label
  Friend WithEvents buttonHalfLoad As System.Windows.Forms.Button
  Friend WithEvents levelBarReserve As MimicControls.LevelBar
  Friend WithEvents levelBarAdd As MimicControls.LevelBar
  Friend WithEvents levelBarTank1 As MimicControls.LevelBar
  Friend WithEvents IO_MachineDrainHot As MimicDeviceValve
  Friend WithEvents IO_FillHot As MimicDeviceValve
  Friend WithEvents IO_FlowReverse As Output
  Friend WithEvents IO_HoldDown As Output
  Friend WithEvents IO_ReserveRunback As MimicDeviceValve
  Friend WithEvents labelMachineLevel As MimicControls.Label
  Friend WithEvents ButtonLevelFlush As Button
  Friend WithEvents ButtonTank1Cancel As Button
  Friend WithEvents CheckBox1 As CheckBox
End Class
