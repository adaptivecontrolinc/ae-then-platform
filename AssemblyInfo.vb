Imports System.Reflection
Imports System.Security.Permissions
Imports System.Runtime.InteropServices

<Assembly: ComVisible(False)> 
<Assembly: CLSCompliant(True)> 
<Assembly: SecurityPermission(SecurityAction.RequestMinimum, Execution:=True)>

<Assembly: AssemblyTitle("AE Then Platform")>
<Assembly: AssemblyCompany("American and Efird")>
<Assembly: AssemblyCopyright("Portions Copyright 1998-2024 Adaptive Control.")>
<Assembly: AssemblyVersion("1.2")>
<Assembly: AssemblyFileVersion("1.2")>


'******************************************************************************************************
'****************                Code Version Notes -- Most Recent At Top              ****************
'******************************************************************************************************
' TODO
' Look Ahead needs work

' Goal to add MimicPages code

' Check AT, AC, AD one more time



'V1.2 [2024-09-29] DH



'V1.1 [2024-09-12] DH
' Rinse fixes so that Vents, then Top Wash, then Fills/Rinses
' Add Transfer fixes
' DrugroomPreview fix for Tank1Calloff : to display correctly


'V1.1 [2024-08-20] DH
' AutoDispenser working with this version.  
' - BatchControl running Sockets and variables are case sensitive.  BatchDyeing (DCOM) variables are case insensitive
'   AutoDispenser reads: 
'   remoteValues_.Connect(CommandSequencer.ConnectType.Sockets, Connection, 5000, "DrugroomDisplayJob1", "DispenseCallOff", "DispenseTank")

' Update Utilities to use Namespace and Modules
' Add Simulation class function


'V1.1 [2024-07-20] DH
' VB6 to .NET update
' Update Control.vb for latest BatchControl
' Use .SDF table format due to local database changes and work.  SDB gives exceptions.


' AE-GMC-ThenPlatform below:

'v1.6 2023-10-30 DH
' Update Graph variables for pressure to expand range for improved visibility and troubleshooting
'

'v1.5 2022-05-23 DH
' Update control code for LevelFlush control run.
' Also add code for local dyelots table update to match Tina13 for working recipe color and style details.

'v1.4 2022-05-10 DH
' Add level flush state to draining sequences when not loosing level.  

'v1.3 2022-04-19 DH
' Add Vessel temp probe offset parameter.  F2-04RTD showing 3.0C offset during startup and Juan and 
'   Humberto want to use a parameter for +/- temperature offset.  0-100, with 50 being 0.0C offset.
' Also update Lid Control to add delay for Carrier Hold Down output to release hold down before opening
'   band to prevent lid bouncing.  Carrier hold down cylinder is slow actuating.

'v1.2 2022-04-13 DH
' A&E GMX Tina14 Production release

'v1.1 2022-04-11 DH
' Updated Timers and flashers to use TickCount instead of LocalDateNow

'v1.0 2022-04-01 DH
' Convert vb6 AEThenPlatform version 2.3.9 to .NET using ThenMultiflex version 1.3 (notes below)
' Define I/O based on vb6 layout
' Remove following:
'   - SampleClosed variable For this machine As Mt. Holly did Not have a sample tank
'   - Kier B properties and logic - Single Kier machine
'   - KierBLoaded, KierBLoadedText, IO.KierUnisolate,
'   - Commands: TC (Heat by Contacts),
'   - Controls: ACManualSteamTest, 
' Updates
'   - Commands: 
'   -   CO, HE, SA
'   -   NP updated to include PackageHeight up to 9 (was 4)


#Region " VB6 VERSION NOTES "
'NOTE: LATEST CHANGES AT TOP


'2/14/2022 - DH
' Version 2.4
' Update for RF command to Step On with Fresh water fill


'10/05/2021 - DH
' version 2.3
'  Update for RF command when filling from vessel.  Previous version would stop/lock-up when stepping into RF command parameter V set.
'  RF V would work if stepping from DR to RF command.  Checked command parameters and updated commands where RF.Cancel was called.
'  Issue does not appear if running in simulation on laptop.  Edited Array size for Dout Write array size in the event that is the
'  cause.
'  Add logic to Depressurize & Stop Pump, then use the air pad to fill back from the machine
'  Update Control Code Status for correct RF.IsForeground flag
'  New parameter 'Parameters_ReserveBackFillStartTime' to delay airpad before runback.  Issue with Airpad and Runback at same time
'       causing pressure safety interlock to drop out.  Time is in seconds.



'06/10/2019 - DH
' version 2.1
'   Update so that MainDrain does not remain open during UL command at request of John Smith.
'   - Issues with vapor from main drain header leaking back into kier vessel during unload command
'       most notable during long UL delays.
'       removed from IO_Drain = UL.IsUnloading Or UL.IsWaitingToUnload
'       and added to IO_TopWash
'   - John Smith wants to only open TopWash during this UL procedure and not Main Drain



'05/17/2018 - DH
' Updated code for output IO_AddPumpStart so that it only starts pump when AF.AddMixing and add level high
'   issue with starting pump with add level > parameter and parent.isprogramrunning, but add mixing valve was closed
'   causing add pump to trip - bad code with last update :(
' Also updated AC command to match AT command a bit better for consistency


'04/17/2018 - DH
' Re-ordered acVersion file for latest at top...
' Update DR and HD commands to flush the level transmitters after a parameter delay, if the level does not reach 0%.
' issue with some machines, especially during HD command, where DP level standard drops or flashes off (HD command)
' resulting in machine level reading higher than calibrated.  Flushing and settling will resolve this issue.
' New parameter 'LevelFlushOverrideTime'
' New alarm - Alarms_CheckLevelTransmitter with DR and HD command levelflushcount > 5
' Update to acAddPrepare class module for KP1.AddMixing flag.  If KP command has mix parameter set to '0', the KP1
'   logic would still set the KP1.AddMixing flag to true upon entering the Heat state.  Updated KP1.Heat state to
'   only set mix flag to 'true' if desired temperature greater than 0.
' Update to IO_AddTankMixing output value for same issue.  Previous version would start mixing add tank if add level
'   was greater than add mixing on level parameter.  This code uses add level parameter with AF.AddMixing flag so that
'   only mixes if AF command calls for it, or if the add transfer commands (AF, AT, etc) call for AddTankMixing output
'   for dosing or circulation reasons as before.

'07/14/2017 - ML
' Added array bounds check on IO_Modbus2.ReadBits
'   We were going outside the upper bound of the Dinp array and generating Null errors (91). This worked
'   for a long time but maybe because of memory map changes it started to fail, we would have been overwriting
'   memory locations adjacent to the Dinp array.

'2014/02/04 - DH
'Add PackageType0 Functionality

'11/13/2013 - DH
'Added new Temp Input for Condensate Return Temp Probe - Installed after the steam trap to monitor Cond Return temp and alarm if
'   rises above P_CondensateTempLimit (212.0F) for P_CondensateTempLimitTime (5sec) to signal a faulty trap
'   Six-Sigma project to reduce utility cose and increase machine efficiency

'1/18/2012 - DH
'  Updating all ControlCode for new Drugroom Preview.  All controllers that use the Drugroom Preview will be updated so that
'    Common drugroom preview will work the same.  All code for display strings will be created/set on controller, as opposed
'    to reading state values and setting a status at Drugroom PC.

'1/13/2011 - DH
'David Hipp calls to Request changing the function of IO_Tank1ToDrain (DOUT 44) such that it's only open when Tank 1 is draining.
'  Currently, This output is always on, except when Tank 1 is transferring to the add or stock tanks.
'  Current IO Configuration:
'  IO_Tank1ToDrain = (NHalt And (Parameters_DDSEnabled <> 1) And Not (KA.IsTransferToAddition Or KA.IsTransferToReserve Or _
'                    KP.KP1.IsTransferToReserve Or KP.KP1.IsTransferToAddition Or LA.KP1.IsTransferToReserve Or LA.KP1.IsTransferToAddition)) _
'                    Or _
'                    (NHalt And (Parameters_DDSEnabled = 1) And Not (IO_DyeDispenseReserve Or IO_DyeDispenseAdd))
'

'1/6/2011 - DH
'Issue with PID.StopIntegral where it was never reset to false once set.  As a result, integral term would fail to calculate and
'   temperature control would not function correctly.  Actual temperature would reach within 2.0 degrees of setpoint and fail to
'   get closer due to lack of integral error sum action.  Correctly reset now at beginning of temperature control as well as when
'   integral term is within range of desired.
'Changed Tank1Operator delay so that mix time included in KP1.IsWaitReady (Slow, Fast, and Mixing) no longer counts agains tank 1 operator.
'   Tank1Operator based on KP1.IsSlow and KP1.IsFast

'9/1/2010 - DH
'New Parameter_AddRunbackPulseTime to configure runback open/close pulse width.  Set to '0' to disable timer

'4/16/2010 - DH
'Fix PID.StopIntegral to work correctly by correctly limiting pidOutputTest to +/-1000 so that stop integral query will work
'Set IO_AddRunback to use NHalt and TempSafe instead of NHSafe (Machine Safe is not true for runback due to pressure required for runback
'   to work successfully...)

'Add Parameters for DeltaRamp and DeltaHold to smooth/restrict heatcool output adjustment for small changes.  While controlling temp,
'   if the delta change between the existing output and the newly calculated output are less than these parameter values for either a
'   hold or rampup/down state, then the output will not change.  tested successfully ramp values of 30 and hold value of 20 with
'   tight control on MC-17, first insulated machine.  This code works best with 4-20ma temp transmitter due to no noise resulting in
'   no temperror bounce as with the RTD inputs.

'Testing:
'Add PLC Logic to filter the RTD signal noise - Method discussed in AutomationDirect manual (6-19) for smoothing individual inputs
'   per channel.  Currently using a filter factor of 0.2 (described range from 0.1 to 0.9 with a smaller filter factor increasing
'   total filtration.
'Some filtering was present, leaving the smooth rate at '0' to see the difference.  Noise was still evident with 4.0F bounce
'   plus/minus.  Reverted back to existing plc logic with increased smoothing in PLC for noise immunity.  4-20mA appears to be the
'   direction to go.

'Enable Reading all spare analog input channels as a testing source for maintance.
'Configure first spare analog input as a temperature transmitter supplied by david hipp - tested first on MC-55
'   transmitter supplies 4-20ma for a range of 0-500F
'4-20mA signal provided stabile temperature input over entire cycle with no bounce (even at smoothing rate of 0) and temperature
'   control to within 0.4F displayed on graph between desired setpoint and actual temperature while running pump.  Variation did
'   occur during flow loss due to flow reversal, but corrected itself once flow was restored.

'3/22/2010
'TODO: Make backfilltimer in AF command a parameter time, not hardcoded at 5 seconds!  cannot troubleshoot histores when less than 10sec
'TODO : Make following parameters attribute changes:
'Tank 1 Drain Time 10
'Current - Time to drain tank 1 during transfer in seconds
'New - time in seconds tank 1 will transfer to the drain when tank 1 level reaches 0% after the transfer is finished
'Tank 1 Rinse Level 200
'Current - Level to fill tank 1 during rinse in tenths of a percent
'New - level in tenths of a percent to fill tank 1 after the transfer is finished for the tank 1 rinsing to the machine
'Tank 1 Rinse Time 3
'Current - Time to rinse tank 1 during transfer in seconds
'New - time in seconds to rinse tank 1 to the machine after the transfer is finished
'Tank 1 Rinse to Drain time 15
'Current - Time to rinse drugroom tank to drain during transfer in seconds
'New - time in seconds to rinse tank 1 to drain after the transfer is finished
'Tank 1 Time after rinse 10
'Current - Time to continue transferring tank 1 after rinsing the tank in seconds
'New - time in seconds to continue transferring to machine after the tank 1 rinse to machine is finished
'Tank 1 Time Before Rinse 10
'Current - Time to continue transferring tank 1 after low level float opened, before rinsing the tank in seconds
'New - time in seconds to continue transferring the kitchen tank to machine, once empty, before rinsing to machine
'LD/UL commands - Set AirpadOn = false at Command start to prevent issue where machine will not depressurize (SA > UL by mistake)

'9/22/2009
'Drugroom Preview modifications:
'
'DrugroomPreview class module gathers recipe date from central database DyelotsBulkedRecipe
' to do so requires the local creation of a new ODBC Connection to the server central database
'Required: Create an OBDC link on each controller to reference the BatchDyeingCentral database
'   Administrative Tools > ODBC Data Source Administrator > System DSN (tab)
'     Add: "BatchDyeingCentral: as SQL Server
'     Server: "Adaptive-Server" or "10.1.21.200"
'     SQL Server Authentication: LoginID: Adaptive  Password: Control
'     Change The Default Database To: "BatchDyeingCentral
'     Test Data Source (ensure working connection) Select OK
' add a spot to la command to get the recipe from the next lot.

'9/1/2009
'Alarms_RedyeIssueCheckMachineTank
'   replication issue where job becomes redye after starting, so that Parent.Job matches original
'     while database records redye.  LA code now uses database values (not parent.job) as fix.
'   need to check for this issue repeat and signal operator to check tank to prevent redispense
'   incorporate P_EnableRedyeAlarm to prevent alarm function once @1 bug resolved due to possible recycle issue
'   Add prepare class module will wait in "WaitIdle" state if alarm active and enabled until operator holds run for 2secs to reset alarm
'   if alarm is not enabled, add prepare flag will still be set and cleared once KP cancels, for history reference & debugging

'8/21/2009
'Add pDyelot, pRedye, pJob to KP command and determine value based on same logic as within LA command using database values for state=2
'Use LA.Job (determined from dyelot, not parent.job) and compare to this new KP.Job to verify current Job corresponds to what the LA prepared for
'   Bug appears when Lot replicated as a redye (job@1) when it's actually the first run, and dyelot value shows redye while parent.job does not
'   May edit preparation to disregard redye based on allen micheal's request, so DispenseJob = pdyelot and not pjob (pdyelot & "@" predye) so that
'     if redye is scheduled, the dispensers will still work based on original recipe.  allen says a&e doesn't only use the batch number for one PO
'     and if it's rescheduled, then they would want to redispense the recipe for the original purchase order recipe.  he's getting approval for that
'     change to make sure it's what everyone wants.  in the event that they want to schedule a job as a manual dispense (no automated recipe) they'll
'     schedule the original PO batch number with an underscore at the end to disable the recipe (job_)
'Issue where program was written using IP's, each having an LA, so that two LA's were back to back before a KP.  Resulted in first LA picking up
'   next KP, then shortly later, the next LA picking up the KP again, resulting in double the drop.

'8/5/2009
'At KP.Start - check to see if currently active and if so, check LA.Job to match Parent.Job and if does, reset the LAActive & LARequest to false
'  Add code to set LA.Job for subsequent KP's within currently running job using database method (not set to Parent.Job due to inconsistency)
'Setting pDyelot & pRedye allows us to test current Job = job prepared for
'  a&e use LA's within treatments (IP's) to prepare next KP in current program
'  if they use the an LA within the same IP back to back, the last LA will pick
'  up the previous calloff instead of the next scheduled job...
'  this will make sure that the next job has a LA.Job to compare instead of an empty string

'6/22/2009
'added RF & AF 300 second timer overrun delay to commands to allow time to fill & heat before signalling delay - allen michael wanted it removed
'  initially and delay if waiting, then decided after seeing the actual cycles that he didn't really believe it was delayed when waiting on the
'  drugroom (kp/la/ka) to dispense to chosen tanks before preparing tanks

'6/15/2009
'Added Parameters_StandardLoadTime and StandardUnloadTime in minutes
'changed StandardReserveTransferTime to be in minutes
'Added Parameters_SevenDaySchedule to set to '1' for full work week and leave at '0' for sunday startup
'   goal is to disregard initial load delays on sunday startup if not working a 7-day week
'   addition requested by allen-micheal due to excessive delays on initial work-week startup and he says to always assume startup
'     will be on sunday (with excpetion of certain holidays that may shift startup to different day, but to not worry about those delays)
'   Reason is that some programs are written with the dispene (kp) before the load command so that dispenses could complete while the
'     new job is scheduled/unblocked and waiting on loading due to all the machines waiting on a dispense at once...
'Fixed the delay interaction for RF.IsOverrun to include RF.IsInterlocked (waiting on KP dispense) to prevent false delay_localprepare
'Fixed Delay_WaitResponse and WaitReady to not wait until the DispenseTimer is finished (timer used for alarm conditions, while the delay
' is true if it's waiting on dispenser and not alarmed...
'Fixed LD.IsOverrun delay, property definition was not correct so that if state = LDOn, overrun was true even though timer wasn't finished
'
'6/12/2009
'Added paramter_FillOpenDelayTime to delay opening fillvalve when requested for this time, in seconds, limited to 10seconds max, to
' ensure that slower opening topwash valve is fully open before opening fill.  Determined that on some machines with slower top wash,
' fill valve opens resulting in machine safety pressure dropping out for the interlock state to depressurize and then restart, resulting
' in a loop cycle of fill/interlock/depressurize until the top wash opens enough for flow (determined by machine auditor and change requested
' by david hipp)

'6/10/2009
'Updating Delay code to help A&E report efficiencies for their cycle "turns" focus.  They want delays to be associated with the actual reason
'   causing the delay (dispenser response as well as which <chemical/dye>, or tank 1 preparation <fill/heat>, or tank 1 operator <wait ready/mixing>,
'   or local reserve mechanical delay.
'Added (and renamed) following delays:
'  Public Delay_Tank1Prepare As Boolean                  '16      <<was Delay_KitchenPrepare>> (mechanical delay due to fill/heat)
'  Public Delay_ReserveTransfer As Boolean               '17      (mechanical delay due to RT transfer overrun - parameter_standardreserver
'  Public Delay_DispenseWaitReady As Boolean             '18      (autodispenser delay)
'  Public Delay_DispenseWaitResultDyes As Boolean        '19      (waiting on dye dispenser)
'  Public Delay_DispenseWaitResultChems As Boolean       '20      (waiting on chemical dispenser)
'  Public Delay_DispenseWaitResultCombined As Boolean    '21      (waiting on both types - could be either dispenser)
'  Public Delay_LocalPrepare As Boolean                  '22      (waiting on local tank to fill/heat/mix <mechanical>)
'  Public Delay_Tank1Operator As Boolean                 '23      (waiting on tank 1 operator to signal ready or tank 1 to finish mixing)
'  Public Delay_Tank1DispenseError As Boolean            '24      (waiting on tank 1 ready/mixed due to unexpected dispense error)

'removed lookaheadrequestdelay and parameters enabling it

'5/22/2009
'Added currentjob string to LA command to always set the LA.Job through the pDyelot and pRedye variables by using the Parent.Job value for a LA
'   from the same job (ProgramHasAnotherKP=true).
'    Setting pDyelot & pRedye will allow us to test that the current job is consistent even within the current program
'     a&e likes to use LA's within treatments (IP's) to prepare next KP in current program, and if they use the an LA within
'     the same IP back to back, the last LA will pick up the previous calloff instead of the next scheduled job...
'     this will make sure that the next job has a LA.Job to compare instead of an empty string ""
'Also set LARequest false after CheckNextDyelot finds a scheduled job (unblocked) to prepare
'Also set appropriate destination tank false (1=add/2=reserve) in KP.Start if LA is active or completed but the job doesn't match Parent.Job
'     this will then drain the destination tanks (manualtankdrainPB's) and the KP will start with the updated parameters
'     *were not clearing the tank ready flags before which would allow the tank to begin transfer
'

'5/13/2009
'Add ControlCode.LARequest boolean to display LA is currently looking to the next scheduled batch for a tank to prepare - a&e want's to be able to sort
'   by this value to ensure that, when scheduling from the intray, they make sure to schedule to currently running machines that are waiting on a
'   new job to prepare (more job turns, better efficiency, less wasted time)
'Add Parameters_LookAheadRequestDelay to enable new Delay_LookAheadRequestActive.  if enabled (set to 1) then as soon as LARequest is set to true,
'   delay will start (Delay=22)


'4/28/2009
'Add PumpRequest to alarms_lidnotlocked to prevent the alarm from being active if the pump is not running (start of program before load command)
'Add Rinse templow/temphigh alarm disable if command parameters are set a 0 (all cold) or 180 (all hot) so that actual temperature deadband is
'   disregarded as not critical
'   Also - RinseTemperature alarm only considers IO_BlendFillTemp when signalling alarm (not VesTemp as it was - false alarms when rinsing after hot drop
'       everyone says that the critical temperature is the fill temp flowing through packages.)
'Edit Alarms_LowFlow if DR/RH/RI/FI.Settle are active...
'  disregardFlow_ = DR.IsOn Or FI.IsOn Or FR.IsReversing Or FC.IsReversing Or RT.IsOn Or _
'                   RI.IsFilling Or RH.IsFilling Or RW.IsOn Or _
'                   (Not IO_MainPumpRunning)

'3/4/2009
'Fixed AF command, when filling from machine where state would be stuck in FillPause to pausetimer not resetting back to 5 when timed out
'Add StateString to several commands to clean up the controlcode.status method as well as improve the visibility, clearness of code
'Add WaitIdle state to AP & RP commands to ensure that when entering the command from another operator command (ie LD or SA...) require the RemoteRun to
'   to be off before using the AdvanceTimer.  From histories, it appears that when going from a LD through and RP to and RT, the held RemoteRun will
'   complete the LD command and also set the ReserveReady in the same instance instead of looking for remoterun to have be off first.
'Change Delay_Dispenser to be Delay_DispenseWaitReady and Delay_DispenseWaitResultDyes & Delay_DispenseWaitResultChems
'   determine which type of drop is active using dispenseproducts string (sent by AutoDispenser) and looking at the product code
'     Split DispenseProduct string into array based on "|" and look for each row with ":" (Row contains product code)
'     If ProductCode is greater than 1000, the product is a chemical and below 1000 is a dye.  We can tell the ProductCode length by if the
'       ":" is at the 5th position of the string after the "|"  Use DispenseDyeOnly/ChemsOnly/DyesChems booleans to display which is active
'       in addprepare class as well in determining delay differentiation.
'Add Beacon Alarm (Red) light on for lid conditions where lid is not fully closed (pin raised) or else lid fully open (lid open Switch made)
'   alarm = Not ((IO_LidLoweredSwitch And IO_LockingPinRaised) Or (IO_LidRaised_SW))
'Add Alarms_TopWashBlocked alarm to detect PressureInterlock input lost while running wash and alarm if enabled.
'   Parameters_TopWashBlockageTimer            (set to 10-30) (or '0' for disabled)
'Variable level control like gaston machines
'   Parameters_NumberOfSpindales               (per machine)
'   Parameters_EnableFillByWorkingLevel        (set to 1)
'   Parameters_EnableReserveFillByWorkingLevel (set to 1)
'   Parameters_PackageTypeLevel(A-H)(1-9)

'7/10/2008
'Set RP command to, when switching to the rp.Ready state, to set ReserverReady to true....
'Set AP command to, when switching to the ap.Ready state, to set AddReady to true...


'6/18/2008  (updated controlcode & mimic)
'Fix the ManualAddDrain & ManualReserveDrain pushbuttons on the mimic so that after the final drain, they pushbuttons turn off by setting a timer in the
'   acmanualadd(reserve)drain modules to remain in the turnoff state until the mimic pushbuttons are off and a timer finishes.  it appears that there was
'   a slight lag between the controller/mimic so that the mimic pushbutton remained on resulting in the drain functions continuing to cycle.

'6/11/2008
'Yet one more change as a result of the RF update - removed the ReserveReady & AddReady reset flags in the Parent. program running state changes section.
'   Controlcode was resetting the tank ready states due to the fact that in the old code (front side) the dispenser only went to the drugroom tanks, not
'   to the downstairs tanks, so the LA would dispense to the kitchen and the next program would transfer the LA prepared kitchen tank downstairs so that
'   reserveready & addready must be false at the beginning of a program.  The platform machines, however, dispense directly to the reserve & add tanks,
'   so we don't want to reset them the same.

'6/11/2008
'RT command waits in State=WaitReady until ReserveReady and ReserveTank is filled and heated - RF was cancelling the ReserveReady flag in controlcode
'   which resulted in a delay due to the tank not being ready....now we don't reset the ReserveReady with RF - we only fill (top up) and heat to desired
'   temp...The RT command makes sure to wait until the tank is filled & heated if the RF.ison.

'6/5/2008
'Delay_Dispenser - Set equal to KP.KP1.IsDispenseResponseOverrun & KP.KP1.IsDispenseReadyOverrun (to distinguish between drugroom operator overruns &
'   dispenser overruns.
'AF - Backfill going to FillPause state, and backfilltimer never counting down - change code to use FillType as long (instead of variant) so that PE will
'  record the value in histories.  AF Param(1) |HCMV| |0-99%| Letter Parameters would normally be a variant type, but Plant explorer doesn't display
'  variants so convert to long
'RF - Did the same update here as with the AF
'Changed Graph Logging Colors to differentiate the variables:
'graph colors
'1- light red   -   SetpointF
'2- pink        -   VesTemp
'3- light blue  -   FlowRatePercent   (was 4)
'4- dark blue   -   VesselLevel
'5 - dark red   -   ReserveLevel
'6- green       -   AdditionLevel
'7- yellow      -   Tank1Level
'8- light red   -   VesselPressure    (was 4)
'9- pink        -   PackageDifferentialPressure   (was 7)

'4/11/2008
'LS - Look Ahead Stop
'Configured IO to achieve reserver backfill from vessel to desired level by using airpad w/ vent closed


'3/5/2008
'Set output for holddown to also be kier coupling (coupling valve only installed on 12-tube machines and could be used with a half-load (HL) command
' to close and run only 6-tubes.  6 & 12-tube machines do not have carrier holddowns, and there are no more available digital outputs for this plc.
'Fixed WK to ensure that it will wait for KP.isOn / LAActive / LA.KP1.IsOn (was only waiting for Kp.ison) Also set the WK to reset both tank ready
' states to false (will only use a WK between two KP's going to the same destination and never preparing both downstairs tanks at the same time).
'Also made it possible to RT after Filling by setting Airpad = false and the beginning of an RT to depressurize the machine and then set it back to
' true at the end.  This may cause a slight delay if a FI is after an RT as the system will need to depressurize again.

'12/19/2007
'AF backfill from machine to fast on 72 package machines due to mechanical valve stops not being set correctly resulting in too much back flow
'  creating a vortex in the add tank where the level is much higher than the recorded value.  Ultimately results in overfilling the add tank with
'  kier volume running out onto the motors below.  Added a FillPause state to fill/pause in 5second intervals when backfilling from the machine to
'  prevent this from occuring - will slow the flow down to a measurable level.
'KR command updated to include parameters to control rinsing to both machine and drain.  if desired, you can prevent rinsing to the machine, drain,
'  or both.  Still cancels KR at completion of KP command as before so the KP only works once for next KP.
'Parameters_PumpSpeedDecelTime given a minimum value of 15seconds to prevent someone setting the parameter value to 0 seconds resulting in pump
'  direction switching immedietly.  Tried to verify other time critical parameters also had a minimum value as well.
'RP advancetimer tied to IO_RemoteRun pushbutton.  Hold for 2seconds to set the RP Ready condition
'AP advancetimer tied to IO_RemoteRun pushbutton.  Hold for 2seconds to set the RP Ready condition
'IO_TestTemp added, channel 4 main plc - maintanence guys requested this channel so that they could connect a simulator and verify the PC temp
'  accuracy.  Not used in code - only display input on IO screen.
'Prevent AT/RT from transferring to machine if lid isn't locked

'11/6/07
'Resetting DispenseCalloff in AddPrepare to '0' if dispense is not enabled, so as to clear the autodispenser variables, which will then allow the
'  kp to be jumped back through, resetting the dispense.
'Add Parameter_ReserveBackfillLevel to use reserve runback valve with airpad during RF command to backfill from V to desired parameter level.
'  due to design of machines (levels equalize between kier and reserve at different levels based on height and diameter of reserve tank) each
'  machine will backfill to a max level that may differ from all other machines.  instead of using a max height reached with a timer, we'll just
'  backfill to a reachable value (set by parameter) and then fill the rest of the RF parameter level with water.  in this way, if a machine stops
'  being able to fill to the parameter value, the supervisor will see this as a delay in the RF filling step (made RF - from vessel) a foreground
'  step to further signify this issue.
'Added code to compare LA.Job to Parent.Job and if a difference is found, the appropriate manualdrainpushbutton is activated to drain the tank
'  before allowing to reset the LA conditions and dispensing the right drop.  Now looking for manualdraincontrol to be active at beginning of
'  addprepare module and wait for it to finish before stepping on.
'If Parameter_LookAheadEnabled <>1, then reset LAActive to false (just to help quickly fix issues where they've jumped around).  Allows redispensing.

'7/23/07
'Added P_DispenseResponseDelayTime & P_DispenseReadyDelayTime used with new alarms: DispenseReadyDelay & DispenseResponseDelay to insure
'that in the event that if either the autodispenser on the server does not give the ready signal or the actual dispenser does not give a response
'within the parameter times (in minutes) then the appropriate alarm will signal and the host will display the issue highlighted in red.  This comes
'as a result of the chemical dispenser hanging up for an hour resulting in 7 lots to become overrun with the desired drops being stuck in the quenue.

'6/12/07
'Correct WK command to set RT or AT Ready flag to false for the instances where you have a KP, WK, and then KP, before a RT or AT.
'Insert a WaitIdle state in AddPrepare steps to prevent dispensing to a tank that's still running the AD or RD background cleaning procedure
'Fix LA command by resetting the IPKP1 string at the beginning of a LA, otherwise it keeps the previous value
'Add SettleTank state in AT and AC commands to allow tank to settle after rinsing before restarting the pump to transfer the tank over

'5/15/07
'Add AC - Add-By-Contacts command
'Add KR - Kitchen Rinse = Turn Off Kitchen Rinse to Machine, only rinse to drain
'Setup the manual switch to work and the mixer button to be a flag in manual - and also prevent transfers (restart prepare procedure)
'Pulsing the AddMix valve while mixingfortime to clean out mix valve chamber dye contamination
'  added a parameter for pulse time to also act as a disable in the event of add pump tripping due to defective valves or solenoids
'Stopped refilling the drugroom tank if the actual level falls below the desired level (MC-36 comms drops out occasionally resulting in
'  the tank overfilling.)  If tank is switched to manual and back to run, it will refill the tank and reheat as desired.
'Changed/Added a lot of the status/display properites to give more information, especially when waiting for drugroom prepare, transfer states
'Stop setting the analog inputs to '0' if noresponse from plc. Prevents the issues from communication noise...
'  *tried setting drugroom baudrate to 19200, but resulted in drugroom plc timeout requiring powering off/on the plc to reset

'5/7/07 edited code for 2-kier (12package) machines, adding lid2switch input & coupling valve output
'add pumprequest mimic pushbutton
'add heatallrequest mimic pushbutton as per david hipp and shop...they want to turn on all heating if maintanance pb pressed
'changed some of the display status to make the states more clear

'4/1/07 changed kp to work with chemical dispense
'and dye disense
'changed tank1dispenseenabled param to dyedispenseenabed
'add param for chemicaldispenseenabled
'changed kp command,la command and add prepare.
'removed the ka command kp now transfers also' not sure if this will work
'had to change rt command and at command

'2/23/2007
'fixed dispense error
'fixed la for ip
'added lamp for drugroom dispense state
'fixed the hc for the small machines changed hcvolume to a double
'changed systemvolume to a double for small machines
'changed rcvolume to a double for small machines

'1/29/07 - dh
'converted then package code to work on small platform then 6pk machines (changed IO layout with addition/removal)

'1/5/07 - mw
'added dispenseing stuff into the code.

'12/08/06
'changed hot drain to not fill turned off airpad on fill.

'09/20/06
'fixed the high temp on a rinse

'09/15/06
'added alarm for temp to high in reserve tank
'added alarm for temp to high in tank 1
'add low flow amount

'08/15/06
'fixed flow reverse command for flow contacts

'08/03/06-mw
'made a change so vessel level has to be lost for 2 seconds before level to low for pump

'07/26/06 - mw
'changed the plc from eaton to automation direct

'04/25/06 - mw
' if in look ahead to next dyelot (overlay) then set commited = 1
'changed kp
' fixed temp low and high alarms for tp and max grad
' if the drugroom tank is mixing and it is made unready go to the slow state added.

'04/11/2006 -mw
'added varible to count pump on and off times if level is to low.
'reset if halt button pushed

'11/14/2005-mw
'need to put look ahead stuff back in
'pot controls for drugroom.

#End Region


'CODE INFORMATION
'******************************************************************************************************
'******            Working with Step Parameters for Gradient, Target, & Time           ****************
'******************************************************************************************************
'if minutes only, standard time = minutes
'if only target - vertical line to temp setpoint and no minutes
'if gradient and target and gradient - calculate time based on two
'if gradient, target, and time - 1st calcualte time based on gradient/target then add time as extra
'******************************************************************************************************
'******************************************************************************************************

'Smaller Touchscreen (800x600)
'Mimic .png size 
'Width: 769 pixels
'Height: 460 pixels
'resol: 96 dpi

'Larger Touchscreen (1024x768)
'Mimic .png size 
'Width: 999 pixels
'Height: 620 pixels
'resol: 96 dpi

'******************************************************************************************************
'****************                     Reasonable Graph Windows Colors                  ****************
'******************************************************************************************************
' Maroon                Teal
' DarkRed               DarkCyan  
' Red                   CadetBlue
' Brown                 SteelBlue
' Firebrick             DodgerBlue
' OrangeRed             MidnightBlue
' Sienna                Blue
' SaddleBrown           DarkSlateBlue
' Peru                  Indigo
' DarkOrange            DarkOrchid
' DarkGoldenrod         DarkViolet
' Goldenrod             BlueViolet
' Olive                 DarkMagenta
' OliveDrab             Crimson              
' DarkOliveGreen        Purple
' ForestGreen           RoyalBlue
' DarkGreen
' Green
' LimeGreen
' Lime
'******************************************************************************************************
