Partial Class ControlCode

  'Add all the commands to this collection in the new section of the command so we have
  '  a convenient way to loop through every command in the control code
  Friend Commands As New List(Of ACCommand)

  ' ADD TANK 
  Public ReadOnly AC As New AC(Me)
  Public ReadOnly AD As New AD(Me)
  Public ReadOnly AF As New AF(Me)
  Public ReadOnly AP As New AP(Me)
  Public ReadOnly AT As New AT(Me)
  ' AV Not used with Then ThenPlatform ProgramGroup

  'KITCHEN
  Public ReadOnly CK As New CK(Me)
  Public ReadOnly KA As New KA(Me)
  Public ReadOnly KP As New KP(Me)
  Public ReadOnly KR As New KR(Me)
  Public ReadOnly WK As New WK(Me)

  'MACHINE
  Public ReadOnly DR As New DR(Me)
  Public ReadOnly FI As New FI(Me)
  ' FC ' TODO?
  Public ReadOnly HC As New HC(Me)
  Public ReadOnly HD As New HD(Me)
  Public ReadOnly PR As New PR(Me)
  Public ReadOnly RC As New RC(Me)
  Public ReadOnly RH As New RH(Me)
  Public ReadOnly RI As New RI(Me)
  Public ReadOnly TM As New TM(Me)

  ' LOOK AHEAD FUNCTIONS
  Public ReadOnly LA As New LA(Me)
  Public ReadOnly LS As New LS(Me)

  ' FLOW CONTROL
  Public ReadOnly CC As New CC(Me)
  Public ReadOnly DP As New DP(Me)
  Public ReadOnly FC As New FC(Me)
  Public ReadOnly FL As New FL(Me)
  Public ReadOnly FP As New FP(Me)
  Public ReadOnly FR As New FR(Me)

  'OPERATOR
  Public ReadOnly LD As New LD(Me)
  Public ReadOnly PH As New PH(Me)
  Public ReadOnly SA As New SA(Me)
  Public ReadOnly UL As New UL(Me)

  'REPORTS
  Public ReadOnly BO As New BO(Me)
  Public ReadOnly BR As New BR(Me)
  Public ReadOnly ET As New ET(Me)

  ' RESERVE TANK
  Public ReadOnly RD As New RD(Me)
  Public ReadOnly RF As New RF(Me)
  Public ReadOnly RP As New RP(Me)
  Public ReadOnly RT As New RT(Me)
  Public ReadOnly RW As New RW(Me)

  ' SETUP
  Public ReadOnly BW As New BW(Me)
  Public ReadOnly HL As New HL(Me)
  Public ReadOnly NP As New NP(Me)
  Public ReadOnly PT As New PT(Me)

  ' TEMPERATURE
  Public ReadOnly CO As New CO(Me)
  Public ReadOnly HE As New HE(Me)
  Public ReadOnly TP As New TP(Me)
  '  Public ReadOnly TC As New TC(Me) ' TODO
  Public ReadOnly WT As New WT(Me)

End Class
