
Public Enum EventValue
  ProgramStarted = 1000
End Enum

Public Enum EDispenseResult As Integer
  Ready = 101
  Busy = 102
  Auto = 201
  Scheduled = 202
  Complete = 301
  Manual = 302
  [Error] = 309
End Enum

Public Enum EDispenseType As Integer
  DyesChems = 1
  DyesOnly = 2
  ChemsOnly = 3
End Enum

Public Enum EFillType As Integer
  Cold = 0
  Hot = 1
  Mix = 2
  Vessel = 3
End Enum

Public Enum EMixState
  Off = 0
  Active = 1
End Enum

Public Enum EKitchenDestination As Integer
  Drain
  Add
  Reserve
End Enum
