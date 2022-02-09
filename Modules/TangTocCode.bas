Sub tat_che_do()
   With Application
      .Calculation = xlCalculationManual
      .ScreenUpdating = False
      .EnableEvents = False
      .DisplayAlerts = False
      .Cursor = xlWait
   End With
End Sub
Sub bat_che_do()
   With Application
      .Calculation = xlCalculationAutomatic
      .ScreenUpdating = True
      .EnableEvents = True
      .DisplayAlerts = True
      .CalculateBeforeSave = True
      .Cursor = xlDefault
   End With
End Sub
