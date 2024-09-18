Public Sub sbrPerformanceOptimierung(ByRef blnStatus As Boolean)
    
    If blnStatus = True Then
    
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
            .AskToUpdateLinks = False
            .DisplayAlerts = False
            .DisplayStatusBar = False
            .Calculation = xlCalculationManual
            ' .Cursor = xlWait
        End With
    Else
    
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .AskToUpdateLinks = False
            .DisplayAlerts = True
            .DisplayStatusBar = True
            .Calculation = xlCalculationAutomatic
            '  .Cursor = xlDefault
        End With
    End If
End Sub