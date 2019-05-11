Attribute VB_Name = "speed_up"
Sub speed_up()

    ' The below functions alters the default behaviour in Excel. Generally speaking this should increase overall macro speed.

    ' if any of these features are needed within the code it is better to temporarily turn on the needed feature and turn it off again once it is no longer needed.

    '   Turning features off
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .EnableEvents = False
    End With

    '   your code here
    '   your code here
    '   your code here
    '   your code here

    '   Turning features back on
    With Application
        .Calculation = xlSemiautomatic
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .EnableEvents = True
    End With

End Sub

Sub slow_feature_toggle(on_or_off As Boolean)

'   Another method for achieving the same results as speedUp(), using a boolean argument here allows for a simple one line call to toggle these features on/off,
'   This can be handy as it allows an easy reset during error trapping without all the extra code.

'   To use this subroutine simple call it with a boolean argument e.g. Call slow_feature_toggle(True)

    Application.ScreenUpdating = on_or_off
    Application.DisplayStatusBar = on_or_off
    Application.EnableEvents = on_or_off
        
'   Calc mode needs to be controlled via an if statement here as the default argument is not a boolean one. If you are looking for micro performance gains then
'   it is probably better to add a simple one liner to control the calc mode to avoid the branching here.
        
    If on_or_off = True Then
        Application.Calculation = xlSemiautomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
        
End Sub
