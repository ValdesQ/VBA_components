Attribute VB_Name = "Formats"
Option Explicit
Sub prcFormat01(ByVal tempRange As Range)
    On Error Resume Next
    With tempRange.Borders(xlDiagonalDown)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tempRange.Borders(xlDiagonalUp)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tempRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tempRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tempRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tempRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    tempRange.Borders(xlInsideVertical).LineStyle = xlNone
    tempRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    On Error GoTo 0


End Sub


Sub prcMiddCells(ByRef oRNG As Range)

    With oRNG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    On Error GoTo 0
    
End Sub
Sub prcCross(ByRef oRNG As Range)
Attribute prcCross.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    oRNG.Borders(xlDiagonalDown).LineStyle = xlNone
    oRNG.Borders(xlDiagonalUp).LineStyle = xlNone
    With oRNG.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRNG.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRNG.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRNG.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRNG.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oRNG.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub prcGrey(ByRef oRNG As Range)
    On Error Resume Next
    
    With oRNG.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    oRNG.Interior.Color = 12566463
    
    On Error GoTo 0
    
End Sub

