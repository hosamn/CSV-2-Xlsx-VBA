Sub AddChart()
'
' AddChart Macro
'

'
    ActiveSheet.UsedRange.Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = _
        "Table1"
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    ActiveChart.SetSourceData Source:=ActiveSheet.ListObjects("Table1").Range
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.PlotArea.Select
    Selection.Height = 197.124
    ActiveChart.Legend.Select
    Selection.Left = 108.992
    Selection.Top = 7.124
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabels.Orientation = xlUpward
    ActiveChart.Axes(xlCategory).MinimumScale = 38718
    ActiveChart.Axes(xlCategory).MaximumScale = 44408
    ActiveSheet.Shapes("Chart 1").ScaleHeight 1.2065974045, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 1").ScaleWidth 1.3468751094, msoFalse, _
        msoScaleFromTopLeft
    ActiveChart.Legend.Select
    Selection.Left = 166.918
    Selection.Top = 21.422
    Application.CommandBars("Format Object").Visible = False
End Sub
