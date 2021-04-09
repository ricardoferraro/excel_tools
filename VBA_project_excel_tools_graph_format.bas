Attribute VB_Name = "Módulo1"
Sub Formatar_grafico()
    Dim cht As Chart
    Dim sr1 As Series
    Dim dl As DataLabels
' Source: https://exceloffthegrid.com/vba-code-for-charts-and-graphs-in-excel/

'Create an empty chart embedded on a worksheet.
'Set cht = Sheets("Graficos").Shapes.AddChart2(201, xlColumnClustered).Chart
Set cht = Sheets("Graficos").Shapes.AddChart2.Chart

'Select source for a chart
Dim rng As Range
'Set rng = Sheets("Graficos").Range("B3:G3")
cht.SetSourceData Source:=Range("Graficos!B3:G3")
cht.SeriesCollection(1).Name = "=""PB"""

'Set the size/position of a ChartObject - method 1
cht.Parent.Height = 200
cht.Parent.Width = 300
cht.Parent.Left = 100
cht.Parent.Top = 100

cht.ChartGroups(1).GapWidth = 120

'=================== AXIS =====================
'Set chart axis min and max
cht.Axes(xlValue).MaximumScale = 25
cht.Axes(xlValue).MinimumScale = 10
cht.Axes(xlValue).MaximumScaleIsAuto = True
cht.Axes(xlValue).MinimumScaleIsAuto = True

'Display axis
cht.HasAxis(xlCategory) = True

'Hide axis
'cht.HasAxis(xlValue, xlSecondary) = False

'Display axis title
'cht.Axes(xlCategory, xlSecondary).HasTitle = True

'Hide axis title
'cht.Axes(xlValue).HasTitle = False

'Change axis title text
'cht.Axes(xlCategory).AxisTitle.Text = "My Axis Title"

'Reverse the order of a catetory axis
'cht.Axes(xlCategory).ReversePlotOrder = True

' ============= GRIDLINES ====================

'Add gridlines
'cht.SetElement (msoElementPrimaryValueGridLinesMajor)
'cht.SetElement (msoElementPrimaryCategoryGridLinesMajor)
'cht.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
'cht.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)

'Delete gridlines
cht.Axes(xlValue).MajorGridlines.Delete
cht.Axes(xlValue).MinorGridlines.Delete
cht.Axes(xlCategory).MajorGridlines.Delete
cht.Axes(xlCategory).MinorGridlines.Delete

'Change colour of gridlines
'cht.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(255, 0, 0)

'Change transparency of gridlines
'cht.Axes(xlValue).MajorGridlines.Format.Line.Transparency = 0.5

' ================= CHART TITLE ======================
'Display chart title
'cht.HasTitle = True

'Hide chart title
cht.HasTitle = False

'Change chart title text
'cht.ChartTitle.Text = "My Chart Title"

'Position the chart title
'cht.ChartTitle.Left = 10
'cht.ChartTitle.Top = 10

'Format the chart title
'cht.ChartTitle.TextFrame2.TextRange.Font.Name = "Calibri"
'cht.ChartTitle.TextFrame2.TextRange.Font.Size = 16
'cht.ChartTitle.TextFrame2.TextRange.Font.Bold = msoTrue
'cht.ChartTitle.TextFrame2.TextRange.Font.Bold = msoFalse
'cht.ChartTitle.TextFrame2.TextRange.Font.Italic = msoTrue
'cht.ChartTitle.TextFrame2.TextRange.Font.Italic = msoFalse

' ================= LEGEND ======================

'Display the legend
cht.HasLegend = True

'Hide the legend
'cht.HasLegend = False

'===================== Legend ====================

'Position the legend
'cht.Legend.Position = xlLegendPositionTop
'cht.Legend.Position = xlLegendPositionRight
'cht.Legend.Position = xlLegendPositionLeft
'cht.Legend.Position = xlLegendPositionCorner
cht.Legend.Position = xlLegendPositionBottom

'Allow legend to overlap the chart.
'False = allow overlap, True = due not overlap
'cht.Legend.IncludeInLayout = False
cht.Legend.IncludeInLayout = True

'Move legend to a specific point
cht.Legend.Left = 40
cht.Legend.Top = 190
cht.Legend.Width = 180
cht.Legend.Height = 25


'Set the size and position of the PlotArea
'cht.PlotArea.Left = 20
'cht.PlotArea.Top = 20
'cht.PlotArea.Width = 200
'cht.PlotArea.Height = 150

'================== Series =================
'Referencing a chart series by name
Dim srs As Series
'Set srs = cht.SeriesCollection("Series Name")

'Change series source data and name
'srs.Values = "=Sheet1!$C$2:$C$6"
'srs.Name = "=""Change Series Name"""

'Add a new chart series
Set srs = cht.SeriesCollection.NewSeries
srs.Values = "=Graficos!$B$4:$G$4"
srs.Name = "=""Reclamações"""
'Set the values for the X axis when using XY Scatter
'srs.XValues = "=Sheet1!$D$2:$D$6"

'cht.FullSeriesCollection(1).ChartType = xlColumnClustered
'cht.FullSeriesCollection(1).AxisGroup = 1
'cht.FullSeriesCollection(2).ChartType = xlLine
'cht.FullSeriesCollection(2).AxisGroup = xlSecondary

'cht.SeriesCollection(1).ChartType = xlColumnClustered
'cht.SeriesCollection(1).AxisGroup = 1
cht.SeriesCollection(2).ChartType = xlLine
cht.SeriesCollection(2).AxisGroup = xlSecondary
cht.SeriesCollection(2).Smooth = True


'cht.HasAxis(xlCategory, xlPrimary) = True
'cht.HasAxis(xlCategory, xlSecondary) = True
cht.HasAxis(xlValue, xlPrimary) = True
cht.HasAxis(xlValue, xlSecondary) = True
cht.Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
cht.Axes(xlCategory, xlSecondary).CategoryType = xlAutomatic
cht.Axes(xlValue, xlPrimary).MinimumScale = 0
cht.Axes(xlValue, xlSecondary).MinimumScale = 0
cht.Axes(xlValue, xlPrimary).MaximumScale = 40000
'cht.Axes(xlValue, xlSecondary).MaximumScale= xxx
cht.Axes(xlCategory).CategoryNames = Range("Graficos!B2:G2")
cht.Axes(xlValue, xlPrimary).TickLabels.Font.Color = RGB(255, 255, 255)
cht.Axes(xlValue, xlSecondary).TickLabels.Font.Color = RGB(255, 255, 255)

'================= Data Labels ======================

'Display data labels on all points in the series
srs.HasDataLabels = True

'Hide data labels on all points in the series
'srs.HasDataLabels = False

'Position data labels
'The label position must be a valid option for the chart type.
srs.DataLabels.Position = xlLabelPositionAbove
'srs.DataLabels.Position = xlLabelPositionBelow
'srs.DataLabels.Position = xlLabelPositionLeft
'srs.DataLabels.Position = xlLabelPositionRight
'srs.DataLabels.Position = xlLabelPositionCenter
'srs.DataLabels.Position = xlLabelPositionInsideEnd
'srs.DataLabels.Position = xlLabelPositionInsideBase
'srs.DataLabels.Position = xlLabelPositionOutsideEnd



'================= Series formatting ======================

' Colors - http://dmcritchie.mvps.org/excel/colors.htm
'cht.SeriesCollection(1).Interior.Color = RGB(192, 192, 192)
cht.SeriesCollection(1).Interior.Color = RGB(128, 128, 128)


    cht.FullSeriesCollection(1).Select
    cht.FullSeriesCollection(1).ApplyDataLabels
    cht.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.Position = xlLabelPositionCenter
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue

'Change fill colour
srs.Format.Fill.ForeColor.RGB = RGB(55, 163, 119)

'Change line colour
srs.Format.Line.ForeColor.RGB = RGB(55, 163, 119)

'Change visibility of line
srs.Format.Line.Visible = msoTrue

'Change line weight
srs.Format.Line.Weight = 2

'Change line style
'srs.Format.Line.DashStyle = msoLineDash
srs.Format.Line.DashStyle = msoLineSolid
'srs.Format.Line.DashStyle = msoLineSysDot
'srs.Format.Line.DashStyle = msoLineSysDash
'srs.Format.Line.DashStyle = msoLineDashDot
'srs.Format.Line.DashStyle = msoLineLongDash
'srs.Format.Line.DashStyle = msoLineLongDashDot
'srs.Format.Line.DashStyle = msoLineLongDashDotDot

'Changer marker type
'srs.MarkerStyle = xlMarkerStyleAutomatic
srs.MarkerStyle = xlMarkerStyleCircle
'srs.MarkerStyle = xlMarkerStyleDash
'srs.MarkerStyle = xlMarkerStyleDiamond
'srs.MarkerStyle = xlMarkerStyleDot
'srs.MarkerStyle = xlMarkerStyleNone
srs.MarkerBackgroundColor = RGB(255, 255, 255)

'Set dl = cht.SeriesCollection(1).DataLabels
'For i = 1 To dl.Count
'    With dl(i)
 '       .Font.Color = RGB(255, 255, 255)
 '   End With
'Next

End Sub

Sub GetRid()
  Dim ASheet As Worksheet
  Dim AChart As Chart


  Application.DisplayAlerts = False
  Application.ScreenUpdating = False

  '** first scan for and delete all non HOME worksheets ***
  'For Each ASheet In ActiveWorkbook.Worksheets
  '  If UCase(ASheet.Name) <> "HOME" Then
  '    ASheet.Delete
  '  End If
  'Next
  nCharts = ActiveSheet.ChartObjects.Count
  '** Now scan and delete any ChartSheets ****
    For iChart = 1 To nCharts
        ActiveSheet.ChartObjects(iChart).Delete
    Next
    
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub

Sub ArrangeMyCharts()
' Source: https://peltiertech.com/Excel/ChartsHowTo/QuickChartVBA.html
'Create an Array of Charts
'Suppose you have a lot of charts on a worksheet, and you'd like to arrange them neatly. The following procedure loops through the charts, resizes them to consistent dimensions, and arranges them in systematic rows and columns:
    
    Dim iChart As Long
    Dim nCharts As Long
    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double
    Dim dWidth As Double
    Dim nColumns As Long

    dTop = 75      ' top of first row of charts
    dLeft = 100    ' left of first column of charts
    dHeight = 225  ' height of all charts
    dWidth = 375   ' width of all charts
    nColumns = 3   ' number of columns of charts
    nCharts = ActiveSheet.ChartObjects.Count

    For iChart = 1 To nCharts
        With ActiveSheet.ChartObjects(iChart)
            .Height = dHeight
            .Width = dWidth
            .Top = dTop + Int((iChart - 1) / nColumns) * dHeight
            .Left = dLeft + ((iChart - 1) Mod nColumns) * dWidth
        End With
    Next
End Sub

