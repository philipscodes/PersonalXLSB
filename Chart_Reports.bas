Attribute VB_Name = "Chart_Reports"
Private Sub chartCreator()
'
' Creates Chart
'

'

Dim chartShape As Shape

    'ActiveSheet.Shapes.AddChart.Select
    Set chartShape = ActiveSheet.Shapes.AddChart
    chartShape.Chart.ChartType = xlColumnClustered
    chartShape.Chart.SetSourceData source:=Sheets("Data Analysis").Range("AQ2:AV14")
    chartShape.ScaleWidth 1.8479166667, msoFalse, _
        msoScaleFromBottomRight
    chartShape.ScaleHeight 1.5329862934, msoFalse, _
        msoScaleFromTopLeft
    chartShape.Chart.SeriesCollection(1).Select
    chartShape.Chart.SeriesCollection(1).Points(1).Select
    chartShape.Chart.SeriesCollection(1).Points(1).ApplyDataLabels
    chartShape.Chart.SeriesCollection(2).Select
    chartShape.Chart.SeriesCollection(2).ApplyDataLabels
    chartShape.Chart.SeriesCollection(3).Select
    chartShape.Chart.SeriesCollection(3).ApplyDataLabels
    chartShape.Chart.SeriesCollection(4).Select
    chartShape.Chart.SeriesCollection(4).ApplyDataLabels
    chartShape.Chart.SeriesCollection(5).Select
    chartShape.Chart.SeriesCollection(5).ApplyDataLabels
    chartShape.Chart.SeriesCollection(4).Select
    chartShape.Chart.SeriesCollection(1).Select
    chartShape.Chart.SeriesCollection(1).Trendlines.Add
    chartShape.Chart.SeriesCollection(1).Trendlines(1).Select
    
    With Selection
        .Type = xlMovingAvg
        .Period = 2
    End With
    
    With Selection.Format.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent1
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6000000238
        .Radius = 8
    End With
    
    chartShape.Chart.Axes(xlValue).Select
    chartShape.Chart.Axes(xlValue).MaximumScale = 800
    chartShape.Chart.Axes(xlValue).MaximumScale = 800
    chartShape.Chart.Axes(xlValue).MajorUnit = 100
    chartShape.Chart.Axes(xlValue).MajorUnit = 100
    chartShape.Chart.ChartArea.Select
    
    With chartShape.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0.3399999738
        .ForeColor.Brightness = 0
        .BackColor.ObjectThemeColor = msoThemeColorAccent1
        .BackColor.TintAndShade = 0.7649999857
        .BackColor.Brightness = 0
        .TwoColorGradient msoGradientHorizontal, 1
    End With
    
    chartShape.Fill.Visible = msoTrue
    chartShape.Fill.Visible = msoTrue
    chartShape.Chart.Parent.RoundedCorners = True
    
    With chartShape.line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    
    With chartShape.line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    
    chartShape.Shadow.Type = msoShadow38
    
    With chartShape.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent4
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6000000238
        .Radius = 18
    End With
    
    chartShape.Chart.PlotArea.Select
    Selection.Format.Fill.Visible = msoFalse
    chartShape.Chart.SetElement (msoElementChartTitleAboveChart)
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "5 Year Sales Depiction"
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 22).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 22).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 18
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    
    chartShape.Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Fiscal Year"
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 11).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 11).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    
    chartShape.Chart.SetElement (msoElementPrimaryValueAxisTitleHorizontal)
    chartShape.Chart.SetElement (msoElementPrimaryValueAxisTitleVertical)
    chartShape.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    chartShape.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Units Sold"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Units Sold"
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    
End Sub


Public Sub fiveYearSalesChart()

    Call chartCreator

End Sub
