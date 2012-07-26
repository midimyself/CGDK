Attribute VB_Name = "Openwindows"
Option Explicit

Sub OpenCalibrateMap()
    If Common.IsWaitingForAction Then Exit Sub
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    CalibrateMap.Calibrate
End Sub

Sub OpenShowGPSWindow()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If GPSLocation.Visible = True Then Unload GPSLocation Else GPSLocation.Show 0
End Sub

Sub OpenPlotOnMap()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not PlotOnMap.Visible Then PlotOnMap.Show 0
End Sub

Sub OpenMoveTo()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not MoveTo.Visible Then MoveTo.Show 0
End Sub

Sub OpenProjection()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not DrawProjection.Visible Then DrawProjection.Show 0
End Sub

Sub OpenCommonTool()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not CommonTools.Visible Then CommonTools.Show 0
End Sub

Sub OpenDrawStratigraphicColumn()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not DrawStratigraphicColumns.Visible Then DrawStratigraphicColumns.Show 0
End Sub

Sub OpenTemplateDesigner()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not TemplateDesigner.Visible Then TemplateDesigner.Show 0
End Sub

Sub OpenPlotGeochemicalDiagram()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not PlotGeochemistryDiagram.Visible Then PlotGeochemistryDiagram.Show 0
End Sub

Sub OpenSeriesEditor()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not SeriesEditor.Visible Then SeriesEditor.Show 0
End Sub

Sub OpenTemplateManager()
    If Documents.Count < 1 Then MsgBox "No CorelDRAW document available.": Exit Sub
    If Not TemplateManager.Visible Then TemplateManager.Show 0
End Sub

Sub OpenAbout()
    If Not About.Visible Then About.Show 0
End Sub

