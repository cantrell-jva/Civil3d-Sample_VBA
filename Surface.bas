Attribute VB_Name = "Surface"
'
' MODULE_ID {Surface.bas}
' {PipeSample.dvb}
'
' Copyright {2026} by Autodesk, Inc.
'
' Permission to use, copy, modify, and distribute this software for
' any purpose and without fee is hereby granted, provided that the
' above copyright notice appears in all copies and  that both that
' copyright notice and the limited warranty and  restricted rights
' notice below appear in all supporting  documentation.
'
' AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
' AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
' MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
' DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
' UNINTERRUPTED OR ERROR FREE.
'
' Use, duplication, or disclosure by the U.S. Government is subject to
' restrictions set forth in FAR 52.227-19 (Commercial Computer
' Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
' (Rights in Technical Data and Computer Software), as applicable.
'
'
'
'
Option Explicit
'
' NOTE: There is a partially hard-coded path that needs to be
' modified if the file is in a different location.
'
'


'
' Load a series of points form a text file and create a
' surface from them.  Return a reference to the surface.
' If there is an error, this function will return Nothing.
'
' NOTE: This method depends on a hard coded path to the
' text file that needs to be modified.
'
Public Function CreateSurfaceByImportTXT() As AeccSurface
    If (g_oPipeDocument Is Nothing) Then
        Set CreateSurfaceByImportTXT = Nothing
        Exit Function
    End If

    Dim oSurfaces As AeccSurfaces
    Set oSurfaces = g_oPipeDocument.Surfaces
    
    ' Get a reference to the existing surface.  If
    ' it does not exist, make a new surface.
    Dim oSurface As AeccSurface
    Dim oTinSurface As AeccTinSurface
    On Error Resume Next
    Set oSurface = oSurfaces.Item(SURFACE_NAME)
    On Error GoTo 0
    If (oSurface Is Nothing) Then
        Dim oTinCreationData As New AeccTinCreationData
        oTinCreationData.Name = SURFACE_NAME
        oTinCreationData.Description = "TIN Surface from TXT file"
        oTinCreationData.Layer = g_oPipeDocument.Layers.Item(0).Name
        oTinCreationData.BaseLayer = g_oPipeDocument.Layers.Item(0).Name
        oTinCreationData.Style = CreateSurfaceStyle(SURFACE_STYLE_NAME).Name
        Set oTinSurface = oSurfaces.AddTinSurface(oTinCreationData)
    Else
        Set oTinSurface = oSurface
    End If
    
    Dim sFileName As String
    sFileName = InputBox("Location of text file with point data", "Surface Sample", g_oPipeApplication.Path & "\C3D\Sample\Civil 3D API\COM\Vba\Pipe\SamplePointFile.txt")
    oTinSurface.PointFiles.Add sFileName, "PENZD (space delimited)", False, False, False
    
    If Err.Number <> 0 Then
        MsgBox "Error importing a point file."
        Set CreateSurfaceByImportTXT = Nothing
        Exit Function
    End If

    ' Fill the screen with the surface, and show the triangles
    ' that make up the surface.
    ThisDrawing.Application.ZoomExtents
    
    Set CreateSurfaceByImportTXT = oTinSurface
End Function
 

'
' Create a new style for the surface.
'
Public Function CreateSurfaceStyle(sStyleName As String) As AeccSurfaceStyle
    Dim oSurfaceStyles As AeccSurfaceStyles
    Dim oSurfaceStyle As AeccSurfaceStyle
    
    ' Get the drawing's collection of surface styles.
    On Error Resume Next
    Set oSurfaceStyles = g_oPipeDocument.SurfaceStyles
    On Error GoTo 0
    If (oSurfaceStyles Is Nothing) Then
        Debug.Print "Error accessing collection of surface styles: " & vbNewLine & Err.Description
        Set CreateSurfaceStyle = Nothing
        Exit Function
    End If
    
    ' Get a particular style from the collection of styles.
    On Error Resume Next
    Set oSurfaceStyle = oSurfaceStyles.Item(sStyleName)
    On Error GoTo 0
    ' If the style does not exist, make a new one with that
    ' name.
    If (oSurfaceStyle Is Nothing) Then
        On Error Resume Next
        Set oSurfaceStyle = oSurfaceStyles.Add(sStyleName)
        On Error GoTo 0
        If (oSurfaceStyle Is Nothing) Then
            Debug.Print "Error trying to make a new surface style: " & vbNewLine & Err.Description
            Set CreateSurfaceStyle = Nothing
            Exit Function
        End If
    End If

    ' Create a style that shows the surface shape while making sure
    ' the pipe is most noticable.
    '
    ' 2D
    oSurfaceStyle.BoundaryStyle.DisplayStylePlan.color = 254
    oSurfaceStyle.BoundaryStyle.DisplayStylePlan.Visible = True
    oSurfaceStyle.TriangleStyle.DisplayStylePlan.color = 252
    oSurfaceStyle.TriangleStyle.DisplayStylePlan.Visible = True
    
    oSurfaceStyle.ContourStyle.MajorContourDisplayStylePlan.Visible = False
    oSurfaceStyle.ContourStyle.MinorContourDisplayStylePlan.Visible = False
    oSurfaceStyle.DirectionStyle.DisplayStylePlan.Visible = False
    oSurfaceStyle.ElevationStyle.DisplayStylePlan.Visible = False
    oSurfaceStyle.PointStyle.DisplayStylePlan.Visible = False
    oSurfaceStyle.WatershedStyle.DisplayStylePlan.Visible = False
    
    ' 3D
    oSurfaceStyle.BoundaryStyle.DisplayStyleModel.color = 254
    oSurfaceStyle.BoundaryStyle.DisplayStyleModel.Visible = True
    oSurfaceStyle.TriangleStyle.DisplayStyleModel.color = 252
    oSurfaceStyle.TriangleStyle.DisplayStyleModel.Visible = True
    
    oSurfaceStyle.ContourStyle.MajorContourDisplayStyleModel.Visible = False
    oSurfaceStyle.ContourStyle.MinorContourDisplayStyleModel.Visible = False
    oSurfaceStyle.DirectionStyle.DisplayStyleModel.Visible = False
    oSurfaceStyle.ElevationStyle.DisplayStyleModel.Visible = False
    oSurfaceStyle.PointStyle.DisplayStyleModel.Visible = False
    oSurfaceStyle.WatershedStyle.DisplayStyleModel.Visible = False
    
    
    Set CreateSurfaceStyle = oSurfaceStyle
End Function
