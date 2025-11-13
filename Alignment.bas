Attribute VB_Name = "Alignment"
'
' MODULE_ID {Alignment.bas}
' {PipeSample.dvb}
'
' Copyright {2016} by Autodesk, Inc.
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
Option Explicit
'
'
Public Function CreateAlignmentByLWPolyline() As AeccAlignment
    ' Check for existance of global objects.
    If (g_oPipeDocument Is Nothing) Then
        Set CreateAlignmentByLWPolyline = Nothing
        Exit Function
    End If

    ' Get the site that the alignment will be part of,
    ' and the collection of the site's alignments.
    Dim oSite As AeccSite
    Dim oAlignments As AeccAlignments
    Dim oAlignment As AeccAlignment
    Set oSite = GetSite()
    If (oSite Is Nothing) Then
        CreateAlignmentByLWPolyline = False
        Exit Function
    End If
    Set oAlignments = oSite.Alignments
    
    ' Create the polyline that the alignment will be based
    ' on.
    Dim objId As LongPtr
    objId = CreateLWPolyline
    If (objId = 0) Then
        Set CreateAlignmentByLWPolyline = Nothing
        Exit Function
    End If

    ' Set the style for the alignment.
    Dim oAlignmentStyle As AeccAlignmentStyle
    If (CreateAlignmentStyle(ALIGNMENT_STYLE_NAME) = False) Then
        Set CreateAlignmentByLWPolyline = Nothing
        Exit Function
    End If
    Set oAlignmentStyle = g_oPipeDocument.AlignmentStyles.Item(ALIGNMENT_STYLE_NAME)

    ' Set the label style for the alignment.
    Dim oLabelStyleSet As AeccAlignmentLabelStyleSet
    If (CreateAlignmentLabelStyleSet(LABEL_STYLE_SET_NAME) = False) Then
        Set CreateAlignmentByLWPolyline = Nothing
        Exit Function
    End If
    Set oLabelStyleSet = g_oPipeDocument.AlignmentLabelStyleSets.Item(LABEL_STYLE_SET_NAME)
    
    ' Create a simple alignment.
    On Error Resume Next
    Set oAlignment = oAlignments.AddFromPolyline(ALIGNMENT_NAME, "0", objId, oAlignmentStyle, oLabelStyleSet, True, True)
    On Error GoTo 0
    If (oAlignment Is Nothing) Then
        Debug.Print "Error alignment: " & Err.Description & " - " & Err.Number
        g_oPipeApplication.ZoomExtents
        Set CreateAlignmentByLWPolyline = Nothing
        Exit Function
    End If

    g_oPipeApplication.Update
    Set CreateAlignmentByLWPolyline = oAlignment
End Function



' Create a polyline using three 2-D points and
' add the polyline to the drawing.
'
Public Function CreateLWPolyline() As LongPtr
    Dim oPoly As AcadLWPolyline
    Dim points(0 To 7) As Double

    points(0) = 3683.4257: points(1) = 3438.8566
    points(2) = 3687.7835: points(3) = 3414.8977
    points(4) = 3717.7157: points(5) = 3390.3219
    points(6) = 3795.8047: points(7) = 3390.3975
    Set oPoly = g_oPipeDocument.Database.ModelSpace.AddLightWeightPolyline(points)
    If (oPoly.ObjectID = 0) Then
        Debug.Print "Error creating a polyline."
        CreateLWPolyline = 0
        Exit Function
    End If
    
    CreateLWPolyline = oPoly.ObjectID
    
End Function

