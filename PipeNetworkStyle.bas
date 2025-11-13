Attribute VB_Name = "PipeNetworkStyle"
'
' MODULE_ID {PipeNetworkStyle.bas}
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
Option Explicit


'
' If a pipe style with the name sStyleName does not exist, create it.
' If it does exist, edit it using new settings.  Return a reference
' to the new style.
'
Public Function CreatePipeStyle(sStyleName As String) As AeccPipeStyle
    Dim oPipeStyle As AeccPipeStyle
    On Error Resume Next
    Set oPipeStyle = g_oPipeDocument.PipeStyles.Add(sStyleName)
    On Error GoTo 0
    If (oPipeStyle Is Nothing) Then
        On Error Resume Next
        Set oPipeStyle = g_oPipeDocument.PipeStyles.Item(sStyleName)
        On Error GoTo 0
        If (oPipeStyle Is Nothing) Then
            MsgBox "Could not create or use a pipe style with the name:" & sStyleName
            Set CreatePipeStyle = Nothing
            Exit Function
        End If
    End If
    
    ' Set the display size of the pipes in plan view.  We will
    ' use absolute drawing units for the inside, outside, and
    ' ends of each pipe.
    oPipeStyle.PlanOption.InnerDiameter = 2.1
    oPipeStyle.PlanOption.OuterDiameter = 2.4
    ' Indicate that we will use our own measurements for the inside
    ' and outside of the pipe, and not base drawing on the actual
    ' type of pipe.
    oPipeStyle.PlanOption.WallSizeType = aeccUserDefinedWallSize
    ' Inidcate what kind of custom sizing to use.
    oPipeStyle.PlanOption.WallSizeOptions = aeccPipeUseAbsoluteUnits
    oPipeStyle.PlanOption.EndLineSize = 2.1
    ' Indicate that we will use our own measurements for the end
    ' line of the pipe, and not base drawing on the actual type
    ' of pipe.
    oPipeStyle.PlanOption.EndSizeType = aeccUserDefinedEndSize
    ' Inidcate what kind of custom sizing to use.
    oPipeStyle.PlanOption.EndSizeOptions = aeccPipeUseAbsoluteUnits

    ' Modify the colors of pipes using this style, as shown in
    ' plan view.
    oPipeStyle.DisplayStylePlan(aeccDispCompPipeOutsideWalls).color = 40 ' orange
    oPipeStyle.DisplayStylePlan(aeccDispCompPipeInsideWalls).color = 200 ' violet
    oPipeStyle.DisplayStylePlan(aeccDispCompPipeEndLine).color = 200 ' violet
    
    ' Set the hatch style for pipes using this style, as shown
    ' in plan view.
    oPipeStyle.HatchStylePlan(aeccHatchPipe).Pattern = "DOTS"
    oPipeStyle.HatchStylePlan(aeccHatchPipe).HatchType = aeccHatchPreDefined
    oPipeStyle.HatchStylePlan(aeccHatchPipe).UseAngleOfObject = False
    oPipeStyle.HatchStylePlan(aeccHatchPipe).ScaleFactor = 9#
    oPipeStyle.PlanOption.HatchOptions = aeccHatchToInnerWalls
    oPipeStyle.DisplayStylePlan(aeccDispCompPipeHatch).color = 120 ' turquose
    oPipeStyle.DisplayStylePlan(aeccDispCompPipeHatch).Visible = True
    
    Set CreatePipeStyle = oPipeStyle
End Function ' CreatePipeStyle


'
' If a structure style with the name sStyleName does not exist, create
' it. If it does exist, edit it using new settings.  Return a reference
' to the new style.
'
Public Function CreateStructureStyle(sStyleName As String) As AeccStructureStyle
    Dim oStructureStyle As AeccStructureStyle
    On Error Resume Next
    Set oStructureStyle = g_oPipeDocument.StructureStyles.Add(sStyleName)
    On Error GoTo 0
    If (oStructureStyle Is Nothing) Then
        On Error Resume Next
        Set oStructureStyle = g_oPipeDocument.StructureStyles.Item(sStyleName)
        On Error GoTo 0
        If (oStructureStyle Is Nothing) Then
            MsgBox "Could not create or use a structure style with the name:" & sStyleName
            Set CreateStructureStyle = Nothing
            Exit Function
        End If
    End If


    oStructureStyle.DisplayStylePlan(aeccDispCompStructure).color = 30 ' orange
    oStructureStyle.DisplayStylePlan(aeccDispCompStructure).Visible = True
    
    oStructureStyle.PlanOption.MaskConnectedObjects = False
    oStructureStyle.PlanOption.SizeType = aeccStructureUseDrawingScale
    oStructureStyle.PlanOption.Size = 3.5

    oStructureStyle.DisplayStyleSection(aeccDispCompStructure).Visible = False
    oStructureStyle.DisplayStyleSection(aeccDispCompStructureHatch).Visible = False
    oStructureStyle.DisplayStylePlan(aeccDispCompStructureHatch).Visible = False
    oStructureStyle.DisplayStyleProfile(aeccDispCompStructure).Visible = False
    oStructureStyle.DisplayStyleProfile(aeccDispCompStructureHatch).Visible = False
    oStructureStyle.DisplayStyleModel(aeccDispCompStructureSolid).Visible = False
    ' oStructureStyle.DisplayStylePlan(aeccDispStructureCompLast).Visible = False
    
    Set CreateStructureStyle = oStructureStyle
End Function ' CreateStructureStyle


'
' Create a style for the display of interferences found
' during an interference check of the pipe networks.
'
Public Function CreateInterferenceStyle(sStyleName As String) As AeccInterferenceStyle
    Dim oInterferenceStyle As AeccInterferenceStyle
    On Error Resume Next
    Set oInterferenceStyle = g_oPipeDocument.InterferenceStyles.Add(sStyleName)
    On Error GoTo 0
    If (oInterferenceStyle Is Nothing) Then
        On Error Resume Next
        Set oInterferenceStyle = g_oPipeDocument.InterferenceStyles.Item(sStyleName)
        On Error GoTo 0
        If (oInterferenceStyle Is Nothing) Then
            MsgBox "Could not create or use an interference style with the name:" & sStyleName
            Set CreateInterferenceStyle = Nothing
            Exit Function
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Below are examples of three different display types
    ' for showing intersections.  Comment out the two blocks
    ' you do not wish to use and uncomment the one you would
    ' like to try.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Draw a 3D red sphere which just circumscribes the
    ' region of intersection.
    oInterferenceStyle.ModelOptions = aeccSphereInterference
    oInterferenceStyle.ModelSolidDisplayStyle2D.color = 10 ' red
    oInterferenceStyle.ModelSolidDisplayStyle2D.Visible = True
    oInterferenceStyle.InterferenceSizeType = aeccSolidExtents
    oInterferenceStyle.PlanSymbolDisplayStyle2D.Visible = False
    oInterferenceStyle.MarkerStyle.MarkerDisplayStyle2d.Visible = False
    
    ' Draw a 3D blue model of the region of intersection.
    'oInterferenceStyle.ModelOptions = aeccTrueSolidInterference
    'oInterferenceStyle.ModelSolidDisplayStyle2D.color = 140 ' blue
    'oInterferenceStyle.ModelSolidDisplayStyle2D.Visible = True
    'oInterferenceStyle.PlanSymbolDisplayStyle2D.Visible = False
    'oInterferenceStyle.MarkerStyle.MarkerDisplayStyle2d.Visible = False
    
    ' Draw a symbol of a violet X with a circle with a
    ' specified drawing size at the point of intersection.
    'oInterferenceStyle.PlanSymbolDisplayStyle2D.Visible = True
    'With oInterferenceStyle.MarkerStyle
    '   .MarkerType = aeccUseCustomMarker
    '   .CustomMarkerStyle = aeccCustomMarkerX
    '   .CustomMarkerSuperimposeStyle = aeccCustomMarkerSuperimposeCircle
    '   .MarkerDisplayStyle2d.color = 200 ' violet
    '   .MarkerDisplayStyle2d.Visible = True
    '   .MarkerSizeType = aeccAbsoluteUnits
    '   .MarkerSize = 5#  ' drawing units
    'End With
    'oInterferenceStyle.ModelSolidDisplayStyle2D.Visible = False
    
    
    Set CreateInterferenceStyle = oInterferenceStyle
End Function ' CreateInterferenceStyle

