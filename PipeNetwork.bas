Attribute VB_Name = "PipeNetwork"
'
' MODULE_ID {PipeNetwork.bas}
' {PipeSample.dvb}
'
' Copyright {2003-2026} by Autodesk, Inc.
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
' Create a pipe network based on hard coded points.
'
Public Function CreatePipeNetwork() As Boolean
    Dim oNetworks As AeccPipeNetworks
    Dim oNetwork As AeccPipeNetwork
    Dim index As Integer
    Dim oSettings As AeccPipeSettingsRoot
    Dim oPartLists As AeccPartLists
    Dim oPartList As AeccPartList
    Dim oPartFamily As AeccPartFamily
    Dim oSizeFilters As AeccPartSizeFilters
    Dim oSizeFilter As AeccPartSizeFilter
    Dim sStructureGuid As String
    Dim sPipeGuid As String
    Dim oStructureFilter As AeccPartSizeFilter
    Dim oPipeFilter As AeccPartSizeFilter
    Dim oPipe As AeccPipe
    
    ' Get the Networks collections
    Set oNetworks = g_oPipeDocument.PipeNetworks
    Set oNetwork = Nothing
    ' looking for the "Sample" network
    On Error Resume Next
    Set oNetwork = oNetworks.Item(NETWORK_NAME)
    On Error GoTo 0
    'if the network isn't there, make it
    If (oNetwork Is Nothing) Then
        Set oNetwork = oNetworks.Add(NETWORK_NAME)
    End If

    ' Create styles for the pipes and structures.
    Dim oPipeStyle As AeccPipeStyle
    Set oPipeStyle = CreatePipeStyle(PIPE_STYLE_NAME)
    Dim oStructureStyle As AeccStructureStyle
    Set oStructureStyle = CreateStructureStyle(STRUCTURE_STYLE_NAME)
    If ((oPipeStyle Is Nothing) Or (oStructureStyle Is Nothing)) Then
        Exit Function
    End If
    
    
    ' We will go through the list of part types and
    ' select the first pipe and the first structure
    ' (which is not a null structure) we find.
    Set oSettings = g_oPipeDocument.Settings
    ' Get all the parts list in the drawing
    Set oPartLists = oSettings.PartLists
    ' Get the first part list found
    Set oPartList = oPartLists.Item(0)
    For Each oPartFamily In oPartList
        ' Look for a pipe family.
        Debug.Print "part family:"; oPartFamily.Name
        If (oPartFamily.Domain = aeccDomPipe) Then
            sPipeGuid = oPartFamily.guid
            ' Get the first size filter list from the family
            Set oPipeFilter = oPartFamily.SizeFilters.Item(0)
            Debug.Print "   size filter:"; oPipeFilter.Name
            Exit For
        End If
    Next
    
    For Each oPartFamily In oPartList
        Debug.Print "part family:"; oPartFamily.Name
        ' Search for the first structure family that is not
        ' named "Null Structure".
        If ((oPartFamily.Domain = aeccDomStructure) And (StrComp(oPartFamily.Name, "Null Structure", vbTextCompare) <> 0)) Then
            sStructureGuid = oPartFamily.guid
            ' grab the first size filter list from the family
            Set oStructureFilter = oPartFamily.SizeFilters.Item(0)
            Debug.Print "   size filter:"; oStructureFilter.Name
            Exit For
        End If
    Next
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Make a pipe network based on hard coded points.
    Dim dEndPoint(0 To 2) As Double
    Dim vStartPoint As Variant
    Dim oStructure As AeccStructure
    
    ' Make a structure at one end.
    dEndPoint(0) = 10: dEndPoint(1) = 70:  dEndPoint(2) = -20
    Set oStructure = oNetwork.Structures.Add(sStructureGuid, oStructureFilter, dEndPoint, 5.2333) ' 305 degrees
    Set oStructure.Style = oStructureStyle
    
    ' Make a pipe from that structure, and add another structure.
    vStartPoint = dEndPoint
    dEndPoint(0) = 35: dEndPoint(1) = 30:  dEndPoint(2) = -20
    Set oPipe = oNetwork.Pipes.Add(sPipeGuid, oPipeFilter, vStartPoint, dEndPoint)
    oPipe.ConnectToStructure aeccPipeStart, oStructure
    Set oStructure = oNetwork.Structures.Add(sStructureGuid, oStructureFilter, dEndPoint, 5.236)
    oPipe.ConnectToStructure aeccPipeEnd, oStructure
    oPipe.FlowDirectionMethod = aeccPipeFlowDirectionMethodStartToEnd
    Debug.Print "Length of pipe 1:"; oPipe.Length2D
    Set oPipe.Style = oPipeStyle
    Set oStructure.Style = oStructureStyle
    
    ' Make a curved pipe from the previous structure.
    vStartPoint = dEndPoint
    dEndPoint(0) = 65: dEndPoint(1) = 30:  dEndPoint(2) = -20
    Set oPipe = oNetwork.Pipes.AddCurvedPipe(sPipeGuid, oPipeFilter, vStartPoint, dEndPoint, 45, False)
    oPipe.ConnectToStructure aeccPipeStart, oStructure
    oPipe.FlowDirectionMethod = aeccPipeFlowDirectionMethodStartToEnd
    Debug.Print "Length of pipe 2:"; oPipe.Length2D
    Set oPipe.Style = oPipeStyle
    
    
    ' Make a pipe and connect it directly to the previous
    ' pipe.  Add a structure to the end.
    vStartPoint = dEndPoint
    dEndPoint(0) = 90: dEndPoint(1) = 70: dEndPoint(2) = -20
    Dim oPipeNew As AeccPipe
    Set oPipeNew = oNetwork.Pipes.Add(sPipeGuid, oPipeFilter, vStartPoint, dEndPoint)
    ' Connect the start of this pipe to the end of the previous pipe.
    ' When two pipes connect together, they are joined by an invisible
    ' structure.
    Set oStructure = oPipeNew.ConnectToPipe(aeccPipeStart, oPipe, aeccPipeEnd)
    ' Even though we are using a null structure to join the pipes,
    ' it is still drawn on the screen so we should still give it a
    ' style.
    'Set oStructureStyle = CreateInvisibleStructureStyle(STRUCTURE_INVISIBLE_STYLE_NAME)
    'If ((oPipeStyle Is Nothing) Or (oStructureStyle Is Nothing)) Then
    '    Exit Function
    'End If
    Set oStructure.Style = oStructureStyle
    ' Add the structure at the end of the last pipe.
    Set oStructure = oNetwork.Structures.Add(sStructureGuid, oStructureFilter, dEndPoint, 0.9959)
    oPipeNew.ConnectToStructure aeccPipeEnd, oStructure
    oPipeNew.FlowDirectionMethod = aeccPipeFlowDirectionMethodStartToEnd
    Debug.Print "Length of pipe 3:"; oPipeNew.Length2D
    Set oPipeNew.Style = oPipeStyle
    Set oStructure.Style = oStructureStyle

    ' Zoom to make the pipe network visible
    g_oPipeApplication.ZoomExtents
    
    CreatePipeNetwork = True
End Function ' CreatePipeNetwork


'
'
' Create a pipe network based on hard coded points,
' using an specific surface to provide depth information
' for the pipes and an alignment to provide stations for
' pipes.
'
Public Function CreatePipeNetworkFromSurface(oAlignment As AeccAlignment, oSurface As AeccSurface) As AeccPipeNetwork
    Dim oNetworks As AeccPipeNetworks
    Dim oNetwork As AeccPipeNetwork
    Dim index As Integer
    Dim oSettings As AeccPipeSettingsRoot
    Dim oPartLists As AeccPartLists
    Dim oPartList As AeccPartList
    Dim oPartFamily As AeccPartFamily
    Dim oSizeFilters As AeccPartSizeFilters
    Dim oSizeFilter As AeccPartSizeFilter
    Dim sStructureGuid As String
    Dim sPipeGuid As String
    Dim oStructureFilter As AeccPartSizeFilter
    Dim oPipeFilter As AeccPartSizeFilter
    Dim oPipe As AeccPipe
    
    ' Check for existance of base global objects.
    If (g_oPipeDocument Is Nothing) Then
        Set CreatePipeNetworkFromSurface = Nothing
        Exit Function
    End If

    ' Check that the parameters are valid
    If (oAlignment Is Nothing) Or (oAlignment Is Nothing) Then
        Set CreatePipeNetworkFromSurface = Nothing
        Exit Function
    End If

    ' Get the Networks collections
    Set oNetworks = g_oPipeDocument.PipeNetworks
    Set oNetwork = Nothing
    ' looking for the "Sample" network
    On Error Resume Next
    Set oNetwork = oNetworks.Item(NETWORK_NAME)
    On Error GoTo 0
    'if the network isn't there, make it
    If (oNetwork Is Nothing) Then
        Set oNetwork = oNetworks.Add(NETWORK_NAME)
    End If

    ' Assign the network to use the surface and alignment
    ' specified.
    oNetwork.ReferenceAlignment = oAlignment
    oNetwork.ReferenceSurface = oSurface
    
    ' Create styles for the pipes and structures.
    Dim oPipeStyle As AeccPipeStyle
    Set oPipeStyle = CreatePipeStyle(PIPE_STYLE_NAME)
    Dim oStructureStyle As AeccStructureStyle
    Set oStructureStyle = CreateStructureStyle(STRUCTURE_STYLE_NAME)
    If (oPipeStyle Is Nothing) Or (oStructureStyle Is Nothing) Then
        Set CreatePipeNetworkFromSurface = Nothing
        Exit Function
    End If
    
    
    ' Get the settings from the PIPES database object.
    Set oSettings = g_oPipeDocument.Settings
    ' Get all the parts list in the drawing.
    Set oPartLists = oSettings.PartLists
    ' Grab the first part list found.
    Set oPartList = oPartLists.Item(0)
    For Each oPartFamily In oPartList
        'Look for a pipe family.
        If (oPartFamily.Domain = aeccDomPipe) Then
            sPipeGuid = oPartFamily.guid
            ' Grab the first size filter list from the family.
            Set oPipeFilter = oPartFamily.SizeFilters.Item(0)
            Debug.Print "Using pipe: "; oPipeFilter.Name; " from family: "; oPartFamily.Name
            Exit For
        End If
    Next
    
    For Each oPartFamily In oPartList
        ' Search for the first structure family that is not
        ' named "Null Structure".
        If ((oPartFamily.Domain = aeccDomStructure) And (StrComp(oPartFamily.Name, "Null Structure", vbTextCompare) <> 0)) Then
            sStructureGuid = oPartFamily.guid
            ' Grab the first size filter list from the family.
            Set oStructureFilter = oPartFamily.SizeFilters.Item(0)
            Debug.Print "Using structure: "; oStructureFilter.Name; " from family: "; oPartFamily.Name
            Exit For
        End If
    Next
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Make a pipe network based on hard coded points.
    Dim dEndPoint(0 To 2) As Double
    Dim vStartPoint As Variant
    Dim oStructure As AeccStructure
    
    ' Make a structure at one end.
    dEndPoint(0) = 3683.4257: dEndPoint(1) = 3438.8566
    ' Place pipe network 6 units below the surface.
    dEndPoint(2) = oSurface.FindElevationAtXY(dEndPoint(0), dEndPoint(1)) - 6#
    Set oStructure = oNetwork.Structures.Add(sStructureGuid, oStructureFilter, dEndPoint, 1.7453) ' 100 degrees
    Set oStructure.Style = oStructureStyle
    
    ' Make a pipe from that structure, and add another structure.
    vStartPoint = dEndPoint
    dEndPoint(0) = 3687.7835: dEndPoint(1) = 3414.8977
    ' Place pipe network 6 units below the surface.
    dEndPoint(2) = oSurface.FindElevationAtXY(dEndPoint(0), dEndPoint(1)) - 6#
    Set oPipe = oNetwork.Pipes.Add(sPipeGuid, oPipeFilter, vStartPoint, dEndPoint)
    oPipe.ConnectToStructure aeccPipeStart, oStructure
    Set oStructure = oNetwork.Structures.Add(sStructureGuid, oStructureFilter, dEndPoint, 0#)
    oPipe.ConnectToStructure aeccPipeEnd, oStructure
    oPipe.FlowDirectionMethod = aeccPipeFlowDirectionMethodBySlope
    Set oPipe.Surface = g_oPipeDocument.Surfaces.Item(SURFACE_NAME)
    Debug.Print "Length pipe 1:"; oPipe.Length2D, oPipe.Surface.Name
    Set oPipe.Style = oPipeStyle
    Set oStructure.Style = oStructureStyle
    
    ' Make a curved pipe from the previous structure.
    vStartPoint = dEndPoint
    dEndPoint(0) = 3717.7157: dEndPoint(1) = 3390.3219
    ' Place pipe network 6 units below the surface.
    dEndPoint(2) = oSurface.FindElevationAtXY(dEndPoint(0), dEndPoint(1)) - 6#
    Set oPipe = oNetwork.Pipes.AddCurvedPipe(sPipeGuid, oPipeFilter, vStartPoint, dEndPoint, 50, False)
    oPipe.ConnectToStructure aeccPipeStart, oStructure
    oPipe.FlowDirectionMethod = aeccPipeFlowDirectionMethodBySlope
    Set oPipe.Surface = g_oPipeDocument.Surfaces.Item(SURFACE_NAME)
    Debug.Print "Length pipe 2:"; oPipe.Length2D, oPipe.Surface.Name
    Set oPipe.Style = oPipeStyle
    
    
    ' Make a pipe and connect it directly to the previous pipe.  Add
    ' a structure to the end.
    vStartPoint = dEndPoint
    dEndPoint(0) = 3795.8047: dEndPoint(1) = 3390.3975
    ' Place pipe network 6 units below the surface.
    dEndPoint(2) = oSurface.FindElevationAtXY(dEndPoint(0), dEndPoint(1)) - 6#
    Dim oPipeNew As AeccPipe
    Set oPipeNew = oNetwork.Pipes.Add(sPipeGuid, oPipeFilter, vStartPoint, dEndPoint)
    Set oPipeNew.Surface = g_oPipeDocument.Surfaces.Item(SURFACE_NAME)
    ' Connect the start of this pipe to the end of the previous pipe.
    ' When two pipes connect together, they are joined by an invisible
    ' structure.
    Set oStructure = oPipeNew.ConnectToPipe(aeccPipeStart, oPipe, aeccPipeEnd)
    ' A null structure is still drawn on the screen.  We will
    ' create a style that makes it invisible.
    Set oStructure.Style = oStructureStyle
    ' Add the structure at the end of the last pipe.
    Set oStructure = oNetwork.Structures.Add(sStructureGuid, oStructureFilter, dEndPoint, 0#)
    oPipeNew.ConnectToStructure aeccPipeEnd, oStructure
    oPipeNew.FlowDirectionMethod = aeccPipeFlowDirectionMethodBySlope
    Debug.Print "Length pipe 3:"; oPipeNew.Length2D, oPipeNew.Surface.Name
    Set oPipeNew.Style = oPipeStyle
    Set oStructure.Style = oStructureStyle

    ' Add an unconnected pipe that crosses the previous pipe
    ' for use with intersection testing.
    vStartPoint(0) = 3775.2: vStartPoint(1) = 3386.9: vStartPoint(2) = oPipeNew.EndPoint.Z
    dEndPoint(0) = 3775.2: dEndPoint(1) = 3393.7: dEndPoint(2) = oPipeNew.EndPoint.Z
    Set oPipeNew = oNetwork.Pipes.Add(sPipeGuid, oPipeFilter, vStartPoint, dEndPoint)
    
    ' Zoom to make the pipe network visible
    g_oPipeApplication.ZoomExtents
    
    Set CreatePipeNetworkFromSurface = oNetwork
End Function ' CreatePipeNetworkFromSurface


'
' Perform an interference check for the given pipe network
' to see if it crosses itself anywhere.
'
Public Function PerformInterferenceCheck(oPipeNetwork As AeccPipeNetwork) As Boolean
    ' Check for global objects.
    If (g_oPipeDocument Is Nothing) Then
        PerformInterferenceCheck = False
        Exit Function
    End If

    ' Get the collection of all interference checks.
    Dim oInterferenceChecks As AeccInterferenceChecks
    Set oInterferenceChecks = g_oPipeDocument.InterferenceChecks
    
    ' Set up the creation data structure for making an
    ' interference check.
    Dim oCreationData As AeccInterferenceCheckCreationData
    Set oCreationData = oInterferenceChecks.GetDefaultCreationData()
    ' oCreationData.InterferenceStyle = style
    oCreationData.Criteria.ApplyProximity = True
    oCreationData.Criteria.CriteriaDistance = 3.5
    oCreationData.Criteria.UseDistanceOrScaleFactor = aeccDistance
    oCreationData.Name = INTERFERENCE_CHECK_NAME
    oCreationData.LayerName = g_oPipeDocument.Layers.Item(0).Name
    ' We will see if pipes in this network cross themselves.
    Set oCreationData.SourceNetwork = oPipeNetwork
    Set oCreationData.TargetNetwork = oPipeNetwork

    ' Create a new check of the pipe network.
    Dim oInterferenceCheck As AeccInterferenceCheck
    Set oInterferenceCheck = oInterferenceChecks.Create(oCreationData)
    
    ' Go through the check, looking at each of the
    ' interferences.  Give a report of the location and
    ' the extent of each interference.
    Dim oInterference As AeccInterference
    For Each oInterference In oInterferenceCheck
        Set oInterference.Style = CreateInterferenceStyle(INTERFERENCE_STYLE_NAME)
        ' Display the 2D x,y location of the interference.
        Dim vLocation As Variant
        Dim sLocation As String
        Dim vExtent As Variant
        vLocation = oInterference.Location
        sLocation = vLocation(0) & ", " & vLocation(1)
        MsgBox "There is an interference at:" & sLocation
        
        ' Display the greatest and least corners of the 3D
        ' rectangle containing the interference.
        vExtent = oInterference.GetExtents()
        Debug.Print
        Debug.Print "The interference takes place between:"
        sLocation = vExtent(0) & ", "
        sLocation = sLocation & vExtent(0 + 1) & ", "
        sLocation = sLocation & vExtent(0 + 2)
        Debug.Print "  "; sLocation; "   and:"
        sLocation = vExtent(3 + 0) & ", "
        sLocation = sLocation & vExtent(3 + 1) & ", "
        sLocation = sLocation & vExtent(3 + 2)
        Debug.Print "  "; sLocation
    Next
    
    If (oInterferenceCheck.Count = 0) Then
        MsgBox "There are no interferences in the network."
    End If
    
    ' Success
    PerformInterferenceCheck = True
End Function


'
' Create a pipe network based on user input through the
' mouse actions - each click places a new structure in the
' document at that location and connects a pipe from any
' previos structure to the latest structure.
'
' Not called by any of the demonstration subroutines exposed
' to the VBARUN dialog box.
'
Public Function CreatePipeNetworkFromMouse() As Boolean
    Dim oNetworks As AeccPipeNetworks
    Dim oNetwork As AeccPipeNetwork
    Dim index As Integer
    Dim oSettings As AeccPipeSettingsRoot
    Dim oPartLists As AeccPartLists
    Dim oPartList As AeccPartList
    Dim oPartFamily As AeccPartFamily
    Dim oSizeFilters As AeccPartSizeFilters
    Dim oSizeFilter As AeccPartSizeFilter
    Dim sStructureGuid As String
    Dim sPipeGuid As String
    Dim oStructureFilter As AeccPartSizeFilter
    Dim oPipeFilter As AeccPartSizeFilter
    Dim vStartPoint As Variant
    Dim vEndPoint  As Variant
    Dim sPrompt As String
    Dim oPipe As AeccPipe
    Dim bContinue As Boolean
    
    ' Check for global objects.
    If (g_oPipeDocument Is Nothing) Then
        CreatePipeNetworkFromMouse = False
        Exit Function
    End If

    ' Get the Networks collections
    Set oNetworks = g_oPipeDocument.PipeNetworks
    Set oNetwork = Nothing
    ' Looking for the "Sample" network.
    On Error Resume Next
    Set oNetwork = oNetworks.Item("Sample")
    On Error GoTo 0
    ' If the network isn't there, make it.
    If (oNetwork Is Nothing) Then
        Set oNetwork = oNetworks.Add("Sample")
    End If

    ' Get the settings from the PIPES database object.
    Set oSettings = g_oPipeDocument.Settings
    ' Get all the parts list in the drawing.
    Set oPartLists = oSettings.PartLists
    ' Grab the first part list found.
    Set oPartList = oPartLists.Item(0)
    For Each oPartFamily In oPartList
        ' Look for a pipe family.
        Debug.Print "part family:"; oPartFamily.Name
        If (oPartFamily.Domain = aeccDomPipe) Then
            sPipeGuid = oPartFamily.guid
            ' Grab the first size filter list from the family.
            Set oPipeFilter = oPartFamily.SizeFilters.Item(0)
            Debug.Print "   size filter:"; oPipeFilter.Name
            Exit For
        End If
    Next
    
    For Each oPartFamily In oPartList
        Debug.Print "part family:"; oPartFamily.Name
        ' Search for the first structure family that is not
        ' named "Null Structure".
        If ((oPartFamily.Domain = aeccDomStructure) And (StrComp(oPartFamily.Name, "Null Structure", vbTextCompare) <> 0)) Then
            sStructureGuid = oPartFamily.guid
            ' Grab the first size filter list from the family.
            Set oStructureFilter = oPartFamily.SizeFilters.Item(0)
            Debug.Print "   size filter:"; oStructureFilter.Name
            Exit For
        End If
    Next
    ' Hard code the z for the structures and pipes.
    Const STRUCTURE_Z = 0
    Const PIPE_Z = -2
    sPrompt = vbCrLf & "Enter a point or input any key to stop:"
    On Error Resume Next
    ' Get the first point from the user.
    vStartPoint = ThisDrawing.Utility.GetPoint(, sPrompt)
    vStartPoint(2) = STRUCTURE_Z
    Dim oStructure As AeccStructure
    ' Add a new structure to the network.
    Set oStructure = oNetwork.Structures.Add(sStructureGuid, oStructureFilter, vStartPoint, 1)
    bContinue = True
    Do While (True)
        ' Use the point entered above as the first start point.
        
        ' Get points until the user stops entering them.
        On Error Resume Next
        vEndPoint = ThisDrawing.Utility.GetPoint(vStartPoint, sPrompt)
        On Error GoTo 0
        If Err Then
            Err.Clear
            Exit Do
        End If
        
        vStartPoint(2) = PIPE_Z
        vEndPoint(2) = PIPE_Z
        ' while the user keeps selecting points, add a pipe and a structure
        Set oPipe = oNetwork.Pipes.Add(sPipeGuid, oPipeFilter, vStartPoint, vEndPoint)
        oPipe.ConnectToStructure aeccPipeStart, oStructure
        vStartPoint = vEndPoint
        vStartPoint(2) = STRUCTURE_Z
        Set oStructure = oNetwork.Structures.Add(sStructureGuid, oStructureFilter, vStartPoint, 1)
        'connect the pipe to the structure
        oPipe.ConnectToStructure aeccPipeEnd, oStructure
    Loop
    
    CreatePipeNetworkFromMouse = True
End Function ' CreatePipeNetworkFromMouse

