Attribute VB_Name = "Subroutines"
'
' MODULE_ID {Subroutines.bas}
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
' This sample demonstrates the creation of a pipe network
' based on an alignment over a surface.  It then performs
' an interference check, and demonstrates pipe properties.
'
' Note: Path to XML file containing surface data is hard coded,
' so it may need to be modified.
'
Sub PerformFullPipeDemonstration()
    Dim oAlignment As AeccAlignment
    Dim oSurface As AeccSurface
    Dim oPipeNetwork As AeccPipeNetwork
    
    ' Always get the objects again since MDI is supported.
    If (GetBasePipeObjects = False) Then
        MsgBox "Error accessing base civil objects."
        Exit Sub
    End If

    Set oSurface = CreateSurfaceByImportTXT()
    If (oSurface Is Nothing) Then
        MsgBox "Error attempting to create a TIN surface."
        Exit Sub
    End If
    
    If (CreateSite = False) Then
        MsgBox "Error creating a site."
        Exit Sub
    End If

    Set oAlignment = CreateAlignmentByLWPolyline()
    If (oAlignment Is Nothing) Then
        MsgBox "Error creating an alignment."
        Exit Sub
    End If

    If (ExportPartsListToWord() = False) Then
        MsgBox "Error printing part catalog."
        Exit Sub
    End If

    Set oPipeNetwork = CreatePipeNetworkFromSurface(oAlignment, oSurface)
    If (oPipeNetwork Is Nothing) Then
        MsgBox "Error creating a pipe network."
        Exit Sub
    End If
    
    If (PerformInterferenceCheck(oPipeNetwork) = False) Then
        MsgBox "Error performing an interference check."
        Exit Sub
    End If
    
    g_oPipeApplication.ZoomExtents
End Sub


'
' This performs a simpler sample - simply listing
' all pipe and structure parts and then creating a
' simple pipe network.
'
Sub PerformSimplePipeDemonstration()
    ' Always get the objects again since MDI is supported.
    If (GetBasePipeObjects = False) Then
        MsgBox "Error accessing base civil objects."
        Exit Sub
    End If

    If (ExportPartsListToWord() = False) Then
        MsgBox "Error printing part catalog."
        Exit Sub
    End If
    
    If (CreatePipeNetwork() = False) Then
        MsgBox "Error creating a pipe network."
        Exit Sub
    End If
End Sub


