Attribute VB_Name = "AlignmentStyle"
'
' MODULE_ID {AlignmentStyle.bas}
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
' Create a plain style for the alignment.
'
Function CreateAlignmentStyle(sStyleName As String) As Boolean
    Dim oAlignmentStyles As AeccAlignmentStyles
    Dim oAlignmentStyle As AeccAlignmentStyle
    
    ' Get the drawing's collection of alignment styles.
    On Error Resume Next
    Set oAlignmentStyles = g_oPipeDocument.AlignmentStyles
    On Error GoTo 0
    If (oAlignmentStyles Is Nothing) Then
        Debug.Print "Error accessing collection of alignment styles: " & vbNewLine & Err.Description
        CreateAlignmentStyle = False
        Exit Function
    End If
    
    ' Get a particular style from the collection of styles.
    On Error Resume Next
    Set oAlignmentStyle = oAlignmentStyles.Item(sStyleName)
    On Error GoTo 0
    ' If the style does not exist, make a new one with that
    ' name.
    If (oAlignmentStyle Is Nothing) Then
        On Error Resume Next
        Set oAlignmentStyle = oAlignmentStyles.Add(sStyleName)
        On Error GoTo 0
        If (oAlignmentStyles Is Nothing) Then
            Debug.Print "Error trying to make a new alignment style: " & vbNewLine & Err.Description
            CreateAlignmentStyle = False
            Exit Function
        End If
    End If

    ' Hide direction arrows and line extensions to make an uncluttered
    ' alignment display.
    oAlignmentStyle.ArrowDisplayStyle2d.Visible = False
    oAlignmentStyle.LineExtensionsDisplayStyle2d.Visible = False
    
    CreateAlignmentStyle = True
End Function


'
' Cratea a default label style for the alignment.
'
Function CreateAlignmentLabelStyleSet(sStyleSetName As String) As Boolean
    Dim oLabelStyleSet As AeccAlignmentLabelStyleSet
    
    On Error Resume Next
    Set oLabelStyleSet = g_oPipeDocument.AlignmentLabelStyleSets.Item(sStyleSetName)
    On Error GoTo 0
    If (oLabelStyleSet Is Nothing) Then
        On Error Resume Next
        Set oLabelStyleSet = g_oPipeDocument.AlignmentLabelStyleSets.Add(sStyleSetName)
        On Error GoTo 0
        If (oLabelStyleSet Is Nothing) Then
            CreateAlignmentLabelStyleSet = False
            Exit Function
        End If
    End If

    CreateAlignmentLabelStyleSet = True
End Function

