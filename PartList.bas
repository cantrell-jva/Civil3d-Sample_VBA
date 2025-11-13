Attribute VB_Name = "PartList"
'
' MODULE_ID {PartList.bas}
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
' Print out the hierarchy of part types available in this
' document, seperated into pipe and structure domains.
'
Public Function ExportPartsListToWord() As Boolean
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
    Dim oPartDataField As AeccPartDataField
    Dim oWordApp As Word.Application
    Dim oWordDoc As Word.Document
    Dim oPara As Word.Paragraph

    'Start Word and open the document template.
    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        On Error Resume Next
        Set oWordApp = CreateObject("Word.Application")
        On Error GoTo 0
        If Err.Number <> 0 Then
            MsgBox "Could not start Microsoft Word."
            Err.Clear
            ExportPartsListToWord = False
            Exit Function
        End If
    End If
    oWordApp.Visible = True
    On Error Resume Next
    Set oWordDoc = oWordApp.Documents.Add
    On Error GoTo 0
    If oWordDoc Is Nothing Then
        MsgBox "Could not create a new Word document."
        ExportPartsListToWord = False
        Exit Function
    End If
   

    ' Get the settings from the Pipes document object, casting
    ' it to type AeccPipeSettingsRoot.
    Set oSettings = g_oPipeDocument.Settings
    
    ' Try retrieving and setting some pipe ambient settings.
    With oSettings.PipeNetworkSettings.RulesSettings
        Debug.Print "Using pipe rules:"; .PipeDefaultRules.Value
    End With
    With oSettings.PipeSettings.AmbientSettings
        .AngleSettings.Unit = aeccAngleUnitRadian
        .CoordinateSettings.Unit = aeccCoordinateUnitFoot
        .DistanceSettings.Unit = aeccCoordinateUnitFoot
    End With
    
    ' Get a reference to all the parts lists in the drawing.
    Set oPartLists = oSettings.PartLists
    Set oPara = oWordDoc.Content.Paragraphs.Add
    oPara.Range.Text = "Number of Part lists: " & oPartLists.Count
    oPara.Range.Font.Bold = True
    oPara.Range.Font.Size = 18
    oPara.Range.Font.Name = "Courier New"
    oPara.Format.SpaceAfter = 24    '24 pt spacing after paragraph.
    oPara.Range.InsertParagraphAfter
    
    ' Get the first part list, whatever it is.
    For Each oPartList In oPartLists
        Set oPara = oWordDoc.Content.Paragraphs.Add(oWordDoc.Bookmarks("\endofdoc").Range)
        oPara.Range.Text = "Part List - " & oPartList.Name
        oPara.Range.Font.Bold = False
        oPara.Range.Font.Size = 14
        oPara.Format.SpaceAfter = 10
        oPara.Range.InsertParagraphAfter
        
        ' From the part list, looking at only those part families
        ' that are pipes, print all the individual parts.
        oPara.Range.Text = "Pipes"
        oPara.Range.Font.Underline = wdUnderlineSingle
        oPara.Range.Font.Size = 10
        oPara.Format.SpaceAfter = 2
        oPara.Range.InsertParagraphAfter
        oPara.Range.Font.Underline = wdUnderlineNone
        
        For Each oPartFamily In oPartList
            ' Look for only pipe families.
            If (oPartFamily.Domain = aeccDomPipe) Then
                sPipeGuid = oPartFamily.guid
                oPara.Range.Text = vbTab & "Family: " & oPartFamily.Name
                oPara.Range.InsertParagraphAfter
                oPara.Range.Text = vbTab & "GUID: " & sPipeGuid
                oPara.Range.InsertParagraphAfter
               ' Go through each part in this family.
                For Each oPipeFilter In oPartFamily.SizeFilters
                    oPara.Range.Text = vbTab & vbTab & "Filter: " & oPipeFilter.Name
                    oPara.Range.InsertParagraphAfter
                    
                    ' Print out all data fields for this pipe size.
                    oPara.Range.Text = vbTab & vbTab & "All data fields for this size:"
                    oPara.Range.InsertParagraphAfter
                    For Each oPartDataField In oPipeFilter.PartDataRecord
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Context name:  " & oPartDataField.ContextString
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Description:   " & oPartDataField.Description
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Internal name: " & oPartDataField.Name
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Value:         " & oPartDataField.Tag
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Type of value: " & oPartDataField.Type
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "------"
                        oPara.Range.InsertParagraphAfter
                    Next
                Next
            End If
        Next

        ' Perform the same action to get a structure filter from
        ' the first structure part family from the first part
        ' family we access.
        Debug.Print: Debug.Print
        oPara.Range.Text = "Structures"
        oPara.Range.Font.Underline = wdUnderlineSingle
        oPara.Range.Font.Size = 10
        oPara.Format.SpaceAfter = 2
        oPara.Range.InsertParagraphAfter
        oPara.Range.Font.Underline = wdUnderlineNone
        
        For Each oPartFamily In oPartList
            ' Look for only structure families.
            If (oPartFamily.Domain = aeccDomStructure) Then
                sStructureGuid = oPartFamily.guid
                oPara.Range.Text = vbTab & "Family: " & oPartFamily.Name
                oPara.Range.InsertParagraphAfter
                oPara.Range.Text = vbTab & "GUID: " & sPipeGuid
                oPara.Range.InsertParagraphAfter
                ' Go through each part in this family.
                For Each oStructureFilter In oPartFamily.SizeFilters
                    oPara.Range.Text = vbTab & vbTab & "Filter: " & oStructureFilter.Name
                    oPara.Range.InsertParagraphAfter
                
                    ' Print out all data fields for this pipe size.
                    oPara.Range.Text = vbTab & vbTab & "All data fields for this size:"
                    oPara.Range.InsertParagraphAfter
                    For Each oPartDataField In oStructureFilter.PartDataRecord
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Context name:  " & oPartDataField.ContextString
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Description:   " & oPartDataField.Description
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Internal name: " & oPartDataField.Name
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Value:         " & oPartDataField.Tag
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "Type of value: " & oPartDataField.Type
                        oPara.Range.InsertParagraphAfter
                        oPara.Range.Text = vbTab & vbTab & vbTab & "------"
                        oPara.Range.InsertParagraphAfter
                    Next
                Next
            End If
        Next
    Next
    Debug.Print
    Debug.Print
    
    ExportPartsListToWord = True
End Function ' ExportPartsListToWord


'
' Simplified version of above function which just uses
' Debug.Print.  However, this tends to output more text
' than the Immediate window of the Visual Basic IDE can
' handle.
'
' Print out the hierarchy of part types available in this
' document, seperated into pipe and structure domains.
'
Public Function PrintPartsList() As Boolean
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
    Dim oPartDataField As AeccPartDataField

    ' Get the settings from the Pipes document object, casting
    ' it to type AeccPipeSettingsRoot.
    Set oSettings = g_oPipeDocument.Settings
    
    ' Try retrieving and setting some pipe ambient settings.
    With oSettings.PipeNetworkSettings.RulesSettings
        Debug.Print "Using pipe rules:"; .PipeDefaultRules.Value
    End With
    With oSettings.PipeSettings.AmbientSettings
        .AngleSettings.Unit = aeccAngleUnitRadian
        .CoordinateSettings.Unit = aeccCoordinateUnitFoot
        .DistanceSettings.Unit = aeccCoordinateUnitFoot
    End With
    
    ' Get a reference to all the parts lists in the drawing.
    Set oPartLists = oSettings.PartLists
    Debug.Print "#Part lists: "; oPartLists.Count
    
    ' Get the first part list, whatever it is.
    For Each oPartList In oPartLists
        Debug.Print: Debug.Print
        Debug.Print "PART LIST - "; oPartList.Name
        Debug.Print "------------------------------------------------"
        
        ' From the part list, looking at only those part families
        ' that are pipes, print all the individual parts.
        Debug.Print "  Pipes"
        Debug.Print "  ====="
        For Each oPartFamily In oPartList
            ' Look for only pipe families.
            If (oPartFamily.Domain = aeccDomPipe) Then
                sPipeGuid = oPartFamily.guid
                Debug.Print "  Family: "; oPartFamily.Name, sPipeGuid
                ' Go through each part in this family.
                For Each oPipeFilter In oPartFamily.SizeFilters
                    Debug.Print "    Filter: "; oPipeFilter.Name
                    
                    ' Print out all data fields for this pipe size.
                    Debug.Print "    All data fields for this size:"
                    Debug.Print "    ======"
                    For Each oPartDataField In oPipeFilter.PartDataRecord
                        Debug.Print "      Context name:  "; oPartDataField.ContextString
                        Debug.Print "      Description:   "; oPartDataField.Description
                        Debug.Print "      Internal name: "; oPartDataField.Name
                        Debug.Print "      Value:         "; oPartDataField.Tag
                        Debug.Print "      Type of value: "; oPartDataField.Type
                        Debug.Print "      ------"
                    Next
                Next
            End If
        Next

        ' Perform the same action to get a structure filter from
        ' the first structure part family from the first part
        ' family we access.
        Debug.Print: Debug.Print
        Debug.Print "  Structures"
        Debug.Print "  =========="
        For Each oPartFamily In oPartList
            ' Look for only structure families.
            If (oPartFamily.Domain = aeccDomStructure) Then
                sStructureGuid = oPartFamily.guid
                Debug.Print "  Family: "; oPartFamily.Name, sStructureGuid
                ' Go through each part in this family.
                For Each oStructureFilter In oPartFamily.SizeFilters
                    Debug.Print "    Filter: "; oStructureFilter.Name
                
                    ' Print out all data fields for this pipe size.
                    Debug.Print "    All data fields for this size:"
                    Debug.Print "    ======"
                    For Each oPartDataField In oStructureFilter.PartDataRecord
                        Debug.Print "      Context name:  "; oPartDataField.ContextString
                        Debug.Print "      Description:   "; oPartDataField.Description
                        Debug.Print "      Internal name: "; oPartDataField.Name
                        Debug.Print "      Value:         "; oPartDataField.Tag
                        Debug.Print "      Type of value: "; oPartDataField.Type
                        Debug.Print "      ------"
                    Next
                Next
            End If
        Next
    Next
    Debug.Print
    Debug.Print
    
    PrintPartsList = True
End Function ' PrintPartsList

