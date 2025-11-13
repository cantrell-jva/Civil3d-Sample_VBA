Attribute VB_Name = "ExcelFunctions"
'
' MODULE_ID {ExcelFunctions.bas}
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
' Start Excel if it is not currently running.
'
Function StartExcel() As Boolean
    Const sAppName = "Excel.Application"
    ' Getting the object will error if Excel is not
    ' running.  Catch the error and start Excel if that
    ' is the case.
    On Error Resume Next
    Set oExcelApp = GetObject(, sAppName)
    On Error GoTo 0
    If (oExcelApp Is Nothing) Then
        On Error Resume Next
        Set oExcelApp = CreateObject(sAppName)
        On Error GoTo 0
        If (oExcelApp Is Nothing) Then
            MsgBox "Error trying to launch Excel using OLE Automation.  Program will end."
            StartExcel = False
            Exit Function
        End If
    End If
    oExcelApp.Visible = True
    StartExcel = True
End Function


'
' Write information about the selected pipe network to a
' spreadsheed in an instance of Excel.
'
Public Sub ExportToExcel()
    Dim oAcadObject As AcadObject
    Dim oPipe As AeccPipe
    Dim oStructure As AeccStructure
    Dim sheetPipes As Worksheet
    Dim sheetStructures As Worksheet
    Dim iRowPipes As Integer
    Dim iColumnPipes As Integer
    Dim iRowStructures As Integer
    Dim iColumnstructures As Integer
    Dim oPartDataField As AeccPartDataField
    Dim oPartDataRecord As AeccPartDataRecord
    Dim dictPipe As Dictionary
    Dim dictStructure As Dictionary
    Dim oExcelBook As Excel.Workbook
    
    If (StartExcel() = False) Then
        Exit Sub
    End If

    Set oExcelBook = oExcelApp.Workbooks.Add
    If (oExcelBook Is Nothing) Then
        MsgBox "Error creating Excel document, exit."
        Exit Sub
    End If

    ' dictionary objects to hold column names and positions
    Set dictPipe = New Dictionary
    Set dictStructure = New Dictionary
    'make a new sheet for structures
    Set sheetStructures = oExcelBook.Sheets.Add
    sheetStructures.Name = "Structures"
    'add a new sheet for pipes
    Set sheetPipes = oExcelBook.Sheets.Add
    sheetPipes.Name = "Pipes"
    
    iRowPipes = 1
    iRowStructures = 1
    For Each oAcadObject In ThisDrawing.ModelSpace
        If (TypeOf oAcadObject Is AeccPipe) Then
            ' Yes, we have a pipe.
            Set oPipe = oAcadObject
            ' get the data attached to the pipe
            Set oPartDataRecord = oPipe.PartDataRecord
            ' if this is the first pipe, then we need to add the headings to Excel
            If (iRowPipes = 1) Then
            ' handle of the pipe so that we can update it
                sheetPipes.Cells(iRowPipes, 1).Value = "Handle"
                'starting point of the pipe. If the pipe is connected to a structure, the position can't be changed
                sheetPipes.Cells(iRowPipes, 2).Value = "Start"
                sheetPipes.Cells(iRowPipes, 3).Value = "End"
                ' slope of the pipe = rise / 2D pipe length
                sheetPipes.Cells(iRowPipes, 4).Value = "Slope"
                dictPipe.Add "Handle", 1
                dictPipe.Add "StartPoint", 2
                dictPipe.Add "endPoint", 3
                dictPipe.Add "Slope", 4
                iRowPipes = iRowPipes + 1
                ' we change change the slope, so color the column yellow
                sheetPipes.Columns(4).Interior.ColorIndex = 6
                sheetPipes.Columns(4).Interior.Pattern = xlSolid
                sheetPipes.Columns(4).Interior.PatternColorIndex = xlAutomatic
            End If
            ' add the values
            sheetPipes.Cells(iRowPipes, 1).Value = oPipe.handle
            sheetPipes.Cells(iRowPipes, 2).Value = oPipe.StartPoint.X & "," & oPipe.StartPoint.Y & "," & oPipe.StartPoint.Z
            sheetPipes.Cells(iRowPipes, 3).Value = oPipe.EndPoint.X & "," & oPipe.EndPoint.Y & "," & oPipe.EndPoint.Z
            sheetPipes.Cells(iRowPipes, 4).Value = (oPipe.EndPoint.Z - oPipe.StartPoint.Z) / oPipe.Length2D
            For Each oPartDataField In oPartDataRecord
                ' make sure the data has a column in Excel, if not, we add the column
                If (Not dictPipe.Exists(oPartDataField.ContextString)) Then
                    sheetPipes.Cells(1, dictPipe.Count + 1).Value = oPartDataField.ContextString
                    dictPipe.Add oPartDataField.ContextString, dictPipe.Count + 1
                    If (oPartDataField.RawDataSource = aeccRuntimeAppend) Then
                    ' run time appended is the user defined data, it will be green in Excel
                        sheetPipes.Columns(dictPipe.Count).Interior.ColorIndex = 4
                        sheetPipes.Columns(dictPipe.Count).Interior.Pattern = xlSolid
                        sheetPipes.Columns(dictPipe.Count).Interior.PatternColorIndex = xlAutomatic
                    End If
                    If ((oPartDataField.ContextString = "Catalog_PartID") Or (oPartDataField.ContextString = "PipeInnerDiameter") Or (oPartDataField.ContextString = "PipeInnerWidth")) Then
                    ' for the dimensions that can be modified by the current import routine
                        sheetPipes.Columns(dictPipe.Count).Interior.ColorIndex = 6
                        sheetPipes.Columns(dictPipe.Count).Interior.Pattern = xlSolid
                        sheetPipes.Columns(dictPipe.Count).Interior.PatternColorIndex = xlAutomatic
                    End If
                End If
                iColumnPipes = dictPipe.Item(oPartDataField.ContextString)
                sheetPipes.Cells(iRowPipes, iColumnPipes).Value = oPartDataField.Tag
            Next
            
            iRowPipes = iRowPipes + 1
        ElseIf (TypeOf oAcadObject Is AeccStructure) Then
            ' we have a structure. The import doesn't use the structure, so the data is just streamed out...
            Set oStructure = oAcadObject
            Set oPartDataRecord = oStructure.PartDataRecord
            If (iRowStructures = 1) Then
                sheetStructures.Cells(iRowStructures, 1).Value = "Handle"
                sheetStructures.Cells(iRowStructures, 2).Value = "Location"
                iColumnstructures = 3
                For Each oPartDataField In oPartDataRecord
                    sheetStructures.Cells(iRowStructures, iColumnstructures).Value = oPartDataField.ContextString
                    iColumnstructures = iColumnstructures + 1
                Next
                iRowStructures = iRowStructures + 1
            End If
            sheetStructures.Cells(iRowStructures, 1).Value = oStructure.handle
            sheetStructures.Cells(iRowStructures, 2).Value = oStructure.Position.X & "," & oStructure.Position.Y & "," & oStructure.Position.Z
            iColumnstructures = 3
            For Each oPartDataField In oPartDataRecord
                sheetStructures.Cells(iRowStructures, iColumnstructures).Value = oPartDataField.Tag
                iColumnstructures = iColumnstructures + 1
            Next
            iRowStructures = iRowStructures + 1
        End If
    Next
    Set oExcelBook = Nothing
    Set oExcelApp = Nothing
End Sub


'
' Import pipe data from the open Excel spreadsheet.
'
Public Sub ImportFromExcel()
    Dim oExcelBook As Workbook
    Dim dictPipes As Dictionary
    Dim columnTitle As String
    Dim i As Integer
    Dim oAcadObject As AcadObject
    Dim oPipe As AeccPipe
    Dim guid As String
    Dim dia As Double
    Dim slope As Double
    Dim col As Integer
    Dim oPipeSheet As Worksheet
    Dim oStructure As AeccStructure
    Dim shape As String
    Dim j As Integer
    Dim oPartDataField As AeccPartDataField
    Dim oPartDataRecord As AeccPartDataRecord

    
    If (StartExcel() = False) Then
        Exit Sub
    End If
    
    Set oExcelBook = oExcelApp.ActiveWorkbook
    Set oPipeSheet = oExcelBook.Sheets.Item("Pipes")
    ' looking for the pipes sheet in the workbook
    Set dictPipes = New Dictionary
    i = 1
    columnTitle = oPipeSheet.Cells(1, i).Value
    Do While columnTitle <> ""
    ' add the column names to the dictonary so that we can find them later
        dictPipes.Add columnTitle, i
        i = i + 1
        columnTitle = oPipeSheet.Cells(1, i).Value
    Loop

    i = 2 ' loop the rows
    Dim handle As String
    handle = oPipeSheet.Cells(i, 1).Value
    Do
    ' from the handle, find the object
        Set oAcadObject = ThisDrawing.HandleToObject(handle)
        If (Not oAcadObject Is Nothing) Then
        ' make sure it's still a pipe
            If (TypeOf oAcadObject Is AeccPipe) Then
                Set oPipe = oAcadObject
                ' this sets the part family for the pipe
                col = dictPipes.Item("Catalog_PartID")
                guid = oPipeSheet.Cells(i, col).Value
                ' this sets the shape for the pipe
                col = dictPipes.Item("Catalog_PartName")
                shape = oPipeSheet.Cells(i, col).Value
                ' this only works with round and egg shaped pipes
                If shape Like "AeccEggShaped*" Then
                ' egg shaped sizes based on inner width
                    col = dictPipes.Item("PipeInnerWidth")
                    dia = oPipeSheet.Cells(i, col).Value
                Else
                ' round sizes based on diameter
                    col = dictPipes.Item("PipeInnerDiameter")
                    dia = oPipeSheet.Cells(i, col).Value
                End If
                
                
                ' Resize the pipe.
                oPipe.ResizeByInnerDiaOrWidth guid, dia, True


                ' adjust the slope of the pipe
                col = dictPipes.Item("Slope")
                slope = oPipeSheet.Cells(i, col).Value
                oPipe.EndPoint.Z = oPipe.StartPoint.Z + (slope * oPipe.Length2D)
                ' this will disconnect and reconnect the pipe to the structure, forcing the structure to update for the new slope
                Set oStructure = oPipe.StartStructure
                If Not (oStructure Is Nothing) Then
                    oPipe.Disconnect aeccPipeStart
                    oPipe.ConnectToStructure aeccPipeStart, oStructure
                End If
                Set oStructure = oPipe.EndStructure
                If Not (oStructure Is Nothing) Then
                    oPipe.Disconnect aeccPipeEnd
                    oPipe.ConnectToStructure aeccPipeEnd, oStructure
                End If
                    

                Set oPartDataRecord = oPipe.PartDataRecord

                For j = 5 To dictPipes.Count ' look over the spreadsheet for Runtime values (green)
                    If ((oPipeSheet.Cells(i, j).Interior.ColorIndex = 4) And (oPipeSheet.Cells(i, j).Value <> "")) Then
                        ' we have a runtime value
                        Set oPartDataField = Nothing
                        Dim context As String
                        context = oPipeSheet.Cells(1, j).Value
                        Set oPartDataField = oPartDataRecord.FindByContext(context)
                        If oPartDataField Is Nothing Then
                        ' append the data to the record
                            Set oPartDataField = oPartDataRecord.Append(context, 0)
                        End If
                        'oPartDataField.Tag = oPipeSheet.Cells(i, j).Value
                    End If
                Next
            End If
        End If
        i = i + 1
        handle = oPipeSheet.Cells(i, 1).Value
    Loop While handle <> ""

    Set oPipeSheet = Nothing
    Set oExcelBook = Nothing
End Sub

