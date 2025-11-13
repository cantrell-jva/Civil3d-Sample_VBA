Attribute VB_Name = "BaseObjects"
'
' MODULE_ID {BaseObjects.bas}
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


' Global objects representing the document and Civil application.
'
Public g_oPipeApplication As AeccPipeApplication
Public g_oPipeDocument As AeccPipeDocument
Public g_oPipeDatabase As AeccPipeDatabase

' This application uses the type libary from Excel 2003.
' If you have a earlier version, you will need to modify
' the reference to Excel
Public oExcelApp As Excel.Application


'
' Function to set up the Civil 3D Pipe application, document,
' and database objects.
'
Function GetBasePipeObjects() As Boolean
    Dim oApp As AcadApplication
    Set oApp = ThisDrawing.Application
    Dim sAppName As String
    ' NOTE - Always specify the version number.
    sAppName = "AeccXUiPipe.AeccPipeApplication.13.8"
    On Error Resume Next
    Set g_oPipeApplication = oApp.GetInterfaceObject(sAppName)
    On Error GoTo 0
    If (g_oPipeApplication Is Nothing) Then
        MsgBox "Error creating " & sAppName & ", exit."
        GetCivilObjects = False
        GetBasePipeObjects = False
        Exit Function
    End If
    Set g_oPipeDocument = g_oPipeApplication.ActiveDocument
    'Set g_oPipeDatabase = g_oPipeDocument.Database ' Do not need this currently
    GetBasePipeObjects = True
End Function


