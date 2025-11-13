Attribute VB_Name = "Site"
'
' MODULE_ID {Site.bas}
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
' Create an AeccSite object.
'
Public Function CreateSite() As Boolean
    On Error Resume Next
    If (g_oPipeDocument Is Nothing) Then
        CreateSite = False
        Exit Function
    End If
    
    Dim oSites As AeccSites
    Set oSites = g_oPipeDocument.Sites
    Dim oSite As AeccSite

    Set oSite = oSites.Item(SITE_NAME)
    If oSite Is Nothing Then
        Set oSite = oSites.Add(SITE_NAME)
        If oSite Is Nothing Then
            Debug.Print "Error creating " & SITE_NAME
            CreateSite = False
            Exit Function
        End If
    End If
    CreateSite = True
End Function


'
' Get a site with the name
'
Public Function GetSite() As AeccSite
    Dim oSites As AeccSites
    Dim oSite As AeccSite
    
    On Error Resume Next
    Set oSites = g_oPipeDocument.Sites
    If (Not oSites Is Nothing) Then
        Set oSite = oSites.Item(SITE_NAME)
        If (oSite Is Nothing) Then
            Set GetSite = Nothing
        Else
            Set GetSite = oSite
        End If
    End If
End Function


