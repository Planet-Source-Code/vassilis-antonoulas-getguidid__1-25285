Attribute VB_Name = "mdlGetGuidID"
'*******************************************************************************
' MODULE:       mdlGetGuidID
' FILENAME:     mdlGetGuidID.mdl
' AUTHOR:       Vassilis Antonoulas
' CREATED:      20-Jul-2001
' COPYRIGHT:    Copyright 2001 XpressWeb Hellas Ltd. All Rights Reserved.
'
' DESCRIPTION:
' This module creates a Genuinely Unique Identification Number.
'
'
' MODIFICATION HISTORY:
' 1.0       20-Jul-2001
'           Vassilis Antonoulas
'           Initial Version
'
'*******************************************************************************
Public Function GetGuidID()

    Dim Guid1 As String
    Dim Guid2 As String
    Dim Guid3 As String
    Dim Guid4 As String
    Dim Guid5 As String
    Dim Unq As String
    Dim RndLetter As Integer
    Dim RndIndex As Integer
    Dim iCount As Integer
    
    Randomize
    
    'Initialize all variables
    Guid1 = "": Guid2 = "": Guid3 = "": Guid4 = "": Guid5 = ""
    Unq = "{"
    
    'Create the first part of the GUID
    For iCount = 1 To 8
        RndIndex = Int(Rnd * 2) + 1
        If RndIndex = 1 Then
            RndLetter = Int(Rnd * 10) + 48 'Choose a random number from 0 to 9
        Else
            RndLetter = Int(Rnd * 6) + 65 'Choose a random letter from A to F
        End If
        Guid1 = Guid1 & Chr(RndLetter)
    Next iCount
    
    Unq = Unq & Guid1 & "-"
    
    'Create the second part of the GUID
    For iCount = 1 To 4
        RndIndex = Int(Rnd * 2) + 1
        If RndIndex = 1 Then
            RndLetter = Int(Rnd * 10) + 48 'Choose a random number from 0 to 9
        Else
            RndLetter = Int(Rnd * 6) + 65 'Choose a random letter from A to F
        End If
        Guid2 = Guid2 & Chr(RndLetter)
    Next iCount

    Unq = Unq & Guid2 & "-"
    
    'Create the third part of the GUID
    For iCount = 1 To 4
        RndIndex = Int(Rnd * 2) + 1
        If RndIndex = 1 Then
            RndLetter = Int(Rnd * 10) + 48 'Choose a random number from 0 to 9
        Else
            RndLetter = Int(Rnd * 6) + 65 'Choose a random letter from A to F
        End If
        Guid3 = Guid3 & Chr(RndLetter)
    Next iCount

    Unq = Unq & Guid3 & "-"
    
    'Create the forth part of the GUID
    For iCount = 1 To 4
        RndIndex = Int(Rnd * 2) + 1
        If RndIndex = 1 Then
            RndLetter = Int(Rnd * 10) + 48 'Choose a random number from 0 to 9
        Else
            RndLetter = Int(Rnd * 6) + 65 'Choose a random letter from A to F
        End If
        Guid4 = Guid4 & Chr(RndLetter)
    Next iCount

    Unq = Unq & Guid4 & "-"
    
    'Create the fifth part of the GUID
    For iCount = 1 To 12
        RndIndex = Int(Rnd * 2) + 1
        If RndIndex = 1 Then
            RndLetter = Int(Rnd * 10) + 48 'Choose a random number from 0 to 9
        Else
            RndLetter = Int(Rnd * 6) + 65 'Choose a random letter from A to F
        End If
        Guid5 = Guid5 & Chr(RndLetter)
    Next iCount

    Unq = Unq & Guid5 & "}"

    GetGuidID = Unq

End Function
