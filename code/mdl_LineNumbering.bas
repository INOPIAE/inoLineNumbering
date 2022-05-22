Attribute VB_Name = "mdl_LineNumbering"
Option Explicit

'Inspired by https://www.herber.de/forum/archiv/1552to1556/1555454_Code_fuer_Zeilennummerierung_von_Nepumuk.html

' The given codes adds line numbering to begining of code line to be used with the Erl() function
' strVBProjects - name of vba project not file name
' strModuleName - name of code module
' vbaComponent - a VBIDE.VBComponent object
' blnNoNumber = True removes line numbering only

Public Function AddLineNumbersToWorkbook(strVBProjects As String, Optional blnNoNumber As Boolean = False) As Long
    ' returns total line numbers added to code in vba project
    Dim lngCount As Long
    With Application.VBE.VBProjects(strVBProjects)
        For intModulCount = 1 To .VBComponents.Count
            lngCount = lngCount + AddLineNumbersToComponent(.VBComponents(intModulCount))
        Next
    End With
    AddLineNumbersToWorkbook = lngCount
End Function

Public Function AddLineNumbersToSingleCodeObject(strVBProjects As String, strModuleName As String, Optional blnNoNumber As Boolean = False) As Long
    ' returns total line numbers added to code of a single code object identified by the module name
    AddLineNumbersToSingleCodeObject = AddLineNumbersToComponent(Application.VBE.VBProjects(strVBProjects).VBComponents(strModuleName), blnNoNumber)
End Function

Public Function AddLineNumbersToComponent(vbaComponent As VBIDE.VBComponent, Optional blnNoNumber As Boolean = False) As Long
    ' returns total line numbers added to code of a single code object as passed to the function
    Dim intModulCount As Integer, intLine As Integer
    Dim intColumn As Integer, intLineCounter As Integer
    Dim strModulname As String
    Dim bolUnderscore As Boolean, bolSelect As Boolean
    Dim lngCount As Long

    With vbaComponent.CodeModule
        For intLine = .CountOfDeclarationLines + 1 To .CountOfLines
            If Trim$(.Lines(intLine, 1)) <> "" And Left$(Trim$(.Lines(intLine, 1)), 1) <> "'" Then
                If .ProcOfLine(intLine, 0) <> strModulname Then
                    strModulname = .ProcOfLine(intLine, 4)
                    intLineCounter = 0
                    
                    If Left$(Trim$(StrReverse(.Lines(intLine, 1))), 1) = "_" Then
                        bolUnderscore = True
                    Else
                        bolUnderscore = False
                    End If
                Else
                    If InStr(1, "End Sub End Function End Property", .Lines(intLine, 1)) = 0 Then
                        If Not bolUnderscore And Not bolSelect Then
                        
                            If Left$(Trim$(StrReverse(.Lines(intLine, 1))), 1) = "_" Then bolUnderscore = True
                            If InStr(1, .Lines(intLine, 1), "Select Case") <> 0 Then bolSelect = True
                            If IsNumeric(Left$(.Lines(intLine, 1), 1)) Then
                                For intColumn = 1 To Len(.Lines(intLine, 1))
                                    If Not IsNumeric(Left$(.Lines(intLine, 1), intColumn)) Then
                                        Exit For
                                    End If
                                Next
                                .ReplaceLine intLine, String(intColumn - 1, " ") & Mid$(.Lines(intLine, 1), intColumn)
                            End If
                            intLineCounter = intLineCounter + 1
                            If blnNoNumber = False Then
                                If Trim$(Left$(.Lines(intLine, 1), Len(Trim(intLineCounter)) + 2)) = "" Then
                                    .ReplaceLine intLine, Mid$(.Lines(intLine, 1), Len(Trim(intLineCounter)) + 2)
                                Else
                                    .ReplaceLine intLine, Trim$(.Lines(intLine, 1))
                                End If
                                .ReplaceLine intLine, Trim$(CStr(intLineCounter)) & " " & .Lines(intLine, 1)
                                lngCount = lngCount + 1
                            End If
                        Else
                            If Left$(Trim$(StrReverse(.Lines(intLine, 1))), 1) <> "_" Then bolUnderscore = False
                            If InStr(1, .Lines(intLine, 1), "Case") <> 0 Then bolSelect = False
                        End If
                    Else
                        strModulname = ""
                    End If
                End If
            End If
        Next
    End With
    AddLineNumbersToComponent = lngCount
End Function

Public Sub LoadVBAReference()
    'Needed to set refernce to Microsoft Visual Basic for Applications Extensibility 5.3 library if not availbale in Tools - References
    Call Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 5, 3)
End Sub


