Attribute VB_Name = "�������"
Option Explicit

Public Function ������_��������_��������(Optional ByVal msg As Variant)
    Cells.Replace What:=".cls", Replacement:="", LookAt:=xlPart, SearchOrder _
                                                                 :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Dim eL As Range
    For Each eL In ActiveSheet.UsedRange
        eL.Value = "Public " & eL.Value & " as New " & eL.Value
    Next
End Function

