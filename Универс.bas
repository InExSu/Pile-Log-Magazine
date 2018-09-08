Attribute VB_Name = "”ниверс"
Option Explicit

Public Function  лассы_ќбъ€вить_ƒиапазон(Optional ByVal msg As Variant)
    Cells.Replace What:=".cls", Replacement:="", LookAt:=xlPart, SearchOrder _
                                                                 :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Dim eL As Range
    For Each eL In ActiveSheet.UsedRange
        eL.Value = "Public " & eL.Value & " as New " & eL.Value
    Next
End Function

