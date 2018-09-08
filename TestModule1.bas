Attribute VB_Name = "TestModule1"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub a_Журнал_Свай_Сформировать_TestMethod1()
    On Error GoTo TestFail
    a_Журнал_Свай_Сформировать
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub cl_ОЗП_TestMethod1()
    On Error GoTo TestFail

    Dim cl_ОЗП As New cl_ОЗП

    cl_ОЗП.Залог_Последний_Отказ_Макс _
            cl_ОЗП.Залог_Последний_Отказ_Мин( _
            cl_ОЗП.Залог_Последний_Ударов_Макс( _
            cl_ОЗП.Залог_Последний_Ударов_Мин))
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



