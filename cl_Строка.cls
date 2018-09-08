VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_Строка"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function расКавычить(ByVal str As String, _
                            Optional msg As Variant) _
                            As String
    ' Возвращает строку, усеченную по кавычки
    ' Отсечь слева до первой кавычки
    'Отсечь справа от первой кавычки

    str = Кавычки_Разные_на_Chr34_Заменить(str)

    If Символ_в_Строке_Раз(Chr$(34), str) > 1 Then

        расКавычить = Отсечь_Справа(Chr$(34), _
                                    Отсечь_Слева(Chr$(34), str))
    Else

        расКавычить = str
    End If
End Function

Public Function Отсечь_Справа(ByVal Symb As String, _
                              str As String) _
                              As String
    Отсечь_Справа = Left$(str, InStrRev(str, Symb) - 1)
End Function

Public Function Отсечь_Слева(ByVal Symb As String, _
                             str As String) _
                             As String
    Отсечь_Слева = Right$(str, _
                          Len(str) - InStr(str, Symb))
End Function

Public Function Кавычки_Разные_на_Chr34_Заменить(ByVal str As String) _
       As String
    'заменить разные кавычки на одинаковые
    str = Replace$(str, "«", Chr$(34))
    str = Replace$(str, "»", Chr$(34))

    Кавычки_Разные_на_Chr34_Заменить = str
End Function

Public Function Символ_в_Строке_Раз(ByVal Symb As String, _
                                    ByVal str As String, _
                                    Optional ByVal Регистр As Long = 1) _
                                    As Long

    Dim Строка_Длина As String, Строка_Длина_без_Символов As String

    Строка_Длина = Len(str)

    Строка_Длина_без_Символов = _
    Len(Replace$(str, Symb, vbNullString, , , Регистр))

    Символ_в_Строке_Раз = Строка_Длина - _
                          Строка_Длина_без_Символов
End Function



