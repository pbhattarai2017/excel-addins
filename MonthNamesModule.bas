Attribute VB_Name = "MonthNamesModule"
Public monthNames(11) As String
Sub initializeMonthNames()
    monthNames(0) = "Baishakh"
    monthNames(1) = "Jestha"
    monthNames(2) = "Ashard"
    monthNames(3) = "Shrawan"
    monthNames(4) = "Bhadra"
    monthNames(5) = "Asoj"
    monthNames(6) = "Kartik"
    monthNames(7) = "Mangsir"
    monthNames(8) = "Poush"
    monthNames(9) = "Magh"
    monthNames(10) = "Falgun"
    monthNames(11) = "Chaitra"
End Sub
' Returns Baishakh for 1, Jestha for 2, etc.
Function NEPALIMONTH(ByVal miti As String) As String
    Dim m As Integer
    Dim count As Integer
    Dim loc As Integer
    Dim ch As Integer
    count = 0
    loc = -1
    m = -1
    For i = 1 To Len(miti)
        ch = Asc(Mid(miti, i, 1))
        If ch < 48 Or ch > 57 Then
            count = count + 1
            loc = i
        End If
    Next
    If count = 2 Then
        m = CInt(Split(miti, Mid(miti, loc, 1))(1))
    End If
    If m <= 12 And m >= 1 Then
        NEPALIMONTH = monthNames(m - 1)
    Else
        NEPALIMONTH = "ERROR"
    End If
End Function
Function NEPALIMONTHNUMBER(ByVal miti As String) As String
    Dim m As Integer
    Dim count As Integer
    Dim loc As Integer
    Dim ch As Integer
    count = 0
    loc = -1
    m = -1
    For i = 1 To Len(miti)
        ch = Asc(Mid(miti, i, 1))
        If ch < 48 Or ch > 57 Then
            count = count + 1
            loc = i
        End If
    Next
    If count = 2 Then
        m = CInt(Split(miti, Mid(miti, loc, 1))(1))
    End If
    If m <= 12 And m >= 1 Then
        NEPALIMONTHNUMBER = "[" & m & "]"
    Else
        NEPALIMONTHNUMBER = "ERROR"
    End If
End Function
Sub togglePrecedents(ByVal showPrecedents As Boolean)
    If showPrecedents Then
        Selection.showPrecedents
    Else
        Selection.showPrecedents Remove:=True
    End If
End Sub
Private Sub Workbook_Open()
    Application.OnKey "^t", "'togglePrecedents True'"
    Application.OnKey "+^t", "'togglePrecedents False'"
    Call initializeMonthNames
End Sub
