Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range
    Set KeyCells = Range("A1")
    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then
        

    Dim sku As String
    sku = Range("A1").Value
    If Len(sku) = 5 Then
        sku = "0" + sku
    ElseIf Left(sku, 1) = "x" And Len(sku) = 7 Then
        sku = Mid(sku, 2, Len(sku))
    ElseIf Len(sku) = 6 Then
        sku = sku
    Else
        sku = 123456
    End If
    Dim sURL As String, sResult As String
    sURL = "http://www.microcenter.com/search/search_results.aspx?Ntt=" & sku
    sResult = GetHTTPResult(sURL)
    Dim unparsedPrice As String, i As Long, j As Long, price As String
    If InStr(sResult, "We're sorry") Then
        unparsedPrice = 0
    Else
        i = InStr(sResult, "$</span>")
        j = InStrRev(sResult, "</span>")
        unparsedPrice = Mid(sResult, i + 1, j - i - 1)
    End If
    price = onlyDigits(unparsedPrice)
    price = FindOpenBox(price)
    Range("F1") = price
    End If
End Sub
        

Function GetHTTPResult(sURL As String) As String
    Dim XMLHTTP As Variant, sResult As String
    Set XMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    XMLHTTP.Open "GET", sURL, False
    XMLHTTP.send
    Debug.Print "Status: " & XMLHTTP.Status & " - " & XMLHTTP.statusText
    sResult = XMLHTTP.responseText
    Debug.Print "Length of response: " & Len(sResult)
    Set XMLHTTP = Nothing
    GetHTTPResult = sResult
End Function


Function onlyDigits(s As String) As String
    Dim retval As String
    Dim i As Integer
    retval = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Or Mid(s, i, 1) = "." Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
    onlyDigits = retval
End Function

Function FindOpenBox(price As String) As String
    Dim splitPoint As Integer, char As String, newPrice As String, OpenboxPrice As String
    If CharCounter(price, ".") = 2 Then
        splitPoint = InStr(1, price, ".")
        char = Mid(price, splitPoint + 3, 1)
        newPrice = Left(price, splitPoint + 2)
        OpenboxPrice = Replace(price, char, " ", splitPoint + 3, 1)
        price = newPrice & "  " & OpenboxPrice
    End If
    FindOpenBox = price
End Function


Function CharCounter(ByVal MyString As String, ByVal CharToSearch As String) As Integer
    Dim X As Integer
    CharCounter = 0
    For X = 1 To Len(MyString)
        If Mid(MyString, X, 1) = CharToSearch Then CharCounter = CharCounter + 1
    Next
End Function






























































