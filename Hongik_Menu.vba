
Sub UpdateMenu()
    Dim http As New WinHttp.WinHttpRequest
    Dim apiURL As String
    Dim responseText As String
    Dim targetShape As Shape
    Dim finalResult As String
    Dim todayStr As String
    
    ' 1. 오늘 날짜 설정 (YYYYMMDD)
    todayStr = Format(Date, "yyyymmdd")
    
    
    ' API 주소
    apiURL = "https://www.hongik.ac.kr/sso/APICipher2.jsp?data=%7B%22url%22%3A%22%2Fhomepage%2Fget_food_list.php%22%2C%22url2%22%3A%22%26CAMPUS%3D%22%2C%22url3%22%3A%220%22%7D"
    
    On Error GoTo ErrorHandler
    
    ' 2. 데이터 요청
    With http
        .Open "GET", apiURL, False
        .Send
        .WaitForResponse
        responseText = BinaryToText(.ResponseBody, "utf-8")
    End With

    ' 3. No.4와 No.5 데이터만 추출하여 조합
    finalResult = GetSpecificMenuPair(responseText, todayStr, 4, 5)

    ' 4. 파워포인트 도형("MenuBox")에 넣기
    On Error Resume Next
    Set targetShape = ActivePresentation.Slides(1).Shapes("MenuBox")
    On Error GoTo 0
    
    If Not targetShape Is Nothing Then
        targetShape.TextFrame.TextRange.Text = finalResult
        
    Else
        MsgBox "슬라이드 1에서 'MenuBox' 도형을 찾을 수 없습니다."
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "오류 발생: " & Err.Description
End Sub

' --- 특정 순번(Index)의 메뉴 2개를 추출하여 합치는 함수 ---
Function GetSpecificMenuPair(jsonText As String, targetDate As String, idx1 As Integer, idx2 As Integer) As String
    Dim items() As String
    Dim i As Integer, count As Integer
    Dim menu1 As String, menu2 As String
    Dim item As String, iDate As String, iMenu As String
    
    items = Split(jsonText, "{")
    count = 0
    
    For i = 1 To UBound(items)
        item = items(i)
        iDate = SimpleExtract(item, "MENU_DATE")
        
        If iDate = targetDate Then
            count = count + 1
            
            ' 4번째 데이터 저장
            If count = idx1 Then
                menu1 = URLDecodeUTF8(SimpleExtract(item, "MENU"))
                menu1 = Replace(menu1, "+", " ")
            End If
            
            ' 5번째 데이터 저장
            If count = idx2 Then
                menu2 = URLDecodeUTF8(SimpleExtract(item, "MENU"))
                menu2 = Replace(menu2, "+", " ")
            End If
        End If
    Next i
    
    ' 결과 조립 (위: No.4 / 아래: No.5)
    Dim combined As String
    If menu1 <> "" Then combined = menu1 Else combined = "No.4 정보를 찾을 수 없습니다."
    
    combined = combined & vbCrLf & "-----------------------" & vbCrLf
    
    If menu2 <> "" Then combined = combined & menu2 Else combined = combined & "No.5 정보를 찾을 수 없습니다."
    
    GetSpecificMenuPair = combined
End Function

' --- (보조 함수들: 이전과 동일) ---
Function SimpleExtract(txt As String, key As String) As String
    On Error Resume Next
    Dim p As String: p = """" & key & """:"""
    Dim s As Long: s = InStr(txt, p)
    If s > 0 Then
        s = s + Len(p)
        SimpleExtract = Mid(txt, s, InStr(s, txt, """") - s)
    End If
End Function

Function URLDecodeUTF8(ByVal str As String) As String
    Dim i As Long: Dim b() As Byte: Dim byteIdx As Long: Dim charCode As String
    If str = "" Then Exit Function
    ReDim b(Len(str)): byteIdx = 0
    For i = 1 To Len(str)
        charCode = Mid(str, i, 1)
        If charCode = "%" Then
            b(byteIdx) = CByte("&H" & Mid(str, i + 1, 2)): i = i + 2
        ElseIf charCode = "+" Then: b(byteIdx) = Asc(" ")
        Else: b(byteIdx) = Asc(charCode)
        End If
        byteIdx = byteIdx + 1
    Next i
    If byteIdx > 0 Then
        ReDim Preserve b(byteIdx - 1)
        URLDecodeUTF8 = BinaryToText(b, "utf-8")
    End If
End Function

Function BinaryToText(BinaryData() As Byte, CharSet As String) As String
    Dim Stream As Object: Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = 1: Stream.Open: Stream.Write BinaryData: Stream.Position = 0
    Stream.Type = 2: Stream.CharSet = CharSet
    BinaryToText = Stream.ReadText
End Function
