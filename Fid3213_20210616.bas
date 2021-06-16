Dim pClient As WebClient
Public Property Get FidClient() As WebClient
    If pClient Is Nothing Then
        Set pClient = New WebClient
        pClient.BaseUrl = "https://www.samsungpop.com/wts/fidBuilder.do"
        
        'Dim Auth As New TodoistAuthenticator
        'Auth.Setup CStr(Credentials.Values("Todoist")("id")), CStr(Credentials.Values("Todoist")("secret")), CStr(Credentials.Values("Todoist")("redirect_url"))
        'Auth.Scope = "data:read"
        'Auth.Login
        
        'Set pClient.Authenticator = Auth
    End If
    
    Set FidClient = pClient
End Property

Public Function GetFid3213()
    Dim code As String
    
    ' See https://developer.todoist.com/#retrieve-data
    Dim Client As New WebClient
    Dim Request As New WebRequest
    Request.Method = HttpPost
    Dim Response As WebResponse
    Dim fidInput As String
    Dim startRow As Integer
    Dim endRow As Integer
    
    Dim RecentClient As New WebClient
    Dim RecentRequest As New WebRequest
    RecentRequest.Method = HttpGet
    Dim RecentResponse As WebResponse
    Dim codes As Object


Try:
            Set Response = Client.GetJson("http://jrimchoi.iptime.org/samsung/fid3213/20210616")
            Dim Json As Object
            If Response.StatusCode = WebStatusCode.Ok Then
                
                Set Json = JsonConverter.ParseJson(Response.Content)
                For i = 1 To Json.Count
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 2).Value = Json.Item(i)("종목코드")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 3).Value = Json.Item(i)("일자")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 4).Value = Json.Item(i)("현재가")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 5).Value = Json.Item(i)("전일대비")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 6).Value = Json.Item(i)("등락율")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 7).Value = Json.Item(i)("거래량")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 8).Value = Json.Item(i)("개인")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 9).Value = Json.Item(i)("기관")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 10).Value = Json.Item(i)("외국인")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 11).Value = Json.Item(i)("프로그램")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 12).Value = Json.Item(i)("연기금")
                    ActiveWorkbook.Worksheets("stockmember").Cells(i + 1, 13).Value = Json.Item(i)("금융투자")
                Next i
            End If

Catch:
            Debug.Print Response.Content


End Function

Public Function GetRecentSecurity(code As String, rowNumber As Integer)
    Dim Client As New WebClient
    Dim Request As New WebRequest
    Request.Method = HttpGet
    Dim Response As WebResponse
    Set Response = Client.GetJson("https://www.stockplus.com/api/securities.json?ids=KOREA-A" + code)
    Dim Json As Object
    If Response.StatusCode = WebStatusCode.Ok Then
        
        Set Json = JsonConverter.ParseJson(Response.Content)
        Debug.Print Response.Content

        Worksheets("stockmember").Cells(rowNumber, 9).Value = Json("recentSecurities").Item(1)("tradePrice")
        Worksheets("stockmember").Cells(rowNumber, 10).Value = Json("recentSecurities").Item(1)("changePriceRate")
        Worksheets("stockmember").Cells(rowNumber, 11).Value = Json("recentSecurities").Item(1)("openingPrice")
        Worksheets("stockmember").Cells(rowNumber, 12).Value = Json("recentSecurities").Item(1)("highPrice")
        Worksheets("stockmember").Cells(rowNumber, 13).Value = Json("recentSecurities").Item(1)("lowPrice")
    End If
End Function

Public Function GetStocks()
    Dim Client As New WebClient
    Dim Request As New WebRequest
    Request.Method = HttpGet
    Dim Response As WebResponse
    Set Response = Client.GetJson("http://jrimchoi.iptime.org/samsung/stocks")
    Dim Json As Object
    If Response.StatusCode = WebStatusCode.Ok Then
        
        Set Json = JsonConverter.ParseJson(Response.Content)
        Debug.Print Response.Content
        For i = 1 To Json.Count
            Worksheets("stocks").Cells(i + 1, 3).Value = Json.Item(i)("code")
            Worksheets("stocks").Cells(i + 1, 2).Value = Json.Item(i)("name")
            Worksheets("stocks").Cells(i + 1, 4).Value = Json.Item(i)("symbol")
            Worksheets("stocks").Cells(i + 1, 5).Value = Json.Item(i)("csname")
            Worksheets("stocks").Cells(i + 1, 6).Value = Json.Item(i)("mktgbcd")
            Worksheets("stocks").Cells(i + 1, 76).Value = Json.Item(i)("upcode")
        Next i

    End If
    Set GetStocks = Json
End Function

