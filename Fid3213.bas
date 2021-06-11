Attribute VB_Name = "Fid3213"
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
    
    startRow = 2
    endRow = 2571
    
    Dim codes(2570) As String
    Dim names(2570) As String
    
    Worksheets("stockinfo").Activate
    For i = 0 To endRow - startRow
        codes(i) = ActiveWorkbook.Worksheets("stockinfo").Cells(i + 2, 2).Value
        names(i) = ActiveWorkbook.Worksheets("stockinfo").Cells(i + 2, 4).Value
    Next i
    
    Worksheets("stockmember").Activate
    Dim rowNumber As Integer
    
    For i = startRow To endRow
        code = codes(i - startRow)
Try:
            rowNumber = i
            Debug.Print i
            fidInput = "[{""idx"":""fid3213"",""gid"":""3212"",""fidCodeBean"":{""3"":""" + code + """,""9104"":""J"",""9220"":""2""},""outFid"":""500,4,5,6,7,8,912,911,913,1547,915,917,914,2121,916,918,919,920,837,839,838,921"",""isList"":""1"",""order"":""ASC"",""reqCnt"":1,""actionKey"":""0"",""saveBufLen"":""1"",""saveBuf"":""1""}]"
            Set Response = Client.PostJson("https://www.samsungpop.com/wts/fidBuilder.do", fidInput)
            Dim Json As Object
            If Response.StatusCode = WebStatusCode.Ok Then
                
                Set Json = JsonConverter.ParseJson(StrConv(Response.Body, vbUnicode))
                
                ActiveWorkbook.Worksheets("stockmember").Cells(i, 2).Value = codes(i - startRow)
                ActiveWorkbook.Worksheets("stockmember").Cells(i, 3).Value = names(i - startRow)
                ActiveWorkbook.Worksheets("stockmember").Cells(i, 4).Value = Json("fid3213")("data").Item(1)("500")
                ActiveWorkbook.Worksheets("stockmember").Cells(i, 5).Value = Json("fid3213")("data").Item(1)("912")
                ActiveWorkbook.Worksheets("stockmember").Cells(i, 6).Value = Json("fid3213")("data").Item(1)("837")
                ActiveWorkbook.Worksheets("stockmember").Cells(i, 7).Value = Json("fid3213")("data").Item(1)("913")
                ActiveWorkbook.Worksheets("stockmember").Cells(i, 8).Value = Json("fid3213")("data").Item(1)("1547")
            End If
            Set RecentResponse = Client.GetJson("https://www.stockplus.com/api/securities.json?ids=KOREA-A" + code)
            Dim RecentJson As Object
            If RecentResponse.StatusCode = WebStatusCode.Ok Then
                
                Set RecentJson = JsonConverter.ParseJson(RecentResponse.Content)
        
                Worksheets("stockmember").Cells(rowNumber, 9).Value = RecentJson("recentSecurities").Item(1)("tradePrice")
                Worksheets("stockmember").Cells(rowNumber, 10).Value = RecentJson("recentSecurities").Item(1)("changePriceRate") * 100
                Worksheets("stockmember").Cells(rowNumber, 11).Value = RecentJson("recentSecurities").Item(1)("changePrice")
                Worksheets("stockmember").Cells(rowNumber, 12).Value = RecentJson("recentSecurities").Item(1)("openingPrice")
                Worksheets("stockmember").Cells(rowNumber, 13).Value = RecentJson("recentSecurities").Item(1)("highPrice")
                Worksheets("stockmember").Cells(rowNumber, 14).Value = RecentJson("recentSecurities").Item(1)("lowPrice")
            End If
Catch:
            Debug.Print "Error : " + code

    Next i

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


