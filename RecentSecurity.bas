Attribute VB_Name = "RecentSecurity"
Public Function GetRecentSecurity()
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
    Dim today As String
    Dim now As String
    Dim strength As String
    Dim url As String

Try:
    today = Worksheets("recent").Cells(1, 2).Text
    now = Worksheets("recent").Cells(1, 4).Text
    strength = Worksheets("recent").Cells(1, 6).Text
    If Int(Left(now, 2)) > 18 Or Int(Left(now, 2)) < 9 Then
        now = "1800"
    End If
    
    url = "http://jrimchoi.iptime.org/samsung/recent/" + today + "/" + now + "/" + strength
    Debug.Print url
    Set Response = Client.GetJson("http://jrimchoi.iptime.org/samsung/recent/20210621/1800/100")
    Dim Json As Object
    If Response.StatusCode = WebStatusCode.Ok Then
                
        Set Json = JsonConverter.ParseJson(Response.Content)
        For i = 1 To Json.Count

            Worksheets("recent").Cells(i + 2, 2).Value = Right(Json.Item(i)("shortCode"), 6)
            Worksheets("recent").Cells(i + 2, 3).Value = Json.Item(i)("date")
            Worksheets("recent").Cells(i + 2, 4).Value = Json.Item(i)("tradeTime")
            Worksheets("recent").Cells(i + 2, 5).Value = Json.Item(i)("tradePrice")
            Worksheets("recent").Cells(i + 2, 6).Value = Json.Item(i)("changePriceRate") * 100
            Worksheets("recent").Cells(i + 2, 7).Value = Json.Item(i)("tradeStrength")
            Worksheets("recent").Cells(i + 2, 8).Value = Json.Item(i)("openingPrice")
            Worksheets("recent").Cells(i + 2, 9).Value = Json.Item(i)("highPrice")
            Worksheets("recent").Cells(i + 2, 10).Value = Json.Item(i)("lowPrice")
            Worksheets("recent").Cells(i + 2, 11).Value = Json.Item(i)("dayChartUrl")
            'Debug.Print Json.Item(i)("lowPrice")
        Next i
    End If

Catch:
    Debug.Print 'End'
End Function

