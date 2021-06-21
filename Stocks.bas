Attribute VB_Name = "Stocks"

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
            Worksheets("stocks").Cells(i + 2, 2).Value = Json.Item(i)("code")
            Worksheets("stocks").Cells(i + 2, 3).Value = Json.Item(i)("name")
            Worksheets("stocks").Cells(i + 2, 4).Value = Json.Item(i)("symbol")
            Worksheets("stocks").Cells(i + 2, 5).Value = Json.Item(i)("csname")
            Worksheets("stocks").Cells(i + 2, 6).Value = Json.Item(i)("mktgbcd")
            Worksheets("stocks").Cells(i + 2, 7).Value = Json.Item(i)("upcode")
            Debug.Print i
        Next i

    End If
    
    Set GetStocks = Json
    
End Function

