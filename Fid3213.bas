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
    Dim codes As Object


Try:
            Set Response = Client.GetJson("http://jrimchoi.iptime.org/samsung/fid3213/20210621")
            Dim Json As Object
            If Response.StatusCode = WebStatusCode.Ok Then
                
                Set Json = JsonConverter.ParseJson(Response.Content)
                For i = 1 To Json.Count
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 2).Value = Json.Item(i)("종목코드")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 3).Value = Json.Item(i)("일자")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 4).Value = Json.Item(i)("현재가")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 5).Value = Json.Item(i)("전일대비")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 6).Value = Json.Item(i)("등락율")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 7).Value = Json.Item(i)("거래량")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 8).Value = Json.Item(i)("개인")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 9).Value = Json.Item(i)("기관")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 10).Value = Json.Item(i)("외국인")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 11).Value = Json.Item(i)("프로그램")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 12).Value = Json.Item(i)("연기금")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 13).Value = Json.Item(i)("금융투자")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 14).Value = Json.Item(i)("보험")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 15).Value = Json.Item(i)("투신")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 16).Value = Json.Item(i)("사모펀드")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 17).Value = Json.Item(i)("은행")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 18).Value = Json.Item(i)("기타금융")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 19).Value = Json.Item(i)("기타법인")
                    ActiveWorkbook.Worksheets("members").Cells(i + 2, 20).Value = Json.Item(i)("기타외국인")
                    Debug.Print i
                Next i
            End If

Catch:
            Debug.Print 'End'


End Function


