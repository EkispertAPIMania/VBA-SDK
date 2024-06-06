VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ekispert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_ApiKey As String    ' APIキー
Private m_EndPoint As String  ' エンドポイントドメイン
Private m_Secure As Boolean   ' TrueならHTTPS/FalseならHTTP
Private m_Error As EkispertError
Private m_Success As Boolean

Private Sub Class_Initialize()
    m_ApiKey = ""
    m_EndPoint = "api.ekispert.jp"
    m_secture = True
    m_Success = False
End Sub

Property Let Secure(ByVal Secure As Boolean)
    m_Secure = Secure
End Property

Property Let ApiKey(ByVal ApiKey As String)
    m_ApiKey = ApiKey
End Property

Property Get Secure() As Boolean
    Secure = m_Secure
End Property

Property Get ApiKey() As String
    ApiKey = m_ApiKey
End Property

' 駅検索クエリーオブジェクトを返す
Function StationQuery() As StationQuery
    Dim Query As StationQuery
    Set Query = New StationQuery
    Query.client = Me
    Set StationQuery = Query
End Function

Property Get Url() As String
    If m_Secure = True Then
        Url = "https"
    Else
        Url = "http"
    End If
    Url = Url & "://" & m_EndPoint
End Property

Function Fetch(ApiPath As String, Query As Dictionary) As Dictionary
    Dim e As EkispertError
    m_Error = e
    m_Success = False
    Dim client As New WebClient
    client.BaseUrl = Url
    
    Dim DirectionsRequest As New WebRequest
    DirectionsRequest.Method = WebMethod.HttpGet
    Query.Add "key", m_ApiKey
    
    Dim QueryString() As String
    ReDim QueryString(UBound(Query.Items))
    
    Dim i As Integer
    For i = 0 To UBound(Query.Items)
        QueryString(i) = Query.Keys(i) & "=" & WorksheetFunction.EncodeURL(Query.Item(Query.Keys(i)))
    Next i
    
    DirectionsRequest.Resource = ApiPath & "?" & Join(QueryString, "&")
    Dim Response As WebResponse
    Set Response = client.Execute(DirectionsRequest)
    Debug.Print client.GetFullUrl(DirectionsRequest)
    
    If Response.StatusCode <> WebStatusCode.Ok Then
        Dim Error As Dictionary
        Set Error = Response.Data("ResultSet")
        BuildError Error
        Exit Function
    End If
    m_Success = True
    Set Fetch = Response.Data("ResultSet")
End Function

Sub BuildError(Error As Dictionary)
    m_Error.ApiVersion = Error("apiVersion")
    m_Error.EngineVersion = Error("engineVersion")
    Dim e As Dictionary
    Set e = Error("Error")
    m_Error.Code = e("code")
    m_Error.Message = e("Message")
End Sub

Property Get Error() As EkispertError
    Error = m_Error
End Property

Property Get Success() As Boolean
    Success = m_Success
End Property


' 指定したキーが配列ではない場合、配列にする
' 配列の場合は、配列で返す
Sub SetArray(ByRef ary() As Dictionary, Data As Dictionary, Key As String)
On Error GoTo NotArray

    Dim Index As Long
    ReDim ary(Data(Key).Count) As Dictionary
    Dim Params As Dictionary
    
    For Each Params In Data(Key)
        Set ary(Index) = Params
        Index = Index + 1
    Next Params
    Exit Sub
NotArray:
    ' 配列じゃない場合
    ReDim ary(1) As Dictionary
    Set ary(0) = Data(Key)
End Sub

