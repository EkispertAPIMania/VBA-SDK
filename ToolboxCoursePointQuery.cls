VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolboxCoursePointQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Properties
Private m_Name As String
Private m_StationCode() As String
Private m_Station() As String
Private m_Time() As String
Private m_Fare() As String
Private m_Traffic() As String
Private m_Distance() As String
Private m_Teiki1() As String
Private m_Teiki3() As String
Private m_Teiki6() As String

Private m_Client As Ekispert

Property Let client(ByVal client As Ekispert)
    Set m_Client = client
End Property

' Property Let methods
Property Let Name(value As String)
    m_Name = value
End Property

Property Let StationCode(i As Long, ByVal value As Integer)
    ReDim Preserve m_StationCode(i) As String
    m_StationCode(i) = CStr(value)
End Property

Property Let Station(i As Long, ByVal value As String)
    ReDim Preserve m_Station(i) As String
    m_Station(i) = value
End Property

Property Let Time(i As Long, ByVal value As Integer)
    ReDim Preserve m_Time(i) As String
    m_Time(i) = CStr(value)
End Property

Property Let Fare(i As Long, ByVal value As Integer)
    ReDim Preserve m_Fare(i) As String
    m_Fare(i) = CStr(value)
End Property

Property Let Traffic(i As Long, ByVal value As String)
    ReDim Preserve m_Traffic(i) As String
    m_Traffic(i) = value
End Property

Property Let Distance(i As Long, ByVal value As Integer)
    ReDim Preserve m_Distance(i) As String
    m_Distance(i) = CStr(value)
End Property

Property Let Teiki1(i As Long, ByVal value As Integer)
    ReDim Preserve m_Teiki1(i) As String
    m_Teiki1(i) = CStr(value)
End Property

Property Let Teiki3(i As Long, ByVal value As Integer)
    ReDim Preserve m_Teiki3(i) As String
    m_Teiki3(i) = CStr(value)
End Property

Property Let Teiki6(i As Long, ByVal value As Integer)
    ReDim Preserve m_Teiki6(i) As String
    m_Teiki6(i) = CStr(value)
End Property


' Initialize default values
Private Sub Class_Initialize()
End Sub

Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/toolbox/course/point"
    Dim Result As Dictionary
    Set Result = m_Client.Fetch(ApiPath, BuildQuery())
    
    Dim Response As ResultSet
    
    If m_Client.Success = False Then
        Response.Success = False
        Response.Error = m_Client.Error
        Find = Response
        Exit Function
    End If
    Response.Point = m_Client.CreatePoint(Result("Point"))
    Response.Success = True
    Find = Response

    ' Debug.Print JsonConverter.ConvertToJson(Result)
End Function


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    
    m_Client.ArrayToString BuildQuery, "stationCode", m_StationCode
    m_Client.ArrayToString BuildQuery, "station", m_Station
    m_Client.ArrayToString BuildQuery, "time", m_Time
    m_Client.ArrayToString BuildQuery, "fare", m_Fare
    m_Client.ArrayToString BuildQuery, "traffic", m_Traffic
    m_Client.ArrayToString BuildQuery, "distance", m_Distance
    m_Client.ArrayToString BuildQuery, "teiki1", m_Teiki1
    m_Client.ArrayToString BuildQuery, "teiki3", m_Teiki3
    m_Client.ArrayToString BuildQuery, "teiki6", m_Teiki6
    
    m_Client.StringParam BuildQuery, "name", m_Name

End Function





