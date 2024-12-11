VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolboxCourseUpdateQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Properties
Private m_Name As String
Private m_StationNameList() As String
Private m_RailNameList() As String
Private m_CorporationNameList() As String
Private m_Date As Date

Private m_Client As Ekispert

Property Let Client(ByVal Client As Ekispert)
    Set m_Client = Client
End Property

' Initialize default values
Private Sub Class_Initialize()
    m_Date = Now
End Sub

Property Let StationNameList(i As Long, ByVal value As String)
    ReDim Preserve m_StationNameList(i) As String
    m_StationNameList(i) = value
End Property

Property Let RailNameList(i As Long, ByVal value As String)
    ReDim Preserve m_RailNameList(i) As String
    m_RailNameList(i) = value
End Property

Property Let CorporationNameList(i As Long, ByVal value As String)
    ReDim Preserve m_CorporationNameList(i) As String
    m_CorporationNameList(i) = value
End Property

Public Property Let DateProp(ByVal value As Date)
    m_Date = value
End Property


Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/toolbox/course/update"
    Dim Result As Dictionary
    Set Result = m_Client.Fetch(ApiPath, BuildQuery())
    
    Dim Response As ResultSet
    
    If m_Client.Success = False Then
        Response.Success = False
        Response.Error = m_Client.Error
        Find = Response
        Exit Function
    End If
    ' Debug.Print JsonConverter.ConvertToJson(Result)
    Dim dicUpdates() As Dictionary
    
    m_Client.SetArray dicUpdates, Result, "Update"
    Response.Updates = m_Client.CreateUpdates(dicUpdates)
    Response.Success = True
    Find = Response
End Function


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    
    BuildQuery.Add "date", Format(m_Date, "yyyymmdd")

    m_Client.ArrayToString BuildQuery, "stationNameList", m_StationNameList
    m_Client.ArrayToString BuildQuery, "railNameList", m_RailNameList
    m_Client.ArrayToString BuildQuery, "corporationNameList", m_CorporationNameList

End Function






