VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchMultipleRangeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_BaseList() As String
Private m_UpperMinute() As String
Private m_UpperTransferCount() As String
Private m_Plane As Boolean
Private m_Shinkansen As Boolean
Private m_LimitedExpress As Boolean
Private m_WaitAverageTime As Boolean
Private m_Limit As Integer
Private m_Date As Date

Private m_Client As Ekispert

' Class Initialize method to set default values
Private Sub Class_Initialize()
    m_Plane = True ' Default value for Plane
    m_Shinkansen = True ' Default value for Shinkansen
    m_LimitedExpress = True ' Default value for LimitedExpress
    m_WaitAverageTime = True ' Default value for WaitAverageTime
    m_Date = Now
End Sub

Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/search/multipleRange"
    Dim Query As Dictionary
    Set Query = BuildQuery()
    Dim Result As Dictionary
    Set Result = m_Client.Fetch(ApiPath, Query)
    
    Dim dicPoints() As Dictionary
    Dim Response As ResultSet
    
    If m_Client.Success = False Then
        Response.Success = False
        Response.Error = m_Client.Error
        Find = Response
        Exit Function
    End If
    If Result.Exists("Point") = True Then
        m_Client.SetArray dicPoints, Result, "Point"
        Response.Points = m_Client.CreatePoints(dicPoints)
    End If
    Dim dicBases() As Dictionary
    If Result.Exists("Base") = True Then
        m_Client.SetArray dicBases, Result, "Base"
        Response.Bases = m_Client.CreateBases(dicBases)
    End If
    
    Response.Success = True
    Find = Response

    ' Debug.Print JsonConverter.ConvertToJson(Result)
End Function


Property Let Client(ByVal Client As Ekispert)
    Set m_Client = Client
End Property


' Property Let for BaseList
Public Property Let BaseList(i As Long, ByVal value As String)
    ReDim Preserve m_BaseList(i) As String
    m_BaseList(i) = value
End Property

Public Property Let UpperMinute(i As Long, ByVal value As Integer)
    ReDim Preserve m_UpperMinute(i) As String
    m_UpperMinute(i) = CStr(value)
End Property

' Property Let for UpperTransferCount
Public Property Let UpperTransferCount(i As Long, ByVal value As Integer)
    ReDim Preserve m_UpperTransferCount(i) As String
    If value = -1 Then
        m_UpperTransferCount(i) = Str("")
    Else
        m_UpperTransferCount(i) = CStr(value)
    End If
End Property

' Property Let for Plane
Public Property Let Plane(value As Boolean)
    m_Plane = value
End Property

' Property Let for Shinkansen
Public Property Let Shinkansen(value As Boolean)
    m_Shinkansen = value
End Property

' Property Let for LimitedExpress
Public Property Let LimitedExpress(value As Boolean)
    m_LimitedExpress = value
End Property

' Property Let for WaitAverageTime
Public Property Let WaitAverageTime(value As Boolean)
    m_WaitAverageTime = value
End Property

' Property Let for Limit
Public Property Let Limit(value As Integer)
    m_Limit = value
End Property

' Property Let for Date
Public Property Let DateProp(value As Date)
    m_Date = value
End Property

Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    BuildQuery.Add "date", Format(m_Date, "yyyymmdd")
    BuildQuery.Add "plane", IIf(m_Plane, "true", "false")
    BuildQuery.Add "shinkansen", IIf(m_Shinkansen, "true", "false")
    BuildQuery.Add "limitedExpress", IIf(m_LimitedExpress, "true", "false")
    BuildQuery.Add "waitAverageTime", IIf(m_WaitAverageTime, "true", "false")
    If m_Limit > 0 Then
        BuildQuery.Add "limit", m_Limit
    End If
    m_Client.ArrayToString BuildQuery, "baseList", m_BaseList
    m_Client.ArrayToString BuildQuery, "upperMinute", m_UpperMinute
    m_Client.ArrayToString BuildQuery, "upperTransferCount", m_UpperTransferCount

End Function

