VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoursePlainQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_From As String
Private m_To As String
Private m_Via As String
Private m_Date As Date
Private m_Plane As Boolean
Private m_Shinkansen As Boolean
Private m_LimitedExpress As Boolean
Private m_Bus As Boolean
Private m_Gcs As GCSEnum

Private m_Client As Ekispert

' Class Initialize method to set default values
Private Sub Class_Initialize()
    m_Date = Now
    m_Shinkansen = True
    m_Plane = True
    m_Bus = True
    m_LimitedExpress = True
    m_Gcs = GCSEnum.Tokyo
End Sub

Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/search/course/plain"
    Dim Query As Dictionary
    Set Query = BuildQuery()
    Dim Result As Dictionary
    Set Result = m_Client.Fetch(ApiPath, Query)
    
    Dim dicCourses() As Dictionary
    Dim Response As ResultSet
    
    If m_Client.Success = False Then
        Response.Success = False
        Response.Error = m_Client.Error
        Find = Response
        Exit Function
    End If
    m_Client.SetArray dicCourses, Result, "Course"
    Response.Courses = m_Client.CreateCourses(dicCourses)
    Response.Success = True
    Find = Response

    ' Debug.Print JsonConverter.ConvertToJson(Result)
End Function

Property Let client(ByVal client As Ekispert)
    Set m_Client = client
End Property


' Property Let for from
Public Property Let FromProp(value As String)
    m_From = value
End Property

' Property Let for to
Public Property Let ToProp(value As String)
    m_To = value
End Property

' Property Let for via
Public Property Let Via(value As String)
    m_Via = value
End Property

' Property Let for date
Public Property Let DateProp(value As Date)
    m_Date = value
End Property

' Property Let for plane
Public Property Let Plane(value As Boolean)
    m_Plane = value
End Property

' Property Let for shinkansen
Public Property Let Shinkansen(value As Boolean)
    m_Shinkansen = value
End Property

' Property Let for limitedExpress
Public Property Let LimitedExpress(value As Boolean)
    m_LimitedExpress = value
End Property

' Property Let for bus
Public Property Let Bus(value As Boolean)
    m_Bus = value
End Property

' Property Let for gcs
Public Property Let Gcs(value As GCSEnum)
    m_Gcs = value
End Property


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    BuildQuery.Add "from", m_From
    BuildQuery.Add "to", m_To
    BuildQuery.Add "date", Format(m_Date, "yyyymmdd")
    BuildQuery.Add "plane", IIf(m_Plane, "true", "false")
    BuildQuery.Add "shinkansen", IIf(m_Shinkansen, "true", "false")
    BuildQuery.Add "limitedExpress", IIf(m_LimitedExpress, "true", "false")
    BuildQuery.Add "bus", IIf(m_Bus, "true", "false")
    If m_Client.GCSString(m_Gcs) <> "" Then
        BuildQuery.Add "gcs", m_Client.GCSString(m_Gcs)
    End If
End Function


