VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimePoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================================
' AxTimeLineIC v2 - Coleccion de Hitos (clsTimePoints)
'========================================================================================

Private m_Col As Collection

Private Sub Class_Initialize()
    Set m_Col = New Collection
End Sub

Private Sub Class_Terminate()
    Clear
    Set m_Col = Nothing
End Sub

Public Function Add(ByVal Title As String, _
                    Optional ByVal Subtitle As String = vbNullString, _
                    Optional ByVal Timestamp As String = vbNullString, _
                    Optional ByVal DateText As String = vbNullString, _
                    Optional ByVal IconChar As Variant = vbNullString, _
                    Optional ByVal Picture As Variant = Empty, _
                    Optional ByVal Status As tlItemStatus = tlStatusNone, _
                    Optional ByVal Key As String = vbNullString) As clsTimePoint
    
    Dim item As New clsTimePoint
    With item
        .Title = Title
        .Subtitle = Subtitle
        .Timestamp = Timestamp
        .DateText = DateText
        .IconChar = IconChar
        If Not IsEmpty(Picture) Then .Picture = Picture
        .Status = Status
        .Key = Key
    End With
    
    If Len(Key) > 0 Then
        m_Col.Add item, Key
    Else
        m_Col.Add item
    End If
    
    Set Add = item
End Function

Public Function Insert(ByVal Index As Long, _
                       ByVal Title As String, _
                       Optional ByVal Subtitle As String = vbNullString, _
                       Optional ByVal Timestamp As String = vbNullString, _
                       Optional ByVal DateText As String = vbNullString, _
                       Optional ByVal IconChar As Variant = vbNullString, _
                       Optional ByVal Picture As Variant = Empty, _
                       Optional ByVal Status As tlItemStatus = tlStatusNone, _
                       Optional ByVal Key As String = vbNullString) As clsTimePoint
    
    Dim item As New clsTimePoint
    With item
        .Title = Title
        .Subtitle = Subtitle
        .Timestamp = Timestamp
        .DateText = DateText
        .IconChar = IconChar
        If Not IsEmpty(Picture) Then .Picture = Picture
        .Status = Status
        .Key = Key
    End With
    
    If Index >= 1 And Index <= m_Col.Count Then
        If Len(Key) > 0 Then
            m_Col.Add item, Key, Index
        Else
            m_Col.Add item, , Index
        End If
    Else
        If Len(Key) > 0 Then
            m_Col.Add item, Key
        Else
            m_Col.Add item
        End If
    End If
    
    Set Insert = item
End Function

Public Sub Remove(ByVal IndexOrKey As Variant)
    On Error Resume Next
    Dim item As clsTimePoint
    Set item = m_Col(IndexOrKey)
    If Not item Is Nothing Then
        item.ClearImageHandle
    End If
    m_Col.Remove IndexOrKey
End Sub

Public Sub Clear()
    Dim i As Long
    For i = 1 To m_Col.Count
        On Error Resume Next
        m_Col(i).ClearImageHandle
    Next i
    Set m_Col = New Collection
End Sub

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Public Property Get Item(ByVal IndexOrKey As Variant) As clsTimePoint
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = m_Col(IndexOrKey)
End Property

Public Function Exists(ByVal Key As String) As Boolean
    On Error Resume Next
    Dim obj As Object
    Set obj = m_Col(Key)
    Exists = (Err.Number = 0 And Not obj Is Nothing)
End Function

Public Function IndexOf(ByVal Key As String) As Long
    Dim i As Long
    For i = 1 To m_Col.Count
        If m_Col(i).Key = Key Then
            IndexOf = i
            Exit Function
        End If
    Next i
    IndexOf = 0
End Function
