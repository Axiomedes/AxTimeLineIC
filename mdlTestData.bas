Attribute VB_Name = "mdlTestData"
Option Explicit

' Generador de datos de prueba para AxTimeLineIC v2 Showcase
Public Enum NameTypeEnum
    ntRandom = 0
    ntMale = 1
    ntFemale = 2
End Enum

Private Const M_FORENAMES = "Alan,Alfie,Andrew,Ben,Bill,Bob,Boris,Brian,Carlos,David,Gavin,Geoff,Hugo,Ian,James,Jon,Lucas,Mateo,Michael,Pablo,Peter,Richard,Robert,Samuel,Simon,Tomas,Victor"
Private Const F_FORENAMES = "Alicia,Alison,Amanda,Barbara,Camila,Carolina,Elena,Hannah,Hayley,Jane,Julia,Karen,Katie,Laura,Lucia,Maria,Paula,Rachel,Sara,Sofia,Susan,Valeria"
Private Const SURNAMES = "Alvarez,Benitez,Castillo,Diaz,Evans,Fernandez,Garcia,Gomez,Hernandez,Jimenez,Lopez,Martinez,Navarro,Perez,Ramirez,Rodriguez,Sanchez,Torres,Vargas"
Private Const JOBS = "Software Architect,UX/UI Designer,DevOps Lead,Product Manager,Data Scientist,Frontend Engineer,Backend Developer,QA Specialist,Security Analyst,Cloud Specialist"
Private Const EVENTS = "Requerimientos Aprobados,Diseno UI/UX Finalizado,Arquitectura de BD Definida,Sprint 1 Completado,Integracion de API Lista,Pruebas QA Exitosas,Despliegue a Produccion,Auditoria y Cierre"

Private mCalled As Boolean
Private mMF() As String
Private mFF() As String
Private mSurnames() As String
Private mJobs() As String
Private mEvents() As String

Private Sub Initialise()
    If Not mCalled Then
        mCalled = True
        Randomize Timer
        mMF = Split(M_FORENAMES, ",")
        mFF = Split(F_FORENAMES, ",")
        mSurnames = Split(SURNAMES, ",")
        mJobs = Split(JOBS, ",")
        mEvents = Split(EVENTS, ",")
    End If
End Sub

Public Function RandomInt(lowerbound As Long, upperbound As Long) As Long
    RandomInt = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Public Function GetPersonName(Optional nType As NameTypeEnum = ntRandom) As String
    Initialise
    Dim fn As String
    Select Case nType
        Case ntMale
            fn = mMF(RandomInt(0, UBound(mMF)))
        Case ntFemale
            fn = mFF(RandomInt(0, UBound(mFF)))
        Case Else
            If RandomInt(0, 1) = 0 Then
                fn = mMF(RandomInt(0, UBound(mMF)))
            Else
                fn = mFF(RandomInt(0, UBound(mFF)))
            End If
    End Select
    GetPersonName = fn & " " & mSurnames(RandomInt(0, UBound(mSurnames)))
End Function

Public Function GetJobTitle() As String
    Initialise
    GetJobTitle = mJobs(RandomInt(0, UBound(mJobs)))
End Function

Public Function GetMilestone(ByVal Index As Long) As String
    Initialise
    If Index >= 0 And Index <= UBound(mEvents) Then
        GetMilestone = mEvents(Index)
    Else
        GetMilestone = "Hito del Proyecto #" & (Index + 1)
    End If
End Function
