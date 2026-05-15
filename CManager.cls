VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.7
'Site:         http://sandsprite.com
'License:      MIT

Option Explicit

Public Users As New Collection 'As CUser
Public Projects As New Collection 'As CProject

Sub addUser(name, age, job)
    Dim u As New CUser
    u.name = name
    u.age = age
    u.job = job
    Users.Add u
End Sub

Sub addProject(name, leadName, budget)
    Dim p As New CProject
    p.name = name
    p.leadName = leadName
    p.budget = budget
    Projects.Add p
End Sub
