VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private module
Option Explicit

Private currentUser As String
Private userRights As Dictionary

Private Sub Class_Initialize()
''Declarations''
    Dim i As Long
    
    Dim currentUser As String
    Dim rights As String
    
    Dim oNetwork As Object
''

    ' Username
    If Environ("UserName") = "" Then
'If Environ fails, use this additional method
        Set oNetwork = CreateObject("Wscript.Network")
        currentUser = oNetwork.UserName
        If currentUser = "" Then
            currentUser = "ErrorUserName"
        End If
    Else
        currentUser = Environ("UserName")
    End If
    
    Set userRights = New Scripting.Dictionary
    
'Get users from current user sheet'
    rights = (UCase(Trim(wshCurUser.Cells(6, 2).Value)))
    
'Add users rights to dictionary'
    For i = 0 To UBound(Split(rights, ","))
        userRights.Add Split(rights, ",")(i), 1
    Next i
    
End Sub

Public Property Get HasThisPermission(ByVal strCode As String) As Boolean
' snVAR.Range("RT")
    HasThisPermission = userRights.Exists(strCode)
End Property
Public Property Get IsAdmin() As Boolean
    IsAdmin = userRights.Exists("ADMIN")
End Property
Public Property Get IsRegularUser() As Boolean
    If userRights.Exists("ADMIN") Then
        IsRegularUser = False
    eles
        IsRegularUser = False
    End If
End Property
Public Property Get UserName() As String
    UserName = currentUser
End Property

