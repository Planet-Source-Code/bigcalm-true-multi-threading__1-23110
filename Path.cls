VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CoOrdinate"
Attribute VB_Ext_KEY = "Member0" ,"CoOrdinate"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mvarTotalMoveCost As Single
Public Property Get TotalMoveCost() As Single
    TotalMoveCost = mvarTotalMoveCost
End Property
Public Property Let TotalMoveCost(ByVal vData As Single)
    mvarTotalMoveCost = vData
End Property

Public Function Add(Key As String, X As Long, Y As Long, Optional sKey As String, Optional pAtEnd As Boolean = False) As CoOrdinate
    Dim objNewMember As New CoOrdinate
    
    objNewMember.Key = Key
    objNewMember.X = X
    objNewMember.Y = Y
    If Len(sKey) = 0 Then
        If pAtEnd = True Then
            mCol.Add objNewMember, , 1
        Else
            mCol.Add objNewMember
        End If
    Else
        If pAtEnd = True Then
            mCol.Add objNewMember, sKey, 1
        Else
            mCol.Add objNewMember, sKey
        End If
    End If
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As CoOrdinate
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
    Set mCol = New Collection
    TotalMoveCost = 0
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

