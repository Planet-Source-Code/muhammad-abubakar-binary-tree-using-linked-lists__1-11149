VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------
'Binary Tree class by Muhammad Abubakar
'                       <joehacker@yahoo.com>
'                       <http://go.to/abubakar>
'------------------------------------------------

'Its a small class but if you understand it, its going to clear lot
'of your concepts about objects of Visual Basic and of Data structures.
Option Explicit

Private t As Integer
Private Parent As Node, Tmp As Node
Private m_List As ListBox
Public Sub Insert(x As Integer)
    Set Tmp = New Node
    Tmp.d = x
    Set Tmp.Left = Nothing
    Set Tmp.Right = Nothing
    t = x
    If Parent Is Nothing Then
        Set Parent = Tmp
    Else
        Recurse Parent
    End If
End Sub

'if the given data is less then current node data then it'll become its left
'node, else it'll become its right node

Private Sub Recurse(tmpx As Node)
    If t >= tmpx.d Then
        If Not (tmpx.Right Is Nothing) Then
            Recurse tmpx.Right
        Else
            Set tmpx.Right = Tmp
        End If
    Else
        If Not (tmpx.Left Is Nothing) Then
            Recurse tmpx.Left
        Else
            Set tmpx.Left = Tmp
        End If
    End If
End Sub
Public Sub EnumThem()
    m_List.Clear
    RecEnum Parent
End Sub

'first visit left then head then right
Private Sub RecEnum(tmpx As Node)
    If Not (tmpx.Left Is Nothing) Then RecEnum tmpx.Left
    m_List.AddItem tmpx.d
    If Not (tmpx.Right Is Nothing) Then RecEnum tmpx.Right
End Sub

Public Property Get SetListBox() As ListBox
    Set SetListBox = m_List
End Property

Public Property Let SetListBox(ByVal vNewValue As ListBox)
    Set m_List = vNewValue
End Property
Public Sub DeleteAll()
    RecursiveDelete Parent
End Sub
'While deleting, first delete left then right then head
Private Sub RecursiveDelete(tmpx As Node)
    If Not (tmpx.Left Is Nothing) Then RecursiveDelete tmpx.Left
    If Not (tmpx.Right Is Nothing) Then RecursiveDelete tmpx.Right
    Set tmpx = Nothing
End Sub

