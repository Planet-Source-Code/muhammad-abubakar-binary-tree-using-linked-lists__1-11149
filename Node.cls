VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------
'Node class by Muhammad Abubakar
'                       <joehacker@yahoo.com>
'                       <http://go.to/abubakar>
'------------------------------------------------

'This is serving as a "Structure" for the Tree class
'I have made this tree to store only integers for the pupose of
'understanding only, but its truely easy if you want to add names
'strings in it, just add a string variable to this Node class below.
'Similarly you can add as many details( datatypes ) in this Node
'as you want to making a quite complex and flexible data storage
'mechanism.

Option Explicit

Public d As Integer, Left As Node, Right As Node
'Left & Right are self-referencing there own Node class.
