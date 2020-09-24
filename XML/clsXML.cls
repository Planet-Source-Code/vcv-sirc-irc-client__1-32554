VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rootNode As typNodes

Private Type typNode
    TagName As String
    Attributes() As typAttr
    Data As String
    NodeNum As Integer
    Nodes() As typNode
    NodeCount As Integer
    parentNode As typNode
End Type

Private Type typAttr
    Key As String
    Value As String
End Type

Public Sub addNode(parentNode As typNode, addNode As typNode)
    parentNode.NodeCount = parentNode.NodeCount + 1
    ReDim parentNode.Nodes(parentNode.NodeCount) As typNode
    parentNode.Nodes(parentNode.NodeCount) = addNode
    parentNode.Nodes(parentNode.NodeCount).NodeNum = parentNode.NodeCount
End Sub

Public Sub ClearData()
    rootNode.TagName = ""
    ReDim rootNode.Attributes(0) As typAttr
    rootNode.Data = ""
    rootNode.NodeNum = 0
    rootNode.Nodes = 0
    rootNode.NodeCount = 0
End Sub
