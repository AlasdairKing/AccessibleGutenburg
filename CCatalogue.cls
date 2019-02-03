VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCatalogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCatalogue As DOMDocument30

Private Sub Class_Initialize()
    On Error Resume Next
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mCatalogue = Nothing
End Sub

Public Sub LoadCatalogueFile(filename As String)
    On Error Resume Next
    Set mCatalogue = New DOMDocument30
    mCatalogue.async = False
    Call mCatalogue.Load(filename)
    If mCatalogue.parseError.errorCode = 0 Then
        'loaded catalogue okay
    Else
        Set mCatalogue = Nothing
    End If
End Sub

Public Sub AddBook(name As String, title As String, path As String)
    On Error Resume Next
    Dim newBook As IXMLDOMNode
    
    If mCatalogue Is Nothing Then
        'nope, haven't got a valid open catalogue, do nothing
    Else
        Set newBook = mCatalogue.createNode(NODE_ELEMENT, "book", "")
        Call newBook.appendChild(mCatalogue.createNode(NODE_ELEMENT, "name", ""))
        newBook.selectSingleNode("name").Text = name
        Call newBook.appendChild(mCatalogue.createNode(NODE_ELEMENT, "title", ""))
        newBook.selectSingleNode("title").Text = name
        Call newBook.appendChild(mCatalogue.createNode(NODE_ELEMENT, "path", ""))
        newBook.selectSingleNode("path").Text = name
    End If
End Sub


