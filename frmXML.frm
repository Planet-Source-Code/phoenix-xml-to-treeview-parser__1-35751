VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmXML 
   Caption         =   "XML To TreeView Parser"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCollapseAll 
      Caption         =   "Collapse All"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExpandAll 
      Caption         =   "Expand All"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin ComctlLib.TreeView tvwXMLDoc 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8493
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      PathSeparator   =   "/"
      Style           =   7
      ImageList       =   "imlIcons"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   1920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmXML.frx":0000
            Key             =   "attribute"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmXML.frx":0552
            Key             =   "comment"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmXML.frx":0AA4
            Key             =   "node"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmXML.frx":0FF6
            Key             =   "text"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmXML.frx":1548
            Key             =   "procinstr"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open File..."
      End
      Begin VB.Menu mnuFileOpenURL 
         Caption         =   "Open &URL..."
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare XML node collection
Public Nodes As Collection

'Declare Node count
Public NodeNumber As Long
'Declare Text node count
Public TextNumber As Long
'Declare Attribute node count
Public AttributeNumber As Long
'Declare Comment node count
Public CommentNumber As Long

'Declare XML Document variable
Public WithEvents XMLDoc As DOMDocument30
Attribute XMLDoc.VB_VarHelpID = -1

'***********************************************
'* Purpose: Collapse all nodes in the TreeView *
'***********************************************
Private Sub cmdCollapseAll_Click()
    Dim Counter As Long
    
    'Go through every node and set their Expanded
    'property to False
    For Counter = 1 To tvwXMLDoc.Nodes.Count
        tvwXMLDoc.Nodes(Counter).Expanded = False
    Next
End Sub

'*********************************************
'* Purpose: Expand all nodes in the TreeView *
'*********************************************
Private Sub cmdExpandAll_Click()
    Do
        'Update the amount of current nodes
        'before expansion occurs
        Dim CurrentNodeCount As Long
        CurrentNodeCount = tvwXMLDoc.Nodes.Count
        
        'Go through every node and set their Expanded
        'property to True
        Dim Counter As Long
        For Counter = 1 To tvwXMLDoc.Nodes.Count
            tvwXMLDoc.Nodes(Counter).Expanded = True
        Next
        
        'As long as there are more nodes after expansion
        'than before, we must loop the process to assure
        'that the new nodes are expanded too.
    Loop Until tvwXMLDoc.Nodes.Count = CurrentNodeCount
End Sub

'****************************************
'* Purpose: Perform pre-view operations *
'****************************************
Private Sub Form_Load()
    'Show user that no file is currently loaded
    '(This isn't necessary, but it looks better than an
    'large, empty box, in my opinion)
    tvwXMLDoc.Nodes.Add , , , "<No XML Document is currently loaded"
End Sub

'*********************************************************
'* Purpose: Resize and reposition the controls as needed *
'*********************************************************
Private Sub Form_Resize()
    'Resize and reposition the control as needed
    On Error Resume Next
    cmdExpandAll.Move 120, 120
    tvwXMLDoc.Move 120, cmdExpandAll.top + cmdExpandAll.Height + 120, Width - 360, Height - cmdExpandAll.Height - 240 - 960
    cmdCollapseAll.Move tvwXMLDoc.left + tvwXMLDoc.Width - cmdCollapseAll.Width, 120
End Sub

'***************************************
'* Purpose: Clear out global variables *
'***************************************
Private Sub Form_Unload(Cancel As Integer)
    'Empty out the Nodes Collection
    Set Nodes = Nothing
    'Empty out the XMLDoc object
    Set XMLDoc = Nothing
End Sub

'*******************************
'* Purpose: Unload the program *
'*******************************
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'***************************************************
'* Purpose: Load an XML document into the TreeView *
'***************************************************
Private Sub mnuFileOpen_Click()
    'Have user open an XML Document
    Dim File As SelectedFile
    File = ShowOpen(Me.hWnd, True)
    
    'If user cancelled, cancel loading
    If File.bCanceled Then Exit Sub
    
    'Put path in a variable
    Dim Path
    Path = File.sLastDirectory & File.sFiles(1)
    
    'Open XML Document
    Set XMLDoc = New DOMDocument30
    XMLDoc.async = False
    XMLDoc.Load Path
    
    'Check if any parseError occured
    If XMLDoc.parseError.errorCode <> 0 Then
        'Error occured, tell the user what's wrong.
        MsgBox "Error occured when parsing the XML file" & vbCrLf _
             & vbCrLf _
             & "Error Code: " & XMLDoc.parseError.errorCode & vbCrLf _
             & "Line Number: " & XMLDoc.parseError.Line & vbCrLf _
             & "Line Position: " & XMLDoc.parseError.linepos & vbCrLf _
             & "Reason: " & XMLDoc.parseError.reason & vbCrLf _
             , vbExclamation, "XML Parse Error"
        'Cancel loading
        Exit Sub
    End If
    
    'Reset global variables
    tvwXMLDoc.Nodes.Clear
    Set Nodes = Nothing
    Set Nodes = New Collection
    NodeNumber = 0
    TextNumber = 0
    AttributeNumber = 0
    CommentNumber = 0
    
    'Start parsing bottom-level nodes
    Dim XMLNode As IXMLDOMNode
    For Each XMLNode In XMLDoc.childNodes
        'Check what kind of node we found
        If XMLNode.nodeType = 7 Then
            'Processing Instruction node found, add one to number of nodes
            NodeNumber = NodeNumber + 1
            
            'Add the node to the node collection
            Nodes.Add XMLNode, "n" & NodeNumber
            
            'Add the node to the TreeView
            tvwXMLDoc.Nodes.Add , , "n" & NodeNumber, XMLNode.nodeName, "procinstr"
            
            'Start AddAttributes Sub
            AddAttributes "n" & NodeNumber
            
            'Start AddChildren Sub
            AddChildren "n" & NodeNumber
        ElseIf XMLNode.nodeName = "#comment" Then
            'Comment found, add one to the number of comments
            CommentNumber = CommentNumber + 1
            
            'Add the comment to the TreeView
            tvwXMLDoc.Nodes.Add , , "c" & CommentNumber, XMLNode.Text, "comment"
        Else
            'Normal node found, add one to number of nodes
            NodeNumber = NodeNumber + 1
            
            'Add the node to the node collection
            Nodes.Add XMLNode, "n" & NodeNumber
            
            'Add the node to the TreeView
            tvwXMLDoc.Nodes.Add , , "n" & NodeNumber, XMLNode.nodeName, "node"
            
            'Start AddAttributes Sub
            AddAttributes "n" & NodeNumber
            
            'Start AddChildren Sub
            AddChildren "n" & NodeNumber
        End If
    Next
    
    Set XMLNode = Nothing
End Sub

'*********************************************************
'* Purpose: Add all attributes of a node to the TreeView
'* Inputs:
'*      NodeKey:    The key to the parent we're adding
'*                  attributes to.
'*********************************************************
Private Sub AddAttributes(NodeKey As String)
    'Collect the node from the node collection using the
    'passed NodeKey
    Dim XMLNode As IXMLDOMNode
    Set XMLNode = Nodes(NodeKey)
        
    'For some (VERY ANNOYING) reason, the MSXML parser can only
    'read one processing instruction correctly, the topmost in
    'a document, that's called "xml". So If it's something else,
    'we must stop this, else we'll get an error
    If XMLNode.nodeType = 7 Then
        If XMLNode.nodeName <> "xml" Or XMLNode.parentNode.nodeName <> "#document" Then
            Exit Sub
        End If
    End If
    
    'Go through all of the node's attributes, and add them
    Dim Counter As Long
    For Counter = 0 To XMLNode.Attributes.length - 1
        'Attribute Node found, add one to the number of attributes
        AttributeNumber = AttributeNumber + 1
        
        'Add the node to the node collection
        Nodes.Add XMLNode, "a" & AttributeNumber
        'Add the node to the TreeView
        tvwXMLDoc.Nodes.Add NodeKey, tvwChild, "a" & AttributeNumber, XMLNode.Attributes(Counter).nodeName, "attribute"
        
        'Since attributes are supposed to contain text, let's
        'add that too, while we're at it
        
        'Add one to the number of text nodes
        TextNumber = TextNumber + 1
        
        'Add the node to the TreeView
        tvwXMLDoc.Nodes.Add "a" & AttributeNumber, tvwChild, "t" & TextNumber, XMLNode.Attributes(Counter).Text, "text"
    Next
    
    'Empty out the XMLNode object
    Set XMLNode = Nothing
End Sub

'*********************************************************
'* Purpose: Add all children nodes of a node to the
'*          TreeView
'* Inputs:
'*      NodeKey:    The key to the parent we're adding
'*                  child nodes to.
'*********************************************************
Private Sub AddChildren(NodeKey As String)
    'Collect the node from the node collection using the
    'passed NodeKey
    Dim XMLNode As IXMLDOMNode
    Set XMLNode = Nodes(NodeKey)
    
    'Go through all of the node's children, and add them
    Dim Counter As Long
    For Counter = 0 To XMLNode.childNodes.length - 1
        'See what kind of node we've got
        If XMLNode.childNodes(Counter).nodeName = "#text" Then
            'Text node, add one to the number of text nodes
            TextNumber = TextNumber + 1
            
            'Add the node to the TreeView
            tvwXMLDoc.Nodes.Add NodeKey, tvwChild, "t" & TextNumber, XMLNode.childNodes(Counter).Text, "text"
        ElseIf XMLNode.childNodes(Counter).nodeName = "#comment" Then
            'Comment found, add one to the number of comments
            CommentNumber = CommentNumber + 1
            
            'Add the comment to the TreeView
            tvwXMLDoc.Nodes.Add NodeKey, tvwChild, "c" & CommentNumber, XMLNode.childNodes(Counter).Text, "comment"
        Else
            'Normal node found, add one to the number of nodes
            NodeNumber = NodeNumber + 1
            
            'Add the node to the node collection
            Nodes.Add XMLNode.childNodes(Counter), "n" & NodeNumber
            'Add the node to the TreeView
            tvwXMLDoc.Nodes.Add NodeKey, tvwChild, "n" & NodeNumber, XMLNode.childNodes(Counter).nodeName, "node"
        End If
    Next
    
    Set XMLNode = Nothing
End Sub

'***************************************************
'* Purpose: Load an XML document from web into the *
'*          TreeView                               *
'***************************************************
Private Sub mnuFileOpenURL_Click()
    'Have user input an URL
    Dim URL As String
    URL = InputBox("Enter the URL to open", "Open URL...")
    
    'If nothing were entered, exit sub
    If URL = "" Then Exit Sub
    
    'Open XML Document
    Set XMLDoc = New DOMDocument30
    XMLDoc.async = False
    XMLDoc.Load URL
    
    'Check if any parseError occured
    If XMLDoc.parseError.errorCode <> 0 Then
        'Error occured, tell the user what's wrong.
        MsgBox "Error occured when parsing the XML file" & vbCrLf _
             & vbCrLf _
             & "Error Code: " & XMLDoc.parseError.errorCode & vbCrLf _
             & "Line Number: " & XMLDoc.parseError.Line & vbCrLf _
             & "Line Position: " & XMLDoc.parseError.linepos & vbCrLf _
             & "Reason: " & XMLDoc.parseError.reason & vbCrLf _
             , vbExclamation, "XML Parse Error"
        'Cancel loading
        Exit Sub
    End If
    
    'Reset global variables
    tvwXMLDoc.Nodes.Clear
    Set Nodes = Nothing
    Set Nodes = New Collection
    NodeNumber = 0
    TextNumber = 0
    AttributeNumber = 0
    CommentNumber = 0
    
    'Start parsing bottom-level nodes
    Dim XMLNode As IXMLDOMNode
    For Each XMLNode In XMLDoc.childNodes
        'Check what kind of node we found
        If XMLNode.nodeType = 7 Then
            'Processing Instruction node found, add one to number of nodes
            NodeNumber = NodeNumber + 1
            
            'Add the node to the node collection
            Nodes.Add XMLNode, "n" & NodeNumber
            
            'Add the node to the TreeView
            tvwXMLDoc.Nodes.Add , , "n" & NodeNumber, XMLNode.nodeName, "procinstr"
            
            'Start AddAttributes Sub
            AddAttributes "n" & NodeNumber
            
            'Start AddChildren Sub
            AddChildren "n" & NodeNumber
        ElseIf XMLNode.nodeName = "#comment" Then
            'Comment found, add one to the number of comments
            CommentNumber = CommentNumber + 1
            
            'Add the comment to the TreeView
            tvwXMLDoc.Nodes.Add , , "c" & CommentNumber, XMLNode.Text, "comment"
        Else
            'Normal node found, add one to number of nodes
            NodeNumber = NodeNumber + 1
            
            'Add the node to the node collection
            Nodes.Add XMLNode, "n" & NodeNumber
            
            'Add the node to the TreeView
            tvwXMLDoc.Nodes.Add , , "n" & NodeNumber, XMLNode.nodeName, "node"
            
            'Start AddAttributes Sub
            AddAttributes "n" & NodeNumber
            
            'Start AddChildren Sub
            AddChildren "n" & NodeNumber
        End If
    Next
    
    Set XMLNode = Nothing
End Sub

'*********************************************************
'* Purpose: Dynamically add new nodes as they are needed *
'*          when the user expands their parent node      *
'*********************************************************
Private Sub tvwXMLDoc_Expand(ByVal Node As ComctlLib.Node)
    'Check if this node already has been expanded before,
    'in that case, we don't need to give it's children
    'attributes and children again.
    If Not Node.Tag = "Expanded" Then
        Dim ChildNode As ComctlLib.Node
        Dim Counter As Long
        
        'Check if the node has got children
        If Node.children <> 0 Then
            'Set Childnode to it's first child
            Set ChildNode = Node.Child
            
            'For each child, add the attributes and it's children
            For Counter = 1 To Node.children
                'If the node is not a regular node, don't add children to it.
                'Only regular nodes can have children
                If left(ChildNode.Key, 1) = "n" Then
                    'Start AddAttributes Sub
                    AddAttributes ChildNode.Key
                    'Start AddChildren Sub
                    AddChildren ChildNode.Key
                End If
                
                'Set ChildNode to the next child
                Set ChildNode = ChildNode.Next
            Next
        End If
        'Tag the node as expanded, so it won't be expanded again.
        Node.Tag = "Expanded"
    End If
End Sub
