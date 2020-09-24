VERSION 5.00
Begin VB.UserControl Split 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   ControlContainer=   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   6615
   Begin VB.PictureBox picSplitter 
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   2790
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4965
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   -135
      Width           =   45
   End
End
Attribute VB_Name = "Split"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim MoveSplitter As Boolean
Dim OldX As Long
Dim OldY As Long
'Default Property Values:
Const m_def_SplitterLeftTop = 0
Const m_def_SplitterWidthHeight = 45
Const m_def_MinimumPaneWidth = 0
Const m_def_SplitColorMouseUp = 0
Const m_def_ResizeContinuos = 0
Const m_def_ControlLeftPane = ""
Const m_def_ControlRightPane = ""
Const m_def_SplitColorMouseDown = &H808080
Const m_def_Orientation = 0
'Property Variables:
Dim m_SplitterLeftTop As Long
Dim m_SplitterWidthHeight As Long
Dim m_MinimumPaneWidth As Long
Dim m_SplitColorMouseUp As OLE_COLOR
Dim m_ResizeContinuos As Boolean
Dim m_ControlLeftPane As String
Dim m_ControlRightPane As String
Dim m_SplitColorMouseDown As OLE_COLOR
Dim m_Orientation As Orientation

Public Enum Orientation
    vertical = 0
    horizontal = 1
End Enum

Public Event PaneEndResize()

Private Sub SetOrientation()

Select Case m_Orientation
 
 Case horizontal
  picSplitter.MousePointer = vbSizeNS
  picSplitter.Top = UserControl.Height / 2
  picSplitter.Height = m_SplitterWidthHeight
  picSplitter.Left = 0
  picSplitter.Width = UserControl.Width
 
 Case vertical
  picSplitter.MousePointer = vbSizeWE
  picSplitter.Left = UserControl.Width / 2
  picSplitter.Width = m_SplitterWidthHeight
  picSplitter.Top = 0
  picSplitter.Height = UserControl.Height
  
End Select
End Sub

Public Property Get Orientation() As Orientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As Orientation)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
    SetOrientation
    SizePanes
End Property

Private Sub FindLeftAndRightPanes(LeftPane As Object, RightPane As Object)
Dim ctl As Object
  
  Set LeftPane = Nothing
  Set RightPane = Nothing
  
  For Each ctl In UserControl.ContainedControls
   If LCase(ctl.Name) = LCase(m_ControlLeftPane) Then Set LeftPane = ctl
   If LCase(ctl.Name) = LCase(m_ControlRightPane) Then Set RightPane = ctl
  Next ctl
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
 OldX = X
 OldY = Y
 MoveSplitter = True
 picSplitter.BackColor = SplitColorMouseDown
 picSplitter.ZOrder 0
End If
End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveSplitter = True Then
  If m_Orientation = vertical Then
   picSplitter.Left = picSplitter.Left - OldX + X
   If ResizeContinuos = True Then SizePanes
  End If
  
  If m_Orientation = horizontal Then
   picSplitter.Top = picSplitter.Top - OldY + Y
   If ResizeContinuos = True Then SizePanes
  End If
End If
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveSplitter = False
 picSplitter.BackColor = SplitColorMouseUp
 If ResizeContinuos = False Then SizePanes
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SplitColorMouseDown() As OLE_COLOR
    SplitColorMouseDown = m_SplitColorMouseDown
End Property

Public Property Let SplitColorMouseDown(ByVal New_SplitColorMouseDown As OLE_COLOR)
    m_SplitColorMouseDown = New_SplitColorMouseDown
    PropertyChanged "SplitColorMouseDown"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SplitColorMouseDown = m_def_SplitColorMouseDown
    m_ControlLeftPane = m_def_ControlLeftPane
    m_ControlRightPane = m_def_ControlRightPane
    m_ResizeContinuos = m_def_ResizeContinuos
    m_SplitColorMouseUp = m_def_SplitColorMouseUp
    m_Orientation = m_def_Orientation
    m_MinimumPaneWidth = m_def_MinimumPaneWidth
    m_SplitterWidthHeight = m_def_SplitterWidthHeight
    m_SplitterLeftTop = m_def_SplitterLeftTop
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_SplitColorMouseDown = PropBag.ReadProperty("SplitColorMouseDown", m_def_SplitColorMouseDown)
    m_ControlLeftPane = PropBag.ReadProperty("ControlLeftPane", m_def_ControlLeftPane)
    m_ControlRightPane = PropBag.ReadProperty("ControlRightPane", m_def_ControlRightPane)
    m_ResizeContinuos = PropBag.ReadProperty("ResizeContinuos", m_def_ResizeContinuos)
    m_SplitColorMouseUp = PropBag.ReadProperty("SplitColorMouseUp", m_def_SplitColorMouseUp)
    picSplitter.BackColor = m_SplitColorMouseUp
    
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_SplitterWidthHeight = PropBag.ReadProperty("SplitterWidthHeight", m_def_SplitterWidthHeight)
    m_SplitterLeftTop = PropBag.ReadProperty("SplitterLeftTop", m_def_SplitterLeftTop)
    
    picSplitter.Left = PropBag.ReadProperty("SplitterLeft", 0)
    m_MinimumPaneWidth = PropBag.ReadProperty("MinimumPaneWidth", m_def_MinimumPaneWidth)
    
    SetOrientation

End Sub

Private Sub SizePanes()
Dim FoundCtlLeftPane As Object
Dim FoundCtlRightPane As Object

On Error Resume Next

 If m_Orientation = vertical Then
  picSplitter.Top = 0
  picSplitter.Height = UserControl.Height
    
  If m_MinimumPaneWidth * 2 + m_SplitterWidthHeight > UserControl.Width Then UserControl.Width = m_MinimumPaneWidth * 2 + m_SplitterWidthHeight
      
  FindLeftAndRightPanes FoundCtlLeftPane, FoundCtlRightPane
  
  If picSplitter.Left < m_MinimumPaneWidth Then
   picSplitter.Left = m_MinimumPaneWidth
   m_SplitterLeftTop = m_MinimumPaneWidth
   
   FoundCtlLeftPane.Refresh
   FoundCtlRightPane.Refresh
  End If
  
  If picSplitter.Left + picSplitter.Width + m_MinimumPaneWidth > UserControl.Width Then
   picSplitter.Left = UserControl.Width - picSplitter.Width - m_MinimumPaneWidth
   m_SplitterLeftTop = UserControl.Width - picSplitter.Width - m_MinimumPaneWidth
      
   FoundCtlLeftPane.Refresh
   FoundCtlRightPane.Refresh
  End If
 
  FoundCtlLeftPane.Move 0, 0
  FoundCtlLeftPane.Width = picSplitter.Left
  FoundCtlLeftPane.Height = UserControl.Height
  
  FoundCtlRightPane.Move picSplitter.Left + picSplitter.Width, 0
  FoundCtlRightPane.Width = UserControl.Width - picSplitter.Left - picSplitter.Width
  FoundCtlRightPane.Height = UserControl.Height
  m_SplitterLeftTop = picSplitter.Left
 End If

 If m_Orientation = horizontal Then
  picSplitter.Left = 0
  picSplitter.Width = UserControl.Width
  
  If m_MinimumPaneWidth * 2 + m_SplitterWidthHeight > UserControl.Height Then UserControl.Height = m_MinimumPaneWidth * 2 + m_SplitterWidthHeight
  
  FindLeftAndRightPanes FoundCtlLeftPane, FoundCtlRightPane
  
  If picSplitter.Top < m_MinimumPaneWidth Then
   picSplitter.Top = m_MinimumPaneWidth
   m_SplitterLeftTop = m_MinimumPaneWidth
   
   FoundCtlLeftPane.Refresh
   FoundCtlRightPane.Refresh
  End If
  
  If picSplitter.Top + picSplitter.Height + m_MinimumPaneWidth > UserControl.Height Then
   picSplitter.Top = UserControl.Height - picSplitter.Height - m_MinimumPaneWidth
   m_SplitterLeftTop = UserControl.Height - picSplitter.Height - m_MinimumPaneWidth
   
   FoundCtlLeftPane.Refresh
   FoundCtlRightPane.Refresh
  End If
 
  FoundCtlLeftPane.Move 0, 0
  FoundCtlLeftPane.Width = UserControl.Width
  FoundCtlLeftPane.Height = picSplitter.Top
  
  FoundCtlRightPane.Move 0, picSplitter.Top + picSplitter.Height
  FoundCtlRightPane.Width = UserControl.Width
  FoundCtlRightPane.Height = UserControl.Height - picSplitter.Top - picSplitter.Height
  m_SplitterLeftTop = picSplitter.Top
 End If
       
       
 RaiseEvent PaneEndResize
 On Error GoTo 0
End Sub

Private Sub UserControl_Resize()
 SizePanes
End Sub

Private Sub UserControl_Show()
 picSplitter.BackColor = m_SplitColorMouseUp
 picSplitter.ZOrder 0
 SizePanes
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SplitColorMouseDown", m_SplitColorMouseDown, m_def_SplitColorMouseDown)
    Call PropBag.WriteProperty("ControlLeftPane", m_ControlLeftPane, m_def_ControlLeftPane)
    Call PropBag.WriteProperty("ControlRightPane", m_ControlRightPane, m_def_ControlRightPane)
    Call PropBag.WriteProperty("ResizeContinuos", m_ResizeContinuos, m_def_ResizeContinuos)
    Call PropBag.WriteProperty("SplitterWidth", picSplitter.Width, 45)
    Call PropBag.WriteProperty("SplitColorMouseUp", m_SplitColorMouseUp, m_def_SplitColorMouseUp)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("SplitterLeft", picSplitter.Left, 0)
    Call PropBag.WriteProperty("MinimumPaneWidth", m_MinimumPaneWidth, m_def_MinimumPaneWidth)
    Call PropBag.WriteProperty("SplitterWidthHeight", m_SplitterWidthHeight, m_def_SplitterWidthHeight)
    Call PropBag.WriteProperty("SplitterLeftTop", m_SplitterLeftTop, m_def_SplitterLeftTop)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ControlLeftPane() As String
    ControlLeftPane = m_ControlLeftPane
End Property

Public Property Let ControlLeftPane(ByVal New_ControlLeftPane As String)
 Dim ctl As Object
 Dim FoundCtl As Object
 Dim bControlFound As Boolean
 
 For Each ctl In UserControl.ContainedControls
  If LCase(ctl.Name) = LCase(New_ControlLeftPane) Then
   bControlFound = True
   Set FoundCtl = ctl
   Exit For
  End If
 Next ctl
 
 If bControlFound = False Then
  MsgBox "Sorry, this control was not Found!", vbCritical, "Error"
  Exit Property
 End If
 
 m_ControlLeftPane = FoundCtl.Name
 PropertyChanged "ControlLeftPane"
 SizePanes
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ControlRightPane() As String
    ControlRightPane = m_ControlRightPane
End Property

Public Property Let ControlRightPane(ByVal New_ControlRightPane As String)
Dim ctl As Object
Dim FoundCtl As Object
Dim bControlFound As Boolean
 
 For Each ctl In UserControl.ContainedControls
  If LCase(ctl.Name) = LCase(New_ControlRightPane) Then
   bControlFound = True
   Set FoundCtl = ctl
   Exit For
  End If
 Next ctl
 
 If bControlFound = False Then
  MsgBox "Sorry, this control was not Found!", vbCritical, "Error"
  Exit Property
 End If

 m_ControlRightPane = FoundCtl.Name
 PropertyChanged "ControlRightPane"
 SizePanes
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ResizeContinuos() As Boolean
    ResizeContinuos = m_ResizeContinuos
End Property

Public Property Let ResizeContinuos(ByVal New_ResizeContinuos As Boolean)
    m_ResizeContinuos = New_ResizeContinuos
    PropertyChanged "ResizeContinuos"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SplitColorMouseUp() As OLE_COLOR
    SplitColorMouseUp = m_SplitColorMouseUp
End Property

Public Property Let SplitColorMouseUp(ByVal New_SplitColorMouseUp As OLE_COLOR)
    m_SplitColorMouseUp = New_SplitColorMouseUp
    picSplitter.BackColor = m_SplitColorMouseUp
    PropertyChanged "SplitColorMouseUp"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinimumPaneWidth() As Long
    MinimumPaneWidth = m_MinimumPaneWidth
End Property

Public Property Let MinimumPaneWidth(ByVal New_MinimumPaneWidth As Long)
    If New_MinimumPaneWidth < 0 Then New_MinimumPaneWidth = 0
    m_MinimumPaneWidth = New_MinimumPaneWidth
    PropertyChanged "MinimumPaneWidth"
    SizePanes False
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SplitterWidthHeight() As Long
    SplitterWidthHeight = m_SplitterWidthHeight
End Property

Public Property Let SplitterWidthHeight(ByVal New_SplitterWidthHeight As Long)
 
 If New_SplitterWidthHeight < 10 Then New_SplitterWidthHeight = 10
 
 If m_Orientation = vertical Then
    If New_SplitterWidthHeight > UserControl.Width / 2 Then
     MsgBox "Splitter Width cannot be greater than (Control Width/2)", vbCritical, "Error"
     Exit Property
    End If
    
    m_SplitterWidthHeight = New_SplitterWidthHeight
    picSplitter.Width = New_SplitterWidthHeight
 End If
     
 If m_Orientation = horizontal Then
    If New_SplitterWidthHeight > UserControl.Height / 2 Then
     MsgBox "Splitter height cannot be greater than (Control height/2)", vbCritical, "Error"
     Exit Property
    End If
    
    m_SplitterWidthHeight = New_SplitterWidthHeight
    picSplitter.Height = New_SplitterWidthHeight
 End If

 PropertyChanged "SplitterWidthHeight"
 SizePanes
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SplitterLeftTop() As Long
    SplitterLeftTop = m_SplitterLeftTop
End Property

Public Property Let SplitterLeftTop(ByVal New_SplitterLeftTop As Long)
    m_SplitterLeftTop = New_SplitterLeftTop
    
    If m_Orientation = vertical Then picSplitter.Left = New_SplitterLeftTop
    If m_Orientation = horizontal Then picSplitter.Top = New_SplitterLeftTop
    
    PropertyChanged "SplitterLeftTop"
    SizePanes
End Property


