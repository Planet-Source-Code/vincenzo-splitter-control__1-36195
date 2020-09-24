VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin Project1.Split Split1 
      Height          =   7080
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   12488
      SplitColorMouseDown=   16761024
      ControlLeftPane =   "Picture1"
      ControlRightPane=   "Picture2"
      SplitColorMouseUp=   -2147483633
      SplitterLeft    =   5250
      SplitterLeftTop =   5250
      Begin VB.PictureBox Picture1 
         Height          =   7080
         Left            =   0
         ScaleHeight     =   7020
         ScaleWidth      =   5190
         TabIndex        =   2
         Top             =   0
         Width           =   5250
         Begin VB.TextBox Text1 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H0000FFFF&
            Height          =   1695
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   3
            Top             =   45
            Width           =   2355
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   7080
         Left            =   5295
         ScaleHeight     =   7020
         ScaleWidth      =   5145
         TabIndex        =   1
         Top             =   0
         Width           =   5205
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim szLine As String

 On Error GoTo errhandler
 
 Split1.SplitterLeftTop = GetSetting(App.Title, "Settings", "SplitLeft", Split1.Width / 2)
 
 Open App.Path + "\readme.txt" For Input As #1
  While Not (EOF(1))
   Input #1, szLine
   Text1.Text = Text1.Text + szLine + vbCrLf
  Wend
 Close #1
 
 
errhandler:
End Sub

Private Sub Split1_PaneEndResize()
On Error GoTo errhandler
 Text1.Move 0, 0, Picture1.Width - 60, Picture1.Height - 60
 SaveSetting App.Title, "Settings", "SplitLeft", Split1.SplitterLeftTop
errhandler:
End Sub
