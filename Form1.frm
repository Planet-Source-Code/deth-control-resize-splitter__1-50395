VERSION 5.00
Object = "*\ASimpleSplitter.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin SimpleSplitter.Splitter Splitter1 
      Height          =   2220
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   3916
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   1845
         TabIndex        =   2
         Top             =   90
         Width           =   1365
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   225
         TabIndex        =   1
         Top             =   90
         Width           =   1230
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
  
'set the left and right control to be split
 Set Splitter1.LeftOrTopControl = Dir1
 Set Splitter1.RightOrBottomControl = File1

'thats all there is to it!!!

End Sub

Private Sub Form_Resize()
   
   'this just resizes the splitter window to this form
   If WindowState <> 1 Then
       Splitter1.Move 0, 0, Form1.Width, Form1.Height
   End If
   
End Sub
