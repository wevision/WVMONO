VERSION 5.00
Begin VB.Form frmHome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SYSTV  - MAIN SCREEN"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.VScrollBar VScroll1 
      Height          =   4455
      Left            =   7920
      Max             =   30
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   1440
      Negotiate       =   -1  'True
      ScaleHeight     =   4455
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000010&
         Caption         =   "Command3"
         Height          =   1095
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click(Index As Integer)
    Form1.Show
End Sub

Private Sub Form_Load()

    With VScroll1 'Si vas a utilizar el Vertical
    .Min = 0
    .SmallChange = 90
    .LargeChange = 300
    .Top = 0
    .ZOrder 0
    End With
    'Cancel = True

End Sub
Private Sub SubWizard1_GotFocus()

End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub VScroll1_Change()
    Picture1.Top = -VScroll1.Value
End Sub

