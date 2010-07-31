VERSION 5.00
Begin VB.MDIForm frmMainScreen 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "SYSTV 2010 - LICENCIADO PARA WEVISION"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13725
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu HHH 
      Caption         =   "Nombre"
   End
End
Attribute VB_Name = "frmMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()

    frmHome.Show
    frmHome.Left = (Me.Width - frmHome.Width) / 2
    frmHome.Top = (Me.Height - frmHome.Height) / 3
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub
