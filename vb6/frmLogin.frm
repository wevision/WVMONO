VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   1500
   ClientLeft      =   2835
   ClientTop       =   3360
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   886.25
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Text            =   "sangabuy_wv"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ingresar"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "bti3elef"
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nombre de usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cnn As ADODB.Connection
Public rs As ADODB.Recordset
Public errorObject As ADODB.Error
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    End
End Sub

Private Sub cmdOK_Click()

 On Error GoTo DisplayErrorInfo
    
   Set cnn = New ADODB.Connection
   Set rs = New ADODB.Recordset

   ' Open a Connection using an ODBC DSN named "Pubs".
   On Error GoTo DisplayErrorInfo

    cnn.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=ituruguay.com;Port=3306;Database=sangabuy_wevision;User=" & txtUserName.Text & "; Password=" & txtPassword.Text & ";Option=3;"
    cnn.ConnectionTimeout = 30
    cnn.CursorLocation = adUseClient
    
    cnn.Open
    
    'comprobar si la contraseña es correcta
   If cnn.State = adStateOpen Then
    
    ' Close the connection.
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        'lo más fácil es establecer una variable global
        
        LoginSucceeded = True
        Me.Hide
        frmMainScreen.Show
        
    Else
        MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If

DisplayErrorInfo:

For Each errorObject In cnn.Errors
    MsgBox "Description :" & errorObject.NativeError & errorObject.Description
    Debug.Print "Description :"; errorObject.Description
    Debug.Print "Number:"; Hex(errorObject.Number)
Next

End Sub

