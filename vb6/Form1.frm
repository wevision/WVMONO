VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E23BBAD-AA35-11D1-ADEA-0000F87734F0}#1.0#0"; "trialoc.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SYSTV - WeVision"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2295
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4048
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Button1 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8453888
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   5
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Usuarios"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin TRIALOCLibCtl.TrialEnd TrialEnd1 
      Left            =   360
      OleObjectBlob   =   "Form1.frx":008B
      Top             =   4080
   End
   Begin VB.Label Label2 
      Caption         =   "APELLIDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Dim rst As New ADODB.Record
   ' Dim cn As New ADODB.Connection
    
    'Set cn = New ADODB.Connection
    
    
   ' Dim sql As String
   
   ' sql = "SELECT * FROM usuario"

   ' cn.CursorLocation = adUseClient
    'cn.Open CADENACONEXION

    'rst.Open sql, cn, adOpenKeyset, adLockOptimistic

    'Set DataGrid1.DataSource = rst
    
    

       ' Create a Recordset by executing an SQL statement.
    Set frmLogin.rs = frmLogin.cnn.Execute("Select * From usuarios")
    
    ' Show the first author.
    'MsgBox rs("nombre") & " " & rs("apellido")
    Set DataGrid1.DataSource = frmLogin.rs
   ' Find out if the attempt to connect worked.


   ' Close the connection.
   'cnn.Close


End Sub
Private Sub Button1_Click()

    Set frmLogin.rs = frmLogin.cnn.Execute("INSERT INTO usuarios(nombre,apellido) VALUES ('" & Text1.Text & "','" & Text2.Text & "');")
    
    Set frmLogin.rs = frmLogin.cnn.Execute("Select * From usuarios")
    
    Set DataGrid1.DataSource = frmLogin.rs
    
    'Call ConsultaSQL("INSERT INTO usuarios(nombre,apellido) VALUES ('" & Text1.Text & "','" & Text2.Text & "');")
           
    'Call Mostrar_Tabla

End Sub
