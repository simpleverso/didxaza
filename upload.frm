VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{899C8CBC-06B4-4162-8AE6-2DB2B703D616}#1.0#0"; "AeroSuite.ocx"
Begin VB.Form upload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "agregar una palabra a la base de datos."
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7125
   Icon            =   "upload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\didxazap\stream.dll;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\didxazap\stream.dll;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "zap_es"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin AeroSuite.AeroButton AeroButton4 
      Height          =   735
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      Caption         =   "insertar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      DataField       =   "zap"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin AeroSuite.AeroButton AeroButton3 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "palabra en zapoteco"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      State           =   3
   End
   Begin AeroSuite.AeroButton AeroButton2 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "palabra en español"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      State           =   3
   End
   Begin AeroSuite.AeroButton AeroButton1 
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1085
      Caption         =   "                     uselo bajo su propio RIESGO                               una vez agregada la palabra NO se puede eliminar."
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      State           =   3
   End
   Begin VB.TextBox Text1 
      DataField       =   "esp"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   0
      Picture         =   "upload.frx":1644A
      Top             =   -7560
      Width           =   20490
   End
End
Attribute VB_Name = "upload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AeroButton4_Click()
Adodc1.Recordset.Update
MsgBox ("el registro se efectuó con éxito")
Load buskeda
buskeda.Show
Unload upload
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load buskeda
buskeda.Show
End Sub
