VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{899C8CBC-06B4-4162-8AE6-2DB2B703D616}#1.0#0"; "AeroSuite.ocx"
Begin VB.Form buskeda 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "traductor español-zapoteco por Gonzalo. BETA"
   ClientHeight    =   6735
   ClientLeft      =   7995
   ClientTop       =   3345
   ClientWidth     =   9000
   Icon            =   "buskeda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin AeroSuite.AeroButton AeroButton4 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "agregar"
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
   Begin AeroSuite.AeroButton AeroButton3 
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "Filtrar"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
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
      ForeColor       =   -2147483625
      State           =   3
   End
   Begin VB.TextBox Text3 
      DataField       =   "zap"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "esp"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "buskeda.frx":1644A
      Height          =   3855
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      Caption         =   "palabras"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "esp"
         Caption         =   "español"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "zap"
         Caption         =   "zapoteco"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin AeroSuite.AeroButton AeroButton1 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      Caption         =   "buscar"
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      ToolTipText     =   "introduce una letra para comenzar a buscar"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   5160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
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
   Begin AeroSuite.AeroButton AeroButton2 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   0
      Picture         =   "buskeda.frx":1645F
      Top             =   -8160
      Width           =   20490
   End
   Begin VB.Menu a3 
      Caption         =   "&como registrar"
      Enabled         =   0   'False
   End
   Begin VB.Menu a1 
      Caption         =   "&ayuda"
      Begin VB.Menu a2 
         Caption         =   "web del autor"
      End
   End
End
Attribute VB_Name = "buskeda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32" ()
'Dim contador As Integer
'Dim letra As Integer


Private Sub a2_Click()
MsgBox ("visite: simplesoft.netii.net")
End Sub

Private Sub a3_Click()
MsgBox ("el registro es gratuito, mas informacion en simplesoft.netii.net")
End Sub

Private Sub AeroButton1_Click()

Adodc1.Recordset.MoveFirst

Text1.Text = "a"

Text1.Text = ""

buscar
''''
'contador = contador + 1
'If contador = 5 Then
'Load registro
'registro.Show
'Unload buskeda
'End If
'''''
End Sub
Private Sub buscar()
Dim criterio
criterio = InputBox$("introduce una palabra en español")
            If Trim$(criterio) <> "" Then
                    'Adodc1.Recordset.MoveFirst
                    criterio = "esp='" + criterio + "'"
                    Adodc1.Recordset.Find criterio
            End If
            If Adodc1.Recordset.EOF Then
                    Adodc1.Recordset.MoveFirst
                    MsgBox ("la palabra no se encontró o está mal escrita")
            End If
End Sub

Private Sub AeroButton4_Click()
buskeda.Hide
Load upload
upload.Show
End Sub

Private Sub Form_Initialize()
    Call SetErrorMode(2)
    Call InitCommonControls
End Sub

Private Sub Adodc1_Error( _
    ByVal ErrorNumber As Long, _
    Description As String, _
    ByVal Scode As Long, _
    ByVal Source As String, _
    ByVal HelpFile As String, _
    ByVal HelpContext As Long, _
    fCancelDisplay As Boolean)
  MsgBox " DEscripción del Error :" & Description
    'Adodc1.Caption = " Registro actual: " & CStr(Adodc1.Recordset.AbsolutePosition)
'End Sub
 'Dim bCancel As Boolean
  'Select Case adReason
  'Case adRsnAddNew
  'Case adRsnClose
  'Case adRsnDelete
  'Case adRsnFirstChange
  'Case adRsnMove
  'Case adRsnRequery
  'Case adRsnResynch
  'Case adRsnUndoAddNew
  'Case adRsnUndoDelete
  'Case adRsnUndoUpdate
  'Case adRsnUpdate
  'End Select
  'If bCancel Then adStatus = adStatusCancel
'End Sub
End Sub


'Private Sub Form_Unload(Cancel As Integer)
 'MsgBox ("gracias por utilizar este programa")
 'End
 'End Sub
 
Private Sub Form_Load()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

CrearAsociacion App.Path & "\" & App.EXEName, "zap", "archivo de datos didxaza", ""
    
    With Adodc1
        .CommandType = adCmdText
        .RecordSource = "Select * From Zap_es"
        .Refresh
        Set DataGrid1.DataSource = Adodc1.Recordset
        DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
        .Visible = False
    End With
    
    With Combo1
        Combo1.Text = "esp"
    End With
    Text1 = ""
    
End Sub

Private Sub Text1_Change()
    On Error GoTo error_handler
    With Adodc1
        If Text1 <> "" Then
            .Recordset.Filter = Combo1 & " LIKE '*" + Text1 + "*'"
             'MsgBox ("la palabra no se encuentra en los datos del diccionario")
           ''''
            'letra = letra + 1
'If letra = 10 Then
'Load registro
'registro.Show
'Unload buskeda
'End If
''''
            Set DataGrid1.DataSource = Adodc1.Recordset
        Else
            .Recordset.Filter = ""
        End If
        .Refresh
    End With
    Exit Sub

error_handler:

    If Err.Number = 3265 Then
        MsgBox "el campo seleccionado no es válido", vbCritical
    Else
        MsgBox Err.Description, vbCritical
    End If
End Sub
