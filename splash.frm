VERSION 5.00
Object = "{899C8CBC-06B4-4162-8AE6-2DB2B703D616}#1.0#0"; "AeroSuite.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10g.ocx"
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   Enabled         =   0   'False
   Icon            =   "splash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1815
      Left            =   -720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8175
      _cx             =   14420
      _cy             =   3201
      FlashVars       =   ""
      Movie           =   "C:\Program Files\didxazap\inf.dll"
      Src             =   "C:\Program Files\didxazap\inf.dll"
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "-1"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin AeroSuite.AeroProgressBar AeroProgressBar1 
      Height          =   270
      Left            =   0
      Top             =   1800
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   476
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   4920
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   14000
      Left            =   2760
      Top             =   840
   End
   Begin VB.Image Image1 
      Height          =   3465
      Left            =   0
      Picture         =   "splash.frx":1644A
      Top             =   0
      Width           =   3870
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
  
    HScroll1.Max = 255
    HScroll1.Min = 50
  
  
  
    HScroll1.Value = 200
  
End Sub
  
Private Sub HScroll1_Change()
  

  
    Call Aplicar_Transparencia(Me.hWnd, CByte(HScroll1.Value))
  
End Sub

Private Sub Timer1_Timer()
If AeroProgressBar1.Value = 100 Then

Load buskeda
buskeda.Show
Unload Me
Else
AeroProgressBar1.Value = Val(AeroProgressBar1.Value) + Val(1)
End If
End Sub
