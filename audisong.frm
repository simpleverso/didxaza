VERSION 5.00
Object = "{899C8CBC-06B4-4162-8AE6-2DB2B703D616}#1.0#0"; "AeroSuite.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form audisong 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reproductor de Didxaza."
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin AeroSuite.AeroGroupBox AeroGroupBox1 
      Height          =   5055
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8916
      BorderColor     =   14408667
      BackColor       =   -2147483633
      BackColor2      =   15395562
      HeadColor1      =   -2147483633
      HeadColor2      =   15000804
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reproductor Fonetico."
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   4575
         Left            =   360
         TabIndex        =   1
         ToolTipText     =   "con este reproductor se puede escuchar la forma correcta de la pronunciacion de las palabras..."
         Top             =   360
         Width           =   6495
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "mini"
         stretchToFit    =   0   'False
         windowlessVideo =   -1  'True
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   -1  'True
         _cx             =   11456
         _cy             =   8070
      End
   End
End
Attribute VB_Name = "audisong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
