VERSION 5.00
Begin VB.Form registro 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   3000
   ClientTop       =   3555
   ClientWidth     =   3855
   Icon            =   "registro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   3240
      Top             =   3000
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "necesita registrar este programa."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   3465
      Left            =   0
      Picture         =   "registro.frx":1644A
      Top             =   0
      Width           =   3870
   End
End
Attribute VB_Name = "registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
End
End Sub
