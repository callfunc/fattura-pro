VERSION 5.00
Begin VB.Form Info 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informazioni su Fattura Pro"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4470
   Icon            =   "Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label LblTipoProg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0033CCFF&
      Caption         =   "Software per la fatturazione"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00007800&
      Height          =   270
      Left            =   1125
      TabIndex        =   2
      Top             =   600
      Width           =   2700
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software scritto e compilato da Birddog"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   615
      TabIndex        =   1
      Top             =   1155
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fattura Pro 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00007800&
      Height          =   345
      Left            =   1125
      TabIndex        =   0
      Top             =   165
      Width           =   2145
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   240
      Picture         =   "Info.frx":4072
      Top             =   195
      Width           =   675
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
