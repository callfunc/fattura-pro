VERSION 5.00
Begin VB.Form OpzioniStampa 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opzioni Stampa"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4980
   Icon            =   "OpzioniStampa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Stampanti 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "OpzioniStampa.frx":4072
      Left            =   1230
      List            =   "OpzioniStampa.frx":4074
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   300
      Width           =   3405
   End
   Begin VB.TextBox Copie 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Top             =   945
      Width           =   615
   End
   Begin VB.OptionButton Colore 
      BackColor       =   &H0033CCFF&
      Caption         =   "Colore"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   2
      Top             =   1530
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Biancoenero 
      BackColor       =   &H0033CCFF&
      Caption         =   "Bianco e Nero"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1290
      TabIndex        =   1
      Top             =   1530
      Width           =   1695
   End
   Begin VB.CommandButton Stampa 
      Caption         =   "Stampa"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   285
      MaskColor       =   &H00FFFFFF&
      Picture         =   "OpzioniStampa.frx":4076
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2100
      UseMaskColor    =   -1  'True
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stampante:"
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
      Left            =   270
      TabIndex        =   5
      Top             =   345
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero di Copie:"
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
      Left            =   270
      TabIndex        =   6
      Top             =   975
      Width           =   1410
   End
End
Attribute VB_Name = "OpzioniStampa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SS As ServiziStampa, StampanteCorr As String, CaricamentoForm As Boolean
Public Scelta As String
Private Sub CambiaCopie_GotFocus()
Stampa.SetFocus
End Sub
Private Sub CambiaCopie_SpinDown()
If Val(Copie.Text) > 1 Then
 Copie.Text = Val(Copie.Text) - 1
End If
End Sub
Private Sub CambiaCopie_SpinUp()
Copie.Text = Val(Copie.Text) + 1
End Sub
Public Sub Inizializza(ClasseStampa As ServiziStampa)
Set SS = ClasseStampa
End Sub
Private Sub Form_Load()
Me.Move 1500, 2000: Scelta = "": CaricamentoForm = True
Dim p As Printer: Stampanti.Clear
If IsObject(Printer) Then
 StampanteCorr = Printer.DeviceName
 If Not StampaConf Then
  Printer.ScaleMode = vbMillimeters: Printer.Orientation = vbPRORPortrait
  Printer.PaperSize = vbPRPSA4: Printer.ColorMode = vbPRCMColor
  StampaConf = True
 End If
End If
For Each p In Printers
 Stampanti.AddItem p.DeviceName
 If StampanteCorr = p.DeviceName Then
  Stampanti.ListIndex = Stampanti.NewIndex
 End If
Next
Copie.Text = 1: CaricamentoForm = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Scelta = "" Then
 If Printer.DeviceName <> StampanteCorr Then
  For Each p In Printers
   If p.DeviceName = StampanteCorr Then
    Set Printer = p: Exit For
   End If
  Next
 End If
End If
End Sub
Private Sub Stampa_Click()
Scelta = "Stampa": Printer.Copies = Copie.Text
If Colore.Value Then
 Printer.ColorMode = vbPRCMColor
Else
 Printer.ColorMode = vbPRCMMonochrome
End If
Unload Me
End Sub
Private Sub Stampanti_Click()
If Not CaricamentoForm Then
 Call SS.ImpostaStampante(Stampanti.Text)
End If
If Printer.DeviceName = "PDFCreator" Then
 Colore.Value = True: Printer.ColorMode = vbPRCMColor
 Exit Sub
End If
If Printer.ColorMode = vbPRCMColor Then
 Colore.Value = True
Else
 Biancoenero.Value = True
End If
End Sub
