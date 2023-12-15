VERSION 5.00
Begin VB.Form InsArticoloMag 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inserisci Articolo"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InsArticoloMag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnAnnulla 
      Caption         =   "Annulla"
      Height          =   405
      Left            =   3810
      TabIndex        =   5
      Top             =   1200
      Width           =   1245
   End
   Begin VB.CommandButton BtnSalva 
      Caption         =   "Salva"
      Height          =   405
      Left            =   2250
      TabIndex        =   4
      Top             =   1200
      Width           =   1245
   End
   Begin VB.TextBox TxtDescr 
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   180
      Width           =   5415
   End
   Begin VB.TextBox TxtGiacenzaIniziale 
      Height          =   315
      Left            =   1590
      TabIndex        =   0
      Top             =   600
      Width           =   1605
   End
   Begin VB.Label LblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrizione:"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   945
   End
   Begin VB.Label LblGiacenza 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Giacenza iniziale:"
      Height          =   225
      Left            =   195
      TabIndex        =   2
      Top             =   630
      Width           =   1335
   End
End
Attribute VB_Name = "InsArticoloMag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsArticoli As ADODB.Recordset
Public ArticoloAggiunto As Boolean
Private Sub BtnAnnulla_Click()
rsArticoli.Close
Unload Me
End Sub
Private Sub BtnSalva_Click()
Dim rsDuplicato As New ADODB.Recordset
rsDuplicato.Open "SELECT * FROM Articoli WHERE descr = '" & TxtDescr.Text & "'", conn, adOpenDynamic
If Not rsDuplicato.EOF Then
 e = True
End If
rsDuplicato.Close
If e Then
 MsgBox "Attenzione, L'articolo esiste già in archivio !", vbExclamation, "Fattura Pro"
 Exit Sub
End If
If TxtGiacenzaIniziale.Text = "" Then
 MsgBox "Attenzione, Inserire la giacenza iniziale !", vbExclamation, "Fattura Pro"
 Exit Sub
End If
rsArticoli.AddNew
rsArticoli("Descr") = TxtDescr.Text
rsArticoli("GiacenzaIniziale") = TxtGiacenzaIniziale.Text
rsArticoli.Update
rsArticoli.Close
ArticoloAggiunto = True
Unload Me
End Sub
Private Sub Form_Load()
Set rsArticoli = New ADODB.Recordset
rsArticoli.Open "Articoli", conn, adOpenDynamic, adLockOptimistic
End Sub
