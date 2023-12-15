VERSION 5.00
Begin VB.Form RicercaArticoli 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ricerca Articoli"
   ClientHeight    =   2460
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "RicercaArticoli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   7860
   Begin VB.CommandButton Ricerca 
      Caption         =   "Ricerca"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2940
      TabIndex        =   8
      Top             =   1800
      Width           =   1515
   End
   Begin VB.TextBox TxtPrezzo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   540
      Width           =   1440
   End
   Begin VB.ListBox LstUm 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      ItemData        =   "RicercaArticoli.frx":4072
      Left            =   6225
      List            =   "RicercaArticoli.frx":4082
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   840
   End
   Begin VB.ListBox LstAliqIva 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      ItemData        =   "RicercaArticoli.frx":4094
      Left            =   4395
      List            =   "RicercaArticoli.frx":4096
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   1050
   End
   Begin VB.TextBox TxtDesc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   6030
   End
   Begin VB.Label Misura_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U.M.:"
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
      Left            =   5730
      TabIndex        =   7
      Top             =   555
      Width           =   420
   End
   Begin VB.Label LblAliqIVA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota IVA:"
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
      Left            =   3300
      TabIndex        =   6
      Top             =   555
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prezzo:"
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
      Left            =   945
      TabIndex        =   5
      Top             =   555
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Prodotto:"
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
      Left            =   210
      TabIndex        =   4
      Top             =   135
      Width           =   1290
   End
End
Attribute VB_Name = "RicercaArticoli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAliquoteIVA As ADODB.Recordset
Private Sub Form_Load()
Me.Move 500, 600
Set rsAliquoteIVA = New ADODB.Recordset
rsAliquoteIVA.Open "SELECT * FROM AliquoteIVA WHERE Aliquota <> 0 ORDER BY Aliquota ASC", conn, adOpenDynamic
While Not rsAliquoteIVA.EOF
 LstAliqIva.AddItem rsAliquoteIVA("Aliquota")
 rsAliquoteIVA.MoveNext
Wend
End Sub
Private Sub TxtPrezzo_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If (TxtPrezzo.SelStart = 0 Or TxtPrezzo.SelStart = Len(TxtPrezzo.Text) - 1) _
 And KeyAscii = 44 Then
  KeyAscii = 0
  Exit Sub
 End If
 If InStr(1, "0123456789,", Chr(KeyAscii)) = 0 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub Ricerca_Click()
Dim ValoriCampi As Variant, StrRicerca As String
ValoriCampi = Array(TxtDesc.Text, TxtPrezzo.Text, LstAliqIva.Text, LstUm.Text)
Dim EsprRicerca As Variant
EsprRicerca = Array("Descr LIKE 'val%'", "Prezzo1=val OR Prezzo2=val OR Prezzo3=val", "AliqIva=val", "Um=val")
For i = 0 To UBound(ValoriCampi)
 If Trim(ValoriCampi(i)) <> "" Then
  StrRicerca = IIf(StrRicerca <> "", StrRicerca & " OR ", " WHERE ")
  EsprRicerca(i) = Replace(EsprRicerca(i), "val", ValoriCampi(i))
  StrRicerca = StrRicerca & EsprRicerca(i)
 End If
Next i
If StrRicerca <> "" Then
 StrRicerca = StrRicerca & " ORDER BY Descr ASC"
 Dim rsArticoli As New ADODB.Recordset
 rsArticoli.Open "SELECT * FROM Articoli" & StrRicerca, conn, adOpenDynamic, adLockOptimistic
 If Not rsArticoli.EOF Then
  Articoli.ImpostaFiltroRicerca rsArticoli
  If Not FormVisibile("Articoli") Then
   Articoli.Show
  ElseIf Not Articoli.FormBloccata Then Articoli.CaricaFiltroRicerca
  Else
   MsgBox "Attenzione, la finestra Articoli è impegnata al momento nella modifica o nell'" _
   & "inserimento di un documento." & vbCrLf & "Prima di eseguire la ricerca completare " _
   & "l'operazione in corso !", vbExclamation, "Fattura Pro"
  End If
 Else
  MsgBox "Attenzione, non è stato trovato nessun articolo !", vbExclamation, "Fattura Pro"
 End If
Else
 MsgBox "Attenzione, tutti i campi di ricerca sono vuoti !", vbExclamation, "Fattura Pro"
End If
End Sub
