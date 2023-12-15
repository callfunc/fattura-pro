VERSION 5.00
Begin VB.Form RicercaFornitori 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ricerca Fornitori"
   ClientHeight    =   4530
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "RicercaFornitori.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   8235
   Begin VB.TextBox TxtStato 
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
      Left            =   1440
      TabIndex        =   17
      Top             =   3345
      Width           =   2850
   End
   Begin VB.TextBox TxtCodFisc 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   915
      Width           =   2850
   End
   Begin VB.TextBox TxtPartIva 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   510
      Width           =   2385
   End
   Begin VB.TextBox TxtLoc 
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
      Left            =   1440
      TabIndex        =   6
      Top             =   2535
      Width           =   4050
   End
   Begin VB.TextBox TxtTel 
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1725
      Width           =   2280
   End
   Begin VB.TextBox TxtCap 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   2130
      Width           =   1380
   End
   Begin VB.TextBox TxtIndirizzo 
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
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   4080
   End
   Begin VB.TextBox TxtDitta 
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
      Left            =   1440
      TabIndex        =   0
      Top             =   105
      Width           =   6555
   End
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
      Height          =   525
      Left            =   3285
      TabIndex        =   8
      Top             =   3915
      Width           =   1515
   End
   Begin VB.TextBox TxtProv 
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
      Left            =   1440
      TabIndex        =   7
      Top             =   2940
      Width           =   645
   End
   Begin VB.Label LblStato 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stato:"
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
      Left            =   225
      TabIndex        =   18
      Top             =   3375
      Width           =   450
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codice Fiscale:"
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
      Left            =   225
      TabIndex        =   16
      Top             =   945
      Width           =   1170
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Partita Iva:"
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
      Left            =   225
      TabIndex        =   15
      Top             =   540
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Località:"
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
      Left            =   225
      TabIndex        =   14
      Top             =   2565
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono:"
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
      TabIndex        =   13
      Top             =   1755
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.A.P.:"
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
      Left            =   225
      TabIndex        =   12
      Top             =   2175
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indirizzo:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   1350
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ditta:"
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
      Left            =   225
      TabIndex        =   10
      Top             =   135
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia:"
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
      Left            =   225
      TabIndex        =   9
      Top             =   2985
      Width           =   780
   End
End
Attribute VB_Name = "RicercaFornitori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TxtCodFisc_KeyPress(KeyAscii As Integer)
If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVXYWZ" & vbBack, UCase(Chr(KeyAscii))) = 0 Then
 KeyAscii = 0
Else
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtPartIva_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtPartIva.Text) = 12 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub TxtTel_KeyPress(KeyAscii As Integer)
If InStr("0123456789" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
If TxtDitta.SelStart = 0 Or Mid(TxtDitta.Text, TxtDitta.SelStart + 1, 1) = "." Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtLoc_KeyPress(KeyAscii As Integer)
If TxtLoc.SelStart = 0 Or Mid(TxtLoc.Text, TxtLoc.SelStart + 1, 1) = "." Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtProv_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtProv.Text) = 2 Then
  KeyAscii = 0
 Else
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End If
End Sub
Private Sub Form_Load()
Me.Move 500, 600
End Sub
Private Sub Ricerca_Click()
Dim ValoriCampi As Variant, EsprRicerca As Variant, StrRicerca As String
ValoriCampi = Array(TxtDitta.Text, TxtCodFisc.Text, TxtPartIva.Text, TxtTel.Text, TxtIndirizzo.Text, _
TxtCap.Text, TxtLoc.Text, TxtProv.Text)
EsprRicerca = Array("Ditta LIKE 'val%'", "CodFiscale=val", "PartIva=val", "Tel=val", "Indirizzo=val", _
"Cap=val", "Loc=val", "ModPag=val", "Prov=val")
For i = 0 To UBound(ValoriCampi)
 If Trim(ValoriCampi(i)) <> "" Then
  StrRicerca = IIf(StrRicerca <> "", StrRicerca & " OR ", " WHERE ")
  EsprRicerca(i) = Replace(EsprRicerca(i), "val", Replace(ValoriCampi(i), "'", "''"))
  StrRicerca = StrRicerca & EsprRicerca(i)
 End If
Next i
If StrRicerca <> "" Then
 StrRicerca = StrRicerca & " AND Rimosso = False ORDER BY Ditta ASC"
 Dim rsFornitori As New ADODB.Recordset
 rsFornitori.Open "SELECT * FROM Fornitori" & StrRicerca, conn, adOpenDynamic, adLockOptimistic
 If Not rsFornitori.EOF Then
  Fornitori.ImpostaFiltroRicerca rsFornitori
  If Not FormVisibile("Fornitori") Then
   Fornitori.Show
  ElseIf Not Fornitori.FormBloccata Then Fornitori.CaricaFiltroRicerca
  Else
   MsgBox "Attenzione, la finestra è impegnata al momento nella modifica o nell'" _
   & "inserimento di un documento." & vbCrLf & "Prima di eseguire la ricerca completare " _
   & "l'operazione in corso !", vbExclamation, "Fattura Pro"
  End If
 Else
  MsgBox "Attenzione, non è stato trovato nessun risultato !", vbExclamation, "Fattura Pro"
 End If
Else
 MsgBox "Attenzione, tutti i campi di ricerca sono vuoti !", vbExclamation, "Fattura Pro"
End If
End Sub
