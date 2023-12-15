VERSION 5.00
Begin VB.Form Fornitori 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fornitori"
   ClientHeight    =   5535
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "Fornitori.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   8145
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
      Left            =   1515
      TabIndex        =   30
      Top             =   3135
      Width           =   2550
   End
   Begin VB.TextBox TxtFax 
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
      Left            =   1515
      TabIndex        =   26
      Top             =   3945
      Width           =   2550
   End
   Begin VB.TextBox TxtEmail 
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
      Left            =   1515
      TabIndex        =   25
      Top             =   4365
      Width           =   2550
   End
   Begin VB.CommandButton BtnPrimo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Fornitori.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnPrec 
      Appearance      =   0  'Flat
      DisabledPicture =   "Fornitori.frx":460C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Fornitori.frx":4BA6
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TxtRecCorr 
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
      Left            =   1635
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4980
      Width           =   1125
   End
   Begin VB.CommandButton BtnSucc 
      Height          =   315
      Left            =   2805
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Fornitori.frx":5140
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnUltimo 
      Height          =   315
      Left            =   3150
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Fornitori.frx":56DA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnNuovo 
      Height          =   315
      Left            =   3495
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Fornitori.frx":5C74
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnCanc 
      Height          =   315
      Left            =   3960
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Fornitori.frx":620E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
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
      Left            =   1515
      TabIndex        =   0
      Top             =   300
      Width           =   6345
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
      Left            =   1515
      TabIndex        =   3
      Top             =   1515
      Width           =   4440
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
      Left            =   1515
      TabIndex        =   5
      Top             =   2325
      Width           =   1710
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
      Left            =   1515
      TabIndex        =   4
      Top             =   3540
      Width           =   2550
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
      Left            =   1515
      TabIndex        =   6
      Top             =   1920
      Width           =   4440
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
      Left            =   1515
      TabIndex        =   1
      Top             =   705
      Width           =   2535
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
      Left            =   1515
      TabIndex        =   2
      Top             =   1110
      Width           =   3150
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
      Left            =   1515
      TabIndex        =   7
      Top             =   2730
      Width           =   540
   End
   Begin VB.Label Label7 
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
      Left            =   255
      TabIndex        =   29
      Top             =   3180
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
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
      TabIndex        =   28
      Top             =   3990
      Width           =   300
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
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
      TabIndex        =   27
      Top             =   4395
      Width           =   555
   End
   Begin VB.Label LblNumRecord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   4395
      TabIndex        =   24
      Top             =   5025
      Width           =   45
   End
   Begin VB.Label Legenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record:"
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
      TabIndex        =   23
      Top             =   5010
      Width           =   600
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
      Left            =   255
      TabIndex        =   15
      Top             =   1950
      Width           =   660
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
      Left            =   255
      TabIndex        =   14
      Top             =   2760
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
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
      TabIndex        =   13
      Top             =   330
      Width           =   540
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
      Left            =   270
      TabIndex        =   12
      Top             =   1545
      Width           =   705
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
      Left            =   270
      TabIndex        =   11
      Top             =   2370
      Width           =   525
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
      Left            =   240
      TabIndex        =   10
      Top             =   3585
      Width           =   750
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
      Left            =   255
      TabIndex        =   9
      Top             =   735
      Width           =   825
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
      Left            =   270
      TabIndex        =   8
      Top             =   1140
      Width           =   1170
   End
End
Attribute VB_Name = "Fornitori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsFornitori As ADODB.Recordset, StatoFornitore As StatoRecord
Dim ElencoControlli As Variant, DescControlli As Variant, PropControlli As Variant
Dim FiltroRicerca As Boolean
Private Sub BtnCanc_Click()
Dim Scelta%
Scelta = MsgBox("Cancellare questo fornitore ?" & vbNewLine & "Non sarà possibile" _
& " annullare questa modifica !", vbYesNo + vbQuestion, "Fattura Pro")
If Scelta = vbYes Then
 Dim PosRec As Integer
 PosRec = rsFornitori.AbsolutePosition
 If StatoFornitore <> inserimento Then
  rsFornitori("Rimosso") = True
  rsFornitori.Update
  rsFornitori.Requery
  If PosRec > rsFornitori.RecordCount Then
   If rsFornitori.RecordCount <> 0 Then
    rsFornitori.MoveLast
   End If
   CreaNuovoRecord
  Else
   rsFornitori.Move PosRec - 1, adBookmarkFirst
   VisualizzaRecord
   TxtRecCorr.Text = rsFornitori.AbsolutePosition
   LblNumRecord.Caption = "di " & rsFornitori.RecordCount
  End If
 Else
  CreaNuovoRecord
 End If
End If
End Sub
Private Sub TxtCap_Change()
ModificaFornitore
End Sub
Private Sub TxtCodFisc_Change()
ModificaFornitore
End Sub
Private Sub TxtCodFisc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(KeyAscii))
If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVXYWZ" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub TxtEmail_Change()
ModificaFornitore
End Sub
Private Sub TxtFax_Change()
ModificaFornitore
End Sub
Private Sub TxtRecCorr_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtRecCorr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtDitta_Change()
ModificaFornitore
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
If TxtDitta.SelStart = 0 Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtLoc_Change()
ModificaFornitore
End Sub
Private Sub TxtLoc_KeyPress(KeyAscii As Integer)
If TxtLoc.SelStart = 0 Or Mid(TxtLoc.Text, TxtLoc.SelStart + 1, 1) = "." Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtPartIva_Change()
ModificaFornitore
End Sub
Private Sub TxtPartIva_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtPartIva.Text) = 18 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub TxtProv_Change()
ModificaFornitore
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
Private Sub TxtStato_Change()
ModificaFornitore
End Sub
Private Sub TxtTel_Change()
ModificaFornitore
End Sub
Private Sub TxtIndirizzo_Change()
ModificaFornitore
End Sub
Private Sub TxtIndirizzo_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If TxtIndirizzo.SelStart = 0 Or Mid(TxtIndirizzo.Text, TxtIndirizzo.SelStart + 1, 1) = "." Then
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End If
End Sub
Private Sub Form_Load()
Me.Move 450, 300
If rsFornitori Is Nothing Then
 Set rsFornitori = New ADODB.Recordset
 rsFornitori.Open "SELECT * FROM Fornitori WHERE Rimosso = False ORDER BY Ditta ASC", conn, adOpenDynamic, adLockOptimistic
End If
BtnPrec.Enabled = False: BtnPrimo.Enabled = False
BtnUltimo.Enabled = rsFornitori.RecordCount > 1
TxtRecCorr.Text = "1"

ElencoControlli = Array("TxtDitta", "TxtPartIva", "TxtCodFisc", "TxtIndirizzo", "TxtTel", _
"TxtFax", "TxtEmail", "TxtLoc", "TxtProv", "TxtStato")
DescControlli = Array("Ditta", "Partita Iva", "Cod. Fiscale", "Indirizzo", "Telefono", _
"C.A.P.", "Località", "Provincia", "Stato")
PropControlli = Array("ao", "o", "a", "a", "n", "n", "a", "ao", "a", "a")

If Not rsFornitori.EOF Then
 If rsFornitori.RecordCount >= 1 Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
  BtnCanc.Enabled = True
 End If
 LblNumRecord.Caption = "di " & rsFornitori.RecordCount
 rsFornitori.MoveFirst
 Call VisualizzaRecord
Else
 LblNumRecord.Caption = "di 1"
 CreaNuovoRecord
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If StatoFornitore <> NonModificato Then
 Dim Scelta As VbMsgBoxResult
 Scelta = MsgBox("Salvare il record corrente ?", _
 vbYesNoCancel + vbQuestion, "Chiusura Archivio Fornitori")
 If Scelta = vbYes Then
  If Not ConvalidaRecord(True) Then
   Cancel = 1
  End If
 ElseIf Scelta = vbCancel Then
  Cancel = 1
 End If
End If
If Cancel <> 1 Then
 Set rsFornitori = Nothing
 Set Fornitori = Nothing
End If
End Sub
Private Sub BtnPrimo_Click()
PosizioneRecord "primo"
End Sub
Private Sub BtnPrec_Click()
PosizioneRecord "precedente"
End Sub
Private Sub BtnSucc_Click()
PosizioneRecord "successivo"
End Sub
Private Sub BtnUltimo_Click()
PosizioneRecord "ultimo"
End Sub
Private Sub BtnNuovo_Click()
PosizioneRecord "nuovo"
End Sub
Private Function ConvalidaRecord(ByVal Salva As Boolean) As Boolean
If StatoFornitore <> NonModificato Then
 For i = 0 To UBound(ElencoControlli)
  If Trim(Me(ElencoControlli(i))) = "" Then
   If InStr(1, PropControlli(i), "o") <> 0 Then
    MsgBox "Attenzione, " & DescControlli(i) & " è un campo obbligatorio !", vbExclamation, _
    "Fattura Pro"
    Exit Function
   End If
  ElseIf PropControlli(i) = "a" And IsNumeric(Trim(Me(ElencoControlli(i)))) Then
   MsgBox "Attenzione, il campo " & DescControlli(i) & " non può contenere valori valori nume" _
   & "rici !", vbExclamation, "Fattura Pro"
   Me(ElencoControlli(i)).Text = "": Exit Function
  ElseIf PropControlli(i) = "n" And Not IsNumeric(Trim(Me(ElencoControlli(i)))) Then
   MsgBox "Attenzione, il campo " & DescControlli(i) & " deve contenere valori numerici !", _
   vbExclamation, "Fattura Pro"
   Me(ElencoControlli(i)).Text = "": Exit Function
  End If
 Next i
 Dim CercaDuplicato As Boolean
 If StatoFornitore = inserimento Then
  CercaDuplicato = True
 ElseIf TxtDitta <> rsFornitori("Ditta") Then
  CercaDuplicato = True
 End If
 If CercaDuplicato Then
  Dim rsDuplicato As New ADODB.Recordset
  rsDuplicato.Open "SELECT * FROM Fornitori WHERE Rimosso = False And UCASE(Ditta) = '" & _
  UCase(Replace(TxtDitta.Text, "'", "''")) & "'", conn, adOpenDynamic
  If Not rsDuplicato.EOF Then
   MsgBox "La ditta che vuoi inserire è già presente in archivio !", vbExclamation, "Fattura Pro"
   TxtDitta.SetFocus: TxtDitta.SelStart = 0: TxtDitta.SelLength = Len(TxtDitta.Text): Exit Function
  End If
 End If
 If TxtProv.Text = "" And TxtStato.Text = "" Then
  MsgBox "Provincia e Stato non possono essere entrambi vuoti !", vbExclamation, "Fattura Pro"
  TxtProv.SetFocus: Exit Function
 End If
 If StatoFornitore = inserimento Then
  rsFornitori.AddNew
 End If
 With rsFornitori
  .Fields("ditta") = EliminaSpazi(TxtDitta.Text)
  .Fields("partitaiva") = Trim(TxtPartIva.Text)
  .Fields("codfiscale") = Trim(TxtCodFisc.Text)
  .Fields("indirizzo") = Trim(TxtIndirizzo.Text)
  .Fields("cap") = Trim(TxtCap.Text)
  .Fields("tel") = Trim(TxtTel.Text)
  .Fields("email") = Trim(TxtEmail.Text)
  .Fields("fax") = Trim(TxtFax.Text)
  .Fields("loc") = Trim(TxtLoc.Text)
  .Fields("prov") = Trim(TxtProv.Text)
  .Fields("stato") = Trim(TxtStato.Text)
  .Fields("coddest") = Trim(TxtCodDest.Text)
  .Update
 End With
 If StatoFornitore = inserimento Then
  CancellaUgualiRimossi
 End If
 EseguiBackup = True
 LblNumRecord.Caption = rsFornitori.RecordCount
 StatoFornitore = NonModificato
End If
ConvalidaRecord = True
End Function
Private Sub CancellaUgualiRimossi()
Dim rsRimossi As New ADODB.Recordset
rsRimossi.Open "SELECT * FROM Fornitori WHERE Rimosso = True And UCASE(Ditta) = '" & _
Replace(TxtDitta.Text, "'", "''") & "'", conn, adOpenDynamic, adLockOptimistic
If Not rsRimossi.EOF Then
 rsRimossi.MoveFirst
 conn.Execute "UPDATE FattureFornitori SET FattureFornitori.IdDitta = " & rsFornitori("Id") & " WHERE " _
 & "IdDitta = " & rsRimossi("Id")
 rsRimossi.Delete
 rsFornitori.Requery
 rsFornitori.Move rsFornitori.RecordCount - 1, adBookmarkFirst
End If
rsRimossi.Close
End Sub
Private Sub PosizioneRecord(PosRecord As String)
Dim valido: valido = ConvalidaRecord(True)
If valido Then
 Select Case PosRecord
 Case "primo"
  rsFornitori.MoveFirst
 Case "ultimo"
  rsFornitori.MoveLast
 Case "precedente"
  If CInt(TxtRecCorr.Text) <= rsFornitori.RecordCount Then
   rsFornitori.MovePrevious
  End If
 Case "successivo"
  If rsFornitori.AbsolutePosition = rsFornitori.RecordCount Then
   CreaNuovoRecord
   Exit Sub
  End If
  rsFornitori.MoveNext
 Case "nuovo"
  CreaNuovoRecord
  Exit Sub
 End Select
 BtnCanc.Enabled = True
 If rsFornitori.AbsolutePosition <= rsFornitori.RecordCount Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
 End If
 If rsFornitori.AbsolutePosition <> rsFornitori.RecordCount Then
  BtnUltimo.Enabled = True
 Else
  BtnUltimo.Enabled = False
 End If
 If rsFornitori.AbsolutePosition <> 1 Then
  BtnPrimo.Enabled = True: BtnPrec.Enabled = True
 Else
  BtnPrimo.Enabled = False: BtnPrec.Enabled = False
 End If
 TxtRecCorr.Text = rsFornitori.AbsolutePosition
 LblNumRecord.Caption = "di " & rsFornitori.RecordCount
 VisualizzaRecord
End If
End Sub
Private Sub CreaNuovoRecord()
TxtRecCorr.Text = rsFornitori.RecordCount + 1: LblNumRecord.Caption = "di " & rsFornitori.RecordCount + 1
For Each NomeControllo In ElencoControlli
 If TypeOf Me(NomeControllo) Is VB.TextBox Then
  Me(NomeControllo).Text = ""
 End If
Next
TxtCap.Text = ""
BtnSucc.Enabled = False
BtnCanc.Enabled = False: BtnNuovo.Enabled = False
If rsFornitori.RecordCount >= 1 Then
 BtnPrimo.Enabled = True: BtnPrec.Enabled = True: BtnUltimo.Enabled = True
End If
StatoFornitore = NonModificato
End Sub
Private Sub VisualizzaRecord()
TxtDitta.Text = rsFornitori.Fields("ditta")
TxtPartIva.Text = rsFornitori.Fields("partitaiva")
TxtCodFisc.Text = rsFornitori.Fields("codfiscale")
TxtIndirizzo.Text = rsFornitori.Fields("indirizzo")
TxtCap.Text = rsFornitori.Fields("cap")
TxtTel.Text = rsFornitori.Fields("tel")
TxtFax.Text = rsFornitori.Fields("fax")
TxtEmail.Text = rsFornitori.Fields("email")
TxtLoc.Text = rsFornitori.Fields("loc")
TxtProv.Text = rsFornitori.Fields("prov")
TxtStato.Text = rsFornitori.Fields("stato")
StatoFornitore = NonModificato
End Sub
Public Sub ImpostaFiltroRicerca(rsRicerca As ADODB.Recordset)
Set rsFornitori = rsRicerca
FiltroRicerca = True
End Sub
Public Sub CaricaFiltroRicerca()
PosizioneRecord "primo"
End Sub
Private Sub ModificaFornitore()
If CInt(TxtRecCorr.Text) > rsFornitori.RecordCount Then
 If Not FiltroRicerca Then
  StatoFornitore = inserimento: BtnSucc.Enabled = True
  BtnCanc.Enabled = True: BtnNuovo.Enabled = True
 End If
Else
 StatoFornitore = modifica
End If
End Sub
Public Function FormBloccata() As Boolean
FormBloccata = StatoFornitore <> NonModificato
End Function
