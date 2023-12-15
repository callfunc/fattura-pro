VERSION 5.00
Begin VB.Form Vettori 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vettori"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8220
   Icon            =   "Vettori.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556.5
   ScaleMode       =   0  'User
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCanc 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3975
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Vettori.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4965
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnNuovo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3510
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Vettori.frx":4162
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4965
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnUltimo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3165
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Vettori.frx":46FC
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4965
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnSucc 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2820
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Vettori.frx":4C96
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4965
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
      Left            =   1650
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4965
      Width           =   1125
   End
   Begin VB.CommandButton BtnPrec 
      Appearance      =   0  'Flat
      DisabledPicture =   "Vettori.frx":5230
      Enabled         =   0   'False
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
      Left            =   1260
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Vettori.frx":57CA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4965
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnPrimo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   915
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Vettori.frx":5D64
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4965
      UseMaskColor    =   -1  'True
      Width           =   345
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
      TabIndex        =   6
      Top             =   2715
      Width           =   540
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
      Top             =   1095
      Width           =   3150
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
      Top             =   690
      Width           =   2535
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
      TabIndex        =   4
      Top             =   1905
      Width           =   4440
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
      TabIndex        =   8
      Top             =   3525
      Width           =   2550
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
      Top             =   2310
      Width           =   1710
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
      Top             =   1500
      Width           =   4440
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
      Top             =   285
      Width           =   6345
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
      TabIndex        =   10
      Top             =   4350
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
      TabIndex        =   9
      Top             =   3930
      Width           =   2550
   End
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
      TabIndex        =   7
      Top             =   3120
      Width           =   2550
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
      Left            =   255
      TabIndex        =   30
      Top             =   4995
      Width           =   600
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
      Left            =   4410
      TabIndex        =   29
      Top             =   5010
      Width           =   45
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
      TabIndex        =   21
      Top             =   1125
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
      Left            =   255
      TabIndex        =   20
      Top             =   720
      Width           =   825
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
      TabIndex        =   19
      Top             =   3570
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
      Left            =   270
      TabIndex        =   18
      Top             =   2355
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
      Left            =   270
      TabIndex        =   17
      Top             =   1530
      Width           =   705
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
      TabIndex        =   16
      Top             =   315
      Width           =   540
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
      TabIndex        =   15
      Top             =   2745
      Width           =   780
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
      TabIndex        =   14
      Top             =   1935
      Width           =   660
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
      TabIndex        =   13
      Top             =   4380
      Width           =   555
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
      TabIndex        =   12
      Top             =   3975
      Width           =   300
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
      TabIndex        =   11
      Top             =   3165
      Width           =   450
   End
End
Attribute VB_Name = "Vettori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVettori As ADODB.Recordset, StatoVettore As StatoRecord
Dim ElencoControlli As Variant, DescControlli As Variant, PropControlli As Variant
Dim FiltroRicerca As Boolean
Public FormChiamante As Form
Private Sub BtnCanc_Click()
Dim Scelta%
Scelta = MsgBox("Cancellare questo vettore ?" & vbNewLine & "Non sarà possibile" _
& " annullare questa modifica !", vbYesNo + vbQuestion, "Fattura Pro")
If Scelta = vbYes Then
 Dim PosRec As Integer
 PosRec = rsVettori.AbsolutePosition
 If StatoVettore <> inserimento Then
  rsVettori("Rimosso") = True
  rsVettori.Update
  rsVettori.Requery
  If PosRec > rsVettori.RecordCount Then
   If rsVettori.RecordCount <> 0 Then
    rsVettori.MoveLast
   End If
   CreaNuovoRecord
  Else
   rsVettori.Move PosRec - 1, adBookmarkFirst
   VisualizzaRecord
   TxtRecCorr.Text = "di " & rsVettori.AbsolutePosition
   LblNumRecord.Caption = "di " & rsVettori.RecordCount
  End If
 Else
  CreaNuovoRecord
 End If
End If
End Sub
Private Sub TxtCap_Change()
ModificaVettore
End Sub
Private Sub TxtCodFisc_Change()
ModificaVettore
End Sub
Private Sub TxtCodFisc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(KeyAscii))
If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVXYWZ" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub TxtEmail_Change()
ModificaVettore
End Sub
Private Sub TxtFax_Change()
ModificaVettore
End Sub
Private Sub TxtRecCorr_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtRecCorr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtDitta_Change()
ModificaVettore
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
If TxtDitta.SelStart = 0 Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtLoc_Change()
ModificaVettore
End Sub
Private Sub TxtLoc_KeyPress(KeyAscii As Integer)
If TxtLoc.SelStart = 0 Or Mid(TxtLoc.Text, TxtLoc.SelStart + 1, 1) = "." Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtPartIva_Change()
ModificaVettore
End Sub
Private Sub TxtPartIva_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtPartIva.Text) = 18 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub TxtProv_Change()
ModificaVettore
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
ModificaVettore
End Sub
Private Sub TxtTel_Change()
ModificaVettore
End Sub
Private Sub TxtIndirizzo_Change()
ModificaVettore
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
If rsVettori Is Nothing Then
 Set rsVettori = New ADODB.Recordset
 rsVettori.Open "SELECT * FROM Vettori WHERE Rimosso = False ORDER BY Ditta ASC", conn, adOpenDynamic, adLockOptimistic
End If

ElencoControlli = Array("TxtDitta", "TxtPartIva", "TxtCodFisc", "TxtIndirizzo", "TxtTel", _
"TxtFax", "TxtEmail", "TxtLoc", "TxtProv", "TxtStato")
DescControlli = Array("Ditta", "Partita Iva", "Cod. Fiscale", "Indirizzo", "Telefono", _
"Fax", "Email", "Località", "Provincia", "Stato")
PropControlli = Array("ao", "o", "a", "ao", "n", "n", "a", "ao", "a", "a")

If FormChiamante Is Nothing Then
 BtnPrec.Enabled = False: BtnPrimo.Enabled = False
 BtnUltimo.Enabled = rsVettori.RecordCount > 1
 TxtRecCorr.Text = "1"

 If Not rsVettori.EOF Then
  If rsVettori.RecordCount >= 1 Then
   BtnSucc.Enabled = True: BtnNuovo.Enabled = True
   BtnCanc.Enabled = True
  End If
  LblNumRecord.Caption = "di " & rsVettori.RecordCount
  rsVettori.MoveFirst
  Call VisualizzaRecord
 Else
  LblNumRecord.Caption = "di 1"
  CreaNuovoRecord
 End If
Else
 Me.Caption = "Inserisci Vettore"
 LblNumRecord.Caption = "di 1"
 CreaNuovoRecord
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If StatoVettore <> NonModificato Then
 Dim Scelta As VbMsgBoxResult
 Scelta = MsgBox("Vuoi salvare il record corrente ?", _
 vbYesNoCancel + vbQuestion, "Chiusura Archivio Vettori")
 If Scelta = vbYes Then
  If Not ConvalidaRecord(True) Then
   Cancel = 1
  End If
 ElseIf Scelta = vbCancel Then
  Cancel = 1
 End If
End If
If Cancel <> 1 Then
 Set rsVettori = Nothing
 Set Vettori = Nothing
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
If StatoVettore <> NonModificato Then
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
 If StatoVettore = inserimento Then
  CercaDuplicato = True
 ElseIf TxtDitta <> rsVettori("Ditta") Then
  CercaDuplicato = True
 End If
 If CercaDuplicato Then
  Dim rsDuplicato As New ADODB.Recordset
  rsDuplicato.Open "SELECT * FROM Vettori WHERE Rimosso = False And UCASE(Ditta) = '" & _
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
 If StatoVettore = inserimento Then
  rsVettori.AddNew
 End If
 With rsVettori
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
  .Update
 End With
 If StatoVettore = inserimento Then
  CancellaUgualiRimossi
 End If
 EseguiBackup = True
 LblNumRecord.Caption = rsVettori.RecordCount
 StatoVettore = NonModificato
End If
ConvalidaRecord = True
End Function
Private Sub CancellaUgualiRimossi()
Dim rsRimossi As New ADODB.Recordset
rsRimossi.Open "SELECT * FROM Vettori WHERE Rimosso = True And UCASE(Ditta) = '" & _
UCase(Replace(TxtDitta.Text, "'", "''")) & "'", conn, adOpenDynamic, adLockOptimistic
If Not rsRimossi.EOF Then
 rsRimossi.MoveFirst
 conn.Execute "UPDATE FattureClienti SET FattureClienti.IdVettore = " & rsVettori("Id") & " WHERE " _
 & "IdDitta = " & rsRimossi("Id")
 rsRimossi.Delete
 rsVettori.Requery
 rsVettori.Move rsVettori.RecordCount - 1, adBookmarkFirst
End If
rsRimossi.Close
End Sub
Private Sub PosizioneRecord(PosRecord As String)
Dim valido: valido = ConvalidaRecord(True)
If valido Then
 Select Case PosRecord
 Case "primo"
  rsVettori.MoveFirst
 Case "ultimo"
  rsVettori.MoveLast
 Case "precedente"
  If CInt(TxtRecCorr.Text) <= rsVettori.RecordCount Then
   rsVettori.MovePrevious
  End If
 Case "successivo"
  If Not FormChiamante Is Nothing Then
   Unload Me
  Else
   If rsVettori.AbsolutePosition = rsVettori.RecordCount Then
    CreaNuovoRecord
    Exit Sub
   End If
   rsVettori.MoveNext
  End If
 Case "nuovo"
  CreaNuovoRecord
  Exit Sub
 End Select
 BtnCanc.Enabled = True
 If rsVettori.AbsolutePosition <= rsVettori.RecordCount Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
 End If
 If rsVettori.AbsolutePosition <> rsVettori.RecordCount Then
  BtnUltimo.Enabled = True
 Else
  BtnUltimo.Enabled = False
 End If
 If rsVettori.AbsolutePosition <> 1 Then
  BtnPrimo.Enabled = True: BtnPrec.Enabled = True
 Else
  BtnPrimo.Enabled = False: BtnPrec.Enabled = False
 End If
 TxtRecCorr.Text = rsVettori.AbsolutePosition
 LblNumRecord.Caption = "di " & rsVettori.RecordCount
 VisualizzaRecord
End If
End Sub
Private Sub CreaNuovoRecord()
TxtRecCorr.Text = rsVettori.RecordCount + 1: LblNumRecord.Caption = "di " & rsVettori.RecordCount + 1
For Each NomeControllo In ElencoControlli
 If TypeOf Me(NomeControllo) Is VB.TextBox Then
  Me(NomeControllo).Text = ""
 End If
Next
TxtCap.Text = ""
BtnSucc.Enabled = False
BtnCanc.Enabled = False: BtnNuovo.Enabled = False
If rsVettori.RecordCount >= 1 And FormChiamante Is Nothing Then
 BtnPrimo.Enabled = True: BtnPrec.Enabled = True: BtnUltimo.Enabled = True
End If
StatoVettore = NonModificato
End Sub
Private Sub VisualizzaRecord()
TxtDitta.Text = rsVettori.Fields("ditta")
TxtPartIva.Text = rsVettori.Fields("partitaiva")
TxtCodFisc.Text = rsVettori.Fields("codfiscale")
TxtIndirizzo.Text = rsVettori.Fields("indirizzo")
TxtCap.Text = rsVettori.Fields("cap")
TxtTel.Text = rsVettori.Fields("tel")
TxtFax.Text = rsVettori.Fields("fax")
TxtEmail.Text = rsVettori.Fields("email")
TxtLoc.Text = rsVettori.Fields("loc")
TxtProv.Text = rsVettori.Fields("prov")
TxtStato.Text = rsVettori.Fields("stato")
StatoVettore = NonModificato
End Sub
Public Sub ImpostaFiltroRicerca(rsRicerca As ADODB.Recordset)
Set rsVettori = rsRicerca
FiltroRicerca = True
End Sub
Public Sub CaricaFiltroRicerca()
PosizioneRecord "primo"
End Sub
Private Sub ModificaVettore()
If CInt(TxtRecCorr.Text) > rsVettori.RecordCount Then
 If Not FiltroRicerca Then
  StatoVettore = inserimento: BtnSucc.Enabled = True
  BtnCanc.Enabled = True:
  If FormChiamante Is Nothing Then
   BtnNuovo.Enabled = True
  End If
 End If
Else
 StatoVettore = modifica
End If
End Sub
Public Function FormBloccata() As Boolean
FormBloccata = StatoVettore <> NonModificato
End Function
