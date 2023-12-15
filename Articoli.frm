VERSION 5.00
Begin VB.Form Articoli 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Articoli"
   ClientHeight    =   2580
   ClientLeft      =   345
   ClientTop       =   645
   ClientWidth     =   7830
   Icon            =   "Articoli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   7830
   Begin VB.TextBox TxtGiacenzaIniziale 
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
      Left            =   5640
      TabIndex        =   24
      Top             =   1620
      Width           =   1605
   End
   Begin VB.TextBox TxtQntDisp 
      Enabled         =   0   'False
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
      Left            =   5640
      TabIndex        =   22
      Top             =   1170
      Width           =   1605
   End
   Begin VB.TextBox TxtPrezzo1 
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
      Left            =   1245
      TabIndex        =   1
      Top             =   555
      Width           =   1365
   End
   Begin VB.TextBox TxtPrezzo2 
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
      Left            =   3555
      TabIndex        =   2
      Top             =   555
      Width           =   1350
   End
   Begin VB.TextBox TxtPrezzo3 
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
      Left            =   5865
      TabIndex        =   3
      Top             =   555
      Width           =   1365
   End
   Begin VB.CommandButton BtnCanc 
      Height          =   315
      Left            =   3855
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Articoli.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.ListBox LstAliqIva 
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
      ItemData        =   "Articoli.frx":4162
      Left            =   1245
      List            =   "Articoli.frx":4164
      TabIndex        =   4
      Top             =   975
      Width           =   810
   End
   Begin VB.ListBox LstUm 
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
      ItemData        =   "Articoli.frx":4166
      Left            =   2760
      List            =   "Articoli.frx":4179
      TabIndex        =   5
      Top             =   990
      Width           =   840
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
      Left            =   885
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Articoli.frx":418F
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnPrec 
      Appearance      =   0  'Flat
      DisabledPicture =   "Articoli.frx":4729
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
      Left            =   1230
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Articoli.frx":4CC3
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
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
      Left            =   1590
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1125
   End
   Begin VB.CommandButton BtnSucc 
      Height          =   315
      Left            =   2745
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Articoli.frx":525D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnUltimo 
      Height          =   315
      Left            =   3090
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Articoli.frx":57F7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnNuovo 
      Height          =   315
      Left            =   3435
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Articoli.frx":5D91
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TxtDescr 
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
      Left            =   1245
      TabIndex        =   0
      Top             =   120
      Width           =   6225
   End
   Begin VB.Label LblGiacenza 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Giacenza iniziale:"
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
      Left            =   3930
      TabIndex        =   23
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantità Disponibile:"
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
      Left            =   3930
      TabIndex        =   21
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prezzo 1:"
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
      Left            =   510
      TabIndex        =   20
      Top             =   570
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prezzo 2:"
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
      Left            =   2805
      TabIndex        =   19
      Top             =   585
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prezzo 3:"
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
      Left            =   5115
      TabIndex        =   18
      Top             =   585
      Width           =   690
   End
   Begin VB.Label LblUM 
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
      Left            =   2280
      TabIndex        =   16
      Top             =   990
      Width           =   420
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
      Left            =   4275
      TabIndex        =   15
      Top             =   2205
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
      Left            =   225
      TabIndex        =   14
      Top             =   2190
      Width           =   600
   End
   Begin VB.Label Label4 
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
      Left            =   165
      TabIndex        =   7
      Top             =   975
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrizione:"
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
      TabIndex        =   6
      Top             =   150
      Width           =   945
   End
End
Attribute VB_Name = "Articoli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAliquoteIVA As ADODB.Recordset, StatoArticolo As StatoRecord
Dim ElencoControlli As Variant, DescControlli As Variant, PropControlli As Variant
Dim FiltroRicerca As Boolean
Dim rsArticoli As ADODB.Recordset
Private Sub BtnCanc_Click()
Dim Scelta%
Scelta = MsgBox("Cancellare l' articolo ?" & vbNewLine & "Non sarà possibile" _
& " annullare questa modifica !", vbYesNo + vbQuestion, "Fattura Pro")
If Scelta = vbYes Then
 Dim PosRec As Integer
 PosRec = rsArticoli.AbsolutePosition
 If StatoArticolo <> inserimento Then
  rsArticoli.Delete
  rsArticoli.Update
  If PosRec > rsArticoli.RecordCount Then
   If rsArticoli.RecordCount <> 0 Then
    rsArticoli.MoveLast
   End If
   CreaNuovoRecord
  Else
   rsArticoli.MoveNext
   VisualizzaRecord
   TxtRecCorr.Text = rsArticoli.AbsolutePosition
   LblNumRecord.Caption = "di " & rsArticoli.RecordCount
  End If
 Else
  CreaNuovoRecord
 End If
End If
End Sub
Private Sub TxtGiacenzaIniziale_Change()
TxtQntDisp.Text = TxtGiacenzaIniziale.Text
ModificaArticolo
End Sub
Private Sub TxtGiacenzaIniziale_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub TxtRecCorr_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtRecCorr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtDescr_Change()
ModificaArticolo
End Sub
Private Sub TxtDescr_KeyPress(KeyAscii As Integer)
If TxtDescr = "" Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
Else
 If Right(TxtDescr.Text, 1) <> "." And Right(TxtDescr.Text, 1) <> " " Then KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If
End Sub
Private Sub LstAliqIva_Click()
ModificaArticolo
End Sub
Private Sub TxtPrezzo1_Change()
ModificaArticolo
End Sub
Private Sub TxtPrezzo1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If (TxtPrezzo1.SelStart = 0 Or TxtPrezzo1.SelStart = Len(TxtPrezzo1.Text) - 1) _
 And KeyAscii = 44 Then
  KeyAscii = 0
  Exit Sub
 End If
 If InStr(1, "0123456789,", Chr(KeyAscii)) = 0 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub TxtPrezzo2_Change()
ModificaArticolo
End Sub
Private Sub TxtPrezzo2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If (TxtPrezzo2.SelStart = 0 Or TxtPrezzo2.SelStart = Len(TxtPrezzo2.Text) - 1) _
 And KeyAscii = 44 Then
  KeyAscii = 0
  Exit Sub
 End If
 If InStr(1, "0123456789,", Chr(KeyAscii)) = 0 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub TxtPrezzo3_Change()
ModificaArticolo
End Sub
Private Sub TxtPrezzo3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If (TxtPrezzo3.SelStart = 0 Or TxtPrezzo3.SelStart = Len(TxtPrezzo3.Text) - 1) _
 And KeyAscii = 44 Then
  KeyAscii = 0
  Exit Sub
 End If
 If InStr(1, "0123456789,", Chr(KeyAscii)) = 0 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub LstUm_Click()
ModificaArticolo
End Sub
Private Sub Form_Load()
Me.Move 600, 450

Set rsAliquoteIVA = New ADODB.Recordset
If rsArticoli Is Nothing Then
 Set rsArticoli = New ADODB.Recordset
 rsArticoli.Open "SELECT * FROM Articoli ORDER BY Descr ASC", conn, adOpenDynamic, adLockOptimistic
End If
rsAliquoteIVA.Open "SELECT * FROM AliquoteIVA WHERE Aliquota <> 0 ORDER BY Aliquota ASC", conn, adOpenDynamic
While Not rsAliquoteIVA.EOF
 LstAliqIva.AddItem rsAliquoteIVA("Aliquota")
 rsAliquoteIVA.MoveNext
Wend

ElencoControlli = Array("TxtDescr", "TxtPrezzo1", "TxtPrezzo2", "TxtPrezzo3", "LstUm", "LstAliqIva")
DescControlli = Array("Descrizione", "Prezzo 1", "Prezzo 2", "Prezzo 3", "U.M.", "Aliquota IVA")
PropControlli = Array("ao", "no", "n", "n", "ao", "no")

BtnPrec.Enabled = False: BtnPrimo.Enabled = False
BtnUltimo.Enabled = False

TxtRecCorr.Text = "1"
If Not rsArticoli.EOF Then
 If rsArticoli.RecordCount >= 1 Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
  BtnCanc.Enabled = True
 End If
 LblNumRecord.Caption = "di " & rsArticoli.RecordCount
 rsArticoli.MoveFirst
 Call VisualizzaRecord
Else
 CreaNuovoRecord
 LblNumRecord.Caption = "di 1"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If StatoArticolo <> NonModificato Then
 Dim Scelta As VbMsgBoxResult
 Scelta = MsgBox("Vuoi salvare il record corrente ?", _
 vbYesNoCancel + vbQuestion, "Chiusura Archivio Articoli")
 If Scelta = vbYes Then
  If Not ConvalidaRecord() Then
   Cancel = 1
  End If
 ElseIf Scelta = vbCancel Then
  Cancel = 1
 End If
End If
If Cancel <> 1 Then
 Set rsArticoli = Nothing
 Set Articoli = Nothing
End If
End Sub
Private Sub BtnPrec_Click()
PosizioneRecord "precedente"
End Sub
Private Sub BtnPrimo_Click()
PosizioneRecord "primo"
End Sub
Private Sub BtnUltimo_Click()
PosizioneRecord "ultimo"
End Sub
Private Sub BtnNuovo_Click()
PosizioneRecord "nuovo"
End Sub
Private Sub BtnSucc_Click()
PosizioneRecord "successivo"
End Sub
Private Function ConvalidaRecord() As Boolean
If StatoArticolo <> NonModificato Then
 For i = 0 To UBound(ElencoControlli)
  If InStr(1, PropControlli(i), "o") <> 0 And Trim(Me(ElencoControlli(i))) = "" Then
   MsgBox "Attenzione, " & DescControlli(i) & " è un campo obbligatorio !", vbExclamation, "Fattura Pro"
   Me(ElencoControlli(i)).Text = "": Exit Function
  ElseIf Trim(Me(ElencoControlli(i))) <> "" Then
   If InStr(1, PropControlli(i), "a") <> 0 And IsNumeric(Trim(Me(ElencoControlli(i)))) Then
    MsgBox "Attenzione, il campo " & DescControlli(i) & " non può contenere valori valori numerici !", _
    vbExclamation, "Fattura Pro": Me(ElencoControlli(i)).Text = "": Exit Function
   ElseIf InStr(1, PropControlli(i), "n") <> 0 And Not IsNumeric(Trim(Me(ElencoControlli(i)))) Then
    MsgBox "Attenzione, il campo " & DescControlli(i) & " deve contenere valori numerici !", vbExclamation, _
    "Fattura Pro"
    Me(ElencoControlli(i)).Text = "": Exit Function
   End If
  End If
 Next i
 Dim CercaDuplicato As Boolean
 If StatoArticolo = inserimento Then
  CercaDuplicato = True
 ElseIf TxtDescr.Text <> rsArticoli("descr") Then
  CercaDuplicato = True
 End If
 If CercaDuplicato Then
  Dim rsDuplicato As New ADODB.Recordset
  rsDuplicato.Open "SELECT * FROM Articoli WHERE descr = '" & TxtDescr.Text & "'", conn, adOpenDynamic
  If Not rsDuplicato.EOF Then
   MsgBox "Attenzione, L'articolo esiste già in archivio !", vbExclamation, "Fattura Pro"
   Exit Function
  End If
 End If
 If StatoArticolo = inserimento Then
  rsArticoli.AddNew
 End If
 With rsArticoli
  .Fields("Descr") = EliminaSpazi(TxtDescr.Text)
  .Fields("Prezzo1") = CDbl(TxtPrezzo1.Text)
  If TxtPrezzo2.Text <> "" Then
   .Fields("Prezzo2") = CDbl(TxtPrezzo2.Text)
  Else
   .Fields("Prezzo2") = 0
  End If
  If TxtPrezzo3.Text <> "" Then
   .Fields("Prezzo3") = CDbl(TxtPrezzo3.Text)
  Else
   .Fields("Prezzo3") = 0
  End If
  .Fields("AliqIva") = LstAliqIva.Text
  .Fields("Um") = Trim(LstUm.Text)
  If StatoArticolo = inserimento Then
   .Fields("QntDisp") = CDbl(TxtQntDisp.Text)
   .Fields("GiacenzaIniziale") = CDbl(TxtGiacenzaIniziale.Text)
  End If
  .Update
 End With
 EseguiBackup = True
 StatoArticolo = NonModificato
End If
ConvalidaRecord = True
End Function
Private Sub PosizioneRecord(PosRecord As String)
Dim valido: valido = ConvalidaRecord()
If valido Then
 Select Case PosRecord
 Case "primo"
  rsArticoli.MoveFirst
 Case "ultimo"
  rsArticoli.MoveLast
 Case "precedente"
  If CInt(TxtRecCorr.Text) <= rsArticoli.RecordCount Then
   rsArticoli.MovePrevious
  End If
 Case "successivo"
  If rsArticoli.AbsolutePosition = rsArticoli.RecordCount Then
   CreaNuovoRecord
   Exit Sub
  End If
  rsArticoli.MoveNext
 Case "nuovo"
  CreaNuovoRecord
  Exit Sub
 End Select
 BtnCanc.Enabled = True
 If rsArticoli.AbsolutePosition <= rsArticoli.RecordCount Then
  BtnSucc.Enabled = True: BtnUltimo.Enabled = True
  BtnNuovo.Enabled = True
 End If
 If rsArticoli.AbsolutePosition <> rsArticoli.RecordCount Then
  BtnUltimo.Enabled = True
 Else
  BtnUltimo.Enabled = False
 End If
 If rsArticoli.AbsolutePosition <> 1 Then
  BtnPrimo.Enabled = True: BtnPrec.Enabled = True
 Else
  BtnPrimo.Enabled = False: BtnPrec.Enabled = False
 End If
 TxtRecCorr.Text = rsArticoli.AbsolutePosition
 LblNumRecord.Caption = "di " & rsArticoli.RecordCount
 VisualizzaRecord
End If
End Sub
Private Sub CreaNuovoRecord()
TxtRecCorr.Text = rsArticoli.RecordCount + 1: LblNumRecord.Caption = rsArticoli.RecordCount + 1
For Each NomeControllo In ElencoControlli
 If TypeOf Me(NomeControllo) Is VB.TextBox Then
  Me(NomeControllo).Text = ""
 End If
Next
LstAliqIva.ListIndex = -1: LstUm.ListIndex = -1
TxtQntDisp.Text = ""
TxtGiacenzaIniziale.Text = ""
TxtGiacenzaIniziale.Enabled = True: StatoArticolo = NonModificato
BtnSucc.Enabled = False: BtnCanc.Enabled = False: BtnNuovo.Enabled = False
If rsArticoli.RecordCount >= 1 Then
 BtnPrimo.Enabled = True: BtnPrec.Enabled = True: BtnUltimo.Enabled = True
End If
End Sub
Private Sub VisualizzaRecord()
TxtDescr.Text = rsArticoli("Descr")
TxtPrezzo1.Text = FormatNumber(rsArticoli("prezzo1"), 2)
TxtPrezzo2.Text = IIf(rsArticoli("prezzo2") <> 0, FormatNumber(rsArticoli("prezzo2"), 2), "")
TxtPrezzo3.Text = IIf(rsArticoli("prezzo3") <> 0, FormatNumber(rsArticoli("prezzo3"), 2), "")
LstAliqIva.ListIndex = SendMessage(LstAliqIva.hWnd, LB_FINDSTRING, -1, ByVal CStr(rsArticoli("AliqIva")))
LstUm.ListIndex = SendMessage(LstUm.hWnd, LB_FINDSTRING, 0&, ByVal CStr(rsArticoli("Um")))
TxtGiacenzaIniziale.Text = FormatNumber(rsArticoli("GiacenzaIniziale"), 3)
TxtQntDisp.Text = FormatNumber(rsArticoli("QntDisp"), 3)
BtnCanc.Enabled = True
TxtGiacenzaIniziale.Enabled = False
StatoArticolo = NonModificato
End Sub
Public Sub ImpostaFiltroRicerca(rsRicerca As ADODB.Recordset)
Set rsArticoli = rsRicerca
FiltroRicerca = True
End Sub
Public Sub CaricaFiltroRicerca()
PosizioneRecord "primo"
End Sub
Private Sub ModificaArticolo()
If CInt(TxtRecCorr.Text) > rsArticoli.RecordCount Then
 If Not FiltroRicerca Then
  StatoArticolo = inserimento: BtnSucc.Enabled = True
  BtnCanc.Enabled = True: BtnNuovo.Enabled = True
 End If
Else
 StatoArticolo = modifica
End If
End Sub
Public Function FormBloccata() As Boolean
FormBloccata = StatoArticolo <> NonModificato
End Function
