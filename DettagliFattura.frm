VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FatturaDiff 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fattura Differita"
   ClientHeight    =   7365
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   14070
   Icon            =   "DettagliFattura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   14070
   Begin VB.TextBox TotFatt 
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
      Left            =   1500
      TabIndex        =   12
      Top             =   5655
      Width           =   1410
   End
   Begin VB.TextBox TotImp 
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
      Left            =   4680
      TabIndex        =   11
      Top             =   5655
      Width           =   1410
   End
   Begin VB.TextBox TotIva 
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
      Left            =   7305
      TabIndex        =   10
      Top             =   5655
      Width           =   1500
   End
   Begin MSFlexGridLib.MSFlexGrid ElencoNote 
      Height          =   5025
      Left            =   90
      TabIndex        =   9
      Top             =   540
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   8864
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   14145495
      BackColorBkg    =   24576
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CreaFattura 
      Caption         =   "Crea Fattura"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   225
      MaskColor       =   &H00D8E9EC&
      Picture         =   "DettagliFattura.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6255
      UseMaskColor    =   -1  'True
      Width           =   1740
   End
   Begin VB.CommandButton Anteprima 
      Caption         =   "Anteprima Fattura"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   4140
      MaskColor       =   &H00D8E9EC&
      Picture         =   "DettagliFattura.frx":4136
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6255
      UseMaskColor    =   -1  'True
      Width           =   1755
   End
   Begin VB.CommandButton Stampa 
      Caption         =   "Stampa Fattura"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   2175
      MaskColor       =   &H00FFFFFF&
      Picture         =   "DettagliFattura.frx":4442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6255
      UseMaskColor    =   -1  'True
      Width           =   1770
   End
   Begin VB.TextBox TxtCliente 
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
      Left            =   6390
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox Data 
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
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox TxtNumFattura 
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
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Fattura:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   225
      TabIndex        =   15
      Top             =   5685
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Imponibile:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3120
      TabIndex        =   14
      Top             =   5685
      Width           =   1500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Iva:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6375
      TabIndex        =   13
      Top             =   5685
      Width           =   870
   End
   Begin VB.Line Line1 
      X1              =   6510
      X2              =   6525
      Y1              =   5490
      Y2              =   5505
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      TabIndex        =   3
      Top             =   150
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
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
      Left            =   3495
      TabIndex        =   2
      Top             =   150
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Documento:"
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
      Left            =   195
      TabIndex        =   0
      Top             =   150
      Width           =   1485
   End
End
Attribute VB_Name = "FatturaDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IntestazioniGriglia As Variant, re As New RegExp, RigaCorr As Integer
Public rsFatture As ADODB.Recordset, rsNoteFattura As ADODB.Recordset, rsVociNote As ADODB.Recordset, _
rsCliente As ADODB.Recordset, TotaliFattura As Collection
Dim inserita As Boolean, cm As CoordinateMouse
Private Sub Anteprima_Click()
If Not ModificaNote Then
 Dim DocValido As Boolean: DocValido = True
 Call ConvalidaDocumento(DocValido)
 If DocValido Then
  CaricaTotaliIva
  Dim SS As New ServiziStampa
  Call CaricaInfoDitta(SS)
  Set SS.rsDoc = rsFatture: rsVociNote.MoveFirst
  Set SS.rsVociDoc = rsVociNote: SS.TipoDoc = FatturaDifferita
  Set SS.TotaliDoc = TotaliFattura: Set SS.rsCliente = rsCliente
  SS.ImpostaAnteprima True: SS.Stampa -1
  AnteprimaDoc.Show vbModal
 End If
End If
End Sub
Private Sub CaricaTotaliIva()
Dim Totale As Double, Iva As Double, Aliquota As Double
Dim TI As TotaleIva
Set TotaliFattura = New Collection
If rsVociNote.RecordCount <> 0 Then
 rsVociNote.MoveFirst
 On Error Resume Next
 While Not rsVociNote.EOF
  If rsVociNote("iva") <> "" Then
   Aliquota = CDbl(rsVociNote("iva"))
   Totale = CDbl(rsVociNote("totale"))
   Iva = CDbl(Arrotonda(Totale * Aliquota / 100))
   Set TI = TotaliFattura(CStr(Aliquota))
   If Err.Number = 0 Then
    TI.Imponibile = TI.Imponibile + Totale
    TI.Iva = TI.Iva + Iva
    TI.Totale = TI.Imponibile + TI.Iva
   Else
   Set TI = New TotaleIva
    TI.Aliquota = Aliquota
    TI.Imponibile = TI.Imponibile + Totale
    TI.Iva = TI.Iva + Iva
    TI.Totale = TI.Imponibile + TI.Iva
    TotaliFattura.Add TI, CStr(Aliquota)
    Err.Clear
   End If
  End If
  rsVociNote.MoveNext
 Wend
End If
If Split(TotDoc, ",")(1) = "99" Then TotDoc = FormatNumber(CDbl(TotDoc) + 0.01, 2)
For Each TotaleIva In TotaliFattura
 If InStr(1, CStr(TotaleIva.Totale), ",99") Then
  TotaleIva.Totale = TotaleIva.Totale + 0.01
 End If
Next
End Sub
Private Sub ElencoNote_Click()
rsNoteFattura.Move ElencoNote.Row - 1, adBookmarkFirst
End Sub
Private Sub TxtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub CreaFattura_Click()
Dim DocValido As Boolean: DocValido = True
Call ConvalidaDocumento(DocValido)
If Not inserita And DocValido Then
 rsFatture.Update
 rsNoteFattura.MoveFirst
 While Not rsNoteFattura.EOF
  rsNoteFattura("IdFatt") = rsFatture("IdDoc")
  rsNoteFattura.Update
  rsNoteFattura.MoveNext
 Wend
 inserita = True
 Note.SalvaNota True
 EseguiBackup = True
 MsgBox "Documento creato con successo !", vbInformation, "Fattura Pro"
End If
End Sub
Private Sub SelezionaRiga(ByVal riga As Long, ByVal Seleziona As Boolean)
ElencoNote.Row = riga
If Seleziona Then
 For i = 0 To ElencoNote.Cols - 1
  ElencoNote.Col = i: ElencoNote.CellBackColor = RGB(49, 106, 197)
  ElencoNote.CellForeColor = RGB(255, 255, 255)
 Next i
 If riga <> ElencoNote.Rows - 1 Then
  Set ElencoNote.CellPicture = CaricaImmagineDaRisorsa("CANCELLA", , , RGB(49, 106, 197))
 End If
Else
 For i = 0 To ElencoNote.Cols - 1
  ElencoNote.Col = i: ElencoNote.CellBackColor = RGB(255, 255, 255)
  ElencoNote.CellForeColor = RGB(0, 0, 0)
 Next i
 If riga <> ElencoNote.Rows - 1 Then Set ElencoNote.CellPicture = ImgCancella
End If
End Sub
Private Sub Form_Load()
FatturaDiff.Move 300, 300
IntestazioniGriglia = Array("N° Doc.", "Data", "Totale")
ElencoNote.ColWidth(0) = 1000: ElencoNote.ColWidth(1) = 1400
ElencoNote.ColWidth(2) = 1300

For i = 0 To ElencoNote.Cols - 1
 ElencoNote.TextMatrix(0, i) = IntestazioniGriglia(i): ElencoNote.ColAlignment(i) = 4
Next i

Dim TipoPag As String
TxtCliente.Text = rsCliente("Ditta")
TotFatt.Text = FormatNumber(rsFatture("TotDoc"), 2)
TotImp.Text = FormatNumber(rsFatture("TotImp"), 2)
TotIva.Text = FormatNumber(rsFatture("TotIva"), 2)
RigaCorr = 0
Set rsVociNota = New ADODB.Recordset
While Not rsNoteFattura.EOF
 With ElencoNote
 .AddItem ""
 .TextMatrix(.Rows - 1, 0) = NumeroDocumento(rsNoteFattura("IdDoc"))
 .TextMatrix(.Rows - 1, 1) = Format$(rsNoteFattura("data"), "dd-mm-yyyy")
 .TextMatrix(.Rows - 1, 2) = FormatNumber(rsNoteFattura("totdoc"), 2)
 .RowHeight(.Rows - 1) = 315
 rsNoteFattura.MoveNext
 End With
Wend
End Sub
Private Sub MostraMessaggioErrore(ByVal Messaggio$, Optional Controllo As TextBox = Nothing)
MsgBox Messaggio, vbExclamation, "Fattura Pro"
If Not Controllo Is Nothing Then
 Controllo.SetFocus: Controllo.SelStart = 0: Controllo.SelLength = Len(Data.Text)
End If
End Sub
Private Sub ConvalidaDocumento(DocValido As Boolean)
Dim d As Variant, id, e As Boolean
d = Split(Data.Text, "-")
If IsDate(Data.Text) Then
 If UBound(d) = 2 Then
  If Len(d(0)) < 2 Then d(0) = "0" & d(0)
  If Len(d(1)) < 2 Then d(1) = "0" & d(1)
  Data.Text = d(0) & "-" & d(1) & "-" & d(2)
 Else: e = True
 End If
Else: e = True
End If
If e Then
 MostraMessaggioErrore "Attenzione, inserire una data valida !", Data
 DocValido = False: Exit Sub
End If
If Not inserita Then
 Dim rsDuplicato As New ADODB.Recordset, IdDoc$
 IdDoc = Mid(Data.Text, 7, 4) & "-" & TxtNumFattura
 rsDuplicato.Open "SELECT * FROM FattureClienti WHERE IdDoc = '" & IdDoc & "'", conn, adOpenDynamic
 If Not rsDuplicato.EOF Then
  MostraMessaggioErrore "Attenzione, il numero documento corrisponde a " _
  & "quello di un altro documento in archivio !", TxtNumFattura
  DocValido = False: Exit Sub
 End If
 Dim re As New RegExp
 re.Global = True
 re.Pattern = "^[1-9][0-9]*/?[1-9]?$"
 If Not re.Test(TxtNumFattura.Text) Then
  MostraMessaggioErrore "Attenzione, inserire un numero documento valido !", TxtNumFattura
  DocValido = False: Exit Sub
 End If
 rsFatture("IdDoc") = Mid(Data.Text, 7) & "-" & TxtNumFattura
 rsFatture("IdDitta") = rsCliente("Id")
 rsFatture("Data") = Data.Text
 rsFatture("Modpag") = rsCliente("Modpag")
 rsFatture("TotDoc") = CDbl(TotFatt.Text)
 rsFatture("TotImp") = CDbl(TotImp.Text)
 rsFatture("TotIva") = CDbl(TotIva.Text)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not inserita Then
 rsFatture.CancelUpdate
 Note.SalvaNota False
End If
rsFatture.Close: rsNoteFattura.Close
rsVociNote.Close
Set FatturaDiff = Nothing
End Sub
Private Sub TxtNumFattura_KeyPress(KeyAscii As Integer)
If InStr("0123456789/" & vbBack, Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub ElencoNote_DblClick()
If MouseInGriglia(ElencoNote) Then
 Dim IdNota$
 IdNota = rsNoteFattura("IdDoc")
 Note.IdDoc = IdNota
 Set Note.FormChiamante = Me
 Note.Show
 Sleep 1
End If
End Sub
Private Sub ElencoNote_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cm.x = x: cm.y = y
End Sub
Public Sub AggiornaTotali(ByVal IdNota As String, NuovaNota As Boolean)
Dim TotImpNota#, TotIvaNota#, riga%
TotImpNota = rsNoteFattura("TotImp"): TotIvaNota = rsNoteFattura("TotIva")
rsNoteFattura.Requery
rsVociNote.Requery
rsNoteFattura.Find "IdDoc = '" & IdNota & "'", , adSearchForward, adBookmarkFirst

If NuovaNota Then
 TotImp = FormatNumber(CDbl(TotImp) + rsNoteFattura("TotImp"), 2)
 TotIva = FormatNumber(CDbl(TotIva) + rsNoteFattura("TotIva"), 2)
 ElencoNote.AddItem ""
 ElencoNote.RowHeight(ElencoNote.Rows - 1) = 315
 riga = ElencoNote.Rows - 1
Else
 TotImp = FormatNumber(CDbl(TotImp) + (rsNoteFattura("TotImp") - TotImpNota), 2)
 TotIva = FormatNumber(CDbl(TotIva) + (rsNoteFattura("TotIva") - TotIvaNota), 2)
 riga = ElencoNote.Row
End If
TotDoc = TotImp + TotIva

ElencoNote.TextMatrix(riga, 0) = NumeroDocumento(IdNota)
ElencoNote.TextMatrix(riga, 1) = Format$(rsNoteFattura("data"), "dd-mm-yyyy")
ElencoNote.TextMatrix(riga, 2) = FormatNumber(rsNoteFattura("totdoc"), 2)
TotFatt.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text), 2)
If InStr(1, TotFatt.Text, ",99") Then TotFatt.Text = FormatNumber(CDbl(TotFatt.Text) + 0.01, 2)
End Sub
Private Sub Stampa_Click()
If Not ModificaNote Then
 Dim DocValido As Boolean: DocValido = True
 Call ConvalidaDocumento(DocValido)
 If DocValido Then
  CaricaTotaliIva
  Dim SS As New ServiziStampa
  Call CaricaInfoDitta(SS)
  Set SS.rsDoc = rsFatture: rsVociNote.MoveFirst
  SS.TipoDoc = FatturaDifferita: Set SS.rsVociDoc = rsVociNote
  Set SS.TotaliDoc = TotaliFattura: Set SS.rsCliente = rsCliente
  OpzioniStampa.Inizializza SS: OpzioniStampa.Show vbModal
  If OpzioniStampa.Scelta = "Stampa" Then
   SS.ImpostaAnteprima False: SS.Stampa -1
  End If
 End If
End If
End Sub
Private Sub TotFatt_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotFatt_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TotImp_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotImp_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TotIva_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotIva_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub CaricaInfoDitta(SS As ServiziStampa)
Dim rsInfoDitta As New ADODB.Recordset
rsInfoDitta.Open "SELECT * FROM InfoDitta", conn, adOpenDynamic, adLockOptimistic
If Not rsInfoDitta.EOF Then
 SS.InfoDitta = rsInfoDitta.GetRows(1, adBookmarkFirst)
End If
rsInfoDitta.Close
End Sub
Private Function MouseInGriglia(Griglia As MSFlexGrid) As Boolean
MouseInGriglia = cm.y >= (Griglia.RowPos(0) + Griglia.RowHeight(0)) And cm.y <= _
(Griglia.RowPos(Griglia.Rows - 1) + Griglia.RowHeight(Griglia.Rows - 1)) And cm.x _
<= (Griglia.ColPos(Griglia.Cols - 1) + Griglia.ColWidth(Griglia.Cols - 1))
End Function
Public Sub ResetModificaNote()
ModificaNote = False
End Sub
