VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FattureClienti 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fatture Clienti"
   ClientHeight    =   7185
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   15465
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FattureClienti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   15465
   Begin VB.CommandButton BtnFatturaElettronica 
      Caption         =   "Crea Fattura Elettronica"
      Height          =   405
      Left            =   9090
      TabIndex        =   45
      Top             =   4830
      Width           =   2340
   End
   Begin VB.Frame RiquadroLC 
      BackColor       =   &H0033CCFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   405
      Left            =   75
      TabIndex        =   40
      Top             =   4350
      Width           =   9660
      Begin VB.TextBox TxtLC 
         Height          =   315
         Left            =   3930
         TabIndex        =   43
         Top             =   15
         Width           =   3885
      End
      Begin VB.CommandButton BtnLuoghiConsegna 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7860
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FattureClienti.frx":4072
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox ChkLuogoConsegna 
         BackColor       =   &H0033CCFF&
         Caption         =   "Luogo di Consegna"
         Height          =   285
         Left            =   60
         TabIndex        =   41
         Top             =   45
         Width           =   1875
      End
      Begin VB.Label LblLC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luogo di Consegna:"
         Height          =   225
         Left            =   2280
         TabIndex        =   44
         Top             =   60
         Width           =   1590
      End
   End
   Begin VB.Frame RiquadroEsenteIva 
      BackColor       =   &H0033CCFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   420
      Left            =   60
      TabIndex        =   35
      Top             =   3855
      Width           =   10845
      Begin VB.CheckBox ChkEsenteIva 
         BackColor       =   &H0033CCFF&
         Caption         =   "Esente I.V.A."
         Height          =   240
         Left            =   75
         TabIndex        =   38
         Top             =   120
         Width           =   1440
      End
      Begin VB.TextBox NonImpIva 
         Height          =   315
         Left            =   3945
         TabIndex        =   37
         Top             =   75
         Width           =   5070
      End
      Begin VB.CommandButton BtnEsenzioniIva 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9060
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FattureClienti.frx":41CD
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Label LblEsenzioneIva 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Esenzione I.V.A.:"
         Height          =   225
         Left            =   2610
         TabIndex        =   39
         Top             =   120
         Width           =   1275
      End
   End
   Begin VB.TextBox TxtNumFattura 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Top             =   180
      Width           =   1380
   End
   Begin VB.ComboBox CmbPag 
      Height          =   345
      ItemData        =   "FattureClienti.frx":4328
      Left            =   7875
      List            =   "FattureClienti.frx":4338
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   180
      Width           =   1650
   End
   Begin VB.TextBox TxtPartIva 
      Height          =   315
      Left            =   7875
      TabIndex        =   33
      Top             =   630
      Width           =   1890
   End
   Begin VB.ListBox PopupDitte 
      Height          =   1185
      Left            =   10260
      TabIndex        =   31
      Top             =   5505
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CommandButton BtnSelDitta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6255
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FattureClienti.frx":436E
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   615
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.PictureBox BoxEsenzioneIva 
      Appearance      =   0  'Flat
      BackColor       =   &H0033CCFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   15480
      ScaleHeight     =   720
      ScaleWidth      =   165
      TabIndex        =   29
      Top             =   4035
      Width           =   165
   End
   Begin VB.CommandButton BtnVai 
      Caption         =   "Vai"
      Height          =   315
      Left            =   3270
      TabIndex        =   28
      Top             =   180
      Width           =   465
   End
   Begin VB.CommandButton BtnDdt 
      Caption         =   "Ddt"
      Height          =   360
      Left            =   10860
      TabIndex        =   27
      Top             =   600
      Width           =   1185
   End
   Begin VB.CheckBox Pagato 
      BackColor       =   &H0033CCFF&
      Caption         =   "Saldato"
      Height          =   240
      Left            =   9855
      TabIndex        =   26
      Top             =   315
      Width           =   975
   End
   Begin VB.CommandButton BtnCanc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3810
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureClienti.frx":44C9
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5460
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton Stampa 
      Caption         =   "Stampa Fattura"
      Height          =   990
      Left            =   165
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FattureClienti.frx":45B9
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6030
      UseMaskColor    =   -1  'True
      Width           =   1635
   End
   Begin VB.CommandButton Anteprima 
      Caption         =   "Anteprima Fattura"
      Height          =   990
      Left            =   1965
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureClienti.frx":4B97
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6030
      UseMaskColor    =   -1  'True
      Width           =   1845
   End
   Begin VB.TextBox TxtModifica 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   7455
      TabIndex        =   21
      Top             =   6690
      Visible         =   0   'False
      Width           =   1170
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
      Left            =   810
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureClienti.frx":4EA3
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5460
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnPrec 
      Appearance      =   0  'Flat
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
      Left            =   1155
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureClienti.frx":543D
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5460
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TxtRecCorr 
      Height          =   315
      Left            =   1545
      TabIndex        =   16
      Top             =   5460
      Width           =   1125
   End
   Begin VB.CommandButton BtnSucc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2715
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureClienti.frx":59D7
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5460
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnUltimo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureClienti.frx":5F71
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5460
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnNuovo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3390
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureClienti.frx":650B
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5460
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin MSFlexGridLib.MSFlexGrid ElencoVoci 
      Height          =   2670
      Left            =   45
      TabIndex        =   12
      Top             =   1140
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   4710
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   315
      BackColor       =   16777215
      BackColorFixed  =   14145495
      BackColorSel    =   12937777
      BackColorBkg    =   24576
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
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
   Begin VB.TextBox TotIva 
      Height          =   315
      Left            =   7245
      TabIndex        =   11
      Top             =   4905
      Width           =   1290
   End
   Begin VB.TextBox TotImp 
      Height          =   315
      Left            =   4875
      TabIndex        =   9
      Top             =   4905
      Width           =   1275
   End
   Begin VB.TextBox TotDoc 
      Height          =   315
      Left            =   1785
      TabIndex        =   7
      Top             =   4905
      Width           =   1365
   End
   Begin VB.TextBox TxtDitta 
      Height          =   315
      Left            =   1845
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   615
      Width           =   4365
   End
   Begin VB.TextBox TxtData 
      Height          =   315
      Left            =   4650
      TabIndex        =   3
      Top             =   180
      Width           =   1560
   End
   Begin VB.Label LblPartIva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Partita IVA:"
      Height          =   225
      Left            =   6855
      TabIndex        =   32
      Top             =   675
      Width           =   870
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00646464&
      Height          =   990
      Left            =   15
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   15345
   End
   Begin VB.Label LblPagamento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pagamento:"
      Height          =   225
      Left            =   6855
      TabIndex        =   24
      Top             =   225
      Width           =   960
   End
   Begin VB.Label Legenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   135
      TabIndex        =   20
      Top             =   5490
      Width           =   600
   End
   Begin VB.Label LblNumRecord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   4245
      TabIndex        =   19
      Top             =   5535
      Width           =   45
   End
   Begin VB.Label LblTotIva 
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
      Left            =   6315
      TabIndex        =   10
      Top             =   4935
      Width           =   870
   End
   Begin VB.Label LblTotImp 
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
      Left            =   3315
      TabIndex        =   8
      Top             =   4935
      Width           =   1500
   End
   Begin VB.Label LblTotDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Documento:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   4935
      Width           =   1605
   End
   Begin VB.Label LblDitta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ditta:"
      Height          =   225
      Left            =   285
      TabIndex        =   4
      Top             =   645
      Width           =   420
   End
   Begin VB.Label LblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   225
      Left            =   4185
      TabIndex        =   2
      Top             =   210
      Width           =   405
   End
   Begin VB.Label LblNumDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Documento:"
      Height          =   225
      Left            =   285
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
End
Attribute VB_Name = "FattureClienti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VociMod As Boolean, CaricamentoDoc As Boolean, TotaliFattura As Collection
Dim NumDdt As String, DataDdt As String, Vettore As Long
Dim RigaCorr As Long
Dim StatoDoc As StatoRecord
Dim ElencoControlli As Variant, DescControlli As Variant
Dim rsFatture As ADODB.Recordset, rsVociFattura As ADODB.Recordset, rsVociNoteDoc As ADODB.Recordset, _
rsClienti As ADODB.Recordset, rsLuoghiConsegna As ADODB.Recordset, Differita As Integer, _
IntestazioniGriglia As Variant, PropCampiVoce As Variant, cm As CoordinateMouse, FiltroRicerca As Boolean
Dim WithEvents ssc As SmartSubClass
Attribute ssc.VB_VarHelpID = -1
Private Sub Anteprima_Click()
If StatoDoc = inserimento Or CInt(TxtRecCorr.Text) <= rsFatture.RecordCount Then
 If ConvalidaRecord(False) Then
  CaricaTotaliIva
  Dim SS As New ServiziStampa: Call CaricaIntestazioneDitta(SS)
  SS.ImpostaAnteprima True
  Set SS.TotaliDoc = TotaliFattura
  Set SS.rsDoc = rsFatture: Set SS.rsCliente = rsClienti
  If Differita = 1 Then
   SS.TipoDoc = FatturaDifferita
   Set SS.rsVociDoc = rsVociNoteDoc
  Else
   SS.TipoDoc = FatturaImmediata
   Set SS.rsVociDoc = rsVociFattura
  End If
  If rsFatture("IdLC") <> 0 Then
   SS.LuogoConsegna = rsLuoghiConsegna.GetRows(1, adBookmarkFirst)
  End If
  SS.Stampa: AnteprimaDoc.Show vbModal
 End If
Else
 MsgBox "Inserire i dati del documento !", vbExclamation, "Fattura Pro"
End If
End Sub
Public Sub AggiornaTotali(ByVal IdNota As String, ByVal NuovaNota As Boolean)
Dim TotImpNota#, TotIvaNota#
TotImpNota = rsVociFattura("TotImp"): TotIvaNota = rsVociFattura("TotIva")
rsVociFattura.Requery
rsVociNoteDoc.Requery
rsVociFattura.Find "IdDoc = '" & IdNota & "'", , adSearchForward, adBookmarkFirst

If NuovaNota Then
 TotImp = FormatNumber(CDbl(TotImp.Text) + rsVociFattura("TotImp"), 2)
 TotIva = FormatNumber(CDbl(TotIva.Text) + rsVociFattura("TotIva"), 2)
 ElencoVoci.AddItem ""
 ElencoVoci.RowHeight(ElencoVoci.Rows - 1) = 315
 riga = ElencoVoci.Rows - 1
Else
 TotImp = FormatNumber(CDbl(TotImp) + (rsVociFattura("TotImp") - TotImpNota), 2)
 TotIva = FormatNumber(CDbl(TotIva) + (rsVociFattura("TotIva") - TotIvaNota), 2)
 riga = ElencoVoci.Row
End If
TotDoc = TotImp + TotIva

ElencoVoci.TextMatrix(riga, 0) = NumeroDocumento(IdNota)
ElencoVoci.TextMatrix(riga, 1) = Format$(rsVociFattura("data"), "dd-mm-yyyy")
ElencoVoci.TextMatrix(riga, 2) = FormatNumber(rsVociFattura("TotDoc"), 2)
TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text), 2)
If InStr(1, TotDoc.Text, ",99") Then TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) + 0.01, 2)
End Sub
Private Sub BtnEsenzioniIva_Click()
Set EsenzioniIva.FormChiamante = Me
EsenzioniIva.Show vbModal
End Sub
Private Sub BtnDdt_Click()
DdtFattura.NumDoc = NumDdt
DdtFattura.DataDoc = DataDdt
DdtFattura.Vettore = Vettore
DdtFattura.Show vbModal
NumDdt = DdtFattura.NumDoc
DataDdt = DdtFattura.DataDoc
Vettore = DdtFattura.Vettore
If DdtFattura.SalvaModifiche Then
 ModificaDoc
End If
Unload DdtFattura
End Sub
Private Sub BtnCanc_Click()
Dim Scelta%
Scelta = MsgBox("Cancellare il documento corrente ?" & vbNewLine & "Non sarà possibile" _
& " annullare questa modifica !", vbYesNo + vbQuestion, "Fattura Pro")
If Scelta = vbYes Then
 Dim PosRec As Integer
 PosRec = rsFatture.AbsolutePosition
 If StatoDoc <> inserimento Then
  rsFatture.Delete
  rsFatture.Update
  If StatoDoc <> NonModificato Then
   conn.CommitTrans
   StatoDoc = NonModificato
  End If
  If PosRec > rsFatture.RecordCount Then
   If rsFatture.RecordCount <> 0 Then
    rsFatture.MoveLast
   End If
   CreaNuovoRecord
  Else
   rsFatture.MoveNext
   VisualizzaRecord
   TxtRecCorr.Text = rsFatture.AbsolutePosition
   LblNumRecord.Caption = "di " & rsFatture.RecordCount
  End If
 Else
  CreaNuovoRecord
 End If
End If
End Sub
Private Sub BtnFatturaElettronica_Click()
Dim DocValido As Boolean, FE As FatturaElettronica
DocValido = ConvalidaRecord(False)
If DocValido Then
 CaricaTotaliIva
 Dim NomeDoc$
 Set FE = New FatturaElettronica
 NomeDoc = FE.CreaFattura(rsClienti, rsFatture, rsVociFattura, TotaliFattura)
 MsgBox "Documento salvato in " & PercorsoApp & "Fatture Elettroniche\" & NomeDoc, vbInformation, _
 "Fattura Pro"
Else
 MsgBox "Documento non valido. Impossibile creare la fattura elettronica", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub BtnLuoghiConsegna_Click()
If ChkLuogoConsegna.Value Then
 If TxtDitta.Tag = "" Then
  ChkLuogoConsegna.Value = 0
  MsgBox "Inserire la ditta destinataria !", vbExclamation, "Fattura Pro"
 Else
  Dim ElencoLuoghiConsegna As New LuoghiConsegna
  If ElencoLuoghiConsegna.CaricaLuoghi(TxtDitta.Tag) Then
   Set ElencoLuoghiConsegna.FormChiamante = Me: ElencoLuoghiConsegna.Show vbModal
  Else
   ChkLuogoConsegna.Value = 0
   MsgBox "La ditta non ha luoghi di consegna associati !", vbExclamation, "Fattura Pro"
  End If
 End If
End If
End Sub
Private Sub BtnSelDitta_Click()
If Differita = 0 Then
 Set SelezioneDitta.FormChiamante = Me
 If TxtDitta.Tag <> "" Then
  SelezioneDitta.TxtDitta.Text = TxtDitta.Text
  SelezioneDitta.TxtDitta.Tag = TxtDitta.Tag
 End If
 SelezioneDitta.Show vbModal
 If Not SelezioneDitta.rsDitte.EOF Then
  TxtDitta.Tag = SelezioneDitta.rsDitte("id")
  TxtDitta.Text = SelezioneDitta.rsDitte("ditta")
  TxtPartIva.Text = SelezioneDitta.rsDitte("partitaiva")
  CmbPag.ListIndex = SelezioneDitta.rsDitte("modpag")
  rsClienti.Requery
  rsClienti.Find "Id = " & TxtDitta.Tag, , adSearchForward, adBookmarkFirst
 End If
End If
End Sub
Private Sub ChkEsenteIva_Click()
ModificaDoc
If Not CaricamentoDoc Then
 If ChkEsenteIva.Value <> 0 Then
  TotDoc = FormatNumber(CDbl(TotDoc) - CDbl(TotIva), 2)
  TotImp = "0,00": TotIva = "0,00"
  If rsVociFattura.RecordCount Then
   rsVociFattura.MoveFirst
   For i = 1 To rsVociFattura.RecordCount
    ElencoVoci.TextMatrix(i, 6) = ""
    rsVociFattura("iva") = ""
    rsVociFattura.Update
    rsVociFattura.MoveNext
   Next i
   ElencoVoci.TextMatrix(i, 6) = ""
  End If
  PropCampiVoce(6) = ""
 Else
  PropCampiVoce(6) = "no"
  NonImpIva.Text = ""
 End If
End If
End Sub
Private Sub ChkLuogoConsegna_Click()
If ChkLuogoConsegna.Value = 0 Then
 TxtLC.Tag = "": TxtLC.Text = ""
End If
ModificaDoc
End Sub
Private Sub ElencoVoci_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cm.x = x: cm.y = y
End Sub
Private Sub Pagato_Click()
ModificaDoc
End Sub
Private Sub TxtData_KeyPress(KeyAscii As Integer)
If InStr("0123456789-" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub PopupDitte_Click()
If PopupDitte.ListIndex <> -1 Then
 TxtDitta.Text = PopupDitte.Text
 TxtDitta.Tag = PopupDitte.ItemData(PopupDitte.ListIndex)
 rsClienti.Find "Id = " & TxtDitta.Tag, , adSearchForward, adBookmarkFirst
 TxtPartIva.Text = rsClienti("partitaiva")
 CmbPag.ListIndex = rsClienti("modpag")
 CaricaLuoghiConsegna
 PopupDitte.Visible = False
End If
End Sub
Private Sub CaricaLuoghiConsegna()
If rsLuoghiConsegna.State <> adStateClosed Then
 rsLuoghiConsegna.Close
End If
rsLuoghiConsegna.Open "SELECT * FROM LuoghiConsegna WHERE IdDitta = " & TxtDitta.Tag, conn, adOpenDynamic
If Not rsLuoghiConsegna.EOF Then
 ChkLuogoConsegna.Value = 0
End If
End Sub
Private Sub TxtDitta_KeyDown(KeyCode As Integer, Shift As Integer)
If Differita = 1 Then KeyCode = 0
End Sub
Private Sub NonImpIva_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub NonImpIva_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtDitta_LostFocus()
PopupDitte.Visible = False
End Sub
Private Sub TxtPartIva_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtPartIva_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtRecCorr_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtRecCorr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub ElencoVoci_Scroll()
TxtModifica.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
If StatoDoc <> NonModificato Then
 Dim Scelta As VbMsgBoxResult
 Scelta = MsgBox("Salvare il documento corrente ?", _
 vbYesNoCancel + vbQuestion, "Chiusura Fatture Clienti")
 If Scelta = vbYes Then
  If Not ConvalidaRecord(True) Then
   Cancel = 1
  End If
 ElseIf Scelta = vbCancel Then
  Cancel = 1
 End If
End If
If Cancel <> 1 Then
 If StatoDoc <> NonModificato Then
  Note.SalvaNota False
  conn.RollbackTrans
 End If
 rsFatture.Close
 ssc.SubClassHwnd ElencoVoci.hWnd, False
 Set rsFatture = Nothing
 Set FattureClienti = Nothing
End If
End Sub
Private Sub TxtNumFattura_Change()
ModificaDoc
End Sub
Private Sub TxtDitta_Change()
ChkLuogoConsegna.Value = 0: TxtLC.Text = ""
ModificaDoc
End Sub
Private Sub TxtData_Change()
ModificaDoc
End Sub
Private Sub TxtNumFattura_KeyPress(KeyAscii As Integer)
If Differita = 0 Then
 If KeyAscii >= 32 Then
  If Len(TxtNumFattura.Text) = 10 Then
   KeyAscii = 0: Exit Sub
  End If
  If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then KeyAscii = 0
 End If
Else
 KeyAscii = 0
End If
End Sub
Private Sub NonImpIva_Change()
ModificaDoc
End Sub
Private Sub CmbPag_Click()
ModificaDoc
End Sub
Private Sub ssc_NewMessage(ByVal hWnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
Static m_bLMousePressed As Boolean, m_bLMouseClicked As Boolean
If m_bLMousePressed And uMsg = WM_LBUTTONUP Then
 m_bLMousePressed = False
 m_bLMouseClicked = True
End If
    
If Not (m_bLMousePressed) And uMsg = WM_LBUTTONDOWN Then
 m_bLMousePressed = True
 m_bLMouseClicked = False
End If
    
If m_bLMouseClicked And uMsg = WM_ERASEBKGND Then
 If ElencoVoci.ColWidth(2) < 800 Then
  ElencoVoci.ColWidth(2) = 800
 End If
 m_bLMouseClicked = False
End If
End Sub
Private Sub TotDoc_Change()
ModificaDoc
End Sub
Private Sub Form_Load()
Me.Move 600, 450
Set rsAliquoteIVA = New ADODB.Recordset
If rsFatture Is Nothing Then
 Set rsFatture = New ADODB.Recordset
 rsFatture.Open "SELECT * FROM FattureClienti ORDER BY IdDoc ASC", conn, adOpenDynamic, adLockOptimistic
End If
Set rsVociFattura = New ADODB.Recordset
Set rsClienti = New ADODB.Recordset
Set rsLuoghiConsegna = New ADODB.Recordset
rsLuoghiConsegna.Open "LuoghiConsegna", conn, adOpenDynamic
rsClienti.Open "Clienti", conn, adOpenDynamic
TxtRecCorr.Text = "1"

BtnPrec.Enabled = False: BtnPrimo.Enabled = False
BtnUltimo.Enabled = rsFatture.RecordCount > 1
Set rsVociNoteDoc = New ADODB.Recordset
If Not rsFatture.EOF Then
 BtnSucc.Enabled = True: BtnNuovo.Enabled = True
 BtnCanc.Enabled = True
 LblNumRecord.Caption = "di " & rsFatture.RecordCount
 rsFatture.MoveFirst
 Call VisualizzaRecord
Else
 Call ImpostazioniStandard
 CreaNuovoRecord
 LblNumRecord.Caption = "di 1"
End If
Set ssc = New SmartSubClass: ssc.SubClassHwnd ElencoVoci.hWnd, True
ElencoControlli = Array("TxtNumFattura", "TxtData", "TxtDitta", "CmbPag", "TotDoc", _
"TotImp", "TotIva")
DescControlli = Array("Num. Documento", "Data", "Ditta", "Pagamento", "Totale Documento", _
"Totale Imponibile", "Totale Iva")
PropCampiVoce = Array("ao", "", "a", "n", "no", "n", "n", "no")
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
If Differita = 0 Then
 If UCase$(Chr(KeyAscii)) <> LCase$(Chr(KeyAscii)) Then
  Dim PosCursore%
  PosCursore = TxtDitta.SelStart + 1
  If PosCursore = 1 Or Mid(TxtDitta.Text, PosCursore + 1, 1) = "." Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
  ElseIf Mid(TxtDitta.Text, PosCursore - 1, 1) = " " Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
 End If
 Dim NumCar As Integer
 NumCar = Len(TxtDitta.Text) + IIf(KeyAscii <> 8, 1, -1)
 If NumCar >= 3 Then
  Dim rsDitte As New ADODB.Recordset
  rsDitte.Open "SELECT * FROM Clienti WHERE Rimosso = False And UCASE(Ditta) LIKE '" & _
  UCase(Replace(TxtDitta.Text, "'", "''")) & "%' ORDER BY Ditta ASC", conn, adOpenDynamic
  If Not rsDitte.EOF Then
   PopupDitte.Clear
   rsDitte.MoveFirst
   While Not rsDitte.EOF
    With PopupDitte
    .AddItem rsDitte("Ditta")
    .ItemData(.NewIndex) = rsDitte("Id")
    End With
    rsDitte.MoveNext
   Wend
   PopupDitte.Visible = True
   PopupDitte.Left = TxtDitta.Left
   PopupDitte.Top = TxtDitta.Top + TxtDitta.Height + 45
   PopupDitte.Width = TxtDitta.Width
   PopupDitte.Height = 2500
  Else
   PopupDitte.Visible = False
  End If
  rsDitte.Close
 Else
  PopupDitte.Visible = False
 End If
Else
 KeyAscii = 0
End If
End Sub
Private Sub TxtModifica_LostFocus()
TxtModifica.Visible = False
End Sub
Private Function SalvaVoceCorr() As Boolean
SalvaVoceCorr = False: Dim e As Boolean
If VociMod Then
 For i = 0 To ElencoVoci.Cols - 2
  If InStr(1, PropCampiVoce(i), "o") <> 0 And ElencoVoci.TextMatrix(RigaCorr, i) = "" Then
    e = True: Exit For
  ElseIf ElencoVoci.TextMatrix(RigaCorr, i) <> "" Then
   If InStr(1, PropCampiVoce(i), "a") And IsNumeric(ElencoVoci.TextMatrix(RigaCorr, i)) Then
    e = True: Exit For
   ElseIf InStr(1, PropCampiVoce(i), "n") <> 0 And Not IsNumeric(ElencoVoci.TextMatrix(RigaCorr, i)) Then
    e = True: Exit For
   End If
  End If
 Next i
 
 If e Then
  ElencoVoci.Row = RigaCorr: ElencoVoci.Col = i
  TxtModifica.Visible = True: Set TxtModifica.Container = ElencoVoci.Container
  With ElencoVoci
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
  End With
  TxtModifica.SetFocus: TxtModifica.Text = ElencoVoci.Text
  MsgBox "Attenzione, uno o più campi della voce n. " & RigaCorr & " del documento contengono valori" _
  & " non validi !", vbExclamation, "Fattura Pro"
  Exit Function
 End If
 
 If RigaCorr > rsVociFattura.RecordCount Then
  rsVociFattura.AddNew
 Else
  rsVociFattura.Move RigaCorr - 1, adBookmarkFirst
 End If
 
 With rsVociFattura
  .Fields("descr") = EliminaSpazi(ElencoVoci.TextMatrix(RigaCorr, 0))
  .Fields("lotto") = ElencoVoci.TextMatrix(RigaCorr, 1)
  .Fields("um") = ElencoVoci.TextMatrix(RigaCorr, 2)
  If ElencoVoci.TextMatrix(RigaCorr, 3) <> "" Then
   .Fields("qnt") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 3))
  Else
   .Fields("qnt") = 0
  End If
  .Fields("prezzo") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 4))
  .Fields("sconto") = ElencoVoci.TextMatrix(RigaCorr, 5)
  .Fields("iva") = ElencoVoci.TextMatrix(RigaCorr, 6)
  .Fields("totale") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7))
  .Update
 End With
 ModificaDoc
 VociMod = False
End If
SalvaVoceCorr = True
End Function
Private Sub TxtModifica_Change()
If TxtModifica.Text <> ElencoVoci.Text Then
 Dim TotNonImp#, Qnt#, Diff#, TotCorr#, NuovoTot#
 Dim NuovaIva#, IvaCorr#, AlIva#
 ModificaDoc
 VociMod = True
 If ElencoVoci.Col = 5 Then
  Dim Sconto#
  If ElencoVoci.TextMatrix(RigaCorr, 7) <> "" Then
   Dim TotNoSconto#: TotCorr = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7))
   If IsNumeric(TxtModifica.Text) Then Sconto = CDbl(TxtModifica.Text)
   If ElencoVoci.Text <> "" Then
    TotNoSconto = TotCorr * 100 / (100 - CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)))
   Else: TotNoSconto = TotCorr
   End If
   NuovoTot = TotNoSconto - (TotNoSconto * Sconto / 100)
   ElencoVoci.TextMatrix(RigaCorr, 7) = FormatNumber(NuovoTot, 2)
   Diff = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) - TotCorr
   TotNonImp = FormatNumber(CDbl(TotDoc.Text) - (CDbl(TotImp.Text) + CDbl(TotIva.Text)), 2)
   ElencoVoci.Text = TxtModifica.Text
   If ElencoVoci.TextMatrix(RigaCorr, 6) <> "" Then
    TotImp.Text = FormatNumber(CDbl(TotImp.Text) + Diff, 2)
    AlIva = CDbl(Replace(ElencoVoci.TextMatrix(RigaCorr, 6), "%", ""))
    IvaCorr = TotCorr * AlIva / 100: NuovaIva = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) * AlIva / 100
    Diff = CDbl(Arrotonda(NuovaIva)) - CDbl(Arrotonda(IvaCorr))
    TotIva.Text = FormatNumber(CDbl(TotIva.Text) + Diff, 2)
    TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text) + TotNonImp, 2)
   Else
    TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) + Diff, 2)
   End If
  End If
 ElseIf ElencoVoci.Col <> 6 Then
  If ElencoVoci.Col = 3 Then
   If IsNumeric(TxtModifica.Text) Then
    ElencoVoci.Text = FormatNumber(CDbl(TxtModifica.Text), 3)
   Else
    ElencoVoci.Text = ""
   End If
  ElseIf ElencoVoci.Col = 4 Then
   If IsNumeric(TxtModifica.Text) Then
    ElencoVoci.Text = FormattaNumero(TxtModifica.Text)
   Else
    ElencoVoci.Text = ""
   End If
  Else: ElencoVoci.Text = Trim(TxtModifica.Text)
  End If
  If (ElencoVoci.Col = 3 Or ElencoVoci.Col = 4) And ElencoVoci.TextMatrix(RigaCorr, 4) <> "" Then
   If ElencoVoci.TextMatrix(RigaCorr, 7) <> "" Then
    TotCorr = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7))
   End If
   Qnt = 1
   If ElencoVoci.TextMatrix(RigaCorr, 3) <> "" Then
    Qnt = CDbl(ElencoVoci.TextMatrix(RigaCorr, 3))
   End If
   NuovoTot = CDbl(ElencoVoci.TextMatrix(RigaCorr, 4)) * Qnt
   If ElencoVoci.TextMatrix(RigaCorr, 5) <> "" Then
    NuovoTot = NuovoTot - (NuovoTot * CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)) / 100)
   End If
   ElencoVoci.TextMatrix(RigaCorr, 7) = FormatNumber(NuovoTot, 2)
   Diff = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) - TotCorr
   TotNonImp = FormatNumber(CDbl(TotDoc.Text) - (CDbl(TotImp.Text) + CDbl(TotIva.Text)), 2)
   If ElencoVoci.TextMatrix(RigaCorr, 6) <> "" Then
    TotImp.Text = FormatNumber(CDbl(TotImp.Text) + Diff, 2)
    IvaCorr = TotCorr * CDbl(ElencoVoci.TextMatrix(RigaCorr, 6)) / 100
    NuovaIva = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) * _
    CDbl(ElencoVoci.TextMatrix(RigaCorr, 6)) / 100
    Diff = CDbl(Arrotonda(NuovaIva)) - CDbl(Arrotonda(IvaCorr))
    TotIva.Text = FormatNumber(CDbl(TotIva.Text) + Diff, 2)
    TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text) + TotNonImp, 2)
   Else
    TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) + Diff, 2)
   End If
  End If
 Else
  If ElencoVoci.TextMatrix(RigaCorr, 7) <> "" Then
   If ElencoVoci.Text <> "" Then
    IvaCorr = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) * CDbl(ElencoVoci.Text) / 100
   End If
   ElencoVoci.Text = TxtModifica.Text
   If ElencoVoci.Text <> "" Then
    NuovaIva = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) * CDbl(ElencoVoci.Text) / 100
   End If
   Diff = CDbl(Arrotonda(NuovaIva)) - CDbl(Arrotonda(IvaCorr))
   If Diff <> 0 Then
    If NuovaIva = 0 Then
     TotImp.Text = FormatNumber(CDbl(TotImp.Text) - CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)), 2)
    ElseIf IvaCorr = 0 Then
     TotImp.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)), 2)
    End If
    TotIva.Text = FormatNumber(CDbl(TotIva.Text) + Diff, 2)
    TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) + Diff, 2)
   End If
  Else
   ElencoVoci.Text = TxtModifica.Text
  End If
 End If
End If
End Sub
Private Sub TxtModifica_KeyPress(KeyAscii As Integer)
Dim NumCar As Integer
If KeyAscii >= 32 Then
 NumCar = Len(TxtModifica.Text) + 1
 If ElencoVoci.Col > 2 Then
  If NumCar > 10 Then
   KeyAscii = 0: Exit Sub
  End If
  Dim CarAmmessi$: CarAmmessi = "0123456789,"
  If ElencoVoci.Col = 5 Or ElencoVoci.Col = 6 Then
   CarAmmessi = "0123456789"
   If InStr(CarAmmessi, Chr(KeyAscii)) = 0 Or NumCar > 2 Then
    KeyAscii = 0: Exit Sub
   End If
   If ElencoVoci.Col = 6 And ChkEsenteIva.Value Then
    KeyAscii = 0: Exit Sub
   End If
  ElseIf InStr(CarAmmessi, Chr(KeyAscii)) = 0 Then
   KeyAscii = 0: Exit Sub
  End If
  If ElencoVoci.Col = 3 Or ElencoVoci.Col = 4 Then
   If Not ControlloCarIns(TxtModifica, Chr(KeyAscii), 7, 3) Then
    KeyAscii = 0: Exit Sub
   End If
  End If
 End If
 If RigaCorr = ElencoVoci.Rows - 1 Then
  Dim ColCorr&: ColCorr = ElencoVoci.Col: ElencoVoci.Col = 8: ElencoVoci.CellPictureAlignment = 4
  Set ElencoVoci.CellPicture = ImgCancella: ElencoVoci.Col = ColCorr: ElencoVoci.AddItem ""
  ElencoVoci.RowHeight(ElencoVoci.Rows - 1) = 315
  If ElencoVoci.Col <> 6 And ChkEsenteIva.Value = 0 Then ElencoVoci.TextMatrix(RigaCorr, 6) = "4"
  If ElencoVoci.Col <> 2 Then ElencoVoci.TextMatrix(RigaCorr, 2) = "Kg"
 End If
 If ElencoVoci.Col = 0 Then
  Dim PosCursore%: PosCursore = TxtModifica.SelStart + 1
  If PosCursore = 1 Or Mid(TxtModifica.Text, PosCursore + 1, 1) = "." Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  If NumCar = 3 Then
   TxtModifica.Text = TxtModifica.Text & Chr(KeyAscii): KeyAscii = 0: TxtModifica.SelStart = Len(TxtModifica.Text)
   Dim rsArticoli As New ADODB.Recordset
   rsArticoli.Open "SELECT * FROM Articoli WHERE descr LIKE '" & TxtModifica.Text & "%' ORDER BY descr ASC", conn, _
   adOpenDynamic, adLockOptimistic
   If Not rsArticoli.EOF Then
    Set ElencoArticoli.FormChiamante = Me
    Set ElencoArticoli.Articoli = rsArticoli: ElencoArticoli.Show vbModal
   End If
  End If
 End If
ElseIf KeyAscii = vbKeyReturn Then
 If ElencoVoci.Col < 6 Then
  ElencoVoci.Col = ElencoVoci.Col + 1
 Else: ElencoVoci.Col = 0
 End If
 With ElencoVoci
  TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
  TxtModifica.Text = .Text: TxtModifica.SetFocus
 End With
End If
End Sub
Private Sub ElencoVoci_Click()
If Differita = 0 Then
 If MouseInGriglia(ElencoVoci, cm) Then
  If ElencoVoci.Row <> RigaCorr Then
   If Not SalvaVoceCorr() Then Exit Sub
  End If
  RigaCorr = ElencoVoci.Row
  With ElencoVoci
   If .Col < 7 And (.Col <> 5 Or NonImpIva.Text = "") Then
    TxtModifica.Visible = True: Set TxtModifica.Container = .Container
    TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
    If .Col = 3 Or .Col = 4 Then
     TxtModifica.Text = Replace(.Text, ".", "")
    Else
     TxtModifica.Text = .Text
    End If
    TxtModifica.SetFocus
   ElseIf .Col = 8 And .Rows > 2 And .Row < .Rows - 1 Then
    Dim Scelta%
    Scelta = MsgBox("Rimuovere la voce dall'elenco ?", vbExclamation + vbYesNo, "Fattura Pro")
    If Scelta = vbYes Then
     Dim VoceIva, VoceTot
     If .TextMatrix(ElencoVoci.Row, 7) <> "" Then
      VoceTot = CDbl(.TextMatrix(.Row, 7))
      If .TextMatrix(.Row, 6) <> "" Then
       TotImp.Text = FormatNumber(CDbl(TotImp.Text) - VoceTot)
       VoceIva = CDbl(Arrotonda(VoceTot * (CDbl(.TextMatrix(.Row, 6)) / 100)))
       TotIva.Text = FormatNumber(CDbl(TotIva.Text) - VoceIva, 2)
       TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text), 2)
      Else
       TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) - VoceTot, 2)
      End If
     End If
     If .Row <= rsVociFattura.RecordCount Then
      rsVociFattura.Move .Row - 1, adBookmarkFirst
      rsVociFattura.Delete
     End If
     .RemoveItem .Row: VociMod = False
    End If
   End If
  End With
 End If
End If
End Sub
Private Sub ElencoVoci_DblClick()
If Differita = 1 Then
 rsVociFattura.Move ElencoVoci.Row - 1, adBookmarkFirst
 Note.IdDoc = rsVociFattura("IdDoc")
 Set Note.FormChiamante = Me
 Note.Show
End If
End Sub
Private Sub BtnPrimo_Click()
PosizioneRecord "primo"
End Sub
Private Sub BtnPrec_Click()
PosizioneRecord "precedente"
End Sub
Private Sub Stampa_Click()
If StatoDoc = inserimento Or CInt(TxtRecCorr.Text) <= rsFatture.RecordCount Then
 If ConvalidaRecord(False) Then
  CaricaTotaliIva
  Dim SS As New ServiziStampa
  Call CaricaIntestazioneDitta(SS): Set SS.rsDoc = rsFatture
  Set SS.rsCliente = rsClienti
  If Differita = 1 Then
   SS.TipoDoc = FatturaDifferita
   Set SS.rsVociDoc = rsVociNoteDoc
  Else
   SS.TipoDoc = FatturaImmediata
   Set SS.rsVociDoc = rsVociFattura
  End If
  If rsFatture("IdLC") <> 0 Then
   SS.LuogoConsegna = rsLuoghiConsegna.GetRows(1, adBookmarkFirst)
  End If
  Set SS.TotaliDoc = TotaliFattura
  OpzioniStampa.Inizializza SS: OpzioniStampa.Show vbModal
  If OpzioniStampa.Scelta = "Stampa" Then
   SS.ImpostaAnteprima False: SS.Stampa
  End If
 End If
Else
 MsgBox "Inserire i dati del documento !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub BtnSucc_Click()
PosizioneRecord "successivo"
End Sub
Private Sub TotDoc_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotDoc_KeyPress(KeyAscii As Integer)
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
Private Sub TxtLC_Change()
ModificaDoc
End Sub
Private Sub TxtLC_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtLC_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub BtnUltimo_Click()
PosizioneRecord "ultimo"
End Sub
Private Sub BtnNuovo_Click()
PosizioneRecord "nuovo"
End Sub
Private Function ConvalidaRecord(Salva As Boolean) As Boolean
Dim IdFattura$, d As Variant, e As Boolean
Dim MsgErroreIva$

If StatoDoc <> NonModificato Then
 For i = 0 To UBound(ElencoControlli) - 4
  If Trim(Me(ElencoControlli(i))) = "" Then
   MsgBox "Attenzione, " & DescControlli(i) & " è un campo obbligatorio !", _
   vbExclamation, "Fattura Pro": Me(ElencoControlli(i)).SetFocus: Exit Function
  ElseIf i = 1 Then
   If IsDate(Me(ElencoControlli(i))) Then
    d = Split(Me(ElencoControlli(i)), "-")
    If UBound(d) = 2 Then
     If Len(d(0)) < 2 Then d(0) = "0" & d(0)
     If Len(d(1)) < 2 Then d(1) = "0" & d(1)
     TxtData.Text = d(0) & "-" & d(1) & "-" & d(2)
    Else: e = True
    End If
   Else: e = True
   End If
   If e Then
    MsgBox "Attenzione, inserire una data valida !", vbExclamation, "Fattura Pro"
    TxtData.SetFocus: TxtData.SelStart = 0: TxtData.SelLength = Len(TxtData.Text): Exit Function
   End If
  End If
 Next i
 Dim re As New RegExp, NumDoc1$
 re.Global = True
 re.Pattern = "^[1-9][0-9]{0,2}/?[1-9]?$"
 NumDoc1 = Split(TxtNumFattura, "/")(0)
 IdFattura = Mid(TxtData.Text, 7) & "-" & String(4 - Len(NumDoc1), "0") & TxtNumFattura
 If re.Test(TxtNumFattura) Then
  Dim CercaDuplicato As Boolean
  If StatoDoc = inserimento Then
   CercaDuplicato = True
  ElseIf IdFattura <> rsFatture("IdDoc") Then
   CercaDuplicato = True
  End If
  If CercaDuplicato Then
   Dim rsDuplicato As New ADODB.Recordset
   rsDuplicato.Open "SELECT * FROM FattureClienti WHERE IdDoc = '" & IdFattura & "'", conn, adOpenDynamic
   If Not rsDuplicato.EOF Then
    MsgBox "Attenzione, il numero documento corrisponde a quello di un altro documento" _
    & " in archivio !" & vbNewLine, vbExclamation, "Fattura Pro"
    TxtNumFattura.SetFocus: TxtNumFattura.SelStart = 0
    TxtNumFattura.SelLength = Len(TxtNumFattura): Exit Function
   End If
  End If
 Else
  MsgBox "Attenzione, il numero documento non è valido !", vbExclamation, "Fattura Pro"
  TxtNumFattura.SetFocus: TxtNumFattura.SelStart = 0
  TxtNumFattura.SelLength = Len(TxtNumFattura): Exit Function
 End If
 If SalvaVoceCorr() Then
  If Differita = 0 Then
   If rsVociFattura.RecordCount = 0 Then
    MsgBox "Attenzione, non è stata inserita nessuna voce !", vbExclamation, "Fattura Pro"
    Exit Function
   End If
  End If
  Dim VerificaNumDoc As Boolean, NumValido As Boolean: NumValido = True
  If StatoDoc = inserimento Then
   VerificaNumDoc = True
  ElseIf IdFattura <> rsFatture("IdDoc") Then
   VerificaNumDoc = True
  End If
  If VerificaNumDoc Then
   id = Split(TxtNumFattura, "/"): NumValido = NumDocValido(id, d(2))
  End If
  If Not NumValido Then
   MsgBox "Attenzione, il numero documento non è valido !", vbExclamation, "Fattura Pro"
   TxtNumFattura.SetFocus: TxtNumFattura.SelStart = 0
   TxtNumFattura.SelLength = Len(TxtNumFattura): Exit Function
  End If
  If Not ControlloIva(MsgErroreIva) Then
   MsgBox MsgErroreIva, vbExclamation, "Fattura Pro"
   ConvalidaRecord = False: Exit Function
  End If
  If TxtDitta.Enabled Then
   If TxtDitta.Tag = "" Or (Not VerificaDestDoc) Then
    MsgBox "Attenzione, il campo ditta non contiene un valore valido !", vbExclamation, _
    "Fattura Pro"
    Exit Function
   End If
  End If
  If Differita = 0 Then
   If NumDdt <> "" And (Not IsDate(DataDdt)) Then
    MsgBox "Attenzione, la data del ddt non è valida !", vbExclamation, _
    "Fattura Pro"
    Exit Function
   ElseIf NumDdt = "" And IsDate(DataDdt) Then
    MsgBox "Attenzione, il numero del ddt non è valido !", vbExclamation, _
    "Fattura Pro"
    Exit Function
   End If
  End If
  If ChkLuogoConsegna.Value And TxtLC.Text = "" Then
   MsgBox "Attenzione, inserire il luogo di consegna per questo cliente !", vbExclamation, _
   "Fattura Pro"
   Exit Function
  End If
  If StatoDoc = inserimento Then
   rsFatture.AddNew
   StatoDoc = modifica
  End If
  rsFatture("IdDoc") = IdFattura
  rsFatture("Data") = TxtData.Text
  rsFatture("IdDitta") = TxtDitta.Tag
  If TxtLC.Text <> "" Then
   rsFatture("IdLC") = TxtLC.Tag
  Else
   rsFatture("IdLC") = Null
  End If
  rsFatture("Pagato") = Pagato.Value
  rsFatture("Modpag") = CmbPag.ListIndex
  rsFatture("TipoDoc") = Differita
  If Differita = 0 Then
   rsFatture("NumDdt") = NumDdt
   If DataDdt <> "" Then
    rsFatture("DataDdt") = DataDdt
   End If
   rsFatture("IdVettore") = Vettore
  End If
  rsFatture("EsenteIva") = NonImpIva.Text
  rsFatture("TotDoc") = CDbl(TotDoc.Text)
  rsFatture("TotImp") = CDbl(TotImp.Text)
  rsFatture("TotIva") = CDbl(TotIva.Text)
  rsFatture.Update
  rsVociFattura.MoveFirst
  While Not rsVociFattura.EOF
   If IsNull(rsVociFattura("IdFatt")) Then
    rsVociFattura("IdFatt") = IdFattura
    rsVociFattura.Update
   End If
   rsVociFattura.MoveNext
  Wend
  If Salva Then
   If Differita Then
    Note.SalvaNota True
   End If
   conn.CommitTrans
   StatoDoc = NonModificato
   EseguiBackup = True
  End If
 Else
  Exit Function
 End If
End If
ConvalidaRecord = True
End Function
Private Sub PosizioneRecord(PosRecord As String)
Dim valido: valido = ConvalidaRecord(True)
If valido Then
 If rsFatture.RecordCount > 1 Then
  Dim NumRec%: NumRec = rsFatture.AbsolutePosition
  rsFatture.Requery
  rsFatture.Move NumRec - 1, adBookmarkFirst
 End If
 Select Case PosRecord
 Case "primo"
  rsFatture.MoveFirst
 Case "ultimo"
  rsFatture.MoveLast
 Case "precedente"
  If CInt(TxtRecCorr.Text) <= rsFatture.RecordCount Then
   rsFatture.MovePrevious
  ElseIf rsFatture.AbsolutePosition <> rsFatture.RecordCount Then
   rsFatture.MoveLast
  End If
 Case "successivo"
  If rsFatture.AbsolutePosition = rsFatture.RecordCount Then
   CreaNuovoRecord
   Exit Sub
  End If
  rsFatture.MoveNext
 Case "nuovo"
  CreaNuovoRecord
  Exit Sub
 End Select
 BtnCanc.Enabled = True
 If rsFatture.AbsolutePosition <= rsFatture.RecordCount Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
 End If
 If rsFatture.AbsolutePosition <> rsFatture.RecordCount Then
  BtnUltimo.Enabled = True
 Else
  BtnUltimo.Enabled = False
 End If
 If rsFatture.AbsolutePosition <> 1 Then
  BtnPrimo.Enabled = True: BtnPrec.Enabled = True
 Else
  BtnPrimo.Enabled = False: BtnPrec.Enabled = False
 End If
 TxtRecCorr.Text = rsFatture.AbsolutePosition
 LblNumRecord.Caption = "di " & rsFatture.RecordCount
 rsVociFattura.Close
 VisualizzaRecord
End If
End Sub
Private Sub VisualizzaRecord()
CaricamentoDoc = True
With rsFatture
 TxtNumFattura = NumeroDocumento(.Fields("iddoc"))
 TxtData = Format$(.Fields("Data"), "dd-mm-yyyy")
 rsClienti.Find "Id = " & .Fields("IdDitta"), , adSearchForward, adBookmarkFirst
 TxtDitta.Text = rsClienti("Ditta")
 TxtDitta.Tag = .Fields("IdDitta")
 TxtPartIva.Text = rsClienti("PartitaIva")
 CmbPag.ListIndex = .Fields("Modpag")
 Pagato.Value = .Fields("Pagato") * -1
 TotDoc = FormatNumber(.Fields("TotDoc"), 2)
 TotImp = FormatNumber(.Fields("TotImp"), 2)
 TotIva = FormatNumber(.Fields("TotIva"), 2)
 Differita = .Fields("TipoDoc")
 If Not IsNull(.Fields("IdLC")) Then
  TxtLC.Tag = .Fields("IdLC")
  rsLuoghiConsegna.Find "Id = " & .Fields("IdLc"), , adSearchForward, adBookmarkFirst
  If Not rsLuoghiConsegna.EOF Then
   ChkLuogoConsegna.Value = 1
   TxtLC.Text = rsLuoghiConsegna("Indirizzo")
  Else
   ChkLuogoConsegna.Value = 0
   TxtLC.Tag = "": TxtLC.Text = ""
  End If
 Else
  TxtLC.Tag = "": TxtLC.Text = ""
  ChkLuogoConsegna.Value = 0
 End If
End With
If rsVociFattura.State <> adStateClosed Then
 rsVociFattura.Close
End If
Dim NumRiga%: ElencoVoci.Rows = 1
If Differita = 1 Then
 RiquadroEsenteIva.Visible = False: RiquadroLC.Visible = False
 ElencoVoci.Cols = 3: IntestazioniGriglia = Array("N° Doc.", "Data", "Totale")
 ElencoVoci.ColWidth(0) = 1100: ElencoVoci.ColWidth(1) = 1400
 ElencoVoci.ColWidth(2) = 1400: ElencoVoci.Width = 6000: ChkEsenteIva.Visible = False
 ElencoVoci.Height = 3600
 BtnDdt.Enabled = False
 conn.IsolationLevel = adXactCursorStability
 For i = 0 To ElencoVoci.Cols - 1
  ElencoVoci.TextMatrix(0, i) = IntestazioniGriglia(i): ElencoVoci.ColAlignment(i) = 4
 Next i
 ElencoVoci.HighLight = flexHighlightWithFocus: ElencoVoci.SelectionMode = flexSelectionByRow
 rsVociFattura.Open "SELECT * FROM NoteConsegna WHERE IdFatt = '" & rsFatture("IdDoc") & "'", _
 conn, adOpenDynamic
 If rsVociNoteDoc.State <> adStateClosed Then
  rsVociNoteDoc.Close
 End If
 rsVociNoteDoc.Open "SELECT NoteConsegna.*, VociNoteConsegna.* FROM NoteConsegna, VociNoteConsegna WHERE IdNota = " _
 & "NoteConsegna.IdDoc And IdFatt = '" & rsFatture("IdDoc") & "' ORDER BY NoteConsegna.IdDoc ASC", conn, adOpenDynamic
 If Not rsVociNoteDoc.EOF Then
  rsVociFattura.MoveFirst
  ElencoVoci.Redraw = False
  While Not rsVociFattura.EOF
   ElencoVoci.Rows = ElencoVoci.Rows + 1: NumRiga = ElencoVoci.Rows - 1
   ElencoVoci.RowHeight(NumRiga) = 315
   ElencoVoci.TextMatrix(NumRiga, 0) = NumeroDocumento(rsVociFattura("IdDoc"))
   ElencoVoci.TextMatrix(NumRiga, 1) = Format$(rsVociFattura("Data"), "dd-mm-yyyy")
   ElencoVoci.TextMatrix(NumRiga, 2) = FormatNumber(rsVociFattura("TotDoc"), 2)
   rsVociFattura.MoveNext
  Wend
  ElencoVoci.Redraw = True
 End If
Else
 BtnFatturaElettronica.Enabled = True
 NumDdt = rsFatture("NumDdt")
 DataDdt = Format$(rsFatture("DataDdt"), "dd-mm-yyyy")
 Vettore = IIf(IsNull(rsFatture("IdVettore")), -1, rsFatture("IdVettore"))
 BtnDdt.Enabled = True
 Call ImpostazioniStandard
 rsVociFattura.Open "SELECT * FROM VociFattureClienti WHERE IdFatt = '" & rsFatture("IdDoc") & "' ORDER BY Id ASC", conn, adOpenDynamic, adLockOptimistic
 If Not rsVociFattura.EOF Then
  rsVociFattura.MoveFirst
 End If
 With ElencoVoci
  Dim CifreDec%
  ElencoVoci.Redraw = False
  While Not rsVociFattura.EOF
   .AddItem "": i = 0
   .RowHeight(.Rows - 1) = 315
   For Each Field In rsVociFattura.Fields
    If Field.Name <> "Id" And Field.Name <> "IdFatt" Then
     If Field.Type = adDouble Then
      If Field.Name = "Qnt" Then
       CifreDec = 3
      Else
       CifreDec = 2
      End If
      If Field.Value <> 0 Then
       .TextMatrix(.Rows - 1, i) = FormatNumber(rsVociFattura(Field.Name), CifreDec)
      End If
     Else
      .TextMatrix(.Rows - 1, i) = rsVociFattura(Field.Name)
     End If
     i = i + 1
    End If
   Next
   .Row = .Rows - 1: .Col = 8: .CellPictureAlignment = 4
   Set .CellPicture = ImgCancella
   rsVociFattura.MoveNext
  Wend
  ElencoVoci.Redraw = True
  .AddItem "": .RowHeight(ElencoVoci.Rows - 1) = 315
 End With
End If
If Not IsNull(rsFatture.Fields("esenteiva")) Then
 NonImpIva.Text = rsFatture.Fields("esenteiva")
Else
 NonImpIva.Text = ""
End If
ChkEsenteIva.Value = IIf(NonImpIva.Text <> "", 1, 0)
If StatoDoc <> NonModificato Then
 StatoDoc = NonModificato
 conn.CommitTrans
End If
CaricamentoDoc = False
End Sub
Private Sub CaricaTotaliIva()
Dim Totale As Double, Iva As Double, Aliquota As Double
Dim TI As TotaleIva
Set TotaliFattura = New Collection
Dim rsVociDoc As New ADODB.Recordset
If Differita = 1 Then
 Set rsVociDoc = rsVociNoteDoc
Else
 Set rsVociDoc = rsVociFattura
End If
If rsVociDoc.RecordCount <> 0 Then
 rsVociDoc.MoveFirst
 On Error Resume Next
 While Not rsVociDoc.EOF
  If rsVociDoc("iva") <> "" Then
   Aliquota = CDbl(rsVociDoc("iva"))
   Totale = CDbl(rsVociDoc("totale"))
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
  rsVociDoc.MoveNext
 Wend
End If
If Split(TotDoc, ",")(1) = "99" Then TotDoc = FormatNumber(CDbl(TotDoc) + 0.01, 2)
For Each TotaleIva In TotaliFattura
 If InStr(1, CStr(TotaleIva.Totale), ",99") Then
  TotaleIva.Totale = TotaleIva.Totale + 0.01
 End If
Next
End Sub
Public Sub ImpostazioniStandard()
RiquadroEsenteIva.Visible = True: RiquadroLC.Visible = True
IntestazioniGriglia = Array("Descrizione", "Lotto", "U.M.", "Quantità", "Prezzo", _
"Sconto %", "IVA", "Totale"): ElencoVoci.Cols = 9: ElencoVoci.Width = 15360
ChkEsenteIva.Visible = True: BoxEsenzioneIva.Visible = True
conn.IsolationLevel = adXactChaos
ElencoVoci.ColWidth(0) = 7000: ElencoVoci.ColWidth(1) = 1100
ElencoVoci.ColWidth(2) = 800: ElencoVoci.ColWidth(3) = 1100
ElencoVoci.ColWidth(4) = 1100: ElencoVoci.ColWidth(5) = 1000
ElencoVoci.ColWidth(6) = 800: ElencoVoci.ColWidth(7) = 1200
ElencoVoci.ColWidth(8) = 375
ElencoVoci.Rows = 1: ElencoVoci.Height = 2670
ElencoVoci.HighLight = flexHighlightNever: ElencoVoci.SelectionMode = flexSelectionFree
TxtDitta.Enabled = True: BtnSelDitta.Enabled = True
For i = 0 To ElencoVoci.Cols - 2
 ElencoVoci.TextMatrix(0, i) = IntestazioniGriglia(i): ElencoVoci.ColAlignment(i) = 4
Next i
BtnDdt.Enabled = True: ChkLuogoConsegna.Enabled = True
End Sub
Private Sub CaricaIntestazioneDitta(SS As ServiziStampa)
Dim rsInfoDitta As New ADODB.Recordset
rsInfoDitta.Open "SELECT * FROM InfoDitta", conn, adOpenDynamic, adLockOptimistic
If Not rsInfoDitta.EOF Then
 SS.InfoDitta = rsInfoDitta.GetRows(1, adBookmarkFirst)
End If
rsInfoDitta.Close
End Sub
Public Function VerificaDestDoc() As Boolean
Dim rsDitte As New ADODB.Recordset
rsDitte.Open "SELECT * FROM Clienti WHERE UCASE(Ditta) = '" & UCase(Replace(TxtDitta.Text, "'", "''")) _
& "'", conn, adOpenDynamic
If Not rsDitte.EOF Then
 TxtDitta.Tag = rsDitte("Id")
 TxtPartIva.Text = rsDitte("PartitaIva")
 VerificaDestDoc = True
End If
End Function
Private Sub CreaNuovoRecord()
Call ImpostazioniStandard
TxtNumFattura.Text = ""
TxtDitta.Text = "": TxtDitta.Tag = ""
TxtData.Text = "": TxtRecCorr.Text = rsFatture.RecordCount + 1
NumDdt = "": DataDdt = "": TxtPartIva.Text = ""
LblNumRecord.Caption = "di " & (rsFatture.RecordCount + 1)
ElencoVoci.Rows = 2: Pagato.Value = 0
ElencoVoci.RowHeight(1) = 315: CmbPag.ListIndex = -1
TotDoc = "0,00": TotImp = "0,00"
TotIva = "0,00": Differita = 0
BtnSucc.Enabled = False: BtnCanc.Enabled = False: BtnNuovo.Enabled = False
If rsFatture.RecordCount >= 1 Then
 BtnPrec.Enabled = True: BtnPrimo.Enabled = True: BtnUltimo.Enabled = True
End If
RigaCorr = 0: ChkEsenteIva = 0: NonImpIva.Text = ""
TxtLC.Tag = "": TxtLC.Text = "": ChkLuogoConsegna.Value = 0
If rsVociFattura.State <> adStateClosed Then
 rsVociFattura.Close
End If
BtnFatturaElettronica.Enabled = False
rsVociFattura.Open "SELECT * FROM VociFattureClienti WHERE IdFatt = '0'", conn, adOpenDynamic, adLockOptimistic
StatoDoc = NonModificato
conn.CommitTrans
End Sub
Private Sub BtnVai_Click()
Dim IdCorr$
If Not rsFatture.EOF Then
 IdCorr = NumeroDocumento(rsFatture("IdDoc")) & "-" & Mid(rsFatture("IdDoc"), 1, 4)
End If
If TxtNumFattura <> "" And TxtNumFattura <> IdCorr Then
 If InStr(TxtNumFattura, "-") <> 0 Then
  Dim IdDoc, IdFatt$, PosRec%
  PosRec = rsFatture.AbsolutePosition
  IdDoc = Split(TxtNumFattura, "-")
  NumDoc1 = Split(IdDoc(0), "/")(0)
  IdFatt = IdDoc(1) & "-" & String(4 - Len(NumDoc1), "0") & IdDoc(0)
  rsFatture.Find "IdDoc = '" & IdFatt & "'", , adSearchForward, adBookmarkFirst
  If Not rsFatture.EOF Then
   TxtRecCorr.Text = rsFatture.AbsolutePosition
   VisualizzaRecord
   If rsFatture.AbsolutePosition > 1 Then
    BtnPrec.Enabled = True: BtnPrimo.Enabled = True
   End If
   If rsFatture.AbsolutePosition < rsFatture.RecordCount Then
    BtnUltimo.Enabled = True
   End If
   BtnNuovo.Enabled = True: BtnSucc.Enabled = True
  Else
   If rsFatture.RecordCount Then
    rsFatture.Move PosRec - 1, adBookmarkFirst
   End If
   MsgBox "Attenzione, documento non presente in archivio !", vbExclamation, "Fattura Pro"
   If StatoDoc <> NonModificato Then
    StatoDoc = NonModificato
    conn.CommitTrans
   End If
  End If
 Else
  If StatoDoc = inserimento Then
   TxtNumFattura.Text = ""
  ElseIf StatoDoc = modifica Then
   TxtNumFattura.Text = IdCorrente
  End If
  MsgBox "Attenzione, devi indicare il documento che vuoi visualizzare nel formato numero-anno (Es: 100-2013) !", _
  vbExclamation, "Fattura Pro"
 End If
End If
End Sub
Private Sub ModificaDoc()
Dim AvviaTrans As Boolean
If StatoDoc = NonModificato Then
 AvviaTrans = True
End If
If CInt(TxtRecCorr.Text) > rsFatture.RecordCount Then
 If Not FiltroRicerca Then
  StatoDoc = inserimento
  BtnSucc.Enabled = True: BtnCanc.Enabled = True
  BtnNuovo.Enabled = True
 Else
  AvviaTrans = False
 End If
Else
 StatoDoc = modifica
End If
If AvviaTrans Then
 conn.BeginTrans
End If
If StatoDoc = inserimento Then
 BtnFatturaElettronica.Enabled = True
End If
End Sub
Public Sub ImpostaFiltroRicerca(rsRicerca As ADODB.Recordset)
Set rsFatture = rsRicerca
FiltroRicerca = True
End Sub
Public Sub CaricaFiltroRicerca()
PosizioneRecord "primo"
End Sub
Public Function FormBloccata() As Boolean
FormBloccata = StatoDoc <> NonModificato
End Function
Private Function VerificaAliquote() As Boolean
Dim re As New RegExp
VerificaAliquote = True: re.Pattern = "Spes[ae] ?(di)? ?Spedizion[ei]|Cost[oi] ?(di)? ?Spedizion[ei]"
re.IgnoreCase = True
For i = 1 To ElencoVoci.Rows - 2
 If ElencoVoci.TextMatrix(i, 6) = "" And (Not re.Test(ElencoVoci.TextMatrix(i, 0))) Then
  VerificaAliquote = False: Exit For
 End If
Next i
End Function
Private Function ControlloIva(MsgErrore As String) As Boolean
ControlloIva = True
If Differita = 0 Then
 If ChkEsenteIva.Value = 0 Then
  If Not VerificaAliquote() Then
   MsgErrore = "Attenzione, per una o più voci del documento corrente non è stata impostata " _
   & "l'aliquota iva !"
   ControlloIva = False
  End If
 ElseIf NonImpIva.Text = "" Then
  MsgErrore = "Inserire il motivo di esenzione IVA !": ControlloIva = False
 End If
End If
End Function
Private Function NumDocValido(id As Variant, ByVal Anno$) As Boolean
If UBound(id) <= 1 Then
 If UBound(id) = 1 Then
  Dim IdPrec As String
  If id(1) = "1" Then
   IdPrec = Anno & "-" & String(4 - Len(id(0)), "0") & id(0)
  Else: IdPrec = Anno & "-" & String(4 - Len(id(0)), "0") & id(0) & "/" & (CInt(id(1)) - 1)
  End If
  If IdPrec <> rsFatture("IdDoc") Then
   Dim rsRecPrec As New ADODB.Recordset
   rsRecPrec.Open "SELECT * FROM FattureClienti WHERE IdDoc = '" & IdPrec & "'", conn, adOpenDynamic
   NumDocValido = Not rsRecPrec.EOF
   rsRecPrec.Close
   If Not NumDocValido Then Exit Function
  Else
   Exit Function
  End If
 End If
Else
 Exit Function
End If
Dim rsRecInvalido As New ADODB.Recordset, IdDoc$, NumDoc1$, DataDoc$
NumDoc1 = id(0)
IdDoc = Anno & "-" & String(4 - Len(NumDoc1), "0") & TxtNumFattura
DataDoc = Format$(TxtData, "yyyy/mm/dd")
rsRecInvalido.Open "SELECT * FROM FattureClienti WHERE (IdDoc < '" & IdDoc & "' AND Data > #" & DataDoc & "#) OR " _
& "(IdDoc > '" & IdDoc & "' AND Data < #" & DataDoc & "#)", conn, adOpenDynamic
NumDocValido = rsRecInvalido.EOF
rsRecInvalido.Close
End Function
