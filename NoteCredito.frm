VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form NoteCredito 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Note di Credito"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15675
   Icon            =   "NoteCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   15675
   Begin VB.TextBox TxtRifDoc 
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
      Left            =   7785
      TabIndex        =   36
      Top             =   165
      Width           =   1320
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
      Left            =   7785
      TabIndex        =   35
      Top             =   615
      Width           =   1695
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
      Left            =   3405
      MaskColor       =   &H00D8E9EC&
      Picture         =   "NoteCredito.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5190
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
      Left            =   3075
      MaskColor       =   &H00D8E9EC&
      Picture         =   "NoteCredito.frx":460C
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5190
      UseMaskColor    =   -1  'True
      Width           =   345
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
      Left            =   2730
      MaskColor       =   &H00D8E9EC&
      Picture         =   "NoteCredito.frx":4BA6
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5190
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
      Left            =   1560
      TabIndex        =   28
      Top             =   5190
      Width           =   1125
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
      Left            =   1170
      MaskColor       =   &H00D8E9EC&
      Picture         =   "NoteCredito.frx":5140
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5190
      UseMaskColor    =   -1  'True
      Width           =   345
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
      Left            =   825
      MaskColor       =   &H00D8E9EC&
      Picture         =   "NoteCredito.frx":56DA
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5190
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnCanc 
      Height          =   315
      Left            =   3825
      MaskColor       =   &H00D8E9EC&
      Picture         =   "NoteCredito.frx":5C74
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5190
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TxtNumNota 
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
      Left            =   1800
      TabIndex        =   0
      Top             =   165
      Width           =   1530
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
      Left            =   4605
      TabIndex        =   16
      Top             =   165
      Width           =   1560
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
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   600
      Width           =   4365
   End
   Begin VB.TextBox TotDoc 
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
      Left            =   1815
      TabIndex        =   14
      Top             =   4605
      Width           =   1365
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
      Left            =   4905
      TabIndex        =   13
      Top             =   4605
      Width           =   1275
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
      Left            =   7290
      TabIndex        =   12
      Top             =   4605
      Width           =   1290
   End
   Begin VB.TextBox TxtModifica 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Left            =   11160
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CheckBox Pagato 
      BackColor       =   &H0033CCFF&
      Caption         =   "Saldato"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9705
      TabIndex        =   9
      Top             =   270
      Width           =   975
   End
   Begin VB.CommandButton BtnVai 
      Caption         =   "Vai"
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
      Left            =   3375
      TabIndex        =   8
      Top             =   165
      Width           =   465
   End
   Begin VB.CheckBox ChkEsenteIva 
      BackColor       =   &H0033CCFF&
      Caption         =   "Esente I.V.A."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   7
      Top             =   3975
      Width           =   1440
   End
   Begin VB.TextBox NonImpIva 
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
      Left            =   3000
      TabIndex        =   6
      Top             =   3930
      Width           =   6285
   End
   Begin VB.CommandButton BtnEsenzioniIva 
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
      Left            =   9330
      MaskColor       =   &H00FFFFFF&
      Picture         =   "NoteCredito.frx":5D64
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3930
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton BtnSelDitta 
      Height          =   315
      Left            =   6210
      MaskColor       =   &H00FFFFFF&
      Picture         =   "NoteCredito.frx":5EBF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.ListBox PopupDitte 
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
      Left            =   9960
      TabIndex        =   3
      Top             =   5370
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CommandButton Anteprima 
      Caption         =   "Anteprima Nota"
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
      MaskColor       =   &H00D8E9EC&
      Picture         =   "NoteCredito.frx":601A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5745
      UseMaskColor    =   -1  'True
      Width           =   1845
   End
   Begin VB.CommandButton Stampa 
      Caption         =   "Stampa Nota"
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
      Left            =   165
      MaskColor       =   &H00FFFFFF&
      Picture         =   "NoteCredito.frx":6326
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5745
      UseMaskColor    =   -1  'True
      Width           =   1770
   End
   Begin MSFlexGridLib.MSFlexGrid ElencoVoci 
      Height          =   2670
      Left            =   45
      TabIndex        =   11
      Top             =   1125
      Width           =   15570
      _ExtentX        =   27464
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
   Begin VB.Label LblPartIva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Partita IVA:"
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
      Left            =   6765
      TabIndex        =   34
      Top             =   660
      Width           =   870
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
      Left            =   4260
      TabIndex        =   33
      Top             =   5265
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
      Left            =   150
      TabIndex        =   32
      Top             =   5220
      Width           =   600
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
      Left            =   240
      TabIndex        =   24
      Top             =   195
      Width           =   1485
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
      Left            =   4140
      TabIndex        =   23
      Top             =   195
      Width           =   405
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   22
      Top             =   615
      Width           =   600
   End
   Begin VB.Label Label5 
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
      Left            =   150
      TabIndex        =   21
      Top             =   4635
      Width           =   1605
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
      Left            =   3345
      TabIndex        =   20
      Top             =   4635
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
      Left            =   6360
      TabIndex        =   19
      Top             =   4635
      Width           =   870
   End
   Begin VB.Label LblRifFatt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rif. Fattura:"
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
      Left            =   6765
      TabIndex        =   18
      Top             =   195
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00646464&
      Height          =   1005
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   15570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esenzione I.V.A.:"
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
      Left            =   1665
      TabIndex        =   17
      Top             =   3975
      Width           =   1275
   End
End
Attribute VB_Name = "NoteCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VociMod As Boolean, RigaCorr As Long, TotaliNota As Collection
Dim StatoDoc As StatoRecord
Dim ElencoControlli As Variant, DescControlli As Variant
Dim rsNoteCredito As ADODB.Recordset, rsVociNota As ADODB.Recordset, rsClienti As ADODB.Recordset, _
IntestazioniGriglia As Variant, PropCampiVoce As Variant, cm As CoordinateMouse
Dim FiltroRicerca As Boolean, CaricamentoDoc As Boolean
Dim WithEvents ssc As SmartSubClass
Attribute ssc.VB_VarHelpID = -1
Private Sub Anteprima_Click()
If StatoDoc = inserimento Or CInt(TxtRecCorr.Text) <= rsNoteCredito.RecordCount Then
 If ConvalidaRecord(False) Then
  CaricaTotaliIva
  Dim SS As New ServiziStampa: Call CaricaIntestazioneDitta(SS)
  Set SS.rsDoc = rsNoteCredito: SS.ImpostaAnteprima True
  Set SS.rsCliente = rsClienti: Set SS.TotaliDoc = TotaliNota
  SS.TipoDoc = NotaCredito: Set SS.rsVociDoc = rsVociNota
  SS.Stampa -1: AnteprimaDoc.Show vbModal
 End If
Else
 MsgBox "Inserire i dati del documento !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub BtnCanc_Click()
Dim Scelta%
Scelta = MsgBox("Cancellare il documento corrente ?" & vbNewLine & "Non sarà possibile" _
& " annullare questa modifica !", vbYesNo + vbQuestion, "Fattura Pro")
If Scelta = vbYes Then
 Dim PosRec As Integer
 PosRec = rsNoteCredito.AbsolutePosition
 If StatoDoc <> inserimento Then
  rsNoteCredito.Delete
  rsNoteCredito.Update
  If StatoDoc <> NonModificato Then
   conn.CommitTrans
   StatoDoc = NonModificato
  End If
  If PosRec > rsNoteCredito.RecordCount Then
   If rsNoteCredito.RecordCount <> 0 Then
    rsNoteCredito.MoveLast
   End If
   CreaNuovoRecord
  Else
   rsNoteCredito.MoveNext
   VisualizzaRecord
   TxtRecCorr.Text = rsNoteCredito.AbsolutePosition
   LblNumRecord.Caption = "di " & rsNoteCredito.RecordCount
  End If
 Else
  CreaNuovoRecord
 End If
End If
End Sub
Private Sub BtnEsenzioniIva_Click()
Set EsenzioniIva.FormChiamante = Me
EsenzioniIva.Show vbModal
End Sub
Private Sub BtnSelDitta_Click()
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
 rsClienti.Requery
 rsClienti.Find "Id = " & TxtDitta.Tag, , adSearchForward, adBookmarkFirst
End If
End Sub
Private Sub ChkEsenteIva_Click()
ModificaDoc
If Not CaricamentoDoc Then
 If ChkEsenteIva.Value <> 0 Then
  TotDoc = FormatNumber(CDbl(TotDoc) - CDbl(TotIva), 2)
  TotImp = "0,00": TotIva = "0,00"
  If rsVociNota.RecordCount Then
   rsVociNota.MoveFirst
   For i = 0 To rsVociNota.RecordCount - 1
    ElencoVoci.TextMatrix(i + 1, 6) = ""
    rsVociNota("iva") = ""
    rsVociNota.Update
    rsVociNota.MoveNext
   Next i
   ElencoVoci.TextMatrix(i + 1, 6) = ""
  End If
  PropCampiVoce(6) = ""
 Else
  NonImpIva.Text = ""
  PropCampiVoce(6) = "no"
 End If
End If
End Sub
Private Sub Data_KeyPress(KeyAscii As Integer)
Dim CarAmmessi$
CarAmmessi = "0123456789-" & vbBack
If InStr(CarAmmessi, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub ElencoVoci_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cm.x = x: cm.y = y
End Sub
Private Sub PopupDitte_Click()
If PopupDitte.ListIndex <> -1 Then
 TxtDitta.Text = PopupDitte.Text
 TxtDitta.Tag = PopupDitte.ItemData(PopupDitte.ListIndex)
 rsClienti.Find "Id = " & TxtDitta.Tag, , adSearchForward, adBookmarkFirst
 TxtPartIva.Text = rsClienti("partitaiva")
 PopupDitte.Visible = False
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
  conn.RollbackTrans
 End If
 rsNoteCredito.Close
 ssc.SubClassHwnd ElencoVoci.hWnd, False
 Set rsNoteCredito = Nothing
 Set NoteCredito = Nothing
End If
End Sub
Private Sub TxtNumNota_Change()
ModificaDoc
End Sub
Private Sub TxtDitta_Change()
ModificaDoc
End Sub
Private Sub Data_Change()
ModificaDoc
End Sub
Private Sub TxtNumNota_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtNumNota.Text) = 10 Then
  KeyAscii = 0: Exit Sub
 End If
 If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End If
End Sub
Private Sub NonImpIva_Change()
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
    
If m_bLMouseClicked And (uMsg = WM_ERASEBKGND) Then
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
Set rsVociNota = New ADODB.Recordset
If rsNoteCredito Is Nothing Then
 Set rsNoteCredito = New ADODB.Recordset
 rsNoteCredito.Open "SELECT * FROM NoteCredito ORDER BY NoteCredito.IdDoc ASC", _
 conn, adOpenDynamic, adLockOptimistic
End If
Set rsClienti = New ADODB.Recordset
rsClienti.Open "Clienti", conn, adOpenDynamic
BtnPrec.Enabled = False: BtnPrimo.Enabled = False
BtnUltimo.Enabled = rsNoteCredito.RecordCount > 1
TxtRecCorr.Text = "1"
If Not rsNoteCredito.EOF Then
 If rsNoteCredito.RecordCount >= 1 Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
  BtnCanc.Enabled = True
 End If
 LblNumRecord.Caption = "di " & rsNoteCredito.RecordCount
 rsNoteCredito.MoveFirst
 Call VisualizzaRecord
Else
 LblNumRecord.Caption = "di 1"
 CreaNuovoRecord
End If
IntestazioniGriglia = Array("Descrizione", "Lotto", "U.M.", "Quantità", "Prezzo", _
"Sconto %", "IVA %", "Totale"): ElencoVoci.Cols = 9: ElencoVoci.Width = 15570
ChkEsenteIva.Visible = True
ElencoVoci.ColWidth(0) = 7000: ElencoVoci.ColWidth(1) = 1100
ElencoVoci.ColWidth(2) = 800: ElencoVoci.ColWidth(3) = 1100
ElencoVoci.ColWidth(4) = 1100: ElencoVoci.ColWidth(5) = 1000
ElencoVoci.ColWidth(6) = 800: ElencoVoci.ColWidth(7) = 1200
ElencoVoci.ColWidth(8) = 315
ElencoVoci.HighLight = flexHighlightNever: ElencoVoci.SelectionMode = flexSelectionFree
TxtDitta.Enabled = True: BtnSelDitta.Enabled = True
For i = 0 To ElencoVoci.Cols - 2
 ElencoVoci.TextMatrix(0, i) = IntestazioniGriglia(i): ElencoVoci.ColAlignment(i) = 4
Next i
Set ssc = New SmartSubClass: ssc.SubClassHwnd ElencoVoci.hWnd, True
ElencoControlli = Array("TxtNumNota", "Data", "TxtDitta", "TxtRifDoc", "TxtPartIva", _
"TotDoc", "TotImp", "TotIva")
DescControlli = Array("Num. Documento", "Data", "Ditta", "Rif. Fattura", "Partita Iva", _
"Totale Documento", "Totale Imponibile", "Totale Iva")
PropCampiVoce = Array("ao", "", "a", "n", "no", "n", "n", "no")
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
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
 Dim rsClienti As New ADODB.Recordset
 rsClienti.Open "SELECT * FROM Clienti WHERE Rimosso = False And UCASE(Ditta) LIKE '" & _
 UCase(Replace(TxtDitta.Text, "'", "''")) & "%' ORDER BY Ditta ASC", conn, adOpenDynamic
 If rsClienti.RecordCount <> 0 Then
  rsClienti.MoveFirst
  PopupDitte.Clear
  While Not rsClienti.EOF
   With PopupDitte
   .AddItem rsClienti("Ditta")
   .ItemData(.NewIndex) = rsClienti("Id") & "," & rsClienti("PartitaIva")
   End With
   rsClienti.MoveNext
  Wend
  PopupDitte.Visible = True
  PopupDitte.Left = TxtDitta.Left
  PopupDitte.Top = TxtDitta.Top + TxtDitta.Height + 45
  PopupDitte.Width = TxtDitta.Width
  PopupDitte.Height = 2500
 Else
  PopupDitte.Visible = False
 End If
 rsClienti.Close
Else
 PopupDitte.Visible = False
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
  MsgBox "Attenzione, uno o più campi della voce corrente del documento contengono valori non validi !", _
  vbExclamation, "Fattura Pro": Exit Function
 End If
 If RigaCorr > rsVociNota.RecordCount Then
  rsVociNota.AddNew
 Else
  rsVociNota.Move RigaCorr - 1, adBookmarkFirst
 End If
 With rsVociNota
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
  VociMod = True
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
  If TxtModifica.SelStart = 0 Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
  ElseIf Right(TxtModifica.Text, 1) <> "." And Right(TxtModifica.Text, 1) <> " " Then
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
  End If
  If NumCar = 3 Then
   TxtModifica.Text = TxtModifica.Text & Chr(KeyAscii): KeyAscii = 0: TxtModifica.SelStart = Len(TxtModifica.Text)
   Dim rsArticoli As New ADODB.Recordset
   rsArticoli.Open "SELECT * FROM Articoli WHERE descr LIKE '" & TxtModifica.Text & "%' ORDER BY descr ASC", conn, _
   adOpenDynamic, adLockOptimistic
   If rsArticoli.RecordCount <> 0 Then
    Set ElencoArticoli.rsArticoli = rsArticoli
    Set ElencoArticoli.FormChiamante = Me: ElencoArticoli.Show vbModal
   End If
   rsArticoli.Close
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
   If .Col < 7 And (.Col <> 6 Or ChkEsenteIva.Value = 0) Then
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
     If .Row <= rsVociNota.RecordCount Then
      rsVociFattura.Move .Row - 1, adBookmarkFirst
      rsVociNota.Delete
     End If
     .RemoveItem .Row: VociMod = False
    End If
   End If
  End With
 End If
End If
End Sub
Private Sub BtnPrimo_Click()
PosizioneRecord "primo"
End Sub
Private Sub BtnPrec_Click()
PosizioneRecord "precedente"
End Sub
Private Sub Stampa_Click()
If StatoDoc = inserimento Or CInt(TxtRecCorr.Text) <= rsNoteCredito.RecordCount Then
 If ConvalidaRecord(False) Then
  CaricaTotaliIva
  Dim SS As New ServiziStampa
  Call CaricaIntestazioneDitta(SS)
  SS.TipoDoc = NotaCredito
  Set SS.rsDoc = rsNoteCredito: Set SS.rsVociDoc = rsVociNota
  Set SS.rsCliente = rsClienti: Set SS.TotaliDoc = TotaliNota
  OpzioniStampa.Inizializza SS: OpzioniStampa.Show vbModal
  If OpzioniStampa.Scelta = "Stampa" Then
   SS.ImpostaAnteprima False: SS.Stampa -1
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
Private Sub BtnUltimo_Click()
PosizioneRecord "ultimo"
End Sub
Private Sub BtnNuovo_Click()
PosizioneRecord "nuovo"
End Sub
Private Function ConvalidaRecord(Salva As Boolean) As Boolean
Dim IdNota$, d As Variant, e As Boolean, id As Variant
Dim re As New RegExp, MsgErroreIva$

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
     Data.Text = d(0) & "-" & d(1) & "-" & d(2)
    Else: e = True
    End If
   Else: e = True
   End If
   If e Then
    MsgBox "Attenzione, inserire una data valida !", vbExclamation, "Fattura Pro"
    Data.SetFocus: Data.SelStart = 0: Data.SelLength = Len(Data.Text): Exit Function
   End If
  End If
 Next i
 re.Global = True
 re.Pattern = "^[1-9][0-9]{0,2}/?[1-9]?$"
 Dim NumDoc1$
 NumDoc1 = Split(TxtNumNota, "/")(0)
 IdNota = Mid(Data.Text, 7) & "-" & String(4 - Len(NumDoc1), "0") & TxtNumNota
 If re.Test(TxtNumNota) Then
  Dim CercaDuplicato As Boolean
  If StatoDoc = inserimento Then
   CercaDuplicato = True
  ElseIf IdNota <> rsNoteCredito("IdDoc") Then
   CercaDuplicato = True
  End If
  If CercaDuplicato Then
   Dim rsDuplicato As New ADODB.Recordset
   rsDuplicato.Open "SELECT * FROM NoteCredito WHERE IdDoc = '" & IdNota & "'", conn, adOpenDynamic
   If Not rsDuplicato.EOF Then
    MsgBox "Attenzione, il numero documento corrisponde a quello di un altro documento" _
    & " in archivio !" & vbNewLine, vbExclamation, "Fattura Pro"
    TxtNumNota.SetFocus: TxtNumNota.SelStart = 0
    TxtNumNota.SelLength = Len(TxtNumNota): Exit Function
   End If
  End If
 Else
  MsgBox "Attenzione, il numero documento non è valido !", vbExclamation, "Fattura Pro"
  TxtNumNota.SetFocus: TxtNumNota.SelStart = 0: TxtNumNota.SelLength = Len(TxtNumNota)
  Exit Function
 End If
 If SalvaVoceCorr() Then
  If rsVociNota.RecordCount = 0 Then
   MsgBox "Ogni documento deve avere almeno una voce !", vbExclamation, "Fattura Pro"
   Exit Function
  End If
  Dim VerificaNumDoc As Boolean, NumValido As Boolean: NumValido = True
  If StatoDoc = inserimento Then
   VerificaNumDoc = True
  ElseIf IdNota <> rsNoteCredito("IdDoc") Then
   VerificaNumDoc = True
  End If
  If VerificaNumDoc Then
   id = Split(TxtNumNota, "/"): NumValido = NumDocValido(id, d(2))
  End If
  If Not NumValido Then
   MsgBox "Attenzione, il numero documento non è valido !", vbExclamation, "Fattura Pro"
   TxtNumNota.SetFocus: TxtNumNota.SelStart = 0
   TxtNumNota.SelLength = Len(TxtNumNota): Exit Function
  End If
  If Not ControlloIva(MsgErroreIva) Then
   MsgBox MsgErroreIva, vbExclamation, "Fattura Pro"
   ConvalidaRecord = False: Exit Function
  End If
  If TxtDitta.Tag = "" Or (Not VerificaDestDoc()) Then
   MsgBox "Attenzione, il campo ditta non contiene un valore valido !", vbExclamation, _
   "Fattura Pro"
   Exit Function
  End If
  re.Global = True
  re.Pattern = "^[1-9][0-9]*/?[0-9]?-[0-9]{4}$"
  If Not re.Test(TxtRifDoc.Text) Then
   e = True
  Else
   Dim rsRifDoc As New ADODB.Recordset
   Dim NumDoc$, AnnoDoc$
   AnnoDoc = Split(TxtRifDoc.Text, "-")(1)
   NumDoc = IdDocumento(Split(TxtRifDoc.Text, "-")(0))
   rsRifDoc.Open "SELECT * FROM FattureClienti WHERE IdDoc = '" & AnnoDoc & "-" & NumDoc & "'", conn, _
   adOpenDynamic
   e = rsRifDoc.EOF
  End If
  If e Then
   MsgBox "Inserire un numero di riferimento fattura valido !", vbExclamation, "Fattura Pro"
   ConvalidaRecord = False: Exit Function
  End If
  If StatoDoc = inserimento Then
   rsNoteCredito.AddNew
   StatoDoc = modifica
  End If
  rsNoteCredito("IdDoc") = IdNota
  rsNoteCredito("Data") = Data.Text
  rsNoteCredito("IdDitta") = TxtDitta.Tag
  rsNoteCredito("RifDoc") = TxtRifDoc.Text
  rsNoteCredito("Pagato") = Pagato.Value
  rsNoteCredito("TotDoc") = CDbl(TotDoc.Text)
  rsNoteCredito("TotImp") = CDbl(TotImp.Text)
  rsNoteCredito("TotIva") = CDbl(TotIva.Text)
  rsNoteCredito.Update
  rsVociNota.MoveFirst
  While Not rsVociNota.EOF
   If IsNull(rsVociNota("IdNota")) Then
    rsVociNota("IdNota") = IdNota
    rsVociNota.Update
   End If
   rsVociNota.MoveNext
  Wend
  If Salva Then
   EseguiBackup = True
   StatoDoc = NonModificato
   conn.CommitTrans
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
 If rsNoteCredito.RecordCount > 1 Then
  Dim NumRec%: NumRec = rsNoteCredito.AbsolutePosition
  rsNoteCredito.Requery
  rsNoteCredito.Move NumRec - 1, adBookmarkFirst
 End If
 Select Case PosRecord
 Case "primo"
  rsNoteCredito.MoveFirst
 Case "ultimo"
  rsNoteCredito.MoveLast
 Case "precedente"
  If CInt(TxtRecCorr.Text) <= rsNoteCredito.RecordCount Then
   rsNoteCredito.MovePrevious
  ElseIf rsNoteCredito.AbsolutePosition <> rsNoteCredito.RecordCount Then
   rsNoteCredito.MoveLast
  End If
 Case "successivo"
  If rsNoteCredito.AbsolutePosition = rsNoteCredito.RecordCount Then
   CreaNuovoRecord
   Exit Sub
  End If
  rsNoteCredito.MoveNext
 Case "nuovo"
  CreaNuovoRecord
  Exit Sub
 End Select
 BtnCanc.Enabled = True
 If rsNoteCredito.AbsolutePosition <= rsNoteCredito.RecordCount Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
 End If
 If rsNoteCredito.AbsolutePosition <> rsNoteCredito.RecordCount Then
  BtnUltimo.Enabled = True
 Else
  BtnUltimo.Enabled = False
 End If
 If rsNoteCredito.AbsolutePosition <> 1 Then
  BtnPrimo.Enabled = True: BtnPrec.Enabled = True
 Else
  BtnPrimo.Enabled = False: BtnPrec.Enabled = False
 End If
 TxtRecCorr.Text = rsNoteCredito.AbsolutePosition
 LblNumRecord.Caption = "di " & rsNoteCredito.RecordCount
 rsVociNota.Close
 VisualizzaRecord
End If
End Sub
Private Sub VisualizzaRecord()
CaricamentoDoc = True
With rsNoteCredito
 TxtNumNota = NumeroDocumento(.Fields("iddoc"))
 Data = Format(.Fields("data"), "dd-mm-yyyy")
 rsClienti.Find "Id = " & .Fields("IdDitta"), , adSearchForward, adBookmarkFirst
 TxtDitta.Text = rsClienti("ditta")
 TxtDitta.Tag = .Fields("idditta")
 TxtPartIva.Text = rsClienti("partitaiva")
 ElencoVoci.Rows = 1
 Pagato.Value = .Fields("pagato")
 TotDoc = FormatNumber(.Fields("totdoc"), 2)
 TotImp = FormatNumber(.Fields("totimp"), 2)
 TotIva = FormatNumber(.Fields("totiva"), 2)
 If Not IsNull(.Fields("esenteiva")) Then
  NonImpIva.Text = .Fields("esenteiva")
 Else
  NonImpIva.Text = ""
 End If
 ChkEsenteIva.Value = IIf(NonImpIva.Text <> "", 1, 0)
End With
If rsVociNota.State <> adStateClosed Then
 rsVociNota.Close
End If
rsVociNota.Open "SELECT * FROM VociNoteCredito WHERE IdNota = '" & rsNoteCredito("IdDoc") & "' ORDER BY Id ASC", conn, adOpenDynamic, adLockOptimistic
rsVociNota.MoveFirst
With ElencoVoci
 Dim CifreDec%, i%
 While Not rsVociNota.EOF
  .AddItem "": i = 0
  .RowHeight(.Rows - 1) = 315
  For Each Field In rsVociNota.Fields
   If Field.Name <> "Id" And Field.Name <> "IdNota" Then
    If Field.Type = adDouble Then
     If Field.Name = "Qnt" Then
      CifreDec = 3
     Else
      CifreDec = 2
     End If
     If Field.Value <> 0 Then
      .TextMatrix(.Rows - 1, i) = FormatNumber(rsVociNota(Field.Name), CifreDec)
     End If
    Else
     .TextMatrix(.Rows - 1, i) = rsVociNota(Field.Name)
    End If
    i = i + 1
   End If
  Next
  .Row = .Rows - 1: .Col = 8: .CellPictureAlignment = 4
  Set .CellPicture = ImgCancella
  rsVociNota.MoveNext
 Wend
 .AddItem "": .RowHeight(ElencoVoci.Rows - 1) = 315
End With
If StatoDoc <> NonModificato Then
 StatoDoc = NonModificato
 conn.CommitTrans
End If
CaricamentoDoc = False
End Sub
Private Sub CaricaTotaliIva()
Dim Totale As Double, Iva As Double, Aliquota As Integer
Dim TI As TotaleIva
Set TotaliNota = New Collection
If rsVociNota.RecordCount <> 0 Then
 rsVociNota.MoveFirst
 On Error Resume Next
 While Not rsVociNota.EOF
  If rsVociNota("Iva") <> "" Then
   Aliquota = CDbl(rsVociNota("iva"))
   Totale = CDbl(rsVociNota("totale"))
   Iva = CDbl(Arrotonda(Totale * Aliquota / 100))
   Set TI = TotaliNota(CStr(Aliquota))
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
    TotaliNota.Add TI, CStr(Aliquota)
    Err.Clear
   End If
  End If
  rsVociNota.MoveNext
 Wend
 TotDoc = FormatNumber(TotImp + TotIva, 2)
 If Split(TotDoc, ",")(1) = "99" Then TotDoc = FormatNumber(CDbl(TotDoc) + 0.01, 2)
 For Each TotaleIva In TotaliNota
  If InStr(1, CStr(TotaleIva.Totale), ",99") Then
   TotaleIva.Totale = TotaleIva.Totale + 0.01
  End If
 Next
End If
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
Dim rsClienti As New ADODB.Recordset
rsClienti.Open "SELECT * FROM Clienti WHERE UCASE(Ditta) = '" & UCase(Replace(TxtDitta.Text, "'", "''")) & "'", conn, adOpenDynamic, adLockOptimistic
If Not rsClienti.EOF Then
 TxtDitta.Tag = rsClienti("Id")
 TxtPartIva.Text = rsClienti("partitaiva")
 VerificaDestDoc = True
End If
End Function
Private Sub CreaNuovoRecord()
TxtNumNota.Text = ""
TxtDitta.Text = "": TxtDitta.Tag = ""
Data.Text = "": LblNumRecord.Caption = "di " & (rsNoteCredito.RecordCount + 1)
ElencoVoci.Rows = 1: ElencoVoci.AddItem "": Pagato.Value = 0
ElencoVoci.RowHeight(1) = 315: TxtPartIva.Text = ""
TotDoc = "0,00": TotImp = "0,00": TotIva = "0,00": Differita = 0
BtnSucc.Enabled = False: BtnCanc.Enabled = False: BtnNuovo.Enabled = False
RigaCorr = 0: TxtRecCorr.Text = rsNoteCredito.RecordCount + 1
ChkEsenteIva.Value = 0: NonImpIva.Text = ""
If rsNoteCredito.RecordCount >= 1 Then
 BtnPrec.Enabled = True: BtnPrimo.Enabled = True: BtnUltimo.Enabled = True
End If
If rsVociNota.State <> adStateClosed Then
 rsVociNota.Close
End If
rsVociNota.Open "SELECT * FROM VociNoteCredito WHERE IdNota = '0'", conn, adOpenDynamic, adLockOptimistic
StatoDoc = NonModificato
conn.CommitTrans
End Sub
Private Sub BtnVai_Click()
If Not rsNoteCredito.EOF Then
 IdCorr = NumeroDocumento(rsNoteCredito("IdDoc")) & "-" & Mid(rsNoteCredito("IdDoc"), 1, 4)
End If
If TxtNumNota <> "" And TxtNumNota <> IdCorr Then
 If InStr(TxtNumNota, "-") <> 0 Then
  Dim IdDoc, NumDoc1$, IdNota$, PosRec%
  PosRec = rsNoteCredito.AbsolutePosition
  IdDoc = Split(TxtNumNota, "-")
  NumDoc1 = Split(IdDoc(0), "/")(0)
  IdNota = IdDoc(1) & "-" & String(4 - Len(NumDoc1), "0") & IdDoc(0)
  rsNoteCredito.Find "IdDoc = '" & IdNota & "'", , adSearchForward, adBookmarkFirst
  If Not rsNoteCredito.EOF Then
   TxtRecCorr.Text = rsNoteCredito.AbsolutePosition
   VisualizzaRecord
   If rsNoteCredito.AbsolutePosition > 1 Then
    BtnPrec.Enabled = True: BtnPrimo.Enabled = True
   End If
   If rsNoteCredito.AbsolutePosition < rsNoteCredito.RecordCount Then
    BtnUltimo.Enabled = True
   End If
   BtnNuovo.Enabled = True: BtnSucc.Enabled = True
  Else
   If rsNoteCredito.RecordCount Then
    rsNoteCredito.Move PosRec - 1, adBookmarkFirst
   End If
   MsgBox "Attenzione, documento non presente in archivio !", vbExclamation, "Fattura Pro"
   If StatoDoc <> NonModificato Then
    StatoDoc = NonModificato
    conn.CommitTrans
   End If
  End If
 Else
  If StatoDoc = inserimento Then
   TxtNumNota.Text = ""
  ElseIf StatoDoc = modifica Then
   TxtNumNota.Text = IdCorrente
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
If CInt(TxtRecCorr.Text) > rsNoteCredito.RecordCount Then
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
If AvviaTrans Then conn.BeginTrans
End Sub
Public Sub ImpostaFiltroRicerca(rsRicerca As ADODB.Recordset)
Set rsNoteCredito = rsRicerca
FiltroRicerca = True
End Sub
Public Sub CaricaFiltroRicerca()
PosizioneRecord "primo"
End Sub
Public Function FormBloccata() As Boolean
FormBloccata = StatoDoc <> NonModificato
End Function
Public Function VerificaAliquote() As Boolean
Dim re As New RegExp
VerificaAliquote = True: re.Pattern = "Spes[ae] ?(di)? ?Spedizion[ei]"
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
   & " l'aliquota iva !"
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
  If IdPrec <> rsNoteCredito("IdDoc") Then
   Dim rsRecPrec As New ADODB.Recordset
   rsRecPrec.Open "SELECT * FROM NoteCredito WHERE IdDoc = " & IdPrec, conn, adOpenStatic
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
IdDoc = Anno & "-" & String(4 - Len(NumDoc1), "0") & TxtNumNota
DataDoc = Format$(Data, "yyyy/mm/dd")
rsRecInvalido.Open "SELECT * FROM NoteCredito WHERE (IdDoc < '" & IdDoc & "' AND Data > #" & DataDoc & "#) OR " _
& "(IdDoc > '" & IdDoc & "' AND Data < #" & DataDoc & "#)", conn, adOpenStatic
NumDocValido = rsRecInvalido.EOF
rsRecInvalido.Close
End Function
Private Sub TxtRifDoc_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtRifDoc.Text) = 8 Then
  KeyAscii = 0: Exit Sub
 End If
 If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End If
End Sub
