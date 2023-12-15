VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FattureFornitori 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fatture Fornitori"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12570
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FattureFornitori.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   12570
   Begin VB.CommandButton BtnFatturaElettronica 
      Caption         =   "Importa Fattura Elettronica"
      Height          =   390
      Left            =   8115
      TabIndex        =   48
      Top             =   810
      Width           =   2325
   End
   Begin MSComDlg.CommonDialog SelezionaFile 
      Left            =   7170
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox CmbPag 
      Height          =   345
      ItemData        =   "FattureFornitori.frx":4072
      Left            =   4935
      List            =   "FattureFornitori.frx":4082
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   675
      Width           =   1650
   End
   Begin VB.TextBox TxtPartIva 
      Height          =   315
      Left            =   1755
      TabIndex        =   43
      Top             =   675
      Width           =   1935
   End
   Begin VB.ListBox PopupDitte 
      Height          =   510
      ItemData        =   "FattureFornitori.frx":40B8
      Left            =   11445
      List            =   "FattureFornitori.frx":40BA
      TabIndex        =   41
      Top             =   735
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox TotNetto 
      Height          =   315
      Left            =   9765
      TabIndex        =   30
      Top             =   4380
      Width           =   1245
   End
   Begin VB.CommandButton BtnSelDitta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11190
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FattureFornitori.frx":40BC
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   255
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin TabDlg.SSTab DatiDoc 
      Height          =   2910
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1140
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   5133
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   556
      BackColor       =   3394815
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Voci Documento"
      TabPicture(0)   =   "FattureFornitori.frx":4217
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PanVoci"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ritenuta d'acconto"
      TabPicture(1)   =   "FattureFornitori.frx":4233
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PanRitenuta"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox PanRitenuta 
         BackColor       =   &H0033CCFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   -74985
         ScaleHeight     =   2565
         ScaleWidth      =   11115
         TabIndex        =   26
         Top             =   330
         Visible         =   0   'False
         Width           =   11115
         Begin VB.CheckBox ChkRitCassa 
            BackColor       =   &H0033CCFF&
            Caption         =   "Contributo Cassa sogg. a Ritenuta"
            Enabled         =   0   'False
            Height          =   225
            Left            =   270
            TabIndex        =   49
            Top             =   1875
            Width           =   3030
         End
         Begin VB.TextBox TxtCassaPro 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8190
            TabIndex        =   47
            Top             =   1230
            Width           =   1140
         End
         Begin VB.CheckBox ChkCalcRit 
            Appearance      =   0  'Flat
            BackColor       =   &H0033CCFF&
            Caption         =   "Calcola Ritenuta"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   210
            TabIndex        =   44
            Top             =   195
            Width           =   1650
         End
         Begin VB.TextBox TxtAliqIva 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4755
            TabIndex        =   40
            Top             =   1230
            Width           =   780
         End
         Begin VB.TextBox TxtSpeseDoc 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1935
            TabIndex        =   37
            Top             =   1230
            Width           =   990
         End
         Begin VB.TextBox TxtAliqCassaPro 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8190
            TabIndex        =   35
            Top             =   780
            Width           =   720
         End
         Begin VB.TextBox TxtRitenuta 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4755
            TabIndex        =   33
            Top             =   780
            Width           =   1215
         End
         Begin VB.TextBox TxtAliqRit 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1935
            TabIndex        =   29
            Top             =   780
            Width           =   975
         End
         Begin VB.Label LblAliqIva 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aliquota Iva:"
            Height          =   225
            Left            =   3705
            TabIndex        =   39
            Top             =   1260
            Width           =   990
         End
         Begin VB.Label LblBaseImp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cassa Professionisti:"
            Height          =   225
            Left            =   6540
            TabIndex        =   38
            Top             =   1275
            Width           =   1590
         End
         Begin VB.Label LblSpeseDoc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spese Documentate:"
            Height          =   225
            Left            =   270
            TabIndex        =   36
            Top             =   1260
            Width           =   1620
         End
         Begin VB.Label LblCassaPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cassa Professionisti (%):"
            Height          =   225
            Left            =   6240
            TabIndex        =   34
            Top             =   810
            Width           =   1905
         End
         Begin VB.Label LblRitenuta 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ritenuta d'acconto:"
            Height          =   225
            Left            =   3150
            TabIndex        =   32
            Top             =   810
            Width           =   1545
         End
         Begin VB.Label LblRitPerc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aliquota Ritenuta:"
            Height          =   225
            Left            =   450
            TabIndex        =   28
            Top             =   810
            Width           =   1425
         End
      End
      Begin VB.PictureBox PanVoci 
         BackColor       =   &H0033CCFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2490
         Left            =   45
         ScaleHeight     =   2490
         ScaleWidth      =   11115
         TabIndex        =   24
         Top             =   330
         Width           =   11115
         Begin MSFlexGridLib.MSFlexGrid ElencoVoci 
            Height          =   2010
            Left            =   15
            TabIndex        =   25
            Top             =   30
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   3545
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   315
            BackColorFixed  =   14737632
            BackColorBkg    =   24576
            GridColorFixed  =   8421504
            FocusRect       =   0
            HighLight       =   0
            GridLinesFixed  =   1
            BorderStyle     =   0
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
      End
   End
   Begin VB.TextBox TxtDitta 
      Height          =   315
      Left            =   7320
      TabIndex        =   21
      Top             =   255
      Width           =   3825
   End
   Begin VB.CheckBox Pagato 
      Appearance      =   0  'Flat
      BackColor       =   &H0033CCFF&
      Caption         =   "Saldato"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6855
      TabIndex        =   20
      Top             =   795
      Width           =   900
   End
   Begin VB.TextBox TxtModifica 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   9645
      TabIndex        =   19
      Top             =   4980
      Visible         =   0   'False
      Width           =   1275
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
      Picture         =   "FattureFornitori.frx":424F
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4980
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
      Left            =   3045
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureFornitori.frx":47E9
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4980
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
      Left            =   2700
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureFornitori.frx":4D83
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TxtRecCorr 
      Height          =   315
      Left            =   1530
      TabIndex        =   13
      Top             =   4980
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
      Left            =   1155
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureFornitori.frx":531D
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4980
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
      Left            =   810
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureFornitori.frx":58B7
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
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
      Left            =   3825
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureFornitori.frx":5E51
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4980
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TxtNumFattura 
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Top             =   255
      Width           =   1920
   End
   Begin VB.TextBox Data 
      Height          =   315
      Left            =   4935
      TabIndex        =   1
      Top             =   255
      Width           =   1395
   End
   Begin VB.TextBox TotDoc 
      Height          =   315
      Left            =   1410
      TabIndex        =   2
      Top             =   4380
      Width           =   1365
   End
   Begin VB.TextBox TotImp 
      Height          =   315
      Left            =   4515
      TabIndex        =   3
      Top             =   4380
      Width           =   1245
   End
   Begin VB.TextBox TotIva 
      Height          =   315
      Left            =   6900
      TabIndex        =   4
      Top             =   4380
      Width           =   1290
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pagamento:"
      Height          =   225
      Left            =   3930
      TabIndex        =   46
      Top             =   720
      Width           =   960
   End
   Begin VB.Label LblPartIva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Partita IVA:"
      Height          =   225
      Left            =   210
      TabIndex        =   42
      Top             =   735
      Width           =   870
   End
   Begin VB.Label LblTotNetto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Dovuto:"
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
      Left            =   8460
      TabIndex        =   31
      Top             =   4410
      Width           =   1245
   End
   Begin VB.Label LblDitta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ditta:"
      Height          =   225
      Left            =   6855
      TabIndex        =   22
      Top             =   285
      Width           =   420
   End
   Begin VB.Label LblNumRecord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   4275
      TabIndex        =   18
      Top             =   5025
      Width           =   45
   End
   Begin VB.Label Legenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   150
      TabIndex        =   17
      Top             =   5010
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Documento:"
      Height          =   225
      Left            =   210
      TabIndex        =   9
      Top             =   285
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   225
      Left            =   4470
      TabIndex        =   8
      Top             =   285
      Width           =   405
   End
   Begin VB.Label Label5 
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
      Left            =   135
      TabIndex        =   7
      Top             =   4410
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
      Left            =   2955
      TabIndex        =   6
      Top             =   4410
      Width           =   1500
   End
   Begin VB.Label Label7 
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
      Height          =   315
      Left            =   5970
      TabIndex        =   5
      Top             =   4410
      Width           =   1035
   End
End
Attribute VB_Name = "FattureFornitori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VociMod As Boolean
Dim RigaCorr As Long, TotaliFattura As Collection
Dim StatoDoc As StatoRecord
Dim ElencoControlli As Variant, DescControlli As Variant
Dim rsFatture As ADODB.Recordset, rsVociFattura As ADODB.Recordset, rsFornitori As ADODB.Recordset, _
rsRitenute As ADODB.Recordset, PropCampiVoce As Variant, cm As CoordinateMouse
Dim FiltroRicerca As Boolean, CaricamentoDocumento As Boolean
Dim WithEvents ssc As SmartSubClass
Attribute ssc.VB_VarHelpID = -1
Private Sub BtnCanc_Click()
Dim Scelta%
Scelta = MsgBox("Cancellare il documento corrente ?" & vbNewLine & "Non sarà possibile" _
& " annullare questa modifica !", vbYesNo + vbQuestion, "Fattura Pro")
If Scelta = vbYes Then
 Dim PosRec%
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
SelezionaFile.FileName = ""
SelezionaFile.Filter = "File XML (*.xml)|*.xml"
SelezionaFile.ShowOpen
If SelezionaFile.FileName <> "" Then
 Dim FE As New FatturaElettronica, Errore As String
 FE.ImportaFattura SelezionaFile.FileName, Me, Errore
 If Errore <> "" Then
  MsgBox "Il documento selezionato non è una fattura elettronica valida" & vbNewLine _
  & Errore, vbCritical, "Fattura Pro"
 Else
  rsFornitori.Requery
 End If
End If
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
 CmbPag.ListIndex = SelezioneDitta.rsDitte("modpag")
 rsFornitori.Requery
 rsFornitori.Find "Id = " & TxtDitta.Tag, , adSearchForward, adBookmarkFirst
End If
End Sub
Private Sub ChkCalcRit_Click()
If ChkCalcRit.Value = 0 Then
 If TxtCassaPro.Text <> "" Then
  TotImp = FormatNumber(CDbl(TotImp.Text) - CDbl(TxtCassaPro.Text), 2)
  TotIva = FormatNumber(CDbl(TotIva.Text) - (CDbl(TxtCassaPro.Text) * (CDbl(TxtAliqIva.Text) / 100)), 2)
  TotDoc.Text = FormatNumber(CDbl(TotImp) + CDbl(TotIva), 2)
 End If
 TotNetto.Text = TotDoc.Text
 TxtAliqRit.Text = "": TxtRitenuta.Text = ""
 TxtAliqCassaPro.Text = "": TxtCassaPro.Text = ""
 TxtSpeseDoc.Text = "": TxtAliqIva.Text = ""
 TxtAliqRit.Enabled = False: TxtRitenuta.Enabled = False
 TxtAliqCassaPro.Enabled = False: TxtCassaPro.Enabled = False
 TxtSpeseDoc.Enabled = False: TxtAliqIva.Enabled = False
 ChkRitCassa.Enabled = False
 If Not rsRitenute.EOF Then
  rsRitenute.Delete
  rsRitenute.Requery
 End If
Else
 TxtAliqRit.Enabled = True: TxtRitenuta.Enabled = True
 TxtAliqCassaPro.Enabled = True: TxtCassaPro.Enabled = True
 TxtSpeseDoc.Enabled = True: TxtAliqIva.Enabled = True
 ChkRitCassa.Enabled = True
End If
ModificaDoc
End Sub
Private Sub ChkRitCassa_Click()
If TxtRitenuta.Text <> "" Then
 CalcolaRitenutaAcconto
End If
End Sub
Private Sub CmbPag_Click()
ModificaDoc
End Sub
Private Sub Data_KeyPress(KeyAscii As Integer)
Dim CarAmmessi$
CarAmmessi = "0123456789-" & vbBack
If InStr(CarAmmessi, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub DatiDoc_Click(PreviousTab As Integer)
If DatiDoc.Tab = 0 Then
 PanVoci.Visible = True
 PanRitenuta.Visible = False
Else
 PanRitenuta.Visible = True
 PanVoci.Visible = False
End If
End Sub
Private Sub ElencoVoci_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cm.x = x: cm.y = y
End Sub
Private Sub PopupDitte_Click()
If PopupDitte.ListIndex <> -1 Then
 TxtDitta.Text = PopupDitte.Text
 TxtDitta.Tag = PopupDitte.ItemData(PopupDitte.ListIndex)
 rsFornitori.Find "Id = " & TxtDitta.Tag, , adSearchForward, adBookmarkFirst
 TxtPartIva.Text = rsFornitori("partitaiva")
 CmbPag.ListIndex = rsFornitori("modpag")
 PopupDitte.Visible = False
End If
End Sub
Private Sub TxtAliqCassaPro_Change()
If ChkCalcRit.Value Then
 CalcolaRitenutaAcconto
End If
End Sub
Private Sub TxtAliqIva_Change()
If ChkCalcRit.Value Then
 CalcolaRitenutaAcconto
End If
End Sub
Private Sub TxtAliqRit_Change()
If ChkCalcRit.Value Then
 CalcolaRitenutaAcconto
End If
End Sub
Private Sub TxtCassaPro_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtCassaPro_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtDitta_LostFocus()
PopupDitte.Visible = False
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
 vbYesNoCancel + vbQuestion, "Chiusura Fatture Fornitori")
 If Scelta = vbYes Then
  If Not ConvalidaRecord Then
   Cancel = 1
  End If
 ElseIf Scelta = vbCancel Then
  Cancel = 1
 End If
End If
If Cancel <> 1 Then
 If StatoDoc <> NonModificato Then
  StatoDoc = NonModificato
  conn.RollbackTrans
 End If
 rsFatture.Close
 Set rsFatture = Nothing
 Set FattureFornitori = Nothing
End If
End Sub
Private Sub TxtNumFattura_Change()
ModificaDoc
End Sub
Private Sub TxtDitta_Change()
ModificaDoc
End Sub
Private Sub Data_Change()
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
 If ElencoVoci.ColWidth(1) < 800 Then
  ElencoVoci.ColWidth(1) = 800
 End If
 m_bLMouseClicked = False
End If
End Sub
Private Sub TxtRitenuta_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtRitenuta_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtSpeseDoc_Change()
If ChkCalcRit.Value Then
 CalcolaRitenutaAcconto
End If
End Sub
Private Sub TotDoc_Change()
ModificaDoc
End Sub
Private Sub Form_Load()
Me.Move 600, 450
Dim IntestazioniElencoVoci
Set rsFornitori = New ADODB.Recordset
Set rsRitenute = New ADODB.Recordset
rsFornitori.Open "Fornitori", conn, adOpenDynamic
rsRitenute.Open "Ritenute", conn, adOpenDynamic, adLockOptimistic
Set rsAliquoteIVA = New ADODB.Recordset
If rsFatture Is Nothing Then
 Set rsFatture = New ADODB.Recordset
 rsFatture.Open "SELECT * FROM FattureFornitori ORDER BY Data ASC", conn, adOpenDynamic, adLockOptimistic
End If
Set rsVociFattura = New ADODB.Recordset
rsAliquoteIVA.Open "SELECT * FROM AliquoteIVA ORDER BY Aliquota ASC", conn, adOpenDynamic, adLockOptimistic
BtnPrec.Enabled = False: BtnPrimo.Enabled = False
BtnUltimo.Enabled = rsFatture.RecordCount > 1
TxtRecCorr.Text = "1"
ElencoVoci.Cols = 7
IntestazioniElencoVoci = Array("Descrizione", "U.M.", "Quantità", "Prezzo", "IVA", _
"Totale", "")
For i = 0 To ElencoVoci.Cols - 1
 ElencoVoci.TextMatrix(0, i) = IntestazioniElencoVoci(i): ElencoVoci.ColAlignment(i) = 4
Next i
If Not rsFatture.EOF Then
 BtnSucc.Enabled = True: BtnNuovo.Enabled = True
 BtnCanc.Enabled = True
 LblNumRecord.Caption = "di " & rsFatture.RecordCount
 rsFatture.MoveFirst
 Call VisualizzaRecord
Else
 LblNumRecord.Caption = "di 1"
 ElencoVoci.AddItem "": ElencoVoci.RowHeight(ElencoVoci.Rows - 1) = 315
 CreaNuovoRecord
End If
DatiDoc.Move 90, DatiDoc.Top, Me.ScaleWidth - 180, DatiDoc.Height
PanVoci.Move 15, 330, DatiDoc.Width - 30, DatiDoc.Height - 345
PanRitenuta.Move 15, 330, DatiDoc.Width - 30, DatiDoc.Height - 345
With ElencoVoci
.Move 0, 0, PanVoci.Width, PanVoci.Height
.ColWidth(0) = 6500
.ColWidth(1) = 800
.ColWidth(2) = 1100
.ColWidth(3) = 1100
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 360
End With
PanRitenuta.Visible = False
DatiDoc.Tab = 0
Set ssc = New SmartSubClass: ssc.SubClassHwnd ElencoVoci.hWnd, True
ElencoControlli = Array("TxtNumFattura", "Data", "TxtDitta", "CmbPag", "TotDoc", "TotImp", "TotIva")
DescControlli = Array("Num. Documento", "Data", "Ditta", "Pagamento", "Totale Documento", _
"Totale Imponibile", "Totale Iva")
PropCampiVoce = Array("ao", "", "n", "no", "n", "no")
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
 Dim rsFornitori As New ADODB.Recordset
 rsFornitori.Open "SELECT * FROM Fornitori WHERE Rimosso = False And UCASE(Ditta) LIKE '" & _
 UCase(Replace(TxtDitta.Text, "'", "''")) & "%' ORDER BY Ditta ASC", conn, adOpenDynamic
 If Not rsFornitori.EOF Then
  PopupDitte.Clear
  rsFornitori.MoveFirst
  While Not rsFornitori.EOF
   With PopupDitte
    .AddItem rsFornitori("Ditta")
    .ItemData(.NewIndex) = rsFornitori("Id")
   End With
   rsFornitori.MoveNext
  Wend
  PopupDitte.Visible = True
  PopupDitte.Left = TxtDitta.Left
  PopupDitte.Top = TxtDitta.Top + TxtDitta.Height + 45
  PopupDitte.Width = TxtDitta.Width
  PopupDitte.Height = 2500
 Else
  PopupDitte.Visible = False
 End If
Else
 PopupDitte.Visible = False
End If
End Sub
Private Sub TxtModifica_LostFocus()
TxtModifica.Visible = False
End Sub
Private Function ConvalidaRigaDoc() As Boolean
ConvalidaRigaDoc = False: Dim e As Boolean
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
 'If RigaCorr > rsVociFattura.RecordCount Then
  'rsVociFattura.AddNew
 'Else
  'rsVociFattura.Move RigaCorr - 1, adBookmarkFirst
 'End If
 'With rsVociFattura
  '.Fields("descr") = EliminaSpazi(ElencoVoci.TextMatrix(RigaCorr, 0))
  '.Fields("um") = ElencoVoci.TextMatrix(RigaCorr, 1)
  'If ElencoVoci.TextMatrix(RigaCorr, 2) <> "" Then
   '.Fields("qnt") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 2))
  'Else
   '.Fields("qnt") = 0
  'End If
  '.Fields("prezzo") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 3))
  '.Fields("iva") = ElencoVoci.TextMatrix(RigaCorr, 4)
  '.Fields("totale") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 5))
  '.Update
 'End With
 VociMod = False
End If
ConvalidaRigaDoc = True
End Function
Private Sub TxtModifica_Change()
If TxtModifica.Text <> ElencoVoci.Text Then
 Dim Qnt#, Diff#, TotCorr#, NuovoTot#, NuovaIva#, IvaCorr#
 ModificaDoc
 VociMod = True
 If ElencoVoci.Col <> 4 Then
  If ElencoVoci.Col = 2 Then
   If IsNumeric(TxtModifica.Text) Then
    ElencoVoci.Text = FormatNumber(CDbl(TxtModifica.Text), 3)
   Else
    ElencoVoci.Text = ""
   End If
  ElseIf ElencoVoci.Col = 3 Then
   If IsNumeric(TxtModifica.Text) Then
    ElencoVoci.Text = FormattaNumero(TxtModifica.Text)
   Else
    ElencoVoci.Text = ""
   End If
  Else: ElencoVoci.Text = Trim(TxtModifica.Text)
  End If
  If (ElencoVoci.Col = 2 Or ElencoVoci.Col = 3) And ElencoVoci.TextMatrix(RigaCorr, 3) <> "" Then
   If ElencoVoci.TextMatrix(RigaCorr, 5) <> "" Then
    TotCorr = CDbl(ElencoVoci.TextMatrix(RigaCorr, 5))
   End If
   Qnt = 1
   If ElencoVoci.TextMatrix(RigaCorr, 2) <> "" Then
    Qnt = CDbl(ElencoVoci.TextMatrix(RigaCorr, 2))
   End If
   NuovoTot = CDbl(ElencoVoci.TextMatrix(RigaCorr, 3)) * Qnt
   ElencoVoci.TextMatrix(RigaCorr, 5) = FormatNumber(NuovoTot, 2)
   Diff = CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)) - TotCorr
   If ElencoVoci.TextMatrix(RigaCorr, 4) <> "" Then
    TotImp.Text = FormatNumber(CDbl(TotImp) + Diff, 2)
    IvaCorr = TotCorr * CDbl(ElencoVoci.TextMatrix(RigaCorr, 4)) / 100
    NuovaIva = CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)) * _
    CDbl(ElencoVoci.TextMatrix(RigaCorr, 4)) / 100
    Diff = Round(NuovaIva, 2) - Round(IvaCorr, 2)
    TotIva.Text = FormatNumber(CDbl(TotIva.Text) + Diff, 2)
    TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text) + TotNonImp, 2)
   Else
    TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) + Diff, 2)
   End If
  End If
 Else
  If ElencoVoci.TextMatrix(RigaCorr, 5) <> "" Then
   If ElencoVoci.Text <> "" And IsNumeric(ElencoVoci.Text) Then
    IvaCorr = CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)) * CDbl(ElencoVoci.Text) / 100
   End If
   ElencoVoci.Text = TxtModifica.Text
   If ElencoVoci.Text <> "" And IsNumeric(ElencoVoci.Text) Then
    NuovaIva = CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)) * CDbl(ElencoVoci.Text) / 100
   End If
   Diff = Round(NuovaIva, 2) - Round(IvaCorr, 2)
   If Diff <> 0 Then
    If NuovaIva = 0 Then
     TotImp.Text = FormatNumber(CDbl(TotImp.Text) - CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)), 2)
    ElseIf IvaCorr = 0 Then
     TotImp.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)), 2)
    End If
    TotIva.Text = FormatNumber(CDbl(TotIva.Text) + Diff, 2)
    TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) + Diff, 2)
   End If
  Else
   ElencoVoci.Text = TxtModifica.Text
  End If
 End If
 TotNetto.Text = TotDoc.Text
 If TxtRitenuta.Text <> "" Then CalcolaRitenutaAcconto
End If
End Sub
Private Sub TxtModifica_KeyPress(KeyAscii As Integer)
Dim NumCar As Integer
If KeyAscii >= 32 Then
 NumCar = Len(TxtModifica.Text) + 1
 If ElencoVoci.Col > 1 Then
  If NumCar > 10 Then
   KeyAscii = 0: Exit Sub
  End If
  Dim CarAmmessi$: CarAmmessi = "0123456789,"
  If InStr(CarAmmessi, Chr(KeyAscii)) = 0 Then
   KeyAscii = 0: Exit Sub
  End If
  If ElencoVoci.Col = 2 Or ElencoVoci.Col = 3 Then
   If Not ControlloCarIns(TxtModifica, Chr(KeyAscii), 7, 3) Then
    KeyAscii = 0: Exit Sub
   End If
  End If
 End If
 If RigaCorr = ElencoVoci.Rows - 1 Then
  Dim ColCorr&: ColCorr = ElencoVoci.Col: ElencoVoci.Col = 6: ElencoVoci.CellPictureAlignment = 4
  Set ElencoVoci.CellPicture = ImgCancella: ElencoVoci.Col = ColCorr: ElencoVoci.AddItem ""
  ElencoVoci.RowHeight(ElencoVoci.Rows - 1) = 315
 End If
 If ElencoVoci.Col = 0 Then
  If TxtModifica.SelStart = 0 Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  If NumCar = 3 Then
   TxtModifica.Text = TxtModifica.Text & Chr(KeyAscii): KeyAscii = 0
   TxtModifica.SelStart = Len(TxtModifica.Text)
   Dim rsArticoli As New ADODB.Recordset
   rsArticoli.Open "SELECT * FROM Articoli WHERE descr LIKE '" & TxtModifica.Text & "%' ORDER BY descr ASC", _
   conn, adOpenDynamic
   If Not rsArticoli.EOF Then
    Set ElencoArticoli.Articoli = rsArticoli
    Set ElencoArticoli.FormChiamante = Me
    ElencoArticoli.Show vbModal
   End If
  End If
 End If
ElseIf KeyAscii = vbKeyReturn Then
 If ElencoVoci.Col < 5 Then
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
If MouseInGriglia(ElencoVoci, cm) Then
 If ElencoVoci.Row <> RigaCorr Then
  If Not ConvalidaRigaDoc() Then Exit Sub
 End If
 RigaCorr = ElencoVoci.Row
 With ElencoVoci
  If .Col < 5 And .Col <> 4 Then
   TxtModifica.Visible = True: Set TxtModifica.Container = .Container
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
   If .Col = 2 Or .Col = 3 Then
    TxtModifica.Text = Replace(.Text, ".", "")
   Else
    TxtModifica.Text = .Text
   End If
   TxtModifica.SetFocus
  ElseIf .Col = 6 And .Rows > 2 And .Row < .Rows - 1 Then
   Dim Scelta%
   Scelta = MsgBox("Rimuovere la voce dall' elenco ?", vbExclamation + vbYesNo, "Fattura Pro")
   If Scelta = vbYes Then
    Dim VoceIva#, VoceTot#
    If .TextMatrix(.Row, 5) <> "" Then
     VoceTot = CDbl(.TextMatrix(.Row, 5))
     If .TextMatrix(.Row, 4) <> "" Then
      VoceIva = Round(VoceTot * (CDbl(.TextMatrix(.Row, 4)) / 100), 2)
      TotImp.Text = FormatNumber(CDbl(TotImp.Text) - VoceTot)
      TotIva.Text = FormatNumber(CDbl(TotIva.Text) - VoceIva, 2)
      TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text), 2)
     Else
      TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) - VoceTot, 2)
     End If
     TotNetto.Text = TotDoc.Text
     If TxtRitenuta.Text <> "" Then CalcolaRitenutaAcconto
    End If
    .RemoveItem .Row: VociMod = False
   End If
  End If
 End With
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
Private Function ConvalidaRecord() As Boolean
Dim d As Variant, e As Boolean, id As Variant
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
     Data.Text = d(0) & "-" & d(1) & "-" & d(2)
    Else: e = True
    End If
   Else: e = True
   End If
   If e Then
    MsgBox "Attenzione, inserire un data valida !", vbExclamation, "Fattura Pro"
    Data.SetFocus: Data.SelStart = 0: Data.SelLength = Len(Data.Text): Exit Function
   End If
  End If
 Next i
 If TxtDitta.Tag = "" Or (Not VerificaDestDoc()) Then
  MsgBox "Attenzione, il campo ditta non contiene un valore valido !", vbExclamation, _
  "Fattura Pro"
  Exit Function
 End If
 Dim CercaDuplicato As Boolean
 If StatoDoc = inserimento Then
  CercaDuplicato = True
 ElseIf TxtNumFattura <> rsFatture("NumDoc") Then
  CercaDuplicato = True
 End If
 If CercaDuplicato Then
  Dim rsDuplicato As New ADODB.Recordset
  rsDuplicato.Open "SELECT * FROM FattureFornitori WHERE NumDoc = '" & TxtNumFattura & "' AND " _
  & "IdDitta = " & TxtDitta.Tag & " AND Year(Data) = " & Mid(Data.Text, 7, 4), _
  conn, adOpenDynamic
  If Not rsDuplicato.EOF Then
   MsgBox "Attenzione, il numero documento corrisponde a quello di un altro documento" _
   & " in archivio !" & vbNewLine, vbExclamation, "Fattura Pro"
   TxtNumFattura.SetFocus: TxNumFattura.SelStart = 0
   TxtNumFattura.SelLength = Len(TxtNumFattura): Exit Function
  End If
 End If
 If ConvalidaRigaDoc() Then
  If Not ControlloIva(MsgErroreIva) Then
   MsgBox MsgErroreIva, vbExclamation, "Fattura Pro"
   ConvalidaRecord = False: Exit Function
  End If
  If StatoDoc = inserimento Then
   rsFatture.AddNew
  End If
  rsFatture("NumDoc") = TxtNumFattura.Text
  rsFatture("Data") = Data.Text
  rsFatture("IdDitta") = TxtDitta.Tag
  rsFatture("Pagato") = Pagato.Value
  rsFatture("Modpag") = CmbPag.ListIndex
  rsFatture("TotDoc") = CDbl(TotDoc.Text)
  rsFatture("TotImp") = CDbl(TotImp.Text)
  rsFatture("TotIva") = CDbl(TotIva.Text)
  rsFatture("TotNetto") = CDbl(TotNetto.Text)
  rsFatture.Update
  SalvaRigheDoc
  If ChkCalcRit.Value And TxtRitenuta.Text <> "" Then
   SalvaDatiRitenuta
  ElseIf ChkCalcRit.Value Then
   MsgBox "Inserire i dati della ritenuta d'acconto", vbExclamation, "Fattura Pro"
   Exit Function
  End If
  StatoDoc = NonModificato
  EseguiBackup = True
  conn.CommitTrans
 Else
  Exit Function
 End If
End If
ConvalidaRecord = True
End Function
Private Function SalvaRigheDoc()
With rsVociFattura
If Not .EOF Then
 .MoveFirst
 While Not .EOF
  .Delete
  .Update
  .MoveNext
 Wend
End If
For RigaDoc = 1 To ElencoVoci.Rows - 2
 With rsVociFattura
 .AddNew
 .Fields("IdFatt") = rsFatture("IdDoc")
 .Fields("descr") = EliminaSpazi(ElencoVoci.TextMatrix(RigaDoc, 0))
 .Fields("um") = ElencoVoci.TextMatrix(RigaDoc, 1)
 If ElencoVoci.TextMatrix(RigaDoc, 2) <> "" Then
  .Fields("qnt") = CDbl(ElencoVoci.TextMatrix(RigaDoc, 2))
 Else
  .Fields("qnt") = 0
 End If
 .Fields("prezzo") = CDbl(ElencoVoci.TextMatrix(RigaDoc, 3))
 .Fields("iva") = ElencoVoci.TextMatrix(RigaDoc, 4)
 .Fields("totale") = CDbl(ElencoVoci.TextMatrix(RigaDoc, 5))
 .Update
 End With
Next
End With
End Function
Public Function VerificaDestDoc() As Boolean
VerificaDestDoc = False
Dim rsDitte As New ADODB.Recordset
rsDitte.Open "SELECT * FROM Fornitori WHERE UCASE(Ditta) = '" & UCase(Replace(TxtDitta.Text, "'", "''")) _
& "'", conn, adOpenDynamic
If Not rsDitte.EOF Then
 TxtDitta.Tag = rsDitte("Id")
 TxtPartIva.Text = rsDitte("PartitaIva")
 VerificaDestDoc = True
End If
End Function
Private Sub SalvaDatiRitenuta()
rsRitenute.Find "IdFatt = " & rsFatture("IdDoc"), adBookmarkFirst
If rsRitenute.EOF Then
 rsRitenute.AddNew
End If
rsRitenute("IdFatt") = rsFatture("IdDoc")
rsRitenute("AliqRit") = TxtAliqRit.Text
rsRitenute("ImpRit") = CDbl(TxtRitenuta.Text)
If TxtAliqCassaPro.Text <> "" Then
 rsRitenute("AliqCassaPro") = CDbl(TxtAliqCassaPro.Text)
 rsRitenute("ImpCassaPro") = CDbl(TxtCassaPro.Text)
End If
If TxtSpeseDoc.Text <> "" Then
 rsRitenute("SpeseDoc") = CDbl(TxtSpeseDoc.Text)
End If
rsRitenute("AliqIVA") = TxtAliqIva.Text
rsRitenute("RitCassa") = ChkRitCassa.Value
rsRitenute.Update
End Sub
Private Sub PosizioneRecord(PosRecord As String)
Dim valido: valido = ConvalidaRecord()
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
CaricamentoDocumento = True
With rsFatture
 TxtNumFattura = .Fields("NumDoc")
 Data = Format$(.Fields("data"), "dd-mm-yyyy")
 rsFornitori.Find "Id = " & .Fields("IdDitta"), , adSearchForward, adBookmarkFirst
 TxtDitta.Text = rsFornitori("ditta")
 TxtDitta.Tag = .Fields("IdDitta")
 TxtPartIva.Text = rsFornitori("partitaiva")
 ElencoVoci.Rows = 1
 CmbPag.ListIndex = .Fields("modpag")
 Pagato.Value = .Fields("pagato")
 TotDoc = FormatNumber(.Fields("totdoc"), 2)
 TotImp = FormatNumber(.Fields("totimp"), 2)
 TotIva = FormatNumber(.Fields("totiva"), 2)
 TotNetto = FormatNumber(.Fields("totnetto"), 2)
 rsRitenute.Find "IdFatt = " & .Fields("IdDoc"), , adSearchForward, adBookmarkFirst
 If Not rsRitenute.EOF Then
  ChkCalcRit.Value = 1
  TxtAliqRit.Text = rsRitenute("AliqRit")
  TxtRitenuta.Text = FormatNumber(rsRitenute("ImpRit"), 2)
  If Not IsNull(rsRitenute("AliqCassaPro")) Then
   TxtAliqCassaPro.Text = rsRitenute("AliqCassaPro")
   TxtCassaPro.Text = FormatNumber(rsRitenute("ImpCassaPro"), 2)
  End If
  If Not IsNull(rsRitenute("SpeseDoc")) Then
   TxtSpeseDoc.Text = FormatNumber(rsRitenute("SpeseDoc"), 2)
  End If
  TxtAliqIva.Text = rsRitenute("AliqIVA")
  ChkRitCassa.Value = rsRitenute("RitCassa")
 Else
  ChkCalcRit.Value = 0
 End If
 If rsVociFattura.State = adStateClosed Then
  rsVociFattura.Open "SELECT * FROM VociFattureFornitori WHERE IdFatt = " & .Fields("IdDoc") & " ORDER BY Id ASC", conn, adOpenDynamic, adLockOptimistic
 End If
End With
rsVociFattura.MoveFirst
With ElencoVoci
 Dim CifreDec%, i%
 While Not rsVociFattura.EOF
  i = 0
  .AddItem ""
  .RowHeight(.Rows - 1) = 315
  For Each Field In rsVociFattura.Fields
   If Field.Name <> "Id" And Field.Name <> "IdFatt" Then
    If Field.Type = adDouble Then
     If rsVociFattura(Field.Name) <> 0 Then
      If Field.Name = "Qnt" Then
       CifreDec = 3
      Else
       CifreDec = 2
      End If
      If Field.Value <> 0 Then
       .TextMatrix(.Rows - 1, i) = FormatNumber(rsVociFattura(Field.Name), CifreDec)
      End If
     End If
    Else
     .TextMatrix(.Rows - 1, i) = rsVociFattura(Field.Name)
    End If
    i = i + 1
   End If
  Next
  .Row = .Rows - 1: .Col = 6: .CellPictureAlignment = 4
  Set .CellPicture = ImgCancella
  rsVociFattura.MoveNext
 Wend
 .AddItem "": .RowHeight(ElencoVoci.Rows - 1) = 315
 rsVociFattura.MoveFirst
End With
DatiDoc.Tab = 0
If StatoDoc <> NonModificato Then
 StatoDoc = NonModificato
 conn.CommitTrans
End If
CaricamentoDocumento = False
End Sub
Private Sub CalcolaRitenutaAcconto()
Static InSub As Boolean
If InSub Then Exit Sub
InSub = True
If TotImp.Text <> "0,00" And TxtAliqRit.Text <> "" Then
 Dim ImpDoc#, BaseImp#, CassaPro#, NuovaIvaCP#, IvaCPCorr#, SpeseDoc#
 ImpDoc = CDbl(TotImp.Text)
 If TxtAliqCassaPro.Text <> "" Then
  If TxtAliqIva.Text = "" Then
   InSub = False: Exit Sub
  End If
  If TxtCassaPro.Text <> "" Then
   ImpDoc = ImpDoc - CDbl(TxtCassaPro.Text)
  End If
  CassaPro = Round(ImpDoc * (CDbl(TxtAliqCassaPro.Text) / 100), 2)
  BaseImp = ImpDoc + CassaPro
  TxtCassaPro.Text = FormatNumber(CassaPro, 2)
  TotImp.Text = FormatNumber(BaseImp, 2)
  NuovaIvaCP = Round(CassaPro * CDbl(TxtAliqIva.Text) / 100, 2)
  If TxtAliqIva.Tag <> "" Then IvaCPCorr = CDbl(TxtAliqIva.Tag)
  TotIva.Text = FormatNumber(CDbl(TotIva.Text) + (NuovaIvaCP - IvaCPCorr), 2)
  TxtAliqIva.Tag = NuovaIvaCP
 Else
  If TxtCassaPro.Text <> "" Then
   ImpDoc = ImpDoc - CDbl(TxtCassaPro.Text)
   TotImp.Text = FormatNumber(ImpDoc, 2)
   TotIva.Text = FormatNumber(CDbl(TotIva.Text) - CDbl(TxtAliqIva.Tag), 2)
  End If
  BaseImp = ImpDoc
  TxtCassaPro.Text = ""
  TxtAliqIva.Tag = ""
 End If
 If TxtSpeseDoc.Text <> "" Then
  SpeseDoc = CDbl(TxtSpeseDoc.Text)
 End If
 If ChkRitCassa.Value Then
  TxtRitenuta.Text = FormatNumber(BaseImp * (CDbl(TxtAliqRit.Text) / 100), 2)
 Else
  TxtRitenuta.Text = FormatNumber(ImpDoc * (CDbl(TxtAliqRit.Text) / 100), 2)
 End If
 TotDoc.Text = FormatNumber(BaseImp + CDbl(TotIva.Text) + SpeseDoc, 2)
 TotNetto.Text = FormatNumber(CDbl(TotDoc.Text) - CDbl(TxtRitenuta.Text), 2)
 ModificaDoc
Else
 ChkCalcRit.Value = 0
 ChkCalcRit.Enabled = False
End If
InSub = False
End Sub
Private Function VerificaAliquote() As Boolean
Dim re As New RegExp
VerificaAliquote = True: re.Pattern = "N[0-7I]"
For i = 1 To ElencoVoci.Rows - 2
 If IsNumeric(ElencoVoci.TextMatrix(i, 4)) = False And (Not re.Test(ElencoVoci.TextMatrix(i, 4))) Then
  VerificaAliquote = False: Exit For
 End If
Next i
End Function
Private Function ControlloIva(MsgErrore As String) As Boolean
ControlloIva = True
If Differita = 0 Then
 If Not VerificaAliquote() Then
  MsgErrore = "Attenzione, per una o più voci del documento corrente non è stata impostata " _
  & " l'aliquota iva !"
  ControlloIva = False
 End If
End If
End Function
Public Sub CreaNuovoRecord(Optional ByVal FatturaElettronica As Boolean = False)
TxtNumFattura.Text = ""
TxtDitta.Text = "": TxtDitta.Tag = ""
Data.Text = "": LblNumRecord.Caption = "di " & (rsFatture.RecordCount + 1)
ElencoVoci.Rows = 1
If Not FatturaElettonica Then
 ElencoVoci.AddItem ""
End If
CmbPag.ListIndex = 0: Pagato.Value = 0
RigaCorr = 0: TxtPartIva.Text = ""
TotDoc = "0,00": TotImp = "0,00"
TotIva = "0,00": TotNetto.Text = "0,00"
TxtRecCorr.Text = rsFatture.RecordCount + 1
BtnSucc.Enabled = False: BtnCanc.Enabled = False: BtnNuovo.Enabled = False
If rsFatture.RecordCount >= 1 Then
 BtnPrec.Enabled = True: BtnPrimo.Enabled = True: BtnUltimo.Enabled = True
End If
If rsVociFattura.State <> adStateClosed Then
 rsVociFattura.Close
End If
rsVociFattura.Open "SELECT * FROM VociFattureFornitori WHERE IdFatt = 0", conn, adOpenDynamic, adLockOptimistic
DatiDoc.Tab = 0
StatoDoc = NonModificato
conn.CommitTrans
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
If AvviaTrans Then conn.BeginTrans
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
