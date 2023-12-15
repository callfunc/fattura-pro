VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Bilancio 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incassi e Spese"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   16620
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Bilancio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   16620
   Begin TabDlg.SSTab Schedario 
      Height          =   8265
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   14579
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabMaxWidth     =   3440
      ShowFocusRect   =   0   'False
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
      TabCaption(0)   =   "Bilancio IVA"
      TabPicture(0)   =   "Bilancio.frx":4072
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "PannelloBilancio(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Incassi e Spese"
      TabPicture(1)   =   "Bilancio.frx":408E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "PannelloBilancio(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox PannelloBilancio 
         BackColor       =   &H0033CCFF&
         BorderStyle     =   0  'None
         Height          =   8325
         Index           =   1
         Left            =   105
         ScaleHeight     =   8325
         ScaleWidth      =   15780
         TabIndex        =   11
         Top             =   390
         Visible         =   0   'False
         Width           =   15780
         Begin VB.CommandButton BtnIncassiSpese 
            Caption         =   "Visualizza Ditta"
            Height          =   480
            Left            =   7860
            TabIndex        =   24
            Top             =   1500
            Width           =   1695
         End
         Begin VB.ComboBox ElencoAnniIncassiSpese 
            Height          =   345
            ItemData        =   "Bilancio.frx":40AA
            Left            =   6495
            List            =   "Bilancio.frx":40AC
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   480
            Width           =   1410
         End
         Begin VB.OptionButton TipoDitta 
            BackColor       =   &H0033CCFF&
            Caption         =   "Mostra Fornitori"
            Height          =   285
            Index           =   1
            Left            =   7500
            TabIndex        =   16
            Top             =   1050
            Width           =   1740
         End
         Begin VB.OptionButton TipoDitta 
            BackColor       =   &H0033CCFF&
            Caption         =   "Mostra Clienti"
            Height          =   285
            Index           =   0
            Left            =   5940
            TabIndex        =   15
            Top             =   1050
            Value           =   -1  'True
            Width           =   1530
         End
         Begin VB.ListBox ElencoCF 
            Height          =   2985
            ItemData        =   "Bilancio.frx":40AE
            Left            =   180
            List            =   "Bilancio.frx":40B0
            TabIndex        =   14
            Top             =   495
            Width           =   5595
         End
         Begin VB.CommandButton BtnIncassiSpeseTutti 
            Caption         =   "Visualizza Tutti"
            Height          =   480
            Left            =   5955
            TabIndex        =   13
            Top             =   1500
            Width           =   1635
         End
         Begin VB.TextBox TotaleCF 
            Height          =   315
            Left            =   5970
            TabIndex        =   12
            Top             =   3150
            Width           =   2070
         End
         Begin MSFlexGridLib.MSFlexGrid TotIncassiSpese 
            Height          =   1170
            Left            =   165
            TabIndex        =   18
            Top             =   6585
            Width           =   14895
            _ExtentX        =   26273
            _ExtentY        =   2064
            _Version        =   393216
            Cols            =   13
            FixedCols       =   0
            BackColorFixed  =   14145495
            BackColorBkg    =   24576
            FocusRect       =   0
            HighLight       =   0
            AllowUserResizing=   3
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
         Begin MSFlexGridLib.MSFlexGrid GrigliaIncassiSpese 
            Height          =   2805
            Left            =   165
            TabIndex        =   19
            Top             =   3630
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   4948
            _Version        =   393216
            Rows            =   1
            Cols            =   13
            FixedCols       =   0
            BackColorFixed  =   14145495
            BackColorSel    =   12937801
            ForeColorSel    =   16777215
            BackColorBkg    =   24576
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
         Begin VB.Label LblAnno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anno:"
            Height          =   225
            Left            =   5955
            TabIndex        =   23
            Top             =   525
            Width           =   480
         End
         Begin VB.Label LblElencoDitte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Elenco Clienti:"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   165
            TabIndex        =   22
            Top             =   210
            Width           =   1125
         End
         Begin VB.Label Stato 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stato"
            Height          =   225
            Left            =   5955
            TabIndex        =   21
            Top             =   2325
            Width           =   405
         End
         Begin VB.Label LblTotCF 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Totale Cliente:"
            Height          =   225
            Left            =   5955
            TabIndex        =   20
            Top             =   2850
            Width           =   1140
         End
      End
      Begin VB.PictureBox PannelloBilancio 
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
         Height          =   7695
         Index           =   0
         Left            =   -74955
         ScaleHeight     =   7695
         ScaleWidth      =   14055
         TabIndex        =   1
         Top             =   420
         Width           =   14055
         Begin VB.CommandButton BtnStampaBilancio 
            Caption         =   "Stampa Bilancio"
            Height          =   1005
            Left            =   2220
            MaskColor       =   &H00D8E9EC&
            Picture         =   "Bilancio.frx":40B2
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   405
            UseMaskColor    =   -1  'True
            Width           =   1710
         End
         Begin VB.TextBox TxtTotBilancio 
            Height          =   315
            Left            =   1905
            TabIndex        =   10
            Top             =   7020
            Width           =   1770
         End
         Begin VB.TextBox TxtBilancioIVA 
            Height          =   315
            Left            =   4995
            TabIndex        =   8
            Top             =   6585
            Width           =   1800
         End
         Begin VB.TextBox TxtBilancioImp 
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   6585
            Width           =   1755
         End
         Begin MSFlexGridLib.MSFlexGrid TotaliBilancio 
            Height          =   4170
            Left            =   270
            TabIndex        =   4
            Top             =   2250
            Width           =   13335
            _ExtentX        =   23521
            _ExtentY        =   7355
            _Version        =   393216
            Rows            =   13
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   315
            BackColorFixed  =   14145495
            BackColorBkg    =   24576
            FocusRect       =   0
            HighLight       =   0
            MergeCells      =   1
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
         Begin VB.CommandButton BtnCalcolaBilancio 
            Caption         =   "Calcola Totali"
            Height          =   435
            Left            =   270
            TabIndex        =   3
            Top             =   1635
            Width           =   1650
         End
         Begin VB.ListBox ElencoAnniBilancio 
            Height          =   1185
            ItemData        =   "Bilancio.frx":45EE
            Left            =   855
            List            =   "Bilancio.frx":45F0
            TabIndex        =   2
            Top             =   225
            Width           =   990
         End
         Begin VB.Label LblAnnoBilancio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anno:"
            Height          =   225
            Left            =   300
            TabIndex        =   25
            Top             =   195
            Width           =   480
         End
         Begin VB.Label LblTotBilancio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Totale Bilancio:"
            Height          =   225
            Left            =   270
            TabIndex        =   9
            Top             =   7065
            Width           =   1215
         End
         Begin VB.Label LblBilancioIva 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bilancio IVA:"
            Height          =   225
            Left            =   3960
            TabIndex        =   7
            Top             =   6615
            Width           =   990
         End
         Begin VB.Label LblTotBilancioImp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bilancio Imponibile:"
            Height          =   225
            Left            =   285
            TabIndex        =   5
            Top             =   6615
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Bilancio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PosMouse As CoordinateMouse, SchedaCorr As Integer, ArrMesi As Variant, DatiCaricati As Boolean
Dim WithEvents ssc As SmartSubClass
Attribute ssc.VB_VarHelpID = -1
Private Sub BtnCalcolaBilancio_Click()
Dim Anno$
Anno = Year(Now)
If ElencoAnniBilancio.Text <> "" Then
 Anno = ElencoAnniBilancio.Text
End If
Dim TotImpAnno#, TotIvaAnno#
Dim rsTotaliVendite As New ADODB.Recordset, rsTotaliAcquisti As New ADODB.Recordset, _
rsTotaliRimborsi As New ADODB.Recordset
rsTotaliVendite.Open "SELECT Format(Data, ""mmmm"") As Mese, Sum(TotImp) As TotImpVen, " _
& "Sum(TotIva) As TotIvaVen FROM FattureClienti WHERE Year(Data) = " & Anno & _
" GROUP BY Format(Data, ""mmmm""), Month(Data) ORDER BY Month(Data) ASC", conn, adOpenDynamic
rsTotaliAcquisti.Open "SELECT Format(Data, ""mmmm"") As Mese, Sum(TotImp) As TotImpAcq, " _
& "Sum(TotIva) As TotIvaAcq FROM FattureFornitori WHERE Year(Data) = " & Anno & " GROUP " _
& "BY Format(Data, ""mmmm""), Month(Data) ORDER BY Month(Data) ASC", conn, adOpenDynamic
rsTotaliRimborsi.Open "SELECT Format(Data, ""mmmm"") As Mese, Sum(TotImp) As TotImpRim, " _
& "Sum(TotIva) As TotIvaRim FROM NoteCredito WHERE Year(Data) = " & Anno & " GROUP " _
& "BY Format(Data, ""mmmm""), Month(Data) ORDER BY Month(Data) ASC ", conn, adOpenDynamic
If Not rsTotaliVendite.EOF Then
 rsTotaliVendite.MoveFirst
End If
If Not rsTotaliAcquisti.EOF Then
 rsTotaliAcquisti.MoveFirst
End If
DatiCaricati = rsTotaliVendite.RecordCount <> 0 Or rsTotaliAcquisti.RecordCount <> 0
If DatiCaricati Then
 For i = 1 To TotaliBilancio.Rows - 1
  With TotaliBilancio
   If Not rsTotaliVendite.EOF Then
    If rsTotaliVendite("Mese") = LCase(.TextMatrix(i, 0)) Then
     .TextMatrix(i, 1) = FormatNumber(rsTotaliVendite("TotImpVen"), 2)
     .TextMatrix(i, 3) = FormatNumber(rsTotaliVendite("TotIvaVen"), 2)
     rsTotaliVendite.MoveNext
    Else
     .TextMatrix(i, 1) = "0,00"
     .TextMatrix(i, 3) = "0,00"
    End If
   Else
    .TextMatrix(i, 1) = "0,00"
    .TextMatrix(i, 3) = "0,00"
   End If
   If Not rsTotaliAcquisti.EOF Then
    If rsTotaliAcquisti("Mese") = LCase(.TextMatrix(i, 0)) Then
     .TextMatrix(i, 2) = FormatNumber(rsTotaliAcquisti("TotImpAcq"), 2)
     .TextMatrix(i, 4) = FormatNumber(rsTotaliAcquisti("TotIvaAcq"), 2)
     rsTotaliAcquisti.MoveNext
    Else
     .TextMatrix(i, 2) = "0,00"
     .TextMatrix(i, 4) = "0,00"
    End If
   Else
    .TextMatrix(i, 2) = "0,00"
    .TextMatrix(i, 4) = "0,00"
   End If
   If Not rsTotaliRimborsi.EOF Then
    If rsTotaliRimborsi("Mese") = LCase(.TextMatrix(i, 0)) Then
     .TextMatrix(i, 2) = FormatNumber(CDbl(.TextMatrix(i, 2)) + rsTotaliRimborsi("TotImpRim"), 2)
     .TextMatrix(i, 4) = FormatNumber(CDbl(.TextMatrix(i, 4)) + rsTotaliRimborsi("TotIvaRim"), 2)
     rsTotaliRimborsi.MoveNext
    End If
   End If
   .TextMatrix(i, 5) = FormatNumber(CDbl(.TextMatrix(i, 1)) - CDbl(.TextMatrix(i, 2)), 2)
   .TextMatrix(i, 6) = FormatNumber(CDbl(.TextMatrix(i, 3)) - CDbl(.TextMatrix(i, 4)), 2)
   TotImpAnno = TotImpAnno + CDbl(.TextMatrix(i, 5))
   TotIvaAnno = TotIvaAnno + CDbl(.TextMatrix(i, 6))
  End With
 Next i
End If
rsTotaliVendite.Close
rsTotaliAcquisti.Close
rsTotaliRimborsi.Close
TxtBilancioImp.Text = FormatNumber(TotImpAnno, 2)
TxtBilancioIVA.Text = FormatNumber(TotIvaAnno, 2)
TxtTotBilancio.Text = FormatNumber(TotImpAnno + TotIvaAnno, 2)
End Sub
Private Sub GestisciSelezioneGriglia(Griglia As MSFlexGrid)
If Griglia.Row <> CLng(Griglia.Tag) Then
 Dim NuovaRiga&: NuovaRiga = Griglia.Row
 If CLng(Griglia.Tag) <> 0 Then SelezionaRigaGriglia Griglia, CLng(Griglia.Tag), False
 SelezionaRigaGriglia Griglia, NuovaRiga, True: Griglia.Tag = NuovaRiga
End If
End Sub
Private Sub SelezionaRigaGriglia(Griglia As MSFlexGrid, ByVal IndiceRiga&, ByVal Seleziona As Boolean)
Griglia.Redraw = False
Griglia.Row = IndiceRiga: Griglia.RowSel = IndiceRiga
Griglia.Col = 0: Griglia.ColSel = Griglia.Cols - 1
If Seleziona Then
 Griglia.CellBackColor = RGB(49, 106, 197): Griglia.CellForeColor = vbWhite
Else
 Griglia.CellBackColor = vbWhite: Griglia.CellForeColor = vbBlack
End If
Griglia.Redraw = True
End Sub
Private Sub BtnStampaBilancio_Click()
If DatiCaricati Then
 Dim MesiBilancio As New Collection, DatiBilancioMese As Variant, MeseCorr$, AnnoCorr$
 Dim SS As New ServiziStampa
 SS.ImpostaAnteprima True
 MeseCorr = Month(Now): AnnoCorr = Year(Now)
 For i = 1 To TotaliBilancio.Rows - 1
  If AnnoCorr = ElencoAnniBilancio.Text And MeseCorr < i Then
   Exit For
  End If
  DatiBilancioMese = Array(TotaliBilancio.TextMatrix(i, 1), TotaliBilancio.TextMatrix(i, 2), _
  TotaliBilancio.TextMatrix(i, 3), TotaliBilancio.TextMatrix(i, 4), TotaliBilancio.TextMatrix(i, 5), _
  TotaliBilancio.TextMatrix(i, 6))
  MesiBilancio.Add DatiBilancioMese, TotaliBilancio.TextMatrix(i, 0)
 Next i
 MesiBilancio.Add Array(TxtBilancioImp.Text, TxtBilancioIVA.Text), "Totali"
 Call CaricaIntestazioneDitta(SS)
 Set SS.ReportBilancio = MesiBilancio
 SS.TipoDoc = ReportBilancio
 SS.Stampa: AnteprimaDoc.Show vbModal
Else
 MsgBox "Attenzione, il report bilancio non contiene dati  !", vbExclamation, "Fattura Pro"
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
Private Sub Form_Load()
Me.Move 450, 100

Dim IntestazioniGriglia As Variant

Set ssc = New SmartSubClass

ArrMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", _
"Settembre", "Ottobre", "Novembre", "Dicembre")

IntestazioniGriglia = Array("Mese", "Imponibile Vendite", "Imponibile Acquisti", _
"Iva Vendite", "Iva Acquisti", "Bilancio Imponibile", "Bilancio IVA")

For i = 0 To TotaliBilancio.Cols - 1
 TotaliBilancio.TextMatrix(0, i) = IntestazioniGriglia(i): TotaliBilancio.ColAlignment(i) = 4
Next i

TotaliBilancio.ColWidth(0) = 2500
TotaliBilancio.ColWidth(1) = Me.TextWidth("Imponibile Acquisti") + 150
TotaliBilancio.ColWidth(2) = Me.TextWidth("Imponibile Vendite") + 150
TotaliBilancio.ColWidth(3) = Me.TextWidth("Iva Acquisti") + 225
TotaliBilancio.ColWidth(4) = Me.TextWidth("Iva Vendite") + 225
TotaliBilancio.ColWidth(5) = Me.TextWidth("Bilancio Imponibile") + 150
TotaliBilancio.ColWidth(6) = Me.TextWidth("Bilancio IVA") + 150

For i = 1 To 12
 TotaliBilancio.TextMatrix(i, 0) = ArrMesi(i - 1)
Next i

GrigliaIncassiSpese.TextMatrix(0, 0) = "Cliente": GrigliaIncassiSpese.ColAlignment(0) = 0
GrigliaIncassiSpese.ColWidth(0) = 4000

For i = 1 To GrigliaIncassiSpese.Cols - 1
 GrigliaIncassiSpese.TextMatrix(0, i) = ArrMesi(i - 1): GrigliaIncassiSpese.ColAlignment(i) = 0
 GrigliaIncassiSpese.ColWidth(i) = 970
Next i

TotIncassiSpese.TextMatrix(0, 0) = "Tot. Clienti Anno"
TotIncassiSpese.ColWidth(0) = Me.TextWidth("Tot. Clienti Anno") + 150
TotIncassiSpese.ColAlignment(0) = 0
For i = 1 To TotIncassiSpese.Cols - 1
 TotIncassiSpese.TextMatrix(0, i) = ArrMesi(i - 1): TotIncassiSpese.ColAlignment(i) = 0
 TotIncassiSpese.ColWidth(i) = 970
Next i

TipoDitta_Click 0

Dim rsAnni As New ADODB.Recordset
rsAnni.Open "SELECT Distinct Year(Data) AS Anno From FattureClienti ORDER BY Year(Data)", conn, adOpenDynamic
If Not rsAnni.EOF Then
 rsAnni.MoveFirst
 While Not rsAnni.EOF
  ElencoAnniBilancio.AddItem rsAnni("Anno")
  ElencoAnniIncassiSpese.AddItem rsAnni("Anno")
  rsAnni.MoveNext
 Wend
End If
rsAnni.Close
ssc.SubClassHwnd GrigliaIncassiSpese.hWnd, True
Schedario.Tab = 0
End Sub
Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
 Schedario.Move 105, 105, Me.ScaleWidth - 210, Me.ScaleHeight - 210
 For i = 0 To PannelloBilancio.Count - 1
  PannelloBilancio(i).Move 15, 315, Schedario.Width - 45, Schedario.Height - 330
 Next i
 GrigliaIncassiSpese.Move GrigliaIncassiSpese.Left, GrigliaIncassiSpese.Top, PannelloBilancio(1).Width - _
 GrigliaIncassiSpese.Left - 150
 TotIncassiSpese.Move TotIncassiSpese.Left, TotIncassiSpese.Top, PannelloBilancio(1).Width - _
 TotIncassiSpese.Left - 150
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
ssc.SubClassHwnd GrigliaIncassiSpese.hWnd, False
Set IncassiSpese = Nothing
End Sub
Private Sub GrigliaIncassiSpese_Click()
If GrigliaIncassiSpese.Row < GrigliaIncassiSpese.Rows And GrigliaIncassiSpese.Row _
<> GrigliaIncassiSpese.Tag Then
 Dim TotaleAnno As Double: GrigliaIncassiSpese.Tag = GrigliaIncassiSpese.Row
 For i = 1 To 12
  TotaleAnno = TotaleAnno + CDbl(GrigliaIncassiSpese.TextMatrix(GrigliaIncassiSpese.Row, i))
 Next i
 TotaleCF.Text = FormatNumber(TotaleAnno, 2)
End If
End Sub
Private Sub BtnIncassiSpese_Click()
If ElencoCF.ListIndex <> -1 Then
 VisualizzaIncassiSpese False
End If
End Sub
Private Sub BtnIncassiSpeseTutti_Click()
VisualizzaIncassiSpese True
End Sub
Private Sub VisualizzaIncassiSpese(ByVal MostraTutti As Boolean)
Dim rsIncassiSpese As New ADODB.Recordset, SQL As String, Anno As String
If TipoDitta(0).Value Then
 SQL = "SELECT Clienti.Ditta, Format(Data, ""mmmm"") AS Mese, Sum(TotDoc) AS TotaleMese FROM FattureClienti" _
 & ", Clienti WHERE Clienti.Id = FattureClienti.IdDitta"
 If Not MostraTutti Then
  SQL = SQL & " AND Clienti.Ditta = '" & ElencoCF.Text & "'"
 End If
 Anno = IIf(ElencoAnniIncassiSpese.Text <> "", ElencoAnniIncassiSpese.Text, Str(Year(Date)))
 SQL = SQL & " AND Year(Data) = " & Anno
 SQL = SQL & " GROUP BY Clienti.Ditta, Format(Data, ""mmmm"") ORDER BY Clienti.Ditta ASC, " _
 & "Format(Data, ""mmmm"") ASC"
Else
 SQL = "SELECT Fornitori.Ditta, Format(Data, ""mmmm"") AS Mese, Sum(TotDoc) AS TotaleMese FROM FattureFornitori" _
 & ", Fornitori WHERE Fornitori.Id = FattureFornitori.IdDitta"
 If Not MostraTutti Then
  SQL = SQL & " AND Fornitori.Ditta = '" & ElencoCF.Text & "'"
 End If
 Anno = IIf(ElencoAnniIncassiSpese.Text <> "", ElencoAnniIncassiSpese.Text, Year(Date))
 SQL = SQL & " AND Year(Data) = " & Anno
 SQL = SQL & " GROUP BY Fornitori.Ditta, Format(Data, ""mmmm"") ORDER BY Fornitori.Ditta ASC, " _
 & "Format(Data, ""mmmm"") ASC"
End If
rsIncassiSpese.Open SQL, conn, adOpenDynamic
GrigliaIncassiSpese.Rows = 1: GrigliaIncassiSpese.Tag = 0
Stato.Caption = "Caricamento dati in corso..."
TotaleCF = "": Me.MousePointer = vbHourglass: GrigliaIncassiSpese.Redraw = False
Dim Ditta$
If Not rsIncassiSpese.EOF Then
 rsIncassiSpese.MoveFirst
 Ditta = rsIncassiSpese("Ditta")
 While Not rsIncassiSpese.EOF
  With GrigliaIncassiSpese
   .AddItem ""
   .TextMatrix(.Rows - 1, 0) = Ditta
   For i = 1 To .Cols - 1
    If Not rsIncassiSpese.EOF Then
     If rsIncassiSpese("Ditta") = Ditta Then
      .TextMatrix(.Rows - 1, i) = FormatNumber(rsIncassiSpese("TotaleMese"), 2)
      rsIncassiSpese.MoveNext
     Else
      .TextMatrix(.Rows - 1, i) = "0,00"
     End If
    Else
     .TextMatrix(.Rows - 1, i) = "0,00"
    End If
   Next i
   If Not rsIncassiSpese.EOF Then
    Ditta = rsIncassiSpese("Ditta")
   End If
  End With
 Wend
 Call MostraTotaliIncassiSpese
Else
 MsgBox "Le ditte selezionate non hanno documenti associati !", vbExclamation, "Fattura Pro"
End If
Me.MousePointer = vbDefault
GrigliaIncassiSpese.Redraw = True
Stato.Caption = "Caricamento dati completato."
End Sub
Private Sub TipoDitta_Click(Index As Integer)
Dim rsDitte As New ADODB.Recordset, Tabella$, LarghezzaStr As Long, LargMax As Long
GrigliaIncassiSpese.Rows = 1: ElencoCF.Clear: TotaleCF.Text = ""
If Index = 0 Then
 LblElencoDitte.Caption = "Elenco Clienti": LblTotCF.Caption = "Totale Cliente:"
 TotIncassiSpese.TextMatrix(0, 0) = "Tot. Clienti Anno"
 TotIncassiSpese.ColWidth(0) = Me.TextWidth("Tot. Clienti Anno") + 150
 SQLQuery = "SELECT * FROM Clienti WHERE Rimosso = False ORDER BY Ditta ASC"
Else
 LblElencoDitte.Caption = "Elenco Fornitori": LblTotCF.Caption = "Totale Fornitore:"
 TotIncassiSpese.TextMatrix(0, 0) = "Tot. Fornitori Anno"
 TotIncassiSpese.ColWidth(0) = Me.TextWidth("Tot. Fornitori Anno") + 150
 SQLQuery = "SELECT * FROM Fornitori WHERE Rimosso = False ORDER BY Ditta ASC"
End If
rsDitte.Open SQLQuery, conn, adOpenDynamic
If Not rsDitte.EOF Then
 While Not rsDitte.EOF
  LarghezzaStr = Me.TextWidth(rsDitte("Ditta"))
  If LarghezzaStr > LargMax Then LargMax = LarghezzaStr
  ElencoCF.AddItem rsDitte("Ditta")
  ElencoCF.ItemData(ElencoCF.NewIndex) = rsDitte("Id")
  rsDitte.MoveNext
 Wend
End If
LargMax = IIf(LargMax <= ElencoCF.Width, 0, LargMax + Me.ScaleX(12, vbPixels, vbTwips))
LargMax = Me.ScaleX(LargMax, vbTwips, vbPixels)
Call SendMessage(ElencoCF.hWnd, LB_SETHORIZONTALEXTENT, LargMax, ByVal 0&)
End Sub
Private Sub Schedario_Click(PreviousTab As Integer)
PannelloBilancio(Schedario.Tab).Visible = True
For i = 0 To PannelloBilancio.Count - 1
 If i <> Schedario.Tab Then PannelloBilancio(i).Visible = False
Next i
End Sub
Private Sub ssc_NewMessage(ByVal hWnd As Long, uMsg As Long, wParam As Long, lParam As Long, _
Cancel As Boolean)
If uMsg = WM_MOUSEWHEEL Then
 Dim Griglia As MSFlexGrid
 Select Case Schedario.Tab
 Case 0:
  Set Griglia = TotIncassiSpese
 Case 1:
  Set Griglia = GrigliaIncassiSpese
 End Select
 If ScrollBarVisibile(Griglia.hWnd) Then
  ScrollGriglia Griglia, wParam / 65536
 End If
End If
End Sub
Private Sub MostraTotaliIncassiSpese()
Dim TotAnno#, TotMese#
For i = 1 To 12
 For j = 1 To GrigliaIncassiSpese.Rows - 1
  TotMese = TotMese + CDbl(GrigliaIncassiSpese.TextMatrix(j, i))
 Next j
 TotAnno = TotAnno + TotMese
 TotIncassiSpese.TextMatrix(1, i) = FormatNumber(TotMese, 2): TotMese = 0
Next i
TotIncassiSpese.TextMatrix(1, 0) = FormatNumber(TotAnno, 2)
End Sub
