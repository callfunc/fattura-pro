VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Scadenzario 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scadenzario"
   ClientHeight    =   8355
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   14250
   Icon            =   "Scadenzario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   14250
   Begin TabDlg.SSTab Scadenze 
      Height          =   10155
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   17912
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
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
      TabCaption(0)   =   "Scadenzario Clienti"
      TabPicture(0)   =   "Scadenzario.frx":4072
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtModifica"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TabFattureClienti"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Scadenzario Fornitori"
      TabPicture(1)   =   "Scadenzario.frx":408E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TabFattureFornitori"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox TabFattureFornitori 
         BackColor       =   &H0033CCFF&
         BorderStyle     =   0  'None
         Height          =   7740
         Left            =   105
         ScaleHeight     =   7740
         ScaleWidth      =   12810
         TabIndex        =   16
         Top             =   435
         Visible         =   0   'False
         Width           =   12810
         Begin VB.TextBox TotInsFor 
            Appearance      =   0  'Flat
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
            Left            =   9315
            TabIndex        =   25
            Top             =   3420
            Width           =   1545
         End
         Begin VB.TextBox TotPagFor 
            Appearance      =   0  'Flat
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
            Left            =   7605
            TabIndex        =   24
            Top             =   3420
            Width           =   1545
         End
         Begin VB.CommandButton AnteprimaFornitori 
            Caption         =   "Anteprima Report"
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
            Left            =   5895
            MaskColor       =   &H00D8E9EC&
            Picture         =   "Scadenzario.frx":40AA
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1455
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin VB.TextBox TotaleFor 
            Appearance      =   0  'Flat
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
            Left            =   5910
            TabIndex        =   22
            Top             =   3420
            Width           =   1515
         End
         Begin VB.OptionButton PagatoFor 
            BackColor       =   &H0033CCFF&
            Caption         =   "Fatture Pagate"
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
            Left            =   9540
            TabIndex        =   21
            Top             =   2700
            Width           =   1740
         End
         Begin VB.OptionButton SospesoFor 
            BackColor       =   &H0033CCFF&
            Caption         =   "Fatture in sospeso"
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
            Left            =   7590
            TabIndex        =   20
            Top             =   2685
            Width           =   1995
         End
         Begin VB.OptionButton TutteFor 
            BackColor       =   &H0033CCFF&
            Caption         =   "Tutte le Fatture"
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
            Left            =   5895
            TabIndex        =   19
            Top             =   2670
            Width           =   1665
         End
         Begin VB.CommandButton StampaFornitori 
            Caption         =   "Stampa Report"
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
            Left            =   5895
            MaskColor       =   &H00D8E9EC&
            Picture         =   "Scadenzario.frx":43B6
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin VB.ListBox ElencoFornitori 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            ItemData        =   "Scadenzario.frx":48F2
            Left            =   165
            List            =   "Scadenzario.frx":48F4
            TabIndex        =   17
            Top             =   300
            Width           =   5490
         End
         Begin MSFlexGridLib.MSFlexGrid ScadenzeFornitori 
            Height          =   3195
            Left            =   165
            TabIndex        =   26
            Top             =   3960
            Width           =   13140
            _ExtentX        =   23178
            _ExtentY        =   5636
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   315
            BackColorFixed  =   14145495
            BackColorSel    =   12937801
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Totale Dovuto:"
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
            Left            =   9285
            TabIndex        =   29
            Top             =   3150
            Width           =   1170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Totale Pagato:"
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
            Left            =   7575
            TabIndex        =   28
            Top             =   3150
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Totale Fatture:"
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
            Left            =   5895
            TabIndex        =   27
            Top             =   3150
            Width           =   1140
         End
      End
      Begin VB.PictureBox TabFattureClienti 
         BackColor       =   &H0033CCFF&
         BorderStyle     =   0  'None
         Height          =   7425
         Left            =   -74775
         ScaleHeight     =   7425
         ScaleWidth      =   13215
         TabIndex        =   2
         Top             =   420
         Width           =   13215
         Begin VB.CommandButton StampaClienti 
            Caption         =   "Stampa Report"
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
            Left            =   5895
            MaskColor       =   &H00D8E9EC&
            Picture         =   "Scadenzario.frx":48F6
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin VB.ListBox ElencoClienti 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            ItemData        =   "Scadenzario.frx":4E32
            Left            =   165
            List            =   "Scadenzario.frx":4E34
            TabIndex        =   10
            Top             =   300
            Width           =   5475
         End
         Begin VB.OptionButton Sospeso 
            BackColor       =   &H0033CCFF&
            Caption         =   "Fatture in sospeso"
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
            Left            =   7650
            TabIndex        =   9
            Top             =   2670
            Width           =   1845
         End
         Begin VB.OptionButton Tutte 
            BackColor       =   &H0033CCFF&
            Caption         =   "Tutte le Fatture"
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
            Left            =   5895
            TabIndex        =   8
            Top             =   2670
            Width           =   1665
         End
         Begin VB.OptionButton Pagato 
            BackColor       =   &H0033CCFF&
            Caption         =   "Fatture Pagate"
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
            Left            =   9660
            TabIndex        =   7
            Top             =   2670
            Width           =   1710
         End
         Begin VB.TextBox Totale 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5895
            TabIndex        =   6
            Top             =   3420
            Width           =   1500
         End
         Begin VB.TextBox TotPag 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7560
            TabIndex        =   5
            Top             =   3420
            Width           =   1500
         End
         Begin VB.TextBox TotIns 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9225
            TabIndex        =   4
            Top             =   3420
            Width           =   1500
         End
         Begin VB.CommandButton AnteprimaClienti 
            Caption         =   "Anteprima Report"
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
            Left            =   5895
            MaskColor       =   &H00D8E9EC&
            Picture         =   "Scadenzario.frx":4E36
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1455
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid ScadenzeClienti 
            Height          =   3150
            Left            =   165
            TabIndex        =   12
            Top             =   3960
            Width           =   13140
            _ExtentX        =   23178
            _ExtentY        =   5556
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   315
            BackColorFixed  =   14145495
            BackColorSel    =   12937801
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Totale Fatture:"
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
            Left            =   5880
            TabIndex        =   15
            Top             =   3150
            Width           =   1140
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Totale Pagato:"
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
            Left            =   7545
            TabIndex        =   14
            Top             =   3150
            Width           =   1140
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Totale Dovuto:"
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
            Left            =   9210
            TabIndex        =   13
            Top             =   3150
            Width           =   1170
         End
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
         Height          =   315
         Left            =   -63495
         TabIndex        =   1
         Top             =   1200
         Visible         =   0   'False
         Width           =   1290
      End
   End
End
Attribute VB_Name = "Scadenzario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClienti As ADODB.Recordset, rsFornitori As ADODB.Recordset, _
rsFattureClienti As ADODB.Recordset, rsFattureFornitori As ADODB.Recordset
Dim ScadenzePag As Variant, GiorniMese As Variant, np As IPictureDisp, ElencoDocumentiReport As Collection, _
RigaCorr As Integer, DatiModificati As Boolean
Dim WithEvents ssc As SmartSubClass
Attribute ssc.VB_VarHelpID = -1
Private Sub AnteprimaClienti_Click()
If ElencoClienti.ListIndex <> -1 And ScadenzeClienti.Rows <> 1 Then
 Call VisualizzaReport(True, "Fatture Clienti")
Else
 MsgBox "Attenzione, l'elenco delle fatture da inserire nel report è vuoto !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub AnteprimaFornitori_Click()
If ElencoFornitori.ListIndex <> -1 And ScadenzeFornitori.Rows <> 1 Then
 Call VisualizzaReport(True, "Fatture Fornitori")
Else
 MsgBox "Attenzione, l'elenco delle fatture da inserire nel report è vuoto !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub TotPag_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotPagFor_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotPagFor_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtModifica_Change()
Dim TabCorr As MSFlexGrid: Set TabCorr = IIf(Scadenze.Tab = 0, ScadenzeClienti, ScadenzeFornitori)
If TxtModifica.Text <> TabCorr.Text Then
 TabCorr.Text = TxtModifica.Text: DatiModificati = True
End If
End Sub
Private Sub TxtModifica_LostFocus()
TxtModifica.Visible = False
End Sub
Private Sub ElencoClienti_Click()
If ElencoClienti.ListIndex <> -1 Then
 If DatiModificati Then SalvaModificheRiga 0
 Dim StrRicerca$: StrRicerca = "IdDitta = " & ElencoClienti.ItemData(ElencoClienti.ListIndex)
 If Sospeso.Value Then StrRicerca = StrRicerca & " AND pagato = False"
 If Pagato.Value Then StrRicerca = StrRicerca & " AND pagato = True"
 rsFattureClienti.Filter = StrRicerca
 MostraTotaliDitta ElencoClienti.Text, "Cliente"
 ScadenzeClienti.Rows = 1
 Call VisualizzaFattureClienti
End If
End Sub
Private Sub MostraTotaliDitta(ByVal Ditta As String, TipoDitta As String)
Dim rsTotali As New ADODB.Recordset, rsTotaliPag As New ADODB.Recordset, rsTotaliIns As New ADODB.Recordset, TabDoc$, TabDitte$
TabDoc = IIf(TipoDitta = "Cliente", "FattureClienti", "FattureFornitori")
TabDitte = IIf(TipoDitta = "Cliente", "Clienti", "Fornitori")
rsTotali.Open "SELECT IdDitta, Sum(TotDoc) As TotFatture FROM " & TabDoc & ", " & TabDitte & " WHERE IdDitta = Id" _
& " And Ditta = '" & Replace(Ditta, "'", "''") & "' GROUP BY IdDitta", conn, adOpenDynamic
rsTotaliIns.Open "SELECT IdDitta, Sum(TotDoc) As TotFatture FROM " & TabDoc & ", " & TabDitte & " WHERE IdDitta = Id" _
& " And Ditta = '" & Replace(Ditta, "'", "''") & "' AND Pagato = False GROUP BY IdDitta", conn, adOpenDynamic
rsTotaliPag.Open "SELECT IdDitta, Sum(TotDoc) As TotFatture FROM " & TabDoc & ", " & TabDitte & " WHERE IdDitta = Id" _
& " And Ditta = '" & Replace(Ditta, "'", "''") & "' AND Pagato = True GROUP BY IdDitta", conn, adOpenDynamic
If TipoDitta = "Cliente" Then
 If Not rsTotali.EOF Then
  rsTotali.MoveFirst
  Totale.Text = FormatNumber(rsTotali("TotFatture"), 2)
 Else
  Totale.Text = "0,00"
 End If
 If Not rsTotaliIns.EOF Then
  rsTotaliIns.MoveFirst
  TotIns.Text = FormatNumber(rsTotaliIns("TotFatture"), 2)
 Else
  TotIns.Text = "0,00"
 End If
 If Not rsTotaliPag.EOF Then
  TotPag.Text = FormatNumber(rsTotaliPag("TotFatture"), 2)
 Else
  TotPag.Text = "0,00"
 End If
Else
 If Not rsTotali.EOF Then
  rsTotali.MoveFirst
  TotaleFor.Text = FormatNumber(rsTotali("TotFatture"), 2)
 Else
  TotaleFor.Text = "0,00"
 End If
 If Not rsTotaliIns.EOF Then
  TotInsFor.Text = FormatNumber(rsTotaliIns("TotFatture"), 2)
 Else
  TotInsFor.Text = "0,00"
 End If
 If Not rsTotaliPag.EOF Then
  TotPagFor.Text = FormatNumber(rsTotaliPag("TotFatture"), 2)
 Else
  TotPagFor.Text = "0,00"
 End If
End If
End Sub
Private Sub ElencoFornitori_Click()
If ElencoFornitori.ListIndex <> -1 Then
 If DatiModificati Then SalvaModificheRiga 0
 Dim StrRicerca$: StrRicerca = "IdDitta = " & ElencoFornitori.ItemData(ElencoFornitori.ListIndex)
 If Sospeso.Value Then StrRicerca = StrRicerca & " AND pagato = False"
 If Pagato.Value Then StrRicerca = StrRicerca & " AND pagato = True"
 rsFattureFornitori.Filter = StrRicerca
 MostraTotaliDitta ElencoFornitori.Text, "Fornitore"
 ScadenzeFornitori.Rows = 1
 Call VisualizzaFattureFornitori
End If
End Sub
Private Sub Form_Load()
Me.Move 450, 450
IntestazioniGriglia = Array("Numero", "Data", "Importo", "Scadenza Pagamento", "Modalità di Pagamento", _
"Pagato")
ScadenzeClienti.ColWidth(0) = 1100: ScadenzeClienti.ColWidth(1) = 1400
ScadenzeClienti.ColWidth(2) = 1400: ScadenzeClienti.ColWidth(3) = 2400
ScadenzeClienti.ColWidth(4) = 5000: ScadenzeClienti.ColWidth(5) = 1000

For i = 0 To ScadenzeClienti.Cols - 1
 ScadenzeClienti.TextMatrix(0, i) = IntestazioniGriglia(i)
 ScadenzeClienti.ColAlignment(i) = 4
Next i

IntestazioniGriglia = Array("Numero", "Data", "Importo", "Scadenza Pagamento", "Modalità di Pagamento", "Pagato")

ScadenzeFornitori.ColWidth(0) = 1800: ScadenzeFornitori.ColWidth(1) = 1400
ScadenzeFornitori.ColWidth(2) = 1400: ScadenzeFornitori.ColWidth(3) = 2400
ScadenzeFornitori.ColWidth(4) = 5000: ScadenzeClienti.ColWidth(5) = 1000

For i = 0 To ScadenzeFornitori.Cols - 1
 ScadenzeFornitori.TextMatrix(0, i) = IntestazioniGriglia(i)
 ScadenzeFornitori.ColAlignment(i) = 4
Next i

Printer.Font.Name = "Arial": Printer.Font.Size = 9: Printer.Font.Bold = False
GiorniMese = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
ScadenzePag = Array(0, 30, 60, 90)

Set np = LoadPicture()

Set rsClienti = New ADODB.Recordset
rsClienti.Open "SELECT * FROM Clienti WHERE Rimosso = False ORDER BY Ditta ASC", conn, adOpenDynamic, adLockOptimistic
If rsClienti.EOF Then
 MsgBox "L'Archivio Clienti è vuoto !", vbExclamation, "Fattura Pro"
End If

Set rsFornitori = New ADODB.Recordset
rsFornitori.Open "SELECT * FROM Fornitori WHERE Rimosso = False ORDER BY Ditta ASC", conn, adOpenDynamic, adLockOptimistic
If rsFornitori.EOF Then
 MsgBox "L'Archivio Fornitori è vuoto !", vbExclamation, "Fattura Pro"
End If

Set rsFattureClienti = New ADODB.Recordset
rsFattureClienti.Open "FattureClienti", conn, adOpenDynamic, adLockOptimistic
If rsFattureClienti.EOF Then
 MsgBox "L'Archivio Fatture Clienti è vuoto !", vbExclamation, "Fattura Pro"
End If

Set rsFattureFornitori = New ADODB.Recordset
rsFattureFornitori.Open "FattureFornitori", conn, adOpenDynamic, adLockOptimistic
If rsFattureFornitori.EOF Then
 MsgBox "L'Archivio Fatture Fornitori è vuoto !", vbExclamation, "Fattura Pro"
End If

Dim LarghezzaStr&, LargMax&, Ditta$: Set Me.Font = ElencoClienti.Font

If Not rsClienti.EOF Then
 While Not rsClienti.EOF
  ElencoClienti.AddItem rsClienti("Ditta")
  ElencoClienti.ItemData(ElencoClienti.NewIndex) = rsClienti("Id")
  LarghezzaStr = Me.TextWidth(rsClienti("Ditta"))
  If LarghezzaStr > LargMax Then LargMax = LarghezzaStr
  rsClienti.MoveNext
 Wend
End If

LargMax = IIf(LargMax <= (ElencoClienti.Width - LargSB), 0, LargMax + Me.ScaleX(10, _
vbPixels, vbTwips))
LargMax = Me.ScaleX(LargMax, vbTwips, vbPixels)
SendMessage ElencoClienti.hWnd, LB_SETHORIZONTALEXTENT, LargMax, ByVal 0&: LargMax = 0

If Not rsFornitori.EOF Then
 While Not rsFornitori.EOF
  ElencoFornitori.AddItem rsFornitori("Ditta")
  ElencoFornitori.ItemData(ElencoFornitori.NewIndex) = rsFornitori("Id")
  LarghezzaStr = Me.TextWidth(rsFornitori("Ditta"))
  If LarghezzaStr > LargMax Then LargMax = LarghezzaStr
  rsFornitori.MoveNext
 Wend
End If

LargMax = IIf(LargMax <= (ElencoFornitori.Width - Me.ScaleX(LargSB, vbPixels, vbTwips)), 0, _
LargMax + Me.ScaleX(10, vbPixels, vbTwips))
LargMax = Me.ScaleX(LargMax, vbTwips, vbPixels)
SendMessage ElencoFornitori.hWnd, LB_SETHORIZONTALEXTENT, LargMax, ByVal 0&
Set ssc = New SmartSubClass: ssc.SubClassHwnd TxtModifica.hWnd, True
ssc.SubClassHwnd ScadenzeClienti.hWnd, True: ssc.SubClassHwnd ScadenzeFornitori.hWnd, True
Scadenze.Tab = 0: Tutte.Value = True
End Sub
Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And FatturaProMDI.WindowState <> 1 Then
 If Me.Width < 13980 Then
  Me.Width = 13980
 ElseIf Me.Height < 8895 Then
  Me.Height = 8895
 Else
  Scadenze.Move 105, 150, Me.ScaleWidth - 210, Me.ScaleHeight - 255
  TabFattureFornitori.Move 15, 315, Scadenze.Width - 45, Scadenze.Height - 330
  TabFattureClienti.Move 15, 315, Scadenze.Width - 45, Scadenze.Height - 330
  ScadenzeClienti.Height = TabFattureClienti.Height - 105 - ScadenzeClienti.Top
  ScadenzeFornitori.Height = TabFattureFornitori.Height - 105 - ScadenzeFornitori.Top
 End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If DatiModificati Then SalvaModificheRiga Scadenze.Tab
If Not ssc Is Nothing Then
 ssc.SubClassHwnd TxtModifica.hWnd, False: ssc.SubClassHwnd ScadenzeClienti.hWnd, False
 ssc.SubClassHwnd ScadenzeFornitori.hWnd, False
End If
Set Scadenzario = Nothing
End Sub
Private Sub Pagato_Click()
If ElencoClienti.ListIndex <> -1 Then
 rsFattureClienti.Filter = "IdDitta = " & ElencoClienti.ItemData(ElencoClienti.ListIndex) & " AND pagato = True"
 ScadenzeClienti.Rows = 1
 If Not rsFattureClienti.EOF Then
  Call VisualizzaFattureClienti
 End If
End If
End Sub
Private Sub PagatoFor_Click()
If ElencoFornitori.ListIndex <> -1 Then
 rsFattureFornitori.Filter = "IdDitta = " & ElencoFornitori.ItemData(ElencoFornitori.ListIndex) & " AND pagato = True"
 ScadenzeFornitori.Rows = 1
 If Not rsFattureFornitori.EOF Then
  Call VisualizzaFattureFornitori
 End If
End If
End Sub
Private Sub ScadenzeClienti_Click()
If ScadenzeClienti.Rows > 1 Then
 If ScadenzeClienti.Row <> RigaCorr And DatiModificati Then SalvaModificheRiga Scadenze.Tab
 RigaCorr = ScadenzeClienti.Row
 rsFattureClienti.Move RigaCorr - 1, adBookmarkFirst
 If ScadenzeClienti.Col = 5 Then
  If Not rsFattureClienti("Pagato") Then
   ScadenzeClienti.CellPictureAlignment = 4: Set ScadenzeClienti.CellPicture = ImgConferma
   rsFattureClienti("Pagato") = True
   If Tutte.Value Then
    TotPag = FormatNumber(CDbl(TotPag) + CDbl(ScadenzeClienti.TextMatrix(ScadenzeClienti.Row, 2)), 2)
    TotIns = FormatNumber(CDbl(Totale) - CDbl(TotPag), 2)
   End If
  Else
   ScadenzeClienti.CellPictureAlignment = 4: Set ScadenzeClienti.CellPicture = np
   rsFattureClienti("Pagato") = False
   If Tutte.Value Then
    TotPag = FormatNumber(CDbl(TotPag) - CDbl(ScadenzeClienti.TextMatrix(ScadenzeClienti.Row, 2)), 2)
    TotIns = FormatNumber(CDbl(Totale) - CDbl(TotPag), 2)
   End If
  End If
  EseguiBackup = True
  rsFattureClienti.Update
 ElseIf ScadenzeClienti.Col = 4 Then
  RigaCorr = ScadenzeClienti.Row
  With ScadenzeClienti
   TxtModifica.Visible = True: Set TxtModifica.Container = .Container
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
   TxtModifica.Text = .Text
   On Error Resume Next
   TxtModifica.SetFocus: DatiModificati = False
  End With
 End If
End If
End Sub
Private Sub ScadenzeClienti_DblClick()
Call ScadenzeClienti_Click
End Sub
Private Sub ScadenzeClienti_Scroll()
TxtModifica.Visible = False
End Sub
Private Sub SalvaModificheRiga(SchedaCorr As Integer)
If SchedaCorr = 0 Then
 rsFattureClienti.Move RigaCorr - 1, adBookmarkFirst
 rsFattureClienti("notepag") = ScadenzeClienti.TextMatrix(RigaCorr, 4)
 rsFattureClienti.Update
Else
 rsFattureFornitori.Move RigaCorr - 1, adBookmarkFirst
 rsFattureFornitori("notepag") = ScadenzeFornitori.TextMatrix(RigaCorr, 4)
 rsFattureFornitori.Update
End If
DatiModificati = False
End Sub
Private Sub ScadenzeFornitori_Click()
If ScadenzeFornitori.Rows > 1 Then
 If ScadenzeFornitori.Row <> RigaCorr And DatiModificati Then SalvaModificheRiga 1
 RigaCorr = ScadenzeFornitori.Row
 rsFattureFornitori.Move RigaCorr - 1, adBookmarkFirst
 If ScadenzeFornitori.Col = 5 Then
  If Not rsFattureFornitori("pagato") Then
   ScadenzeFornitori.CellPictureAlignment = 4: Set ScadenzeFornitori.CellPicture = ImgConferma
   rsFattureFornitori("pagato") = True
   If TutteFor.Value Then
    TotPagFor = FormatNumber(CDbl(TotPagFor) + CDbl(ScadenzeFornitori.TextMatrix(ScadenzeFornitori.Row, 2)), 2)
    TotInsFor = FormatNumber(CDbl(TotaleFor) - CDbl(TotPagFor), 2)
   End If
  Else
   ScadenzeFornitori.CellPictureAlignment = 4: Set ScadenzeFornitori.CellPicture = np
   rsFattureFornitori("pagato") = False
   If TutteFor.Value Then
    TotPagFor = FormatNumber(CDbl(TotPagFor) - CDbl(ScadenzeFornitori.TextMatrix(ScadenzeFornitori.Row, 2)), 2)
    TotInsFor = FormatNumber(CDbl(TotaleFor) - CDbl(TotPagFor), 2)
   End If
  End If
  rsFattureFornitori.Update
  EseguiBackup = True
 ElseIf ScadenzeFornitori.Col = 4 Then
  RigaCorr = ScadenzeFornitori.Row
  With ScadenzeFornitori
   TxtModifica.Visible = True: Set TxtModifica.Container = .Container
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
   TxtModifica.Text = .Text
   On Error Resume Next
   TxtModifica.SetFocus: DatiModificati = False
  End With
 End If
End If
End Sub
Private Sub ScadenzeFornitori_DblClick()
Call ScadenzeFornitori_Click
End Sub
Private Sub TotPag_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Sospeso_Click()
If ElencoClienti.ListIndex <> -1 Then
 If DatiModificati Then SalvaModificheRiga 0
 rsFattureClienti.Filter = "IdDitta = " & ElencoClienti.ItemData(ElencoClienti.ListIndex) & " AND pagato = False"
 ScadenzeClienti.Rows = 1: Call VisualizzaFattureClienti
End If
End Sub
Private Sub SospesoFor_Click()
If ElencoFornitori.ListIndex <> -1 Then
 If DatiModificati Then SalvaModificheRiga 1
 rsFattureFornitori.Filter = "IdDitta = " & ElencoFornitori.ItemData(ElencoFornitori.ListIndex) & " AND pagato = False"
 ScadenzeFornitori.Rows = 1: Call VisualizzaFattureFornitori
End If
End Sub
Private Sub ssc_NewMessage(ByVal hWnd As Long, uMsg As Long, wParam As Long, lParam As Long, _
Cancel As Boolean)
If uMsg = WM_MOUSEWHEEL Then
 Dim Griglia As MSFlexGrid
 Set Griglia = IIf(Scadenze.Tab = 0, ScadenzeClienti, ScadenzeFornitori)
 If TxtModifica.hWnd = hWnd And ScrollBarVisibile(Griglia.hWnd) Then
  TxtModifica.Visible = False: Griglia.SetFocus
 ElseIf ScrollBarVisibile(Griglia.hWnd) Then
  ScrollGriglia Griglia, wParam / 65536
 End If
End If
End Sub
Private Sub Scadenze_Click(PreviousTab As Integer)
If DatiModificati Then SalvaModificheRiga PreviousTab
If Scadenze.Tab = 0 Then
 Tutte.Value = True: TabFattureClienti.Visible = True
 TabFattureFornitori.Visible = False
Else
 TutteFor.Value = True: TabFattureClienti.Visible = False
 TabFattureFornitori.Visible = True
End If
End Sub
Private Sub StampaClienti_Click()
If ElencoClienti.ListIndex <> -1 And ScadenzeClienti.Rows <> 1 Then
 Call VisualizzaReport(False, "Fatture Clienti")
Else
 MsgBox "Attenzione, l'elenco delle fatture da inserire nel report è vuoto !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub StampaFornitori_Click()
If ElencoFornitori.ListIndex <> -1 And ScadenzeFornitori.Rows <> 1 Then
 Call VisualizzaReport(False, "Fatture Fornitori")
Else
 MsgBox "Attenzione, l'elenco delle fatture da inserire nel report è vuoto !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub Totale_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub Totale_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TotaleFor_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotaleFor_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TotIns_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotIns_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TotInsFor_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotInsFor_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Tutte_Click()
If ElencoClienti.ListIndex <> -1 Then
 If DatiModificati Then SalvaModificheRiga 0
 rsFattureClienti.Filter = "IdDitta = " & ElencoClienti.ItemData(ElencoClienti.ListIndex)
 ScadenzeClienti.Rows = 1: Call VisualizzaFattureClienti
End If
End Sub
Private Sub VisualizzaFattureClienti()
Dim GiorniScadenza%, TotFatt#, TotPag#, NumDoc$, DataDoc$
While Not rsFattureClienti.EOF
 With ScadenzeClienti
  NumDoc = NumeroDocumento(rsFattureClienti("iddoc"))
  GiorniScadenza = ScadenzePag(rsFattureClienti("modpag"))
  DataDoc = rsFattureClienti("data")
  .AddItem ""
  .TextMatrix(.Rows - 1, 0) = NumDoc
  .TextMatrix(.Rows - 1, 1) = Format(DataDoc, "dd-mm-yyyy")
  .TextMatrix(.Rows - 1, 2) = FormatNumber(rsFattureClienti("totdoc"), 2)
  .TextMatrix(.Rows - 1, 3) = Format(DateAdd("d", GiorniScadenza, DataDoc), "dd-mm-yyyy")
  .TextMatrix(.Rows - 1, 4) = IIf(IsNull(rsFattureClienti("notepag")), "", _
  rsFattureClienti("notepag"))
  If rsFattureClienti("pagato") Then
   .Row = .Rows - 1: .Col = 5
   .CellPictureAlignment = 4: Set .CellPicture = ImgConferma
  End If
  .RowHeight(.Rows - 1) = 315
 End With
 rsFattureClienti.MoveNext
Wend
If Not rsFattureClienti.EOF Then
 rsFattureClienti.MoveFirst
End If
End Sub
Private Sub TutteFor_Click()
If ElencoFornitori.ListIndex <> -1 Then
 If DatiModificati Then SalvaModificheRiga 1
 rsFattureFornitori.Filter = "IdDitta = " & ElencoFornitori.ItemData(ElencoFornitori.ListIndex)
 ScadenzeFornitori.Rows = 1: Call VisualizzaFattureFornitori
End If
End Sub
Private Sub VisualizzaFattureFornitori()
Dim NumGiorniScadenza%, DataDoc$
While Not rsFattureFornitori.EOF
 With ScadenzeFornitori
  GiorniScadenza = ScadenzePag(rsFattureFornitori("modpag"))
  DataDoc = rsFattureFornitori("data")
  .AddItem ""
  .TextMatrix(.Rows - 1, 0) = rsFattureFornitori("iddoc")
  .TextMatrix(.Rows - 1, 1) = Format(rsFattureFornitori("data"), "dd-mm-yyyy")
  .TextMatrix(.Rows - 1, 2) = FormatNumber(rsFattureFornitori("totdoc"), 2)
  .TextMatrix(.Rows - 1, 3) = Format(DateAdd("d", GiorniScadenza, DataDoc), "dd-mm-yyyy")
  .TextMatrix(.Rows - 1, 4) = IIf(IsNull(rsFattureFornitori("notepag")), "", rsFattureFornitori("notepag"))
  If rsFattureFornitori("pagato") Then
   .Row = .Rows - 1: .Col = 5
   .CellPictureAlignment = 4: Set .CellPicture = ImgConferma
  End If
  End With
 rsFattureFornitori.MoveNext
Wend
If Not rsFattureFornitori.EOF Then
 rsFattureFornitori.MoveFirst
End If
End Sub
Private Sub CaricaInfoDitta(SS As ServiziStampa)
Dim rsInfoDitta As New ADODB.Recordset, InfoDitta As New ADODB.Record
rsInfoDitta.Open "SELECT * FROM InfoDitta", conn, adOpenDynamic, adLockOptimistic
rsInfoDitta.MoveFirst
SS.InfoDitta = rsInfoDitta.GetRows(1, 1)
rsInfoDitta.Close
End Sub
Private Sub CreaReport(SS As ServiziStampa, ByVal Tipo As String)
Dim ArrDatiDoc As Variant, ElencoDocReport As New Collection
Dim rs As New ReportScadenzario
Call CaricaInfoDitta(SS)
If Tipo = "Fatture Clienti" Then
 rsClienti.Find "Id = " & ElencoClienti.ItemData(ElencoClienti.ListIndex), , adSearchForward, adBookmarkFirst
 rsFattureClienti.MoveFirst
 For i = 1 To ScadenzeClienti.Rows - 1
  ArrDatiDoc = Array(ScadenzeClienti.TextMatrix(i, 0), ScadenzeClienti.TextMatrix(i, 1), _
  ScadenzeClienti.TextMatrix(i, 2), ScadenzeClienti.TextMatrix(i, 3), _
  ScadenzeClienti.TextMatrix(i, 4), IIf(rsFattureClienti("pagato"), "Pagato", "Non Pagato"))
  ElencoDocReport.Add ArrDatiDoc
  rsFattureClienti.MoveNext
 Next i
 SS.TipoDoc = ReportFattureClienti: Set rs.Ditta = rsClienti
 rs.TotalePagato = TotPag.Text: rs.TotaleInsoluto = TotIns.Text
 Set rs.ElencoFatture = ElencoDocReport
Else
 rsFornitori.Find "Id = " & ElencoFornitori.ItemData(ElencoFornitori.ListIndex), , adSearchForward, adBookmarkFirst
 rsFattureFornitori.MoveFirst
 For i = 1 To ScadenzeFornitori.Rows - 1
  ArrDatiDoc = Array(ScadenzeFornitori.TextMatrix(i, 0), ScadenzeFornitori.TextMatrix(i, 1), _
  ScadenzeFornitori.TextMatrix(i, 2), ScadenzeFornitori.TextMatrix(i, 3), ScadenzeFornitori.TextMatrix(i, 4), _
  IIf(rsFattureFornitori("pagato"), "Pagato", "Non Pagato"))
  ElencoDocReport.Add ArrDatiDoc
  rsFattureFornitori.MoveNext
 Next i
 rsFornitori.Find "Id = " & ElencoFornitori.ItemData(ElencoFornitori.ListIndex)
 SS.TipoDoc = ReportFattureFornitori: Set rs.Ditta = rsFornitori
 rs.TotalePagato = TotPagFor.Text: rs.TotaleInsoluto = TotInsFor.Text
 Set rs.ElencoFatture = ElencoDocReport
 rsFattureFornitori.MoveFirst
End If
Set SS.ReportScadenze = rs
End Sub
Private Sub VisualizzaReport(Anteprima As Boolean, ByVal Tipo As String)
Dim SS As New ServiziStampa
If Anteprima Then
 Call CreaReport(SS, Tipo): SS.ImpostaAnteprima True
 SS.Stampa -1: AnteprimaDoc.Show vbModal
Else
 OpzioniStampa.Inizializza SS: OpzioniStampa.Show vbModal
 If OpzioniStampa.Scelta = "Stampa" Then
  Call CreaReport(SS, Tipo): SS.ImpostaAnteprima False
  SS.Stampa -1
 End If
End If
End Sub
