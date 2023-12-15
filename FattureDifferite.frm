VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CreaFatturaDiff 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fattura Differita"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9360
   Icon            =   "FattureDifferite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9360
   Begin MSFlexGridLib.MSFlexGrid GrigliaLuoghi 
      Height          =   1695
      Left            =   225
      TabIndex        =   10
      Top             =   3840
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   14145495
      BackColorSel    =   16711680
      ForeColorSel    =   16777215
      BackColorBkg    =   24576
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
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
   Begin VB.CommandButton BtnNoteAgg 
      Caption         =   "Note Aggiuntive"
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
      Height          =   390
      Left            =   7095
      TabIndex        =   5
      Top             =   1815
      Width           =   1935
   End
   Begin VB.Frame RiquadroPeriodo 
      BackColor       =   &H0033CCFF&
      Caption         =   "Periodo di Riferimento"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   225
      TabIndex        =   2
      Top             =   5730
      Width           =   5550
      Begin MSComCtl2.DTPicker DataFinale 
         Height          =   330
         Left            =   3870
         TabIndex        =   9
         Top             =   405
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   24576
         CalendarTitleBackColor=   24576
         CalendarTitleForeColor=   3394815
         CalendarTrailingForeColor=   10526880
         CustomFormat    =   "dd-MM-yyy"
         Format          =   152043523
         CurrentDate     =   41480
      End
      Begin MSComCtl2.DTPicker DataIniziale 
         Height          =   330
         Left            =   1185
         TabIndex        =   8
         Top             =   405
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   24576
         CalendarTitleBackColor=   24576
         CalendarTitleForeColor=   3394815
         CalendarTrailingForeColor=   10526880
         CustomFormat    =   "dd-MM-yyy"
         Format          =   152043523
         CurrentDate     =   41480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data iniziale:"
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
         Left            =   150
         TabIndex        =   4
         Top             =   435
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data finale:"
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
         Left            =   2925
         TabIndex        =   3
         Top             =   435
         Width           =   885
      End
   End
   Begin VB.CommandButton BtnDettagli 
      Caption         =   "Dettagli Fattura"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   7095
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FattureDifferite.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   1935
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
      Height          =   3210
      ItemData        =   "FattureDifferite.frx":4136
      Left            =   210
      List            =   "FattureDifferite.frx":4138
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   6540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elenco Ditte"
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
      TabIndex        =   7
      Top             =   165
      Width           =   945
   End
   Begin MSForms.Label Label3 
      Height          =   60
      Left            =   345
      TabIndex        =   6
      Top             =   315
      Width           =   1545
      BackColor       =   3394815
      Size            =   "2725;106"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "CreaFatturaDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClienti As ADODB.Recordset, rsFattureClienti As ADODB.Recordset, rsNoteCliente As ADODB.Recordset, _
rsLuoghiConsegna As ADODB.Recordset, rsVociNote As ADODB.Recordset, NoteCaricate As Boolean, IndCorr As Integer
Public ElencoNoteAgg As Collection
Dim TotaliFattura As Collection
Dim TotFatt#, TotImp#, TotIva#, GiorniMese As Variant
Private Sub DataFinale_Change()
ElencoClienti.ListIndex = -1
IndCorr = -1
End Sub
Private Sub DataIniziale_Change()
ElencoClienti.ListIndex = -1
IndCorr = -1
Dim MeseCorr%, AnnoCorr%
MeseCorr = Month(DataIniziale.Value)
AnnoCorr = Year(DataIniziale.Value)
DataFinale.Value = DateValue(GiorniMese(MeseCorr - 1) & "/" & MeseCorr & "/" & AnnoCorr)
End Sub
Private Sub BtnDettagli_Click()
If NoteCaricate Then
 CreaFatturaTemp
 Set FatturaDiff.rsFatture = rsFattureClienti
 Set FatturaDiff.rsNoteFattura = rsNoteCliente
 Set FatturaDiff.rsVociNote = rsVociNote
 Set FatturaDiff.rsCliente = rsClienti
 FatturaDiff.Show
 Unload Me
End If
End Sub
Private Sub CaricaIntestazioneDitta(SS As ServiziStampa)
Dim rsInfoDitta As New ADODB.Recordset
rsInfoDitta.Open "SELECT * FROM InfoDitta", conn, adOpenDynamic, adLockOptimistic
If Not rsInfoDitta.EOF Then
 SS.InfoDitta = rsInfoDitta.GetRows(1, 1)
End If
rsInfoDitta.Close
End Sub
Private Sub ElencoClienti_Click()
If ElencoClienti.ListIndex <> -1 And ElencoClienti.ListIndex <> IndCorr Then
 rsClienti.Move ElencoClienti.ListIndex, adBookmarkFirst
 Dim DataDa$, DataA$
 DataDa = Format$(DataIniziale.Value, "yyyy/mm/dd")
 DataA = Format$(DataFinale.Value, "yyyy/mm/dd")
 If rsNoteCliente.State <> adStateClosed Then
  rsNoteCliente.Close
 End If
 If rsVociNote.State <> adStateClosed Then
  rsVociNote.Close
 End If
 rsNoteCliente.Open "SELECT * FROM NoteConsegna WHERE IdDitta = " & rsClienti("Id") & _
 " AND Data BETWEEN #" & DataDa & "# AND #" & DataA & "# ORDER BY IdDoc ASC", conn, _
 adOpenDynamic, adLockOptimistic
 rsVociNote.Open "SELECT * FROM NoteConsegna, VociNoteConsegna WHERE IdNota = IdDoc AND " _
 & "IdDitta = " & rsClienti("Id") & " AND Data BETWEEN #" & DataDa & "# AND #" & DataA & "# " _
 & "ORDER BY IdDoc ASC", conn, adOpenDynamic
 
 IndCorr = ElencoClienti.ListIndex: GrigliaLuoghi.Visible = False
 NoteCaricate = False: BtnNoteAgg.Enabled = False

 If rsNoteCliente.EOF Then
  Me.Height = 5910: RiquadroPeriodo.Top = 3840
  MsgBox "Il cliente selezionato non ha note di consegna associate per il periodo indicato !", _
  vbExclamation, "Fattura Pro"
 Else
  Set ElencoNoteAgg = New Collection: Set SS = New ServiziStampa
  Set rsLuoghiConsegna = New ADODB.Recordset
  rsLuoghiConsegna.Open "SELECT * FROM LuoghiConsegna WHERE IdDitta = " & rsClienti("Id") _
  & " ORDER BY Indirizzo ASC", conn, adOpenDynamic
  If Not rsLuoghiConsegna.EOF Then
   GrigliaLuoghi.Visible = True: GrigliaLuoghi.Enabled = True
   GrigliaLuoghi.HighLight = flexHighlightWithFocus
   Me.Height = 7785: RiquadroPeriodo.Top = 6150
   GrigliaLuoghi.Rows = 1: CaricaLuoghiConsegna
  Else
   Me.Height = 5910: RiquadroPeriodo.Top = 3840
   GrigliaLuoghi.Visible = False
   BtnNoteAgg.Enabled = True: NoteCaricate = True
  End If
 End If
End If
End Sub
Private Sub Form_Load()
Me.Move 1000, 300, Me.Width, Me.Height
Dim Ditta$: IndCorr = -1
Set rsClienti = New ADODB.Recordset
Set rsFattureClienti = New ADODB.Recordset
Set rsNoteCliente = New ADODB.Recordset
Set rsVociNote = New ADODB.Recordset
rsFattureClienti.Open "FattureClienti", conn, adOpenDynamic, adLockOptimistic
rsClienti.Open "SELECT * FROM Clienti WHERE Rimosso = False ORDER BY Ditta ASC", conn, adOpenDynamic, adLockOptimistic
If Not rsClienti.EOF Then
 While Not rsClienti.EOF
  Ditta = rsClienti("Ditta")
  LarghezzaStr = Me.TextWidth(Ditta)
  If LarghezzaStr > LargMax Then LargMax = LarghezzaStr
  ElencoClienti.AddItem Ditta
  ElencoClienti.ItemData(ElencoClienti.NewIndex) = rsClienti("Id")
  rsClienti.MoveNext
 Wend
Else
 MsgBox "L'Archivio Clienti è vuoto !", vbExclamation, "Fattura Pro"
End If

Dim IntestazioniGriglia As Variant
IntestazioniGriglia = Array("Descrizione", "Indirizzo")
GrigliaLuoghi.ColWidth(0) = 3000
GrigliaLuoghi.ColWidth(1) = 3750

For i = 0 To UBound(IntestazioniGriglia)
 GrigliaLuoghi.TextMatrix(0, i) = IntestazioniGriglia(i)
 GrigliaLuoghi.ColAlignment(i) = 4
Next i

TotImp = 0: TotIva = 0: TotFatt = 0
GrigliaLuoghi.Visible = False
RiquadroPeriodo.Top = 3840
Me.Height = 5910

LargMax = IIf(LargMax <= (ElencoClienti.Width - Me.ScaleX(LargSB, vbPixels, vbTwips)), 0, LargMax + _
Me.ScaleX(10, vbPixels, vbTwips))
LargMax = Me.ScaleX(LargMax, vbTwips, vbPixels)
If LargMax > ElencoClienti.Width Then
 SendMessage ElencoClienti.hWnd, LB_SETHORIZONTALEXTENT, LargMax, ByVal 0&
End If

GiorniMese = Array("31", "28", "31", "30", "31", "30", "31", "31", "30", "31", "30", "31")
Dim MeseCorr%, AnnoCorr%: MeseCorr = Month(Date): AnnoCorr = Year(Date)
DataIniziale.Value = DateValue("1/" & MeseCorr & "/" & AnnoCorr)
If CInt(AnnoCorr) Mod 400 = 0 Or (AnnoCorr Mod 4 = 0 And (AnnoCorr Mod 100 <> 0)) Then
 GiorniMese(MeseCorr - 1) = 29
End If
DataFinale.Value = DateValue(GiorniMese(MeseCorr - 1) & "/" & MeseCorr & "/" & AnnoCorr)
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set FattureDifferite = Nothing
End Sub
Private Sub CaricaLuoghiConsegna()
With GrigliaLuoghi
 While Not rsLuoghiConsegna.EOF
  .AddItem ""
  .TextMatrix(.Rows - 1, 0) = rsLuoghiConsegna("descr")
  .TextMatrix(.Rows - 1, 1) = rsLuoghiConsegna("indirizzo")
 Wend
End With
End Sub
Private Sub GrigliaLuoghi_Click()
If GrigliaLuoghi.Rows > 1 Then
 CreaElencoNoteLuogo
 If Not rsNoteCliente.EOF Then
  BtnNoteAgg.Enabled = True: NoteCaricate = True
 Else
  MsgBox "Il cliente selezionato non ha alcuna voce associata ad esso per le date " _
  & "e il luogo di consegna indicati !", vbExclamation, "Fattura Pro"
  NoteCaricate = False
 End If
End If
End Sub
Private Sub CreaFatturaTemp()
Dim PartIVA$, IdCliente$, CodFisc$
rsFattureClienti.AddNew
Call CaricaTotaliIva
Call InserisciNoteAgg
rsFattureClienti("IdDitta") = rsClienti("Id")
rsFattureClienti("TotDoc") = TotFatt
rsFattureClienti("TotImp") = TotImp
rsFattureClienti("TotIva") = TotIva
rsFattureClienti("TipoDoc") = 1
rsFattureClienti("Pagato") = False
End Sub
Private Sub CaricaTotaliIva()
Dim rsVociNota As ADODB.Recordset, Totale As Double, Iva As Double, Aliquota As Double
Dim TI As TotaleIva
Set TotaliFattura = New Collection
Set rsVociNota = New ADODB.Recordset
rsVociNota.Open "VociNoteConsegna", conn, adOpenDynamic
While Not rsNoteCliente.EOF
 rsVociNota.Filter = "IdNota = " & rsNoteCliente("IdDoc")
 rsVociNota.MoveFirst
 On Error Resume Next
 While Not rsVociNota.EOF
  Totale = CDbl(rsVociNota("totale"))
  Aliquota = CDbl(rsVociNota("iva"))
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
  TotImp = TotImp + Totale
  TotIva = TotIva + Iva
  TotFatt = TotImp + TotIva
  rsVociNota.MoveNext
 Wend
 rsNoteCliente.MoveNext
Wend
rsNoteCliente.MoveFirst
For Each TotaleIva In TotaliFattura
 If InStr(1, CStr(TotaleIva.Totale), ",99") Then
  TotaleIva.Totale = TotaleIva.Totale + 0.01
 End If
Next
End Sub
Private Sub BtnNoteAgg_Click()
Set NoteAggiuntive.NoteCliente = rsNoteCliente
NoteAggiuntive.IdDitta = ElencoClienti.ItemData(ElencoClienti.ListIndex)
NoteAggiuntive.Show
End Sub
Private Sub InserisciNoteAgg()
If ElencoNoteAgg.Count <> 0 Then
 rsNoteCliente.Close
 QuerySQL = "SELECT * FROM NoteConsegna WHERE IdDitta = " & rsClienti("Id") & " AND Data BETWEEN #" & _
 Format$(DataIniziale.Value, "yyyy/mm/dd") & "# AND #" & Format$(DataFinale.Value, "yyyy/mm/dd") & "# OR IdDoc IN ("
 For i = 1 To ElencoNoteAgg.Count
  QuerySQL = QuerySQL & "'" & ElencoNoteAgg(i) & "'"
  If i <> ElencoNoteAgg.Count Then QuerySQL = QuerySQL & ","
 Next
 QuerySQL = QuerySQL & ") ORDER BY IdDoc ASC"
 rsNoteCliente.Open QuerySQL, conn, adOpenDynamic, adLockOptimistic
 rsNoteCliente.MoveFirst
End If
End Sub
Private Sub CreaElencoNoteLuogo()
rsNoteCliente.Close
Dim DataDa$, DataA$
DataDa = Format$(DataIniziale.Value, "yyyy/mm/dd")
DataA = Format$(DataFinale.Value, "yyyy/mm/dd")
QuerySQL = "SELECT * FROM NoteConsegna, LuoghiConsegna WHERE IdDitta = " & rsClienti("Id") & _
" AND Data BETWEEN #" & DataDa & "# AND #" & DataA & "# AND IdLC = LuoghiConsegna.Id AND " _
& "LuoghiConsegna.Nome = '" & GrigliaLuoghi.TextMatrix(GrigliaLuoghi.Row, 0) & _
"' AND IdNota = NoteConsegna.IdDoc ORDER BY IdDoc ASC"
rsNoteCliente.Open QuerySQL, conn, adOpenDynamic, adLockOptimistic
rsNoteCliente.MoveFirst
rsVociNote.Open "SELECT VociNoteConsegna.* FROM NoteConsegna, VociNoteConsegna WHERE IdNota = IdDoc AND " _
& "IdDitta = " & rsClienti("Id") & " AND Data BETWEEN #" & DataDa & "# AND #" & DataA & "# " _
& "ORDER BY IdDoc ASC", conn, adOpenDynamic, adLockOptimistic
rsVociNote.MoveFirst
End Sub
