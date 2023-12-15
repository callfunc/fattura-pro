VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Magazzino 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestione Magazzino"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14310
   Icon            =   "Magazzino.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   14310
   Begin VB.Frame RiquadroRicerca 
      Appearance      =   0  'Flat
      BackColor       =   &H0033CCFF&
      Caption         =   "Ricerca Movimento"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   3930
      TabIndex        =   3
      Top             =   165
      Width           =   8775
      Begin VB.ComboBox CmbTipoMov 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Magazzino.frx":4072
         Left            =   6390
         List            =   "Magazzino.frx":407F
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   495
         Width           =   1560
      End
      Begin VB.TextBox TxtArticolo 
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
         Left            =   195
         TabIndex        =   7
         Top             =   510
         Width           =   3015
      End
      Begin VB.CommandButton BtnRicerca 
         Caption         =   "Ricerca"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3480
         TabIndex        =   6
         Top             =   1020
         Width           =   1710
      End
      Begin MSComCtl2.DTPicker DataInizio 
         Height          =   360
         Left            =   3360
         TabIndex        =   4
         Top             =   495
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
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
         Format          =   114753537
         CurrentDate     =   43152
      End
      Begin MSComCtl2.DTPicker DataFine 
         Height          =   360
         Left            =   4875
         TabIndex        =   9
         Top             =   495
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   635
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
         Format          =   114753537
         CurrentDate     =   43152
      End
      Begin VB.Label LblTipoMov 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Movimento"
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
         Left            =   6375
         TabIndex        =   11
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label LblDataFine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data A:"
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
         Left            =   4875
         TabIndex        =   10
         Top             =   225
         Width           =   570
      End
      Begin VB.Label LblProdotto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Articolo:"
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
         TabIndex        =   8
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LblDataInizio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Da:"
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
         Left            =   3360
         TabIndex        =   5
         Top             =   225
         Width           =   660
      End
   End
   Begin MSFlexGridLib.MSFlexGrid ElencoMovimenti 
      Height          =   4740
      Left            =   165
      TabIndex        =   2
      Top             =   1965
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   8361
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   14145495
      BackColorSel    =   12937777
      ForeColorSel    =   16777215
      GridColorFixed  =   8421504
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
   Begin VB.CommandButton BtnScarico 
      Caption         =   "Scarica Prodotto"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   1965
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Magazzino.frx":409F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1665
   End
   Begin VB.CommandButton BtnCarico 
      Caption         =   "Carica Prodotto"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   195
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Magazzino.frx":4148
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1515
   End
End
Attribute VB_Name = "Magazzino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMovimenti As ADODB.Recordset, RigaCorr%
Private Sub BtnCarico_Click()
Dim MovMag As New MovimentoMagazzino
MovMag.Tipo = "Carico"
MovMag.NuovoMovimento
MovMag.Show vbModal
If MovMag.RecordModificato Then
 EseguiBackup = True
 CaricaMovimenti "SELECT MovimentiMagazzino.*, Descr FROM MovimentiMagazzino, Articoli WHERE" _
 & " Id = IdArticolo ORDER BY Data ASC"
End If
End Sub
Private Sub BtnRicerca_Click()
If BtnRicerca.Caption = "Ricerca" Then
 Dim ValoriCampi As Variant, EsprRicerca As Variant, StrRicerca$
 ValoriCampi = Array(TxtArticolo.Text, Format$(DataInizio.Value, "yyyy/mm/dd"), Format$(DataFine.Value, _
 "yyyy/mm/dd"), CmbTipoMov.Text)
 EsprRicerca = Array("Descr LIKE 'val%'", "Data>=#val#", "Data<=#val#", "TipoMov='val'")
 For i = 0 To UBound(ValoriCampi)
  If Trim(ValoriCampi(i)) <> "" Then
   StrRicerca = IIf(StrRicerca <> "", StrRicerca & " AND ", " WHERE ")
   EsprRicerca(i) = Replace(EsprRicerca(i), "val", ValoriCampi(i))
   StrRicerca = StrRicerca & EsprRicerca(i)
  End If
 Next
 If StrRicerca <> "" Then
  CaricaMovimenti "SELECT MovimentiMagazzino.*, Descr FROM MovimentiMagazzino, Articoli" & StrRicerca _
  & " AND Id = IdArticolo ORDER BY Data ASC", True
 Else
  MsgBox "Attenzione, tutti i campi di ricerca sono vuoti !", vbExclamation, "Fattura Pro"
 End If
Else
 BtnRicerca.Caption = "Ricerca"
 CaricaMovimenti "SELECT MovimentiMagazzino.*, Descr FROM MovimentiMagazzino, Articoli WHERE" _
 & " Id = IdArticolo ORDER BY Data ASC"
End If
End Sub
Private Sub BtnScarico_Click()
Dim MovMag As New MovimentoMagazzino
MovMag.Tipo = "Scarico"
MovMag.NuovoMovimento
MovMag.Show vbModal
If MovMag.RecordModificato Then
 EseguiBackup = True
 CaricaMovimenti "SELECT MovimentiMagazzino.*, Descr FROM MovimentiMagazzino, Articoli WHERE" _
 & " Id = IdArticolo ORDER BY Data ASC"
End If
End Sub
Private Sub ElencoMovimenti_Click()
If ElencoMovimenti.Rows <> 1 Then
 Dim ColCorr%
 ColCorr = ElencoMovimenti.Col
 If ElencoMovimenti.Row <> RigaCorr Then
  Dim NuovaRiga%
  NuovaRiga = ElencoMovimenti.Row
  ElencoMovimenti.Redraw = False
  If RigaCorr <> 0 Then
   ElencoMovimenti.Row = RigaCorr
   For i = 0 To ElencoMovimenti.Cols - 2
    ElencoMovimenti.Col = i: ElencoMovimenti.CellBackColor = RGB(255, 255, 255)
    ElencoMovimenti.CellForeColor = RGB(0, 0, 0)
   Next i
   Set ElencoMovimenti.CellPicture = CaricaImmagineDaRisorsa("CANCELLA", , , RGB(255, 255, 255))
  End If
  ElencoMovimenti.Row = NuovaRiga
  For i = 0 To ElencoMovimenti.Cols - 2
   ElencoMovimenti.Col = i: ElencoMovimenti.CellBackColor = RGB(49, 106, 197)
   ElencoMovimenti.CellForeColor = RGB(255, 255, 255)
  Next i
  Set ElencoMovimenti.CellPicture = CaricaImmagineDaRisorsa("CANCELLA", , , RGB(49, 106, 197))
  ElencoMovimenti.Redraw = True
  RigaCorr = NuovaRiga
 End If
 ElencoMovimenti.Col = ColCorr
 If ElencoMovimenti.Col = 6 Then
  Dim Scelta%
  Scelta = MsgBox("Rimuovere la voce dall'elenco ?", vbExclamation + vbYesNo, "Fattura Pro")
  If Scelta = vbYes Then
   Dim rsArticoli As New ADODB.Recordset
   rsArticoli.Open "SELECT * FROM Articoli WHERE Descr = '" & ElencoMovimenti.TextMatrix(ElencoMovimenti.Row, _
   1) & "'", conn, adOpenDynamic, adLockOptimistic
   rsArticoli.MoveFirst
   If ElencoMovimenti.TextMatrix(ElencoMovimenti.Row, 4) = "Carico" Then
    rsArticoli("QntDisp") = rsArticoli("QntDisp") - CDbl(ElencoMovimenti.TextMatrix(ElencoMovimenti.Row, 2))
   Else
    rsArticoli("QntDisp") = rsArticoli("QntDisp") + CDbl(ElencoMovimenti.TextMatrix(ElencoMovimenti.Row, 2))
   End If
   rsArticoli.Update
   conn.Execute "DELETE * FROM MovimentiMagazzino WHERE IdMov = " & ElencoMovimenti.TextMatrix(ElencoMovimenti.Row, 7)
   rsArticoli.Close
   If ElencoMovimenti.Rows > 2 Then
    ElencoMovimenti.RemoveItem ElencoMovimenti.Row
   Else
    ElencoMovimenti.Rows = 1
   End If
   RigaCorr = 0
  End If
 End If
End If
End Sub
Private Sub ElencoMovimenti_DblClick()
Dim MovMag As New MovimentoMagazzino
MovMag.Tipo = ElencoMovimenti.TextMatrix(ElencoMovimenti.Row, 4)
MovMag.CaricaMovimento ElencoMovimenti.TextMatrix(ElencoMovimenti.Row, 7)
MovMag.Show vbModal
If MovMag.RecordModificato Then
 EseguiBackup = True
 CaricaMovimenti "SELECT MovimentiMagazzino.*, Descr FROM MovimentiMagazzino, Articoli WHERE" _
 & " Id = IdArticolo ORDER BY Data ASC"
End If
End Sub
Private Sub ElencoMovimenti_SelChange()
ElencoMovimenti_Click
End Sub
Private Sub Form_Load()
Me.Move 600, 450
Dim IntestazioniGriglia As Variant
IntestazioniGriglia = Array("Data", "Descrizione", "Quantità", "Cliente / Fornitore", "Movimento", _
"Rif. Movimento", "")
ElencoMovimenti.ColWidth(0) = 1000: ElencoMovimenti.ColWidth(1) = 4000
ElencoMovimenti.ColWidth(2) = 1000: ElencoMovimenti.ColWidth(3) = 4000
ElencoMovimenti.ColWidth(4) = 1200: ElencoMovimenti.ColWidth(5) = 2300
ElencoMovimenti.ColWidth(6) = 400: ElencoMovimenti.ColWidth(7) = 0
For i = 0 To ElencoMovimenti.Cols - 2
 ElencoMovimenti.TextMatrix(0, i) = IntestazioniGriglia(i): ElencoMovimenti.ColAlignment(i) = 4
Next i
CmbTipoMov.ListIndex = 2
Set rsMovimenti = New ADODB.Recordset
CaricaMovimenti "SELECT MovimentiMagazzino.*, Descr FROM MovimentiMagazzino, Articoli WHERE" _
& " Id = IdArticolo ORDER BY Data ASC"
End Sub
Private Sub CaricaMovimenti(StrSQL As String, Optional Ricerca As Boolean = False)
If rsMovimenti.State <> adStateClosed Then
 rsMovimenti.Close
End If
rsMovimenti.Open StrSQL, conn, adOpenDynamic
ElencoMovimenti.Rows = 1
If Ricerca Then
 BtnRicerca.Caption = "Annulla Ricerca"
End If
If Not rsMovimenti.EOF Then
 rsMovimenti.MoveFirst
 ElencoMovimenti.Redraw = False
 While Not rsMovimenti.EOF
  With ElencoMovimenti
  .AddItem ""
  .TextMatrix(.Rows - 1, 0) = Format$(rsMovimenti("Data"), "dd-mm-yyyy")
  .TextMatrix(.Rows - 1, 1) = rsMovimenti("Descr")
  .TextMatrix(.Rows - 1, 2) = FormatNumber(rsMovimenti("Qnt"), 3)
  .TextMatrix(.Rows - 1, 3) = rsMovimenti("CliFor")
  .TextMatrix(.Rows - 1, 4) = rsMovimenti("TipoMov")
  .TextMatrix(.Rows - 1, 5) = rsMovimenti("RifMov")
  .TextMatrix(.Rows - 1, 7) = rsMovimenti("IdMov")
  .Row = .Rows - 1: .Col = 6: .CellPictureAlignment = 4
  Set .CellPicture = ImgCancella
  End With
  rsMovimenti.MoveNext
 Wend
 ElencoMovimenti.Redraw = True
 RigaCorr = 0
End If
End Sub
