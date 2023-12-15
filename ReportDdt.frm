VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ReportDdt 
   BackColor       =   &H0033CCFF&
   Caption         =   "Report Ddt"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15105
   Icon            =   "ReportDdt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   15105
   Begin VB.CheckBox ChkPagCorr 
      BackColor       =   &H0033CCFF&
      Caption         =   "Stampa Schermata Corrente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   9
      Top             =   2100
      Width           =   2940
   End
   Begin VB.CommandButton PagFin 
      Height          =   315
      Left            =   8265
      MaskColor       =   &H00D8E9EC&
      Picture         =   "ReportDdt.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Pagina Finale"
      Top             =   345
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.TextBox PagCorr 
      Height          =   315
      Left            =   6030
      TabIndex        =   6
      Top             =   345
      Width           =   540
   End
   Begin VB.CommandButton PagIniz 
      Height          =   315
      Left            =   4215
      MaskColor       =   &H00D8E9EC&
      Picture         =   "ReportDdt.frx":0AA0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Pagina Iniziale"
      Top             =   345
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.CommandButton PagPrec 
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
      Left            =   4755
      MaskColor       =   &H00D8E9EC&
      Picture         =   "ReportDdt.frx":11B6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Pagina Precedente"
      Top             =   345
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.CommandButton PagSucc 
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
      Left            =   7725
      MaskColor       =   &H00D8E9EC&
      Picture         =   "ReportDdt.frx":18CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Pagina Successiva"
      Top             =   345
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.CommandButton StampaReport 
      Caption         =   "Stampa Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ReportDdt.frx":1FE2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   870
      UseMaskColor    =   -1  'True
      Width           =   1785
   End
   Begin MSFlexGridLib.MSFlexGrid TabellaDdt 
      Height          =   4770
      Left            =   195
      TabIndex        =   0
      Top             =   870
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   8414
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   24576
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label PagCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6690
      TabIndex        =   7
      Top             =   375
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pagina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5310
      TabIndex        =   5
      Top             =   375
      Width           =   645
   End
End
Attribute VB_Name = "ReportDdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ListaDdt As MSXML2.IXMLDOMNodeList, IndiceElenco As Integer, IndiceFinale As Integer, _
NumPag As Integer, RecPag As Integer
Private Sub Form_Load()
Dim FattureClientiDoc As MSXML2.DOMDocument40, i
Me.Move 300, 300, 15345, 6435: IntestazioniGriglia = Array("Numero", "Data", "Cliente", "N° Fattura")
TabellaDdt.ColWidth(0) = 2500: TabellaDdt.ColWidth(1) = 1500
TabellaDdt.ColWidth(2) = 6000: TabellaDdt.ColWidth(3) = 1300

For i = 0 To TabellaDdt.Cols - 1
 TabellaDdt.TextMatrix(0, i) = IntestazioniGriglia(i)
 TabellaDdt.ColAlignment(i) = 4
Next i

If Dir(PercorsoApp & "FattureClienti.xml") <> "" Then
 Set FattureClientiDoc = New MSXML2.DOMDocument40
 FattureClientiDoc.setProperty "SelectionLanguage", "XPath"
 FattureClientiDoc.Load "FattureClienti.xml"
 Set ListaDdt = FattureClientiDoc.selectNodes("/FattureClienti/FatturaCliente/" _
 & "ElencoDdt/Ddt[substring(@data,7,4)=" & Year(Date) & "]")

 If ListaDdt.Length <> 0 Then
  Dim Cliente$, AltezzaRighe%, AltezzaRiga%
  PagPrec.Enabled = False: PagCorr.Text = "1": AltezzaRiga = TabellaDdt.RowHeight(0)
 
  For i = 0 To ListaDdt.Length - 1
   With TabellaDdt
    AltezzaRighe = AltezzaRighe + AltezzaRiga + 1
    If AltezzaRighe + 45 >= TabellaDdt.Height Then Exit For
    .Rows = .Rows + 1: .Row = .Rows - 1: RecPag = RecPag + 1
    TabellaDdt.TextMatrix(.Row, 0) = ListaDdt(i).Attributes.getNamedItem("num").Text
    TabellaDdt.TextMatrix(.Row, 1) = ListaDdt(i).Attributes.getNamedItem("data").Text
    Cliente = ListaDdt(i).parentNode.parentNode.Attributes.getNamedItem("cliente").Text
    Cliente = Split(Cliente, "$")(0): TabellaDdt.TextMatrix(.Row, 2) = Cliente
    TabellaDdt.TextMatrix(.Row, 3) = ListaDdt(i).parentNode.parentNode.Attributes.getNamedItem("id").Text
   End With
  Next i

  If i = ListaDdt.Length Then PagSucc.Enabled = False
  NumPag = -Int(-(ListaDdt.Length / RecPag)): PagCount = "di " & NumPag
 Else
  MsgBox "Attenzione, non sono stati trovati ddt in archivio !", vbExclamation, "Fattura Pro"
 End If
Else
 MsgBox "Attenzione, L'Archivio Fatture Clienti non esiste !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set ReportDdt = Nothing
End Sub
Private Sub PagPrec_Click()
IndiceElenco = IndiceElenco - RecPag: IndiceFinale = IndiceElenco + RecPag - 1
TabellaDdt.Rows = 1: PagCorr.Text = CInt(PagCorr.Text) - 1
If IndiceElenco = 0 Then PagPrec.Enabled = False
Dim r: r = 1: PagSucc.Enabled = True
For i = IndiceElenco To IndiceFinale
 TabellaDdt.AddItem ""
 TabellaDdt.TextMatrix(r, 0) = ListaDdt(i).Attributes.getNamedItem("num").Text
 TabellaDdt.TextMatrix(r, 1) = ListaDdt(i).Attributes.getNamedItem("data").Text
 Cliente = ListaDdt(i).parentNode.parentNode.Attributes.getNamedItem("cliente").Text
 Cliente = Split(Cliente, "$")(0): TabellaDdt.TextMatrix(r, 2) = Cliente
 TabellaDdt.TextMatrix(r, 3) = ListaDdt(i).parentNode.parentNode.Attributes.getNamedItem("id").Text
 r = r + 1
Next i
End Sub
Private Sub PagSucc_Click()
IndiceElenco = IndiceElenco + RecPag: IndiceFinale = IndiceElenco + RecPag - 1
If ListaDdt.Length - IndiceFinale < 0 Then
 IndiceFinale = ListaDdt.Length - 1: PagSucc.Enabled = False
End If
TabellaDdt.Rows = 1: PagCorr.Text = CInt(PagCorr.Text) + 1
Dim r: r = 1: PagPrec.Enabled = True
For i = IndiceElenco To IndiceFinale
 TabellaDdt.AddItem ""
 TabellaDdt.TextMatrix(r, 0) = ListaDdt(i).Attributes.getNamedItem("num").Text
 TabellaDdt.TextMatrix(r, 1) = ListaDdt(i).Attributes.getNamedItem("data").Text
 Cliente = ListaDdt(i).parentNode.parentNode.Attributes.getNamedItem("cliente").Text
 Cliente = Split(Cliente, "$")(0): TabellaDdt.TextMatrix(r, 2) = Cliente
 TabellaDdt.TextMatrix(r, 3) = ListaDdt(i).parentNode.parentNode.Attributes.getNamedItem("id").Text
 r = r + 1
Next i
End Sub
Private Sub StampaReport_Click()
If TabellaDdt.Rows > 1 Then
 Dim SS As New ServiziStampa: Set SS.ElencoDdt = ListaDdt
 SS.TipoDoc = ReportDocumentiTrasporto: SS.ImpostaAnteprima True, True
 If ChkPagCorr.value Then
  SS.StampaReportDdt IndiceElenco, RecPag
 Else: SS.StampaReportDdt -1, , 0
 End If
Else
 MsgBox "Attenzione, il report da stampare è vuoto !", vbExclamation, "Fattura Pro"
End If
End Sub
