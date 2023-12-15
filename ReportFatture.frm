VERSION 5.00
Begin VB.Form ReportFatture 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Fatture"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8175
   Icon            =   "ReportFatture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   8175
   Begin VB.CheckBox ModificaNome 
      BackColor       =   &H0033CCFF&
      Caption         =   "Modifica Nome File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3135
      TabIndex        =   7
      Top             =   2190
      Width           =   2115
   End
   Begin VB.TextBox NomeFile 
      Enabled         =   0   'False
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
      Left            =   4215
      TabIndex        =   6
      Top             =   1605
      Width           =   3615
   End
   Begin VB.CommandButton Invia 
      Caption         =   "Invia Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5535
      TabIndex        =   4
      Top             =   300
      Width           =   1680
   End
   Begin VB.ListBox Anni 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   3720
      TabIndex        =   2
      Top             =   300
      Width           =   1500
   End
   Begin VB.ListBox Mesi 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "ReportFatture.frx":4072
      Left            =   870
      List            =   "ReportFatture.frx":409A
      TabIndex        =   1
      Top             =   300
      Width           =   1800
   End
   Begin VB.Label StatoInvio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "StatoInvio"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3135
      Width           =   885
   End
   Begin VB.Label StatoCreazione 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "StatoCreazione"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2835
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Report:"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   1635
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anno:"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   285
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mese:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   555
   End
End
Attribute VB_Name = "ReportFatture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NomeArchivio = "FattureClienti.xml"
Const Query = "/FattureClienti/FatturaCliente["

Private WithEvents CPDFCreator As PDFCreator.clsPDFCreator
Attribute CPDFCreator.VB_VarHelpID = -1
Private WithEvents MailReport As vbSendMail.clsSendMail
Attribute MailReport.VB_VarHelpID = -1
Dim COpzioniPDFCreator As PDFCreator.clsPDFCreatorOptions, _
ListaFatture As MSXML2.IXMLDOMNodeList, SS As ServiziStampa, NomiColonne As Variant, _
CoordRigaY As Single, CoordIntX As Single, CoordIntY As Single, PosColonne As Variant, _
Logo As StdPicture, NumPag%, Titolo$, StampanteCorrente As String, ProcessoInCorso As Boolean
Private Sub Anni_Click()
If Mesi.Text <> "" And Anni.Text <> "" Then
 NomeFile.Text = "Fatture " & Mesi.Text & " " & Anni.Text & ".pdf"
End If
End Sub
Private Sub CPDFCreator_eReady()
StatoCreazione.Caption = "Creazione documento pdf completata."
CPDFCreator.cPrinterStop = True: SS.ImpostaStampante StampanteCorrente
Call InviaReport
End Sub
Private Sub Form_Load()
Me.Move 1000, 800
If Dir(PercorsoApp & NomeArchivio) <> "" Then
 StatoCreazione = "": StatoInvio = "": Titolo = "FATTURE "
 Set DocFatture = New MSXML2.DOMDocument: DocFatture.async = False
 DocFatture.Load NomeArchivio: DocFatture.setProperty "SelectionLanguage", "XPath"
 Dim Anno%, AnnoFinale%, DataStr$, Controllo: AnnoFinale = Year(Date)
 DataStr = DocFatture.documentElement.firstChild.Attributes.getNamedItem("data").Text
 Anno = CInt(Split(DataStr, "-")(2)): Set SS = New ServiziStampa
 While Anno <= AnnoFinale
  Set Controllo = DocFatture.selectNodes(Query & "substring(@data,7,4)=" & Anno & "]")
  If Controllo.Length <> 0 Then Anni.AddItem Anno
  Anno = Anno + 1
 Wend
Else
 MsgBox "Attenzione, L'Archivio Fatture Clienti non esiste !", vbExclamation, "Fattura Pro"
 Exit Sub
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not ProcessoInCorso Then
 If Not CPDFCreator Is Nothing Then
  If InStr(CPDFCreator.cWindowsVersion, "Vista") Then
   CPDFCreator.cClose
   While CPDFCreator.cProgramIsRunning
    DoEvents
   Wend
  Else
   CPDFCreator.cClose: DoEvents
  End If
  Set ReportFatture = Nothing
 End If
Else: Cancel = 1
End If
End Sub
Private Sub Invia_Click()
If Mesi.Text <> "" And Anni.Text <> "" And (Not ProcessoInCorso) Then
 Set ListaFatture = DocFatture.selectNodes(Query & "substring(@data,4,2)=" & IdMese() _
 & " and substring(@data,7,4)=" & Anni.Text & "]")
 
 If ListaFatture.Length <> 0 Then
  ProcessoInCorso = True
  Set CPDFCreator = CreateObject("PDFCreator.clsPDFCreator")
  With CPDFCreator
  .cVisible = False
  If .cStart("/NoProcessingAtStartup") = False Then
   If .cStart("/NoProcessingAtStartup", True) = False Then
    MsgBox "Si è verificato un' errore durante l'inizializzazione di PDFCreator !", _
    vbExclamation, "Fattura Pro"
    ProcessoInCorso = False: Exit Sub
   End If
  End If
  Set COpzioniPDFCreator = .cOptions: .cClearCache
  End With
  StatoCreazione = "Creazione documento pdf in corso...": StatoInvio = ""
  With COpzioniPDFCreator
   .AutosaveDirectory = PercorsoApp: .AutosaveFilename = NomeFile
   .UseAutosave = 1: .UseAutosaveDirectory = 1
   .AutosaveFormat = 0
  End With
  Set CPDFCreator.cOptions = COpzioniPDFCreator
  Invia.Enabled = False: Titolo = Titolo & UCase(Mesi.Text) & " " & Anni.Text
  CreaReportFatture
 Else
  MsgBox "Non è stato trovato nessun documento relativo al mese e all'anno indicati !", _
  vbExclamation, "Fattura Pro"
 End If
Else
 MsgBox "Seleziona il mese e l'anno !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub CreaReportFatture()
Dim OffsetStampaX As Single, OffsetStampaY As Single, NomiCampi, campo$, j
StampanteCorrente = Printer.DeviceName
SS.ImpostaStampante "PDFCreator": SS.ImpostaDimensioniPagina
OffsetStampaX = SS.OffsetStampaX: OffsetStampaY = SS.OffsetStampaY
NomiColonne = Array("N. Documento", "Data", "Cliente", "Importo")
NomiCampi = Array("id", "data", "cliente", "tdoc"): CoordIntX = (SS.AreaStampaX - 180) / 2
PosColonne = Array(CoordIntX, CoordIntX + 24, CoordIntX + 44, CoordIntX + 160, CoordIntX + 180)
Set Logo = LoadResPicture(101, vbResBitmap)
Printer.Font.Name = "Arial": Printer.Font.Size = 9: NumPag = 1
DisegnaReport OffsetStampaX, OffsetStampaY: Printer.Font.Bold = False: CoordRigaY = CoordIntY

For r = 0 To ListaFatture.Length - 1
 If CoordRigaY + 6 < 230 Then
  For c = 1 To 4
   j = c - 1: campo = ListaFatture(r).Attributes.getNamedItem(NomiCampi(j)).Text
   If NomiCampi(j) = "cliente" Then campo = Split(campo, "$")(0)
   Printer.CurrentX = (PosColonne(c) - PosColonne(j) - Printer.TextWidth(campo)) / 2 _
   + PosColonne(j) + OffsetStampaX
   Printer.CurrentY = CoordRigaY + OffsetStampaY: Printer.Print campo
  Next c
 Else
  Printer.NewPage: NumPag = NumPag + 1: DisegnaReport OffsetStampaX, OffsetStampaY
  Printer.Font.Bold = False
 End If
 CoordRigaY = CoordRigaY + Printer.TextHeight("a") + 1
Next r

Printer.EndDoc: CPDFCreator.cPrinterStop = False
End Sub
Private Sub InviaReport()
Set MailReport = New vbSendMail.clsSendMail
With MailReport
 .SMTPHost = "mail.casalgismondo.it"
 .Username = "info@casalgismondo.it"
 .Password = "maragnao79"
 .UseAuthentication = True
 .From = "info@casalgismondo.it"
 .FromDisplayName = "Fattura Pro"
 .Recipient = "info@casalgismondo.it"
 .RecipientDisplayName = "Azienda AgroZootecnica Biologica Casalgismondo"
 .ServerReceipt = True
 .Priority = HIGH_PRIORITY
 .Subject = "Fatture " & Mesi.Text & " " & Anni.Text
 .Attachment = PercorsoApp & NomeFile
 StatoInvio = "Invio documento pdf in corso...completato al 0%."
 .send
End With
End Sub
Private Sub MailReport_Progress(PercentComplete As Long)
StatoInvio = "Invio documento pdf in corso...completato al " & PercentComplete & "%."
End Sub
Private Sub MailReport_SendFailed(Explanation As String)
StatoInvio = "Invio documento pdf fallito : " & Explanation
Call CancellaFileReport: ProcessoInCorso = False: Invia.Enabled = True
End Sub
Private Sub MailReport_SendSuccesful()
StatoInvio = "Invio documento pdf completato."
Call CancellaFileReport: ProcessoInCorso = False: Invia.Enabled = True
End Sub
Private Sub Mesi_Click()
If Mesi.Text <> "" And Anni.Text <> "" Then
 NomeFile.Text = "Fatture " & Mesi.Text & " " & Anni.Text & ".pdf"
End If
End Sub
Private Sub ModificaNome_Click()
NomeFile.Enabled = ModificaNome.Value
If (Not NomeFile.Enabled) And Mesi.Text <> "" And Anni.Text <> "" Then
 NomeFile.Text = "Fatture " & Mesi.Text & " " & Anni.Text & ".pdf"
 End If
End Sub
Private Sub DisegnaReport(ByVal OffsetStampaX As Single, OffsetStampaY As Single)
Printer.Font.Bold = True
Printer.PaintPicture Logo, ((SS.AreaStampaX - (Logo.Width / 100)) / 2) + OffsetStampaX, 15 + OffsetStampaY
Printer.CurrentX = (SS.AreaStampaX - Printer.TextWidth(Titolo)) / 2 + OffsetStampaX
CoordIntY = 15 + (Logo.Height / 100) + 10
Printer.CurrentY = CoordIntY + OffsetStampaY: Printer.Print Titolo
CoordIntY = CoordIntY + 15
Printer.Line (CoordIntX + OffsetStampaX, CoordIntY + OffsetStampaY)-(CoordIntX + 180 + OffsetStampaX, 230 + _
OffsetStampaY), RGB(0, 0, 0), B
Printer.Line (CoordIntX + OffsetStampaX, CoordIntY + 8 + OffsetStampaY)-(CoordIntX + 180 + OffsetStampaX, _
CoordIntY + 8 + OffsetStampaY)
CoordRigaY = (8 - Printer.TextHeight("a")) / 2 + CoordIntY + OffsetStampaY

For i = 1 To 4
 j = i - 1
 If i < 4 Then
  Printer.Line (PosColonne(i) + OffsetStampaX, CoordIntY + OffsetStampaY)-(PosColonne(i) + _
  OffsetStampaX, CoordIntY + 8 + OffsetStampaY)
 End If
 Printer.CurrentX = (PosColonne(i) - PosColonne(j) - Printer.TextWidth(NomiColonne(j))) _
 / 2 + PosColonne(j) + OffsetStampaX
 Printer.CurrentY = CoordRigaY: Printer.Print NomiColonne(j)
Next i

Printer.CurrentX = (SS.AreaStampaX - Printer.TextWidth("Pagina " & NumPag)) / 2 + OffsetStampaX
Printer.CurrentY = SS.AreaStampaY - Printer.TextHeight("a") - 3 + OffsetStampaY
CoordIntY = CoordIntY + 10: Printer.Print "Pagina " & NumPag
End Sub
Private Sub CancellaFileReport()
Dim fso As FileSystemObject: Set fso = New FileSystemObject
fso.DeleteFile PercorsoApp & NomeFile, True
End Sub
Private Function IdMese() As String
Dim NumMese%: NumMese = Mesi.ListIndex + 1
If NumMese < 10 Then
 IdMese = "0" & NumMese
Else: IdMese = NumMese
End If
End Function
