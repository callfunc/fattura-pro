VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.MDIForm FatturaProMDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Fattura Pro 1.0"
   ClientHeight    =   9435
   ClientLeft      =   225
   ClientTop       =   510
   ClientWidth     =   19980
   Icon            =   "FattMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Visible         =   0   'False
   Begin MSComctlLib.ImageList IconeMenu 
      Left            =   5070
      Top             =   3765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   14215660
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":4072
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":415C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":4251
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":42CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":434D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":43CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":44BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":45BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":462F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":48A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":4904
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":49B5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarraFunzioni 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19980
      _ExtentX        =   35243
      _ExtentY        =   1482
      ButtonWidth     =   2990
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImmaginiBarra"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Archivio Clienti"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   120
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fatture Clienti"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   120
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Archivio Fornitori"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   120
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fatture Fornitori"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   120
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Note di Credito"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fatture Differite"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   120
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Note di Consegna"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   120
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Articoli"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   120
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gestione Magazzino"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Scadenzario"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImmaginiBarra 
      Left            =   6540
      Top             =   3825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":4B23
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":5007
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":50DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":5A99
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":5FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":60A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":639F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FattMain.frx":648C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Opzioni 
      Caption         =   "Opzioni Ricerca"
      NegotiatePosition=   1  'Left
      Begin VB.Menu MenuRicercaArticoli 
         Caption         =   "Ricerca Articoli"
      End
      Begin VB.Menu MenuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuRicercaClienti 
         Caption         =   "Ricerca Clienti"
      End
      Begin VB.Menu MenuRicercaNote 
         Caption         =   "Ricerca Note di Consegna"
      End
      Begin VB.Menu MenuRicercaFattureCli 
         Caption         =   "Ricerca Fatture Clienti"
      End
      Begin VB.Menu MenuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuRicercaNoteCredito 
         Caption         =   "Ricerca Note di Credito"
      End
      Begin VB.Menu MenuRicercaFornitori 
         Caption         =   "Ricerca Fornitori"
      End
      Begin VB.Menu MenuRicercaFattureFor 
         Caption         =   "Ricerca Fatture Fornitori"
      End
      Begin VB.Menu MenuRicercaVettori 
         Caption         =   "Ricerca Vettori"
      End
   End
   Begin VB.Menu Tabelle 
      Caption         =   "Tabelle"
      Begin VB.Menu MenuAliquoteIva 
         Caption         =   "Aliquote Iva"
      End
      Begin VB.Menu MenuVettori 
         Caption         =   "Vettori"
      End
   End
   Begin VB.Menu Strumenti 
      Caption         =   "Strumenti"
      Begin VB.Menu MenuImpostazioni 
         Caption         =   "Impostazioni"
      End
      Begin VB.Menu MenuBilancio 
         Caption         =   "Bilancio"
      End
   End
   Begin VB.Menu MenuAiuto 
      Caption         =   "?"
      Begin VB.Menu MenuInfo 
         Caption         =   "Informazioni su Fattura Pro"
      End
   End
End
Attribute VB_Name = "FatturaProMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IMenu As IconMenu6.cIconMenu, fso As FileSystemObject
Private Sub MenuAliquoteIva_Click()
AliquoteIVA.Show
End Sub
Private Sub MenuImpostazioni_Click()
InfoDitta.Show
End Sub
Private Sub MenuBilancio_Click()
Bilancio.Show
End Sub
Private Sub MenuInfo_Click()
Info.Show
End Sub
Private Sub MDIForm_Load()
Dim re As New RegExp, DatiIcone, InfoIcona

DatiIcone = Array("Ricerca Articoli:1", "Ricerca Clienti:2", "Ricerca Note di Consegna:3", _
"Ricerca Fatture Clienti:4", "Ricerca Note di Credito:5", "Ricerca Fornitori:6", _
"Ricerca Fatture Fornitori:7", "Ricerca Vettori:13", "Aliquote Iva:12", "Vettori:13", _
"Impostazioni:8", "Bilancio:9", "Informazioni su Fattura Pro:11")

Set IMenu = New IconMenu6.cIconMenu: Logo.Show
With IMenu
 .ImageList = IconeMenu
 .Attach Me.hWnd: .OfficeXpStyle = False
 .HighlightStyle = ECPHighlightStyleGradient
 For i = 0 To UBound(DatiIcone)
  InfoIcona = Split(DatiIcone(i), ":")
  .IconIndex(InfoIcona(0)) = CInt(InfoIcona(1))
 Next i
End With
End Sub
Public Sub Esci()
Static in_esecuzione As Boolean
If in_esecuzione Then Exit Sub
in_esecuzione = True
Dim f As Form

For Each f In Forms
 If f.Name <> "FatturaProMDI" Then
  Unload f: Set f = Nothing
 End If
Next

Set rsImpostazioni = New ADODB.Recordset
rsImpostazioni.Open "Impostazioni", conn, adOpenDynamic, adLockOptimistic

If Not rsImpostazioni.EOF Then
 EseguiBackup = EseguiBackup Or rsImpostazioni("backup")
End If

If EseguiBackup Then
 Set fso = New FileSystemObject
 If Not fso.FolderExists("C:\FP") Then
  fso.CreateFolder "C:\FP"
 End If
 rsImpostazioni("backup") = True
 rsImpostazioni.Update
 fso.CopyFile PercorsoApp & "FatturaPro.mdb", "C:\FP\FatturaPro.mdb", True
 'Backup.Show vbModal
End If

FreeGDIPlus

If conn.State <> adStateClosed Then
 conn.Close
End If

If SepMigCorr <> "." Then SetLocaleInfo LCID, LOCALE_STHOUSAND, SepMigCorr
If SepDecCorr <> "," Then SetLocaleInfo LCID, LOCALE_SDECIMAL, SepDecCorr
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
Esci
End Sub
Private Sub MenuRicercaArticoli_Click()
RicercaArticoli.Show
End Sub
Private Sub MenuRicercaClienti_Click()
RicercaClienti.Show
End Sub
Private Sub MenuRicercaFattureCli_Click()
RicercaFattureClienti.Show
End Sub
Private Sub MenuRicercaFattureFor_Click()
RicercaFattureFornitori.Show
End Sub
Private Sub MenuRicercaFornitori_Click()
RicercaFornitori.Show
End Sub
Private Sub MenuRicercaNoteCredito_Click()
RicercaNoteCredito.Show
End Sub
Private Sub MenuRicercaNote_Click()
RicercaNote.Show
End Sub
Private Sub BarraFunzioni_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2: Clienti.Show
Case 4: FattureClienti.Show
Case 6: Fornitori.Show
Case 8: FattureFornitori.Show
Case 10: NoteCredito.Show
Case 12: CreaFatturaDiff.Show
Case 14: Note.Show
Case 16: Articoli.Show
Case 18: Magazzino.Show
Case 20: Scadenzario.Show
End Select
End Sub
Private Sub MenuRicercaVettori_Click()
RicercaVettori.Show
End Sub
Private Sub MenuVettori_Click()
Vettori.Show
End Sub
