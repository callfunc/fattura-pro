VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Impostazioni 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup e Info Ditta"
   ClientHeight    =   8175
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   11985
   Icon            =   "Impostazioni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11985
   Begin TabDlg.SSTab SSTab1 
      Height          =   8025
      Left            =   15
      TabIndex        =   0
      Top             =   90
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14155
      _Version        =   393216
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      Enabled         =   0   'False
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
      TabCaption(0)   =   "Backup"
      TabPicture(0)   =   "Impostazioni.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SchedaDitta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dati Ditta"
      TabPicture(1)   =   "Impostazioni.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SchedaVarie"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Varie"
      TabPicture(2)   =   "Impostazioni.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.PictureBox SchedaVarie 
         Appearance      =   0  'Flat
         BackColor       =   &H0033CCFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7125
         Left            =   -74850
         ScaleHeight     =   7125
         ScaleWidth      =   11145
         TabIndex        =   33
         Top             =   390
         Visible         =   0   'False
         Width           =   11145
         Begin VB.TextBox CartellaRemotaFtp 
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
            Left            =   345
            TabIndex        =   36
            Top             =   675
            Width           =   7080
         End
         Begin VB.CommandButton SalvaVarie 
            Caption         =   "Salva Modifiche"
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
            Left            =   360
            Picture         =   "Impostazioni.frx":03DE
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1365
            Width           =   1770
         End
         Begin VB.TextBox Edit 
            Alignment       =   2  'Center
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
            Height          =   270
            Left            =   8175
            TabIndex        =   34
            Top             =   3825
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cartella Dati Remota:"
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
            Left            =   330
            TabIndex        =   37
            Top             =   405
            Width           =   1665
         End
      End
      Begin VB.PictureBox SchedaDitta 
         BackColor       =   &H0033CCFF&
         BorderStyle     =   0  'None
         Height          =   7485
         Left            =   120
         ScaleHeight     =   7485
         ScaleWidth      =   11610
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   11610
         Begin VB.CommandButton Cerca 
            Height          =   315
            Left            =   7905
            MaskColor       =   &H00000000&
            Picture         =   "Impostazioni.frx":1020
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   825
            Width           =   945
         End
         Begin VB.CommandButton SalvaDati 
            Caption         =   "Salva Dati"
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
            Left            =   285
            Picture         =   "Impostazioni.frx":141E
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   6150
            Width           =   1770
         End
         Begin VB.TextBox Tel 
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
            Left            =   2115
            TabIndex        =   11
            Tag             =   "O"
            Top             =   4560
            Width           =   4395
         End
         Begin VB.TextBox SedeAz 
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
            Left            =   2115
            TabIndex        =   10
            Tag             =   "AO"
            Top             =   3615
            Width           =   7395
         End
         Begin VB.TextBox Piva 
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
            Left            =   2115
            TabIndex        =   9
            Tag             =   "NO"
            Top             =   2685
            Width           =   5160
         End
         Begin VB.TextBox RagSoc 
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
            Left            =   2115
            TabIndex        =   8
            Tag             =   "AO"
            Top             =   2220
            Width           =   6795
         End
         Begin VB.TextBox LogoAzienda 
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
            Left            =   2115
            TabIndex        =   7
            Tag             =   "AO"
            Top             =   825
            Width           =   5745
         End
         Begin VB.TextBox Fax 
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
            Left            =   2115
            TabIndex        =   6
            Tag             =   "o"
            Top             =   5040
            Width           =   3900
         End
         Begin VB.TextBox NomeAzienda 
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
            Left            =   2115
            TabIndex        =   5
            Tag             =   "AO"
            Top             =   1755
            Width           =   6105
         End
         Begin VB.TextBox Cfisc 
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
            Left            =   2115
            TabIndex        =   4
            Tag             =   "A"
            Top             =   3150
            Width           =   5790
         End
         Begin VB.TextBox SedeLeg 
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
            Left            =   2115
            TabIndex        =   3
            Tag             =   "AO"
            Top             =   4080
            Width           =   6000
         End
         Begin VB.PictureBox Logo 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00006000&
            Height          =   900
            Left            =   2655
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   340
            TabIndex        =   2
            Top             =   6150
            Width           =   5100
         End
         Begin MSComDlg.CommonDialog Sfoglia 
            Left            =   10770
            Top             =   4560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Cerca File"
            Filter          =   "File Immagine  (*.jpg;*.bmp;*.gif) | *.jpg;*.bmp;*.gif"
            FontName        =   "Ms Sans Serif"
            FontSize        =   10
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono:"
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
            Left            =   285
            TabIndex        =   32
            Top             =   4575
            Width           =   750
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sede Aziendale:"
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
            Left            =   285
            TabIndex        =   31
            Top             =   3645
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partita Iva:"
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
            Left            =   270
            TabIndex        =   30
            Top             =   2715
            Width           =   825
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ragione Sociale:"
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
            Left            =   270
            TabIndex        =   29
            Top             =   2250
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00006000&
            BackStyle       =   0  'Transparent
            Caption         =   "Logo Ditta:"
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
            Left            =   270
            TabIndex        =   28
            Top             =   855
            Width           =   870
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
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
            Left            =   285
            TabIndex        =   27
            Top             =   5055
            Width           =   300
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Ditta:"
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
            Left            =   270
            TabIndex        =   26
            Top             =   1800
            Width           =   960
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codice Fiscale:"
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
            Left            =   270
            TabIndex        =   25
            Top             =   3210
            Width           =   1170
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sede Legale:"
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
            Left            =   270
            TabIndex        =   24
            Top             =   4095
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   255
            Left            =   1320
            TabIndex        =   23
            Top             =   840
            Width           =   75
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   240
            Left            =   1395
            TabIndex        =   22
            Top             =   1785
            Width           =   75
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   240
            Left            =   1830
            TabIndex        =   21
            Top             =   3630
            Width           =   75
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   240
            Left            =   1905
            TabIndex        =   20
            Top             =   2235
            Width           =   75
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   210
            Left            =   1560
            TabIndex        =   19
            Top             =   4080
            Width           =   75
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   240
            Left            =   1275
            TabIndex        =   18
            Top             =   2715
            Width           =   75
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   165
            Left            =   1215
            TabIndex        =   17
            Top             =   4575
            Width           =   75
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   210
            Left            =   750
            TabIndex        =   16
            Top             =   5040
            Width           =   75
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I campi con l'asterisco sono obbligatori"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C8&
            Height          =   225
            Left            =   270
            TabIndex        =   15
            Top             =   270
            Width           =   3105
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota: Se il campo Tipo Ditta viene riempito il campo Nome Ditta verrà visualizzato sotto di esso"
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
            Left            =   315
            TabIndex        =   14
            Top             =   5625
            Width           =   7515
         End
      End
   End
End
Attribute VB_Name = "Impostazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocImpostazioni As New MSXML2.DOMDocument40, DatiDitta As MSXML2.IXMLDOMElement
Dim SezioneEsenzioneIva As MSXML2.IXMLDOMElement
Dim DocNote As New MSXML2.DOMDocument40, DocFattureFornitori As New MSXML2.DOMDocument40, _
DocFattureClienti As New MSXML2.DOMDocument40, CartellaBackup As String, DiscoCorr As String
Dim fso As New FileSystemObject, pic As IPictureDisp, Dati As Variant, _
Selezionato As Boolean, DatiModificati As Boolean, AnnoCorr As Integer
Private Sub Carica_Click()
If Selezionato And AnniArchiviati.ListIndex <> -1 Then
 Dim Anno$: Anno = AnniArchiviati.Text
 If RdbNote.Value Or RdbFattureClienti.Value Then
  Set BackupNote = New MSXML2.DOMDocument40: BackupNote.async = False
  Set BackupFattureClienti = Nothing
  
  If Dir(PercorsiBackup.List(0) & "\Note_" & Anno & ".xml") <> "" Then
   BackupNote.Load PercorsiBackup.List(0) & "\Note_" & Anno & ".xml"
  ElseIf PercorsiBackup.ListCount <> 1 Then
   BackupNote.Load PercorsoApp & PercorsiBackup.List(1) & "\Note_" & Anno & ".xml"
  End If
  
  If RdbFattureClienti.Value Then
   Dim PercorsoBackup$
   If Dir(PercorsiBackup.List(0) & "\FattureClienti_" & Anno & ".xml") <> "" Then
    PercorsoBackup = PercorsiBackup.List(0) & "\FattureClienti_" & Anno & ".xml"
   ElseIf PercorsiBackup.ListCount <> 1 Then
    If Dir(PercorsoApp & PercorsiBackup.List(1) & "\FattureClienti_" & Anno & ".xml") <> "" Then
     PercorsoBackup = PercorsoApp & PercorsiBackup.List(1) & "\FattureClienti_" & Anno & ".xml"
    End If
   End If
   If PercorsoBackup <> "" Then
    Set BackupFattureClienti = New MSXML2.DOMDocument40
    BackupFattureClienti.async = False: BackupFattureClienti.Load PercorsoBackup
   Else
    MsgBox "Attenzione, il file di backup delle fatture clienti che hai selezionato non esiste." _
    & vbCrLf & "Impossibile effettuare il caricamento dei dati !", vbExclamation, _
    "Fattura Pro"
   End If
  End If
 Else
  Set BackupFattureFornitori = New MSXML2.DOMDocument40: BackupFattureFornitori.async = False
  If Dir(PercorsiBackup.List(0) & "\FattureFornitori_" & Anno & ".xml") <> "" Then
   BackupFattureFornitori.Load PercorsiBackup.List(0) & "\FattureFornitori_" & Anno & ".xml"
  ElseIf PercorsiBackup.ListCount <> 1 Then
   BackupFattureFornitori.Load PercorsoApp & PercorsiBackup.List(1) & "\FattureFornitori_" & Anno & ".xml"
  End If
 End If
Else
 MsgBox "Attenzione, per effettuare il caricamento dei dati di backup devi selezionare" _
 & " il tipo di documento e scegliere un anno dall'elenco !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub Cerca_Click()
Sfoglia.ShowOpen
If Sfoglia.FileName <> "" Then
 Set pic = LoadPicture(Sfoglia.FileName)
 If (pic.Width / 100) > 92 Or (pic.Height / 100) > 16 Then
  MsgBox "Attenzione, il file immagine del logo deve avere una larghezza massima di " _
  & "92 mm (340 pixel) ed un'altezza massima di 16 mm (60 pixel) !", vbExclamation, _
  "Fattura Pro"
  LogoAzienda.Text = ""
 Else
  LogoAzienda.Text = Sfoglia.FileName: Set Logo.Picture = pic
 End If
End If
End Sub
Private Sub Edit_LostFocus()
Edit.Visible = False
End Sub
Private Sub RdbNote_Click()
Selezionato = True: CaricaAnni 0
End Sub
Private Sub RdbFattureClienti_Click()
Selezionato = True: CaricaAnni 0
End Sub
Private Sub RdbFattureFornitori_Click()
Selezionato = True: CaricaAnni 1
End Sub
Private Sub Form_Load()
Dim d As Drive, IntestazioniGriglia
Me.Move 350, 450
AnnoCorr = Year(Date): Dati = Array("LogoAzienda", "TipoAzienda", "NomeAzienda", "RagSoc", _
"PIva", "Cfisc", "SedeAz", "SedeLeg", "Tel", "Fax")

Logo.Font.Name = "Arial": Logo.Font.Size = 14: Logo.Font.Bold = True

IntestazioniGriglia = Array("Descrizione", "")
CasiEsenzione.ColWidth(0) = 8000: CasiEsenzione.ColWidth(1) = 360

For i = 0 To CasiEsenzione.Cols - 1
 CasiEsenzione.TextMatrix(0, i) = IntestazioniGriglia(i)
 CasiEsenzione.ColAlignment(i) = 4
Next i

For Each d In fso.Drives
 If d.DriveType = Fixed Then
  ElencoDischi.AddItem d.DriveLetter & ":" & "  Disco Fisso"
 ElseIf d.DriveType = Removable Then
  ElencoDischi.AddItem d.DriveLetter & ":" & "  Disco Rimovibile"
 End If
Next d
AnnoCorr = Year(Date)
If Dir(PercorsoApp & "Note.xml") <> "" Then
 DocNote.async = False: DocNote.setProperty "SelectionLanguage", "XPath"
 DocNote.Load "Note.xml"
End If
If Dir(PercorsoApp & "FattureFornitori.xml") <> "" Then
 DocFattureFornitori.async = False: DocFattureFornitori.setProperty "SelectionLanguage", "XPath"
 DocFattureFornitori.Load "FattureFornitori.xml"
End If

If AnnoBackup <> "" Then
 AnnoCaricato.Text = AnnoBackup
 If Not BackupFattureClienti Is Nothing Then DocBackup.Text = "Fatture Clienti"
 If Not BackupNote Is Nothing Then DocBackup.Text = "Note"
 If Not BackupFattureFornitori Is Nothing Then DocBackup.Text = "Fatture Fornitori"
End If

Dim ImpostazioniApp As MSXML2.IXMLDOMElement
If Dir(PercorsoApp & "Impostazioni.xml") = "" Then
 Set ImpostazioniApp = DocImpostazioni.createElement("Impostazioni")
 DocImpostazioni.appendChild ImpostazioniApp
Else
 DocImpostazioni.async = False: DocImpostazioni.setProperty "SelectionLanguage", "XPath"
 DocImpostazioni.Load "Impostazioni.xml"
 If DocImpostazioni.documentElement Is Nothing Then
  Set ImpostazioniApp = DocImpostazioni.createElement("Impostazioni")
  DocImpostazioni.appendChild ImpostazioniApp
 End If
 
 Set DatiDitta = DocImpostazioni.selectSingleNode("/Impostazioni/DatiAzienda")
 If Not DatiDitta Is Nothing Then
  LogoAzienda = DatiDitta.getAttribute("logo")
  If Dir(LogoAzienda) <> "" Then
   Set Logo.Picture = LoadPicture(LogoAzienda)
  End If
  NomeAzienda = DatiDitta.getAttribute("nomeaz")
  TipoAzienda = DatiDitta.getAttribute("tipoaz")
  RagSoc = DatiDitta.getAttribute("ragsoc")
  Piva = DatiDitta.getAttribute("piva"): Cfisc = DatiDitta.getAttribute("cfisc")
  SedeAz = DatiDitta.getAttribute("sedeaz"): SedeLeg = DatiDitta.getAttribute("sedeleg")
  Tel = DatiDitta.getAttribute("tel"): Fax = DatiDitta.getAttribute("fax")
 End If
 
 Dim CartellaDatiRemota As MSXML2.IXMLDOMNode
 Set SezioneEsenzioneIva = DocImpostazioni.selectSingleNode("/Impostazioni/EsenzioneIva")
 If Not SezioneEsenzioneIva Is Nothing Then
  For i = 0 To SezioneEsenzioneIva.childNodes.Length - 1
   CasiEsenzione.AddItem ""
   CasiEsenzione.TextMatrix(i + 1, 0) = SezioneEsenzioneIva.childNodes(i).Attributes.getNamedItem("desc").Text
   CasiEsenzione.Row = i + 1: CasiEsenzione.Col = 1
   CasiEsenzione.CellPictureAlignment = 4: Set CasiEsenzione.CellPicture = ImgCancella
  Next i
 Else
  Set SezioneEsenzioneIva = DocImpostazioni.createElement("EsenzioneIva")
 End If
 Set CartellaDatiRemota = DocImpostazioni.selectSingleNode("/Impostazioni/CartellaDatiRemota")
 If Not CartellaDatiRemota Is Nothing Then
  CartellaRemotaFtp.Text = CartellaDatiRemota.Attributes.getNamedItem("percorso").Text
 End If
End If
CasiEsenzione.AddItem ""
If LogoAzienda = "" Or Logo.Picture.Width = 0 Then
 Logo.CurrentX = (Logo.ScaleWidth - Logo.TextWidth("Logo Ditta")) / 2
 Logo.CurrentY = (Logo.ScaleHeight - Logo.TextHeight("a")) / 2
 Logo.Print "Logo Ditta"
End If
SSTab1.Tab = 0
End Sub
Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
 SSTab1.Move 105, 150, Me.ScaleWidth - 210, Me.ScaleHeight - 255
 SchedaBackup.Move 15, 315, SSTab1.Width - 45, SSTab1.Height - 330
 SchedaDitta.Move 15, 315, SSTab1.Width - 45, SSTab1.Height - 330
 SchedaVarie.Move 15, 315, SSTab1.Width - 45, SSTab1.Height - 330
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Impostazioni = Nothing
End Sub
Private Sub SalvaDati_Click()
For i = 0 To UBound(Dati)
 If InStr(1, Me(Dati(i)).Tag, "O") <> 0 And Me(Dati(i)) = "" Then
  MsgBox "Attenzione, uno dei campi obbligatori è vuoto !", _
  vbExclamation, "Fattura Pro": Me(Dati(i)).SetFocus: Exit Sub
 ElseIf Me(Dati(i)) <> "" Then
  If InStr(1, Me(Dati(i)).Tag, "A") <> 0 And IsNumeric(Me(Dati(i))) Then
   MsgBox "Attenzione, il campo " & Dati(i) & " non puo contenere valori numerici !", _
   vbExclamation, "Fattura Pro": Me(Dati(i)).SetFocus: Exit Sub
  ElseIf InStr(1, Me(Dati(i)).Tag, "N") <> 0 And Not IsNumeric(Me(Dati(i))) Then
   MsgBox "Attenzione, il campo " & Dati(i) & " deve contenere valori numerici !", _
   vbExclamation, "Fattura Pro": Me(Dati(i)).SetFocus: Exit Sub
  End If
 End If
Next i
If DatiDitta Is Nothing Then
 Set DatiDitta = DocImpostazioni.createElement("DatiAzienda")
 DocImpostazioni.documentElement.appendChild DatiDitta
End If
DatiDitta.setAttribute "logo", Trim(LogoAzienda)
DatiDitta.setAttribute "ragsoc", EliminaSpazi(RagSoc)
DatiDitta.setAttribute "tipoaz", EliminaSpazi(TipoAzienda)
DatiDitta.setAttribute "nomeaz", EliminaSpazi(NomeAzienda)
DatiDitta.setAttribute "piva", Trim(Piva)
DatiDitta.setAttribute "cfisc", Trim(Cfisc)
DatiDitta.setAttribute "sedeaz", EliminaSpazi(SedeAz)
DatiDitta.setAttribute "sedeleg", EliminaSpazi(SedeLeg)
DatiDitta.setAttribute "tel", Trim(Tel)
DatiDitta.setAttribute "fax", Trim(Fax): DocImpostazioni.save PercorsoApp & "Impostazioni.xml"
End Sub
Private Sub CasiEsenzione_Click()
If CasiEsenzione.Col < 1 Then
 With CasiEsenzione
  Edit.Visible = True: Set Edit.Container = .Container
  Edit.Move .CellLeft + .Left + 40, .CellTop + .Top + 15, .CellWidth - 70, .CellHeight - 15
  Edit.Text = .Text: Edit.SetFocus
 End With
ElseIf CasiEsenzione.Rows > 2 And CasiEsenzione.Row < CasiEsenzione.Rows - 1 Then
 Dim Scelta%
 Scelta = MsgBox("Cancellare questa voce ?" & vbNewLine, vbYesNo + vbQuestion, "Fattura Pro")
 If Scelta = vbYes Then
  If Not SezioneEsenzioneIva Is Nothing Then
   If CasiEsenzione.Row <= SezioneEsenzioneIva.childNodes.Length Then
    SezioneEsenzioneIva.removeChild SezioneEsenzioneIva.childNodes(CasiEsenzione.Row - 1)
   End If
   If SezioneEsenzioneIva.childNodes.Length = 0 And (Not SezioneEsenzioneIva.parentNode Is Nothing) Then
    DocImpostazioni.documentElement.removeChild SezioneEsenzioneIva
   End If
  End If
  CasiEsenzione.RemoveItem CasiEsenzione.Row: DatiModificati = True
 End If
End If
End Sub
Private Sub CasiEsenzione_Scroll()
Edit.Visible = False
End Sub
Private Sub Edit_Change()
DatiModificati = True: CasiEsenzione.Text = Edit.Text
End Sub
Private Sub Edit_KeyPress(KeyAscii As Integer)
If CasiEsenzione.Row = CasiEsenzione.Rows - 1 And KeyAscii <> 0 Then
 CasiEsenzione.AddItem "": CasiEsenzione.Col = 1
 CasiEsenzione.CellPictureAlignment = 4: Set CasiEsenzione.CellPicture = ImgCancella
 CasiEsenzione.Col = 0
End If
End Sub
Private Sub SalvaVarie_Click()
If DatiModificati Then
 Dim CasiEsenzioneIva As MSXML2.IXMLDOMNodeList, CasoEsenzione As MSXML2.IXMLDOMNode
 Dim attr, r&, msg$, valcampo$
 Set CasiEsenzioneIva = DocImpostazioni.selectNodes("/Impostazioni/EsenzioneIva/CasiEsenzione")
 
 For r = 1 To CasiEsenzione.Rows - 2
  valcampo = CasiEsenzione.TextMatrix(r, 0)
  If valcampo = "" Then
    e = True: msg = "Descrizione è un campo obbligatorio !": Exit For
  End If
  If r > CasiEsenzioneIva.Length Then
   Set CasoEsenzione = DocImpostazioni.createElement("CasoEsenzione")
   Set attr = DocImpostazioni.createAttribute("desc"): attr.Text = CasiEsenzione.TextMatrix(r, 0)
   CasoEsenzione.Attributes.setNamedItem attr: SezioneEsenzioneIva.appendChild CasoEsenzione
   If SezioneEsenzioneIva.parentNode Is Nothing Then DocImpostazioni.documentElement.appendChild SezioneEsenzioneIva
  Else
   Set CasoEsenzione = CasiEsenzioneIva(r - 1): Set attr = CasoEsenzione.Attributes.getNamedItem("num")
   attr.Text = CasiEsenzione.TextMatrix(r, 0)
  End If
 Next r
 
 If e Then
  CasiEsenzione.Row = r: CasiEsenzione.Col = 0
  With CasiEsenzione
   Edit.Visible = True: Set Edit.Container = .Container
   Edit.Move .CellLeft + .Left + 40, .CellTop + .Top + 15, .CellWidth - 70, .CellHeight - 15
   Edit.Text = .Text: Edit.SetFocus
  End With
  MsgBox msg, vbExclamation, "Fattura Pro": Exit Sub
 End If
 
 If CartellaRemotaFtp.Text = "" Then
  MsgBox "Cartella Remota non può essere vuoto !", vbExclamation, "Fattura Pro"
  Exit Sub
 End If
 
 Dim CartellaDatiRemota As MSXML2.IXMLDOMNode
 Set CartellaDatiRemota = DocImpostazioni.selectSingleNode("/Impostazioni/CartellaDatiRemota")
 If CartellaDatiRemota Is Nothing Then
  Set CartellaDatiRemota = DocImpostazioni.createElement("CartellaDatiRemota")
  Set attr = DocImpostazioni.createAttribute("percorso"): attr.Text = CartellaRemotaFtp.Text
  CartellaDatiRemota.Attributes.setNamedItem attr
 Else
  Set attr = CartellaDatiRemota.Attributes.getNamedItem("percorso")
  attr.Text = CartellaRemotaFtp.Text
 End If

 DocImpostazioni.save "Impostazioni.xml": DatiModificati = False
End If
End Sub
Private Sub Selezione_Click()
If Selezionato And AnniPresenti.ListIndex <> -1 And CartellaBackup <> "" Then
 Dim DocumentiBackup As MSXML2.IXMLDOMNodeList, SchedaBackup As IXMLDOMElement, _
 outstream As Scripting.TextStream, outstream2 As Scripting.TextStream, DaitXml$
 StatoBackup.Caption = "Backup in corso, attendere prego ..."
 If RdbNote.Value Or RdbFattureClienti.Value Then
  If Not fso.FolderExists("Backup") Then fso.CreateFolder "Backup"
  If Not fso.FolderExists(Path & "Note") Then fso.CreateFolder CartellaBackup & "Note"
  Set DocumentiBackup = DocNote.selectNodes("/Note/Nota[substring(@data,7,4)='" & _
  AnniPresenti.Text & "']")
  Set outstream = fso.CreateTextFile(CartellaBackup & "Note\Note_" & AnniPresenti.Text & ".xml")
  Set outstream2 = fso.CreateTextFile("Backup\Note_" & AnniPres.Text & ".xml")
  outstream.Write "<Note>"
  For i = 0 To DocumentiBackup.Length - 1
   DatiXml = Replace(DocumentiBackup(i).xml, "€", "&#8364;")
   outstream.Write DatiXml: outstream2.Write DatiXml
   DocNote.documentElement.removeChild DocumentiBackup(i): DoEvents
  Next i
  outstream.Write "</Note>": outstream2.Write "</Note>"
  outstream.Close: outstream2.Close: DocNote.save "Note.xml"
  If Not fso.FolderExists(CartellaBackup & "FattureClienti") Then
   fso.CreateFolder CartellaBackup & "FattureClienti"
  End If
  DocFattureClienti.async = False: DocFattureClienti.Load "FattureClienti.xml"
  Set DocumentiBackup = FattureClientiDoc.selectNodes("/FattureClienti/FatturaCliente[substring(@data," _
  & "7,4)='" & AnniPresenti.Text & "']")
  Set outstream = fso.CreateTextFile(CartellaBackup & "FattureClienti\FattureClienti_" & _
  AnniPresenti.Text & ".xml")
  Set outstream2 = fso.CreateTextFile("Backup\FattureClienti_" & AnniPresenti.Text & ".xml")
  outstream.Write "<FattureClienti>": outstream2.Write "<FattureClienti>"
  For i = 0 To DocumentiBackup.Length - 1
   DatiXml = Replace(DocumentiBackup(i).xml, "€", "&#8364;")
   outstream.Write DatiXml: outstream2.Write DatiXml
   DocFattureClienti.documentElement.removeChild DocumentiBackup(i): DoEvents
  Next i
  outstream.Write "</FattureClienti>": outstream2.Write "</FattureClienti>"
  outstream.Close: outstream2.Close: DocFattureClienti.save "FattureClienti.xml"
 Else
  If Not fso.FolderExists(CartellaBackup & "FattureFornitori") Then fso.CreateFolder CartellaBackup & _
  "FattureFornitori"
  Set DocumentiBackup = DocFattureFornitori.selectNodes("/FattureFornitori/FatturaFornitore[substring" _
  & "(@data,7,4)='" & AnniPresenti.Text & "']")
  Set outstream = fso.CreateTextFile(CartellaBackup & "FattureFornitori\FattureFornitori_" _
  & AnniPresenti.Text & ".xml")
  Set outstream2 = fso.CreateTextFile("Backup\FattureFornitori_" & AnniPresenti.Text & ".xml")
  outstream.Write "<FattureFornitori>": outstream2.Write "<FattureFornitori>"
  For i = 0 To DocumentiBackup.Length - 1
   DatiXml = Replace(DocumentiBackup(i).xml, "€", "&#8364;")
   outstream.Write DatiXml: outstream2.Write DatiXml
   DocFattureFornitori.documentElement.removeChild DocumentiBackup(i): DoEvents
  Next i
  outstream.Write "</FattureFornitori>": outstream2.Write "</FattureFornitori>"
  outstream.Close: outstream2.Close: DocFattureFornitori.save "FattureFornitori.xml"
 End If
 StatoBackup.Caption = "Backup completato"
Else
 MsgBox "Attenzione, per effettuare il backup su selezione devi selezionare un tipo di " _
 & "documento e scegliere un anno ed un'unità disco in cui memorizzare i dati", vbExclamation, _
 "Fattura Pro"
End If
End Sub
Private Sub Semplice_Click()
If CartellaBackup <> "" Then
 If App.Path <> CartellaBackup Then
  Dim files As Variant, copiati As Integer, Data As String, instream As Scripting.TextStream, _
  outstream As Scripting.TextStream, AvvisoBackup As String, FileMancanti As Boolean
  NomiFile = Array("FatturaPro.exe", "FatturaPro.exe.manifest", "Articoli.xml", "Clienti.xml", _
  "Note.xml", "FattureClienti.xml", "NoteCredito.xml", "Fornitori.xml", "FattureFornitori.xml", _
  "Dipendenti.xml", "Stipendi.xml", "AP.xml", "Fisco.xml", "Tributi.xml")
  AvvisoBackup = "Non è stato possibile effettuare il backup dei seguenti file perchè non sono " _
  & "stati trovati:" & vbCrLf
  If Not fso.FolderExists(CartellaBackup) Then fso.CreateFolder CartellaBackup
   StatoBackup.Caption = "Stato Backup: " & copiati & " di " & UBound(NomiFile) + 1 & " files copiati"
  For i = 0 To UBound(NomiFile)
   If fso.FileExists(NomiFile(i)) Then
    Set instream = fso.OpenTextFile(NomiFile(i))
    Set outstream = fso.CreateTextFile(CartellaBackup & NomiFile(i))
    While Not instream.AtEndOfStream
     Data = instream.Read(50): outstream.Write Data: DoEvents
    Wend
    instream.Close: outstream.Close: copiati = copiati + 1
   Else
    AvvisoBackup = AvvisoBackup & NomiFile(i) & vbCrLf: FileMancanti = True
   End If
   StatoBackup.Caption = "Stato Backup: " & copiati & " di " & UBound(NomiFile) + 1 & " files copiati"
  Next i
  If FileMancanti Then MsgBox AvvisoBackup, vbExclamation, "Fattura Pro"
 Else
  MsgBox "Il backup non può essere effettuato quando la directory corrente corrisponde a quella di backup !", _
  vbExclamation, "Fattura Pro"
 End If
Else
 MsgBox "Attenzione, per effettuare il backup semplice devi selezionare un'unità in cui " _
 & "memorizzare i dati", vbExclamation, "Fattura Pro"
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Dim ElencoSchede: ElencoSchede = Array(SchedaBackup, SchedaDitta, SchedaVarie)
For i = 0 To UBound(ElencoSchede)
 If SSTab1.Tab = i Then
  ElencoSchede(i).Visible = True
 Else
  ElencoSchede(i).Visible = False
 End If
Next i
End Sub
Private Sub ElencoDischi_Click()
If ElencoDischi.ListIndex <> -1 And ElencoDischi.Text <> DiscoCorr Then
 CartellaBackup = Mid(ElencoDischi.Text, 1, InStr(ElencoDischi.Text, ":")) & "\Backup Fattura Pro\"
 DiscoCorr = ElencoDischi.Text
 If RdbNote.Value Or RdbFattureClienti.Value Then CaricaAnni 0
 If RdbFattureFornitori.Value Then CaricaAnni 1
End If
End Sub
Private Sub VisualizzaDati_Click()
Dim DatiBackup As New Collection
If RdbNote.Value And (Not BackupNote Is Nothing) Then
 For i = 0 To BackupNote.documentElement.childNodes.Length
  DatiBackup.Add BackupNote.documentElement.childNodes(i)
 Next i
 Note.ImpostaFiltroDati DatiBackup
 If Not FormVisibile("Note") Then
  Note.Show
 Else: Note.CaricaFiltro
 End If
ElseIf RdbFattureClienti.Value And (Not BackupFattureClienti Is Nothing) Then
 For i = 0 To BackupFattureClienti.documentElement.childNodes.Length
  DatiBackup.Add BackupFattureClienti.documentElement.childNodes(i)
 Next i
 FattureClienti.ImpostaFiltroDati DatiBackup, BackupNote
 If Not FormVisibile("FattureClienti") Then
  FattureClienti.Show
 Else: FattureClienti.CaricaFiltro
 End If
ElseIf RdbFattureFornitori.Value And (Not BackupFattureFornitori Is Nothing) Then
 For i = 0 To BackupFattureFornitori.documentElement.childNodes.Length
  DatiBackup.Add BackupFattureFornitori.documentElement.childNodes(i)
 Next i
 FattureFornitori.ImpostaFiltroDati DatiBackup
 If Not FormVisibile("FattureFornitori") Then
  FattureFornitori.Show
 Else: FattureFornitori.CaricaFiltro
 End If
End If
End Sub
Public Sub CaricaAnni(Sorgente As Integer)
Dim TipoDoc As String, Anno As Integer, DataStr As String, ChkData As MSXML2.IXMLDOMNodeList, _
trovati As Boolean

If Sorgente = 0 Then
 DataStr = DocNote.documentElement.firstChild.Attributes.getNamedItem("data").Text
 TipoDoc = "Note"
Else
 DataStr = DocFattureFornitori.documentElement.firstChild.Attributes.getNamedItem("data").Text
 TipoDoc = "FattureFornitori"
End If
Anno = CInt(Split(DataStr, "-")(2)): AnniPres.Clear
If Anno <= AnnoCorr - 1 Then
 AnniPresenti.AddItem CStr(Anno): Anno = Anno + 1
 While Anno <= AnnoCorr - 1
  Set ChkData = DocNote.selectNodes("/Note/Nota[substring(@data,7,4)=" & Anno & "][position()=1]")
  If ChkData.Length > 0 Then AnniPresenti.AddItem Anno
  Anno = Anno + 1
 Wend
End If

PercorsiBackup.Clear
If ElencoDischi.ListIndex <> -1 Then
 Dim f As Scripting.File, d As Scripting.Folder, NomeTipo As String, AnnoNum As String
 
 AnniArchiviati.Clear
 If fso.FolderExists(ElencoDischi.Text & "\Backup Fattura Pro\" & TipoDoc) Then
  Set d = fso.GetFolder(ElencoDischi.Text & "\Backup Fattura Pro\" & TipoDoc)
  For Each f In d.files
   If UBound(Split(f.Name, "_")) = 1 Then
    NomeTipo = Split(f.Name, "_")(0): AnnoNum = Split(f.Name, "_")(1)
    If NomeTipo = TipoDoc And IsNumeric(AnnoNum) Then
     AnniArchiviati.AddItem AnnoNum: trovati = True
    End If
   End If
  Next f
  If trovati Then PercorsiBackup.AddItem ElencoDischi.Text & "\Backup Fattura Pro" & TipoDoc
 End If
End If

If PercorsiBackup.ListCount = 0 Then
 AnniArchiviati.Clear
Else: trovati = False
End If

If fso.FolderExists("Backup") Then
 Set d = fso.GetFolder("Backup")
 For Each f In d.files
  If UBound(Split(f.Name, "_")) = 1 Then
   NomeTipo = Split(f.Name, "_")(0): AnnoNum = Split(f.Name, "_")(1)
   If NomeTipo = TipoDoc And IsNumeric(AnnoNum) And _
   SendMessage(AnniArchiviati.hWnd, LB_FINDSTRING, -1, ByVal AnnoNum) = -1 Then
    AnniArchiviati.AddItem AnnoNum: trovati = True
   End If
  End If
 Next f
 If trovati Then PercorsiBackup.AddItem "Backup"
End If
End Sub
