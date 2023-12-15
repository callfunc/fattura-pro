VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form InfoDitta 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Info Ditta"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InfoDitta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCodTrasmittente 
      Height          =   315
      Left            =   285
      TabIndex        =   27
      Top             =   7035
      Width           =   4305
   End
   Begin VB.TextBox TxtIBAN 
      Height          =   330
      Left            =   285
      TabIndex        =   25
      Top             =   5535
      Width           =   4290
   End
   Begin VB.PictureBox PicLogo 
      BackColor       =   &H0033CCFF&
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   300
      ScaleHeight     =   1065
      ScaleWidth      =   1635
      TabIndex        =   23
      Top             =   8010
      Width           =   1635
   End
   Begin VB.TextBox TxtSedeLeg 
      Height          =   315
      Left            =   300
      TabIndex        =   21
      Top             =   2610
      Width           =   5865
   End
   Begin VB.TextBox TxtSedeAz 
      Height          =   315
      Left            =   300
      TabIndex        =   19
      Top             =   1875
      Width           =   5865
   End
   Begin MSComDlg.CommonDialog SfogliaFile 
      Left            =   5565
      Top             =   3990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtLogoAzienda 
      Height          =   315
      Left            =   285
      TabIndex        =   17
      Tag             =   "AO"
      Top             =   6270
      Width           =   5745
   End
   Begin VB.CommandButton BtnCerca 
      Height          =   315
      Left            =   6075
      MaskColor       =   &H00000000&
      Picture         =   "InfoDitta.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6270
      Width           =   735
   End
   Begin VB.CommandButton BtnAnnulla 
      Caption         =   "Annulla"
      Height          =   480
      Left            =   3675
      TabIndex        =   15
      Top             =   9345
      Width           =   1575
   End
   Begin VB.TextBox TxtDitta 
      Height          =   315
      Left            =   300
      TabIndex        =   13
      Top             =   1170
      Width           =   5865
   End
   Begin VB.TextBox TxtEmail 
      Height          =   315
      Left            =   285
      TabIndex        =   11
      Top             =   4080
      Width           =   3210
   End
   Begin VB.CommandButton BtnOk 
      Caption         =   "Ok"
      Height          =   480
      Left            =   1830
      TabIndex        =   10
      Top             =   9345
      Width           =   1545
   End
   Begin VB.TextBox TxtFax 
      Height          =   315
      Left            =   3165
      TabIndex        =   9
      Top             =   3360
      Width           =   2610
   End
   Begin VB.TextBox TxtTel 
      Height          =   315
      Left            =   300
      TabIndex        =   7
      Top             =   3360
      Width           =   2610
   End
   Begin VB.TextBox TxtAzienda 
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   480
      Width           =   5865
   End
   Begin VB.TextBox TxtPartIva 
      Height          =   315
      Left            =   3105
      TabIndex        =   1
      Top             =   4800
      Width           =   2265
   End
   Begin VB.TextBox TxtCodFisc 
      Height          =   315
      Left            =   285
      TabIndex        =   0
      Top             =   4800
      Width           =   2610
   End
   Begin VB.Label LblCodiceTrasmittente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codice Trasmittente (Fattura Elettronica)"
      Height          =   225
      Left            =   285
      TabIndex        =   26
      Top             =   6750
      Width           =   3210
   End
   Begin VB.Label LblIBAN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coordinate Bancarie"
      Height          =   225
      Left            =   285
      TabIndex        =   24
      Top             =   5280
      Width           =   1605
   End
   Begin VB.Label LblSedeLeg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sede Legale:"
      Height          =   225
      Left            =   285
      TabIndex        =   22
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label LblSedeAziendale 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sede Aziendale:"
      Height          =   225
      Left            =   300
      TabIndex        =   20
      Top             =   1620
      Width           =   1230
   End
   Begin VB.Label LblLogoAzienda 
      AutoSize        =   -1  'True
      BackColor       =   &H00006000&
      BackStyle       =   0  'Transparent
      Caption         =   "Logo Ditta:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   285
      TabIndex        =   18
      Top             =   6000
      Width           =   870
   End
   Begin VB.Label LblDitta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ditta:"
      Height          =   225
      Left            =   285
      TabIndex        =   14
      Top             =   915
      Width           =   420
   End
   Begin VB.Label LblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   225
      Left            =   285
      TabIndex        =   12
      Top             =   3825
      Width           =   480
   End
   Begin VB.Label LblFax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      Height          =   225
      Left            =   3165
      TabIndex        =   8
      Top             =   3105
      Width           =   300
   End
   Begin VB.Label LblTel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono:"
      Height          =   225
      Left            =   315
      TabIndex        =   6
      Top             =   3105
      Width           =   750
   End
   Begin VB.Label LblAzienda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Azienda:"
      Height          =   225
      Left            =   285
      TabIndex        =   5
      Top             =   225
      Width           =   1215
   End
   Begin VB.Label LblPartitaIva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Partita Iva:"
      Height          =   225
      Left            =   3105
      TabIndex        =   4
      Top             =   4545
      Width           =   825
   End
   Begin VB.Label LblCodFisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codice Fiscale:"
      Height          =   225
      Left            =   285
      TabIndex        =   3
      Top             =   4545
      Width           =   1170
   End
End
Attribute VB_Name = "InfoDitta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsInfoDitta As ADODB.Recordset
Dim Controlli As Variant, DescControlli As Variant, AttrCampiDitta As Variant
Dim stm As ADODB.Stream, FileLogo As String, TipoFile As String
Private Sub BtnAnnulla_Click()
Dim EsciProg As Boolean, CampiObbligatori As Variant
CampiObbligatori = Array("Azienda", "Ditta", "SedeAziendale", "SedeLegale", "PartitaIVA")
If Not rsInfoDitta.EOF Then
 For i = 0 To UBound(CampiObbligatori)
  If rsInfoDitta(CampiObbligatori(i)) = "" Then
   EsciProg = True: Exit For
  End If
 Next
Else
 EsciProg = True
End If
If EsciProg Then
 MsgBox "Uno o più campi obbligatori sono vuoti. Il programma verrà chiuso", vbExclamation, _
 "Fattura Pro"
 FatturaProMDI.Esci
Else
 Unload Me
End If
End Sub
Private Sub BtnCerca_Click()
SfogliaFile.FileName = ""
SfogliaFile.Filter = "File Immagini (*.jpg;*.jpeg;*.gif)|*.jpg;*.jpeg;*.gif"
SfogliaFile.ShowOpen
If SfogliaFile.FileName <> "" Then
 Dim Logo As StdPicture
 Set Logo = LoadPicture(SfogliaFile.FileName)
 If (Logo.Width / 100) > 92 Or (Logo.Height / 100) > 16 Then
  MsgBox "Attenzione, il file immagine del logo deve avere una larghezza massima di " _
  & "92 mm (340 pixel) ed un'altezza massima di 16 mm (60 pixel) !", vbExclamation, _
  "Fattura Pro"
  TxtLogoAzienda.Text = ""
 Else
  FileLogo = Mid(SfogliaFile.FileName, InStrRev(SfogliaFile.FileName, "\") + 1)
  TipoFile = Mid(FileLogo, InStrRev(FileLogo, ".") + 1)
  TxtLogoAzienda.Text = FileLogo
  rap_img = ScaleX(Logo.Width / 100, vbMillimeters, vbTwips) / ScaleY(Logo.Height / 100, vbMillimeters, vbTwips)

  LargRiquadro = PicLogo.ScaleWidth
  AltRiquadro = PicLogo.ScaleHeight

  If LargRiquadro / AltRiquadro > rap_img Then
   LargRiquadro = rap_img * AltRiquadro
  Else
   AltRiquadro = LargRiquadro / rap_img
  End If

  PicLogo.Cls
  PicLogo.PaintPicture Logo, (PicLogo.ScaleWidth - LargRiquadro) / 2, _
  (PicLogo.ScaleHeight - AltRiquadro) / 2, LargRiquadro, AltRiquadro
 End If
End If
End Sub
Private Sub BtnOk_Click()
Dim e As Boolean, err_msg$
For i = 0 To UBound(Controlli)
 If Right(AttrCampiDitta(i), 1) = "o" And Me(Controlli(i)).Text = "" Then
  e = True
  err_msg = DescControlli(i) & " è un campo obbligatorio !"
  Exit For
 ElseIf Me(Controlli(i)).Text <> "" Then
  If Left(AttrCampiDitta(i), 1) = "a" And IsNumeric(Me(Controlli(i)).Text) Then
   e = True
   err_msg = DescControlli(i) & " contiene un valore non valido"
   Exit For
  ElseIf Left(AttrCampiDitta(i), 1) = "n" And Not IsNumeric(Me(Controlli(i)).Text) Then
   e = True
   err_msg = DescControlli(i) & " contiene un valore non valido"
   Exit For
  End If
 End If
Next i
If TxtLogoAzienda.Text = "" Then
 err_msg = "Inserire il logo della ditta !"
End If

If e Then
 MsgBox err_msg, vbExclamation, "Fattura Pro"
 Exit Sub
End If

If FileLogo <> "" Then
 Set stm = New ADODB.Stream
 'set the type to binary to load the image as a binary stream
 stm.Type = adTypeBinary
 stm.Open
 'Load the content of the picture into the stream object
 stm.LoadFromFile FileLogo
End If

If rsInfoDitta.EOF Then
 rsInfoDitta.AddNew
End If
rsInfoDitta("Azienda") = TxtAzienda.Text
rsInfoDitta("Ditta") = TxtDitta.Text
rsInfoDitta("Tel") = TxtTel.Text
rsInfoDitta("Fax") = TxtFax.Text
rsInfoDitta("Email") = TxtEmail.Text
rsInfoDitta("SedeAziendale") = TxtSedeAz.Text
rsInfoDitta("SedeLegale") = TxtSedeLeg.Text
rsInfoDitta("Codfiscale") = TxtCodFisc.Text
rsInfoDitta("PartitaIva") = TxtPartIva.Text
rsInfoDitta("IBAN") = TxtIBAN.Text
rsInfoDitta("CodTrasmittente") = TxtCodTrasmittente.Text
rsInfoDitta("Logo") = stm.Read
rsInfoDitta("ImgLogo") = TipoFile
rsInfoDitta.Update
Unload Me
End Sub
Private Sub Form_Load()
Controlli = Array("TxtAzienda", "TxtDitta", "TxtSedeAz", "TxtSedeLeg", "TxtTel", _
"TxtFax", "TxtEmail", "TxtCodFisc", "TxtPartIVA", "TxtCodTrasmittente")
DescControlli = Array("Azienda", "Ditta", "Sede Aziendale", "Sede Legale", "Telefono", _
"Fax", "Email", "Codice fiscale", "Partita Iva", "Codice Trasmittente")
AttrCampiDitta = Array("ao", "ao", "ao", "ao", "n", "n", "a", "a", "no", "o")
If rsInfoDitta Is Nothing Then
 Set rsInfoDitta = New ADODB.Recordset
 rsInfoDitta.Open "SELECT * FROM InfoDitta", conn, adOpenDynamic, adLockOptimistic
End If
If Not rsInfoDitta.EOF Then
 rsInfoDitta.MoveFirst
 TxtAzienda.Text = rsInfoDitta("Azienda")
 TxtDitta.Text = rsInfoDitta("Ditta")
 TxtTel.Text = rsInfoDitta("Tel")
 TxtFax.Text = rsInfoDitta("Fax")
 TxtEmail.Text = rsInfoDitta("Email")
 TxtSedeAz.Text = rsInfoDitta("SedeAziendale")
 TxtSedeLeg.Text = rsInfoDitta("SedeLegale")
 TxtCodFisc.Text = rsInfoDitta("CodFiscale")
 TxtPartIva.Text = rsInfoDitta("PartitaIVA")
 TxtIBAN.Text = IIf(IsNull(rsInfoDitta("IBAN")), "", rsInfoDitta("IBAN"))
 TxtCodTrasmittente.Text = IIf(IsNull(rsInfoDitta("CodTrasmittente")), "", rsInfoDitta("CodTrasmittente"))
 Set stm = New ADODB.Stream
 'set the type to binary to load the image as a binary stream
 stm.Type = adTypeBinary
 stm.Open
 'Load the binary image data from the DB into the stream object
 stm.Write rsInfoDitta.Fields("Logo")
 'Check the size of the ado stream to make sure there is data
 If stm.Size > 0 Then
  FileLogo = App.Path & "\logo.tmp"
  'Write the content of the stream object to a file
  'The file will br created if doesn't exists. Otherwise over writes the existing file
  stm.SaveToFile FileLogo, adSaveCreateOverWrite
  'Load the temp Picture into the Image control
  PicLogo.Picture = LoadPicture(FileLogo)
  TipoFile = rsInfoDitta("ImgLogo")
  TxtLogoAzienda.Text = "logo." & TipoFile
  Kill FileLogo
  FileLogo = ""
 End If
End If
End Sub
Private Sub TxtIBAN_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TxtLogoAzienda_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
