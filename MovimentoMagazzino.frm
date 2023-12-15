VERSION 5.00
Begin VB.Form MovimentoMagazzino 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MovimentoMagazzino.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnAnnulla 
      Caption         =   "Annulla"
      CausesValidation=   0   'False
      Height          =   465
      Left            =   3120
      TabIndex        =   12
      Top             =   3930
      Width           =   1365
   End
   Begin VB.CommandButton BtnInsArt 
      Height          =   345
      Left            =   5220
      MaskColor       =   &H00FFFFFF&
      Picture         =   "MovimentoMagazzino.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1110
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.ComboBox CmbArticoli 
      Height          =   345
      Left            =   300
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1110
      Width           =   4875
   End
   Begin VB.CommandButton BtnSalva 
      Caption         =   "Salva"
      Height          =   465
      Left            =   1650
      TabIndex        =   9
      Top             =   3930
      Width           =   1185
   End
   Begin VB.TextBox TxtRifMov 
      Height          =   330
      Left            =   300
      TabIndex        =   8
      Top             =   3210
      Width           =   4815
   End
   Begin VB.TextBox TxtCliFor 
      Height          =   330
      Left            =   300
      TabIndex        =   6
      Top             =   2520
      Width           =   4815
   End
   Begin VB.TextBox TxtQnt 
      Height          =   315
      Left            =   300
      TabIndex        =   4
      Top             =   1830
      Width           =   1695
   End
   Begin VB.TextBox TxtData 
      Height          =   330
      Left            =   300
      TabIndex        =   1
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label LblRifMov 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rif. Movimento:"
      Height          =   225
      Left            =   300
      TabIndex        =   7
      Top             =   2940
      Width           =   1275
   End
   Begin VB.Label LblCliFor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente / Fornitore:"
      Height          =   225
      Left            =   285
      TabIndex        =   5
      Top             =   2250
      Width           =   1455
   End
   Begin VB.Label LblQnt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantità:"
      Height          =   225
      Left            =   285
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label LblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   225
      Left            =   300
      TabIndex        =   2
      Top             =   165
      Width           =   405
   End
   Begin VB.Label LblArticolo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Articolo:"
      Height          =   225
      Left            =   285
      TabIndex        =   0
      Top             =   840
      Width           =   675
   End
End
Attribute VB_Name = "MovimentoMagazzino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMovimento As ADODB.Recordset, rsArticoli As ADODB.Recordset, ElencoControlli As Variant, _
ContControlli As Variant, DescControlli As Variant
Dim IdMov As Long, DatiMod As Boolean
Public RecordModificato As Boolean, Tipo As String
Public Sub CaricaMovimento(ByVal IdMovimento As Long)
IdMov = IdMovimento
Set rsMovimento = New ADODB.Recordset
rsMovimento.Open "SELECT * FROM MovimentiMagazzino WHERE IdMov = " & IdMovimento, conn, adOpenDynamic, adLockOptimistic
TxtData.Text = Format$(rsMovimento("Data"), "dd-mm-yyyy")
rsArticoli.Find "Id = " & rsMovimento("IdArticolo"), , adSearchForward, adBookmarkFirst
CmbArticoli.ListIndex = SendMessage(CmbArticoli.hWnd, CB_FINDSTRING, -1, ByVal CStr(rsArticoli("Descr")))
TxtQnt.Text = rsMovimento("Qnt")
TxtCliFor.Text = rsMovimento("CliFor")
TxtRifMov.Text = rsMovimento("RifMov")
DatiMod = False
End Sub
Public Sub NuovoMovimento()
Set rsMovimento = New ADODB.Recordset
rsMovimento.Open "MovimentiMagazzino", conn, adOpenDynamic, adLockOptimistic
IdMov = 0
TxtData.Text = ""
CmbArticoli.ListIndex = -1
TxtQnt.Text = ""
TxtCliFor.Text = ""
TxtRifMov.Text = ""
DatiMod = False
End Sub
Private Sub BtnAnnulla_Click()
Unload Me
End Sub
Private Sub BtnInsArt_Click()
Dim InsArticolo As New InsArticoloMag
InsArticolo.Show vbModal
If InsArticolo.ArticoloAggiunto Then
 CaricaArticoli
End If
End Sub
Private Sub BtnSalva_Click()
If DatiMod Then
 Dim re As New RegExp
 re.Global = True
 For i = 0 To UBound(ElencoControlli)
  re.Pattern = ContControlli(i)
  If Not re.Test(Me(ElencoControlli(i))) Then
   e = True
  Exit For
  End If
  If ElencoControlli(i) = "TxtData" Then
   If Not IsDate(Me(ElencoControlli(i))) Then
    e = True
   Else
    TxtData.Text = Format$(TxtData.Text, "dd-mm-yyyy")
   End If
  End If
 Next
 If e Then
  If TypeOf Me(ElencoControlli(i)) Is VB.TextBox Then
   Me(ElencoControlli(i)).SetFocus
   Me(ElencoControlli(i)).SelStart = 0
   Me(ElencoControlli(i)).SelLength = Len(Me(ElencoControlli(i)))
  End If
  MsgBox "Attenzione, " & DescControlli(i) & " contiene un valore non valido !", vbExclamation, _
  "Fattura Pro"
  Exit Sub
 End If
 If IdMov <> 0 Then
  If rsMovimento("IdArticolo") <> CmbArticoli.ItemData(CmbArticoli.ListIndex) Then
   rsArticoli.Open "SELECT * FROM Articoli WHERE Id = " & rsMovimento("IdArticolo"), conn, adOpenDynamic, _
   adLockOptimistic
   rsArticoli("QntDisp") = IIf(Tipo = "Carico", rsArticoli("QntDisp") - CDbl(TxtQnt.Text), _
   rsArticoli("QntDisp") + CDbl(TxtQnt.Text))
  End If
 End If
 If rsArticoli.State <> adStateClosed Then
  rsArticoli.Close
 End If
 rsArticoli.Open "SELECT * FROM Articoli WHERE Id = " & CmbArticoli.ItemData(CmbArticoli.ListIndex), conn, adOpenDynamic, adLockOptimistic
 rsArticoli.MoveFirst
 If IdMov = 0 Then
  If Tipo = "Carico" Then
   rsArticoli("QntDisp") = rsArticoli("QntDisp") + CDbl(TxtQnt.Text)
  Else
   rsArticoli("QntDisp") = rsArticoli("QntDisp") - CDbl(TxtQnt.Text)
  End If
 Else
  Dim Diff#
  Diff = CDbl(TxtQnt.Text) - rsMovimento("Qnt")
  If Tipo = "Carico" Then
   rsArticoli("QntDisp") = rsArticoli("QntDisp") + Diff
  Else
   rsArticoli("QntDisp") = rsArticoli("QntDisp") - Diff
  End If
 End If
 rsArticoli.Update
 If rsArticoli("QntDisp") < 0 Then
  MsgBox "Attenzione, i dati inseriti hanno prodotto una quantità disponibile negativa !", vbExclamation, _
  "Fattura Pro"
 End If
 rsArticoli.Close
 If IdMov = 0 Then
  rsMovimento.AddNew
 End If
 rsMovimento("Data") = TxtData.Text
 rsMovimento("IdArticolo") = CmbArticoli.ItemData(CmbArticoli.ListIndex)
 rsMovimento("Qnt") = TxtQnt.Text
 rsMovimento("CliFor") = TxtCliFor.Text
 rsMovimento("TipoMov") = Tipo
 rsMovimento("RifMov") = TxtRifMov.Text
 rsMovimento.Update
 RecordModificato = True
End If
Unload Me
End Sub
Private Sub CmbArticoli_Click()
DatiMod = True
End Sub
Private Sub Form_Load()
Me.Caption = "Dati " & Tipo
Set rsArticoli = New ADODB.Recordset
CaricaArticoli
ElencoControlli = Array("TxtData", "CmbArticoli", "TxtQnt", "TxtCliFor", "TxtRifMov")
ContControlli = Array("^[0-9][0-9]?-[0-9][0-9]?-[0-9]{4}$", "^.+$", "^[1-9][0-9]*$", _
"^.+$", "^.+$", "^.+$")
DescControlli = Array("Data", "Descrizione", "Quantità", "Cliente / Fornitore", "Rif. Movimento")
End Sub
Private Sub CaricaArticoli()
If rsArticoli.State <> adStateClosed Then
 rsArticoli.Close
End If
rsArticoli.Open "SELECT * FROM Articoli ORDER BY Descr ASC", conn, adOpenDynamic
If Not rsArticoli.EOF Then
 rsArticoli.MoveFirst
 CmbArticoli.Clear
 While Not rsArticoli.EOF
  CmbArticoli.AddItem rsArticoli("Descr")
  CmbArticoli.ItemData(CmbArticoli.NewIndex) = rsArticoli("Id")
  rsArticoli.MoveNext
 Wend
End If
End Sub
Private Sub TxtCliFor_Change()
DatiMod = True
End Sub
Private Sub TxtData_Change()
DatiMod = True
End Sub
Private Sub TxtData_KeyPress(KeyAscii As Integer)
If InStr("0123456789-" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub TxtQnt_Change()
DatiMod = True
End Sub
Private Sub TxtQnt_KeyPress(KeyAscii As Integer)
If InStr("0123456789" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub TxtRifMov_Change()
DatiMod = True
End Sub
