VERSION 5.00
Begin VB.Form SelezioneDitta 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selezione Ditta"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6540
   Icon            =   "SelezioneDitta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbPag 
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
      ItemData        =   "SelezioneDitta.frx":4072
      Left            =   330
      List            =   "SelezioneDitta.frx":4082
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5760
      Width           =   2010
   End
   Begin VB.ListBox PopupDitte 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   3675
      TabIndex        =   21
      Top             =   375
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox TxtStato 
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
      Left            =   330
      TabIndex        =   11
      Top             =   4320
      Width           =   3585
   End
   Begin VB.TextBox TxtCodFisc 
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
      Left            =   330
      TabIndex        =   12
      Top             =   5040
      Width           =   2475
   End
   Begin VB.TextBox TxtPartIva 
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
      Left            =   2985
      TabIndex        =   13
      Top             =   5040
      Width           =   2265
   End
   Begin VB.CommandButton BtnAnnulla 
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3375
      TabIndex        =   18
      Top             =   6450
      Width           =   1290
   End
   Begin VB.CommandButton BtnOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1965
      TabIndex        =   17
      Top             =   6450
      Width           =   1125
   End
   Begin VB.TextBox TxtPaese 
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
      Left            =   1770
      TabIndex        =   9
      Top             =   3615
      Width           =   3450
   End
   Begin VB.TextBox TxtProv 
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
      Left            =   5415
      TabIndex        =   10
      Top             =   3615
      Width           =   780
   End
   Begin VB.TextBox TxtCAP 
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
      Left            =   330
      TabIndex        =   8
      Top             =   3615
      Width           =   1245
   End
   Begin VB.TextBox TxtIndirizzo 
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
      Left            =   315
      TabIndex        =   7
      Top             =   2880
      Width           =   5445
   End
   Begin VB.TextBox TxtNome 
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
      Left            =   300
      TabIndex        =   5
      Top             =   2160
      Width           =   5865
   End
   Begin VB.OptionButton OptNuovaDitta 
      BackColor       =   &H0033CCFF&
      Caption         =   "Crea nuova ditta"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   255
      TabIndex        =   3
      Top             =   1470
      Width           =   1665
   End
   Begin VB.TextBox TxtDitta 
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
      Left            =   285
      TabIndex        =   2
      Top             =   870
      Width           =   5745
   End
   Begin VB.OptionButton OptDittaEsistente 
      BackColor       =   &H0033CCFF&
      Caption         =   "Usa Ditta esistente "
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
      Left            =   255
      TabIndex        =   0
      Top             =   195
      Width           =   1785
   End
   Begin VB.Label LblPag 
      AutoSize        =   -1  'True
      BackColor       =   &H0033CCFF&
      Caption         =   "Pagamento:"
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
      TabIndex        =   24
      Top             =   5505
      Width           =   960
   End
   Begin VB.Label LblStato 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stato:"
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
      Left            =   345
      TabIndex        =   22
      Top             =   4065
      Width           =   450
   End
   Begin VB.Label LblCodFisc 
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
      Left            =   315
      TabIndex        =   20
      Top             =   4785
      Width           =   1170
   End
   Begin VB.Label LblPartitaIva 
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
      Left            =   2985
      TabIndex        =   19
      Top             =   4785
      Width           =   825
   End
   Begin VB.Label LblPaese 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Città:"
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
      Left            =   1755
      TabIndex        =   16
      Top             =   3360
      Width           =   420
   End
   Begin VB.Label LblProv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia"
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
      Left            =   5415
      TabIndex        =   15
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label LblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CAP:"
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
      Top             =   3360
      Width           =   390
   End
   Begin VB.Label LbIndirizzo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indirizzo:"
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
      TabIndex        =   6
      Top             =   2625
      Width           =   705
   End
   Begin VB.Label LblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
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
      Left            =   300
      TabIndex        =   4
      Top             =   1905
      Width           =   540
   End
   Begin VB.Label LblDitta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ditta:"
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
      TabIndex        =   1
      Top             =   615
      Width           =   420
   End
End
Attribute VB_Name = "SelezioneDitta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsDitte As ADODB.Recordset
Public UsaDittaEsistente As Boolean
Public FormChiamante As Form
Dim archivio As String
Private Sub BtnAnnulla_Click()
Unload Me
End Sub
Private Sub BtnOk_Click()
If OptDittaEsistente.Value Then
 If TxtDitta.Text = "" Then
  MsgBox "Selezionare una ditta !", vbExclamation, "Fattura Pro"
  Exit Sub
 End If
Else
 If TxtNome.Text = "" Then
  MsgBox "Nome è un campo obbligatorio !", vbExclamation, "Fattura Pro"
  Exit Sub
 End If
 If TxtIndirizzo.Text = "" Then
  MsgBox "Indirizzo è un campo obbligatorio !", vbExclamation, "Fattura Pro"
  Exit Sub
 End If
 If TxtPartIva.Text = "" Then
  MsgBox "Partita IVA è un campo obbligatorio !", vbExclamation, "Fattura Pro"
  Exit Sub
 End If
 If TxtPaese.Text = "" Then
  MsgBox "Città è un campo obbligatorio !", vbExclamation, "Fattura Pro"
  Exit Sub
 End If
 If TxtProv.Text = "" And TxtStato.Text = "" Then
  MsgBox "Provincia e Stato non possono essere entrambi vuoti !", vbExclamation, "Fattura Pro"
  Exit Sub
 End If
 rsDitte.Filter = ""
 rsDitte.Filter = "Ditta = '" & Replace(TxtNome.Text, "'", "''") & "' AND PartitaIva = '" _
 & TxtPartIva.Text & "'"
 If Not rsDitte.EOF Then
  MsgBox "Attenzione, ditta già esistente!", vbExclamation, "Fattura Pro"
  rsDitte.MoveNext
  Exit Sub
 End If
 Dim rsNuovoId As New ADODB.Recordset
 rsDitte.AddNew
 rsDitte("Ditta") = TxtNome.Text
 rsDitte("Indirizzo") = TxtIndirizzo.Text
 rsDitte("Cap") = TxtCap.Text
 rsDitte("Loc") = TxtPaese.Text
 rsDitte("Prov") = TxtProv.Text
 rsDitte("Stato") = TxtStato.Text
 rsDitte("Codfiscale") = TxtCodFisc.Text
 rsDitte("PartitaIva") = TxtPartIva.Text
 rsDitte("ModPag") = CmbPag.ListIndex
 rsDitte("Rimosso") = False
 rsDitte.Update
 CancellaUgualiRimossi
End If
UsaDittaEsistente = OptDittaEsistente.Value
Unload Me
End Sub
Private Sub CancellaUgualiRimossi()
Dim rsRimossi As New ADODB.Recordset
rsRimossi.Open "SELECT * FROM " & archivio & " WHERE UCASE(Ditta) = '" & UCase(Replace(TxtDitta.Text, "'", "''")) _
& "' AND PartitaIva = '" & TxtPartIva.Text & "'", conn, adOpenDynamic, adLockOptimistic
If Not rsRimossi.EOF Then
 rsRimossi.MoveFirst
 If archivio = "Clienti" Then
  conn.Execute "UPDATE FattureClienti, NoteConsegna, NoteCredito SET FattureClienti.IdDitta = " & rsDitte("Id") _
  & ", NoteConsegna.IdDitta = " & rsDitte("Id") & ", NoteCredito.IdDitta = " & rsDitte("Id") & _
  " WHERE FattureClienti.IdDitta = " & rsRimossi("Id") & " Or NoteConsegna.IdDitta=" & rsRimossi("Id") & " Or " _
  & "NoteCredito.IdDitta=" & rsRimossi("Id")
 Else
  conn.Execute "UPDATE FattureFornitori SET FattureFornitori.IdDitta = " & rsDitte("Id") _
  & " WHERE FattureFornitori.IdDitta = " & rsRimossi("Id")
 End If
 rsRimossi.Delete
End If
rsRimossi.Close
End Sub
Public Property Get Ditta() As String
Ditta = TxtDitta.Text
End Property
Private Sub Form_Load()
Set rsDitte = New ADODB.Recordset
If FormChiamante.Name = "FattureFornitori" Then
 archivio = "Fornitori"
Else
 archivio = "Clienti"
End If
rsDitte.Open "SELECT * FROM " & archivio & " WHERE Rimosso = False", conn, adOpenDynamic, adLockOptimistic
If Not rsDitte.EOF Then
 rsDitte.MoveFirst
 While Not rsDitte.EOF
  PopupDitte.AddItem rsDitte("Ditta")
  PopupDitte.ItemData(PopupDitte.NewIndex) = rsDitte("Id")
  rsDitte.MoveNext
 Wend
Else
 OptDittaEsistente.Enabled = False
End If
CmbPag.ListIndex = 0
End Sub
Private Sub PopupDitte_Click()
If PopupDitte.ListIndex <> -1 Then
 TxtDitta.Text = PopupDitte.List(PopupDitte.ListIndex)
 TxtDitta.Tag = PopupDitte.ItemData(PopupDitte.ListIndex)
 PopupDitte.Visible = False
 rsDitte.Move PopupDitte.ListIndex, adBookmarkFirst
End If
End Sub
Private Sub ModificaCarIns(ControlloTesto As VB.TextBox, Car As String)
Dim PosCursore%
PosCursore = ControlloTesto.SelStart + 1
If PosCursore = 1 Or Mid(ControlloTesto.Text, PosCursore + 1, 1) = "." Then
 Car = UCase(Car)
End If
End Sub
Private Sub TxtCAP_KeyPress(KeyAscii As Integer)
If InStr("0123456789" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub TxtCodFisc_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVXYWZ", UCase(Chr(KeyAscii))) = 0 Then
  KeyAscii = 0
 Else
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End If
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
Dim NumCar As Integer
If KeyAscii >= 32 Then
 NumCar = Len(TxtDitta.Text) + 1
Else
 NumCar = Len(TxtDitta.Text)
End If
If NumCar = 3 Then
 If KeyAscii >= 32 Then
  TxtDitta.Text = TxtDitta.Text & Chr(KeyAscii): KeyAscii = 0
  TxtDitta.SelStart = Len(TxtDitta.Text)
 End If
 rsDitte.Filter = "Ditta LIKE '" & TxtDitta.Text & "*'"
 If Not rsDitte.EOF Then
  PopupDitte.Clear
  rsDitte.MoveFirst
  While Not rsDitte.EOF
   With PopupDitte
   .AddItem rsDitte("Ditta")
   .ItemData(.NewIndex) = rsDitte("Id")
   End With
   rsDitte.MoveNext
  Wend
  PopupDitte.Visible = True
  PopupDitte.Left = TxtDitta.Left
  PopupDitte.Top = TxtDitta.Top + TxtDitta.Height + 45
  PopupDitte.Width = TxtDitta.Width
  PopupDitte.Height = 2500
 End If
ElseIf UCase$(Chr(KeyAscii)) <> LCase$(Chr(KeyAscii)) Then
 Dim Car$: Car = Chr(KeyAscii)
 ModificaCarIns TxtDitta, Car
 KeyAscii = Asc(Car)
End If
End Sub
Private Sub TxtIndirizzo_KeyPress(KeyAscii As Integer)
If UCase$(Chr(KeyAscii)) <> LCase$(Chr(KeyAscii)) Then
 Dim Car$: Car = Chr(KeyAscii)
 ModificaCarIns TxtIndirizzo, Car
 KeyAscii = Asc(Car)
End If
End Sub
Private Sub TxtNome_KeyPress(KeyAscii As Integer)
If UCase$(Chr(KeyAscii)) <> LCase$(Chr(KeyAscii)) Then
 Dim Car$: Car = Chr(KeyAscii)
 ModificaCarIns TxtNome, Car
 KeyAscii = Asc(Car)
End If
End Sub
Private Sub TxtPaese_KeyPress(KeyAscii As Integer)
If UCase$(Chr(KeyAscii)) <> LCase$(Chr(KeyAscii)) Then
 Dim Car$: Car = Chr(KeyAscii)
 ModificaCarIns TxtPaese, Car
 KeyAscii = Asc(Car)
End If
End Sub
Private Sub TxtPartIva_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtPartIva.Text) = 18 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub TxtProv_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtProv.Text) = 2 Then
  KeyAscii = 0
  Exit Sub
 End If
 If InStr("ABCDEFGHIJKLMNOPQRSTUVXYWZ", UCase(Chr(KeyAscii))) = 0 Then
  KeyAscii = 0
 Else
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End If
End Sub
Private Sub TxtStato_KeyPress(KeyAscii As Integer)
If UCase$(Chr(KeyAscii)) <> LCase$(Chr(KeyAscii)) Then
 Dim Car$: Car = Chr(KeyAscii)
 ModificaCarIns TxtStato, Car
 KeyAscii = Asc(Car)
End If
End Sub
