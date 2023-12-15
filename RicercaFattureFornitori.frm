VERSION 5.00
Begin VB.Form RicercaFattureFornitori 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ricerca Fatture Fornitori"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "RicercaFattureFornitori.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   7815
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
      Height          =   960
      ItemData        =   "RicercaFattureFornitori.frx":4072
      Left            =   615
      List            =   "RicercaFattureFornitori.frx":4074
      TabIndex        =   9
      Top             =   585
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.ListBox LstAnni 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      ItemData        =   "RicercaFattureFornitori.frx":4076
      Left            =   5760
      List            =   "RicercaFattureFornitori.frx":4078
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   105
      Width           =   1080
   End
   Begin VB.CommandButton Ricerca 
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
      Height          =   510
      Left            =   3180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3270
      Width           =   1545
   End
   Begin VB.ListBox LstMesi 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      ItemData        =   "RicercaFattureFornitori.frx":407A
      Left            =   3180
      List            =   "RicercaFattureFornitori.frx":40A5
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   1815
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
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   1860
      Width           =   6450
   End
   Begin VB.TextBox TxtNumFattura 
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
      Left            =   930
      TabIndex        =   0
      Top             =   90
      Width           =   1500
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anno:"
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
      Left            =   5220
      TabIndex        =   7
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornitore:"
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
      Left            =   225
      TabIndex        =   6
      Top             =   1890
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mese:"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N� Doc.:"
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
      TabIndex        =   4
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "RicercaFattureFornitori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move 500, 550
Dim rsElencoAnni As New ADODB.Recordset
rsElencoAnni.Open "SELECT Year(Data) As Anno FROM FattureFornitori GROUP BY Year(Data) " _
& "ORDER BY Year(Data) ASC", conn, adOpenDynamic
While Not rsElencoAnni.EOF
 LstAnni.AddItem rsElencoAnni("Anno")
 rsElencoAnni.MoveNext
Wend
rsElencoAnni.Close
End Sub
Private Sub TxtDitta_Change()
TxtDitta.Tag = ""
End Sub
Private Sub TxtDitta_LostFocus()
PopupDitte.Visible = False
End Sub
Private Sub TxtNumFattura_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
 KeyAscii = 0
End If
End Sub
Private Sub LstMesi_DblClick()
LstMesi.ListIndex = -1
End Sub
Private Sub Ricerca_Click()
Dim ValoriCampi As Variant, EsprRicerca As Variant, StrRicerca As String
ValoriCampi = Array(TxtNumFattura.Text, TxtDitta.Tag, NumMese(), LstAnni.Text)
EsprRicerca = Array("IdDoc=val", "IdDitta='val'", "Month(Data)=val", "Year(Data)=val")
For i = 0 To UBound(ValoriCampi)
 If Trim(ValoriCampi(i)) <> "" Then
  StrRicerca = IIf(StrRicerca <> "", StrRicerca & " AND ", " WHERE ")
  EsprRicerca(i) = Replace(EsprRicerca(i), "val", ValoriCampi(i))
  StrRicerca = StrRicerca & EsprRicerca(i)
 End If
Next i
If StrRicerca <> "" Then
 StrRicerca = StrRicerca & " ORDER BY IdDoc ASC"
 Dim rsFattureFornitori As New ADODB.Recordset
 rsFattureFornitori.Open "SELECT * FROM FattureFornitori " & StrRicerca, conn, adOpenDynamic, adLockOptimistic
 If Not rsFattureFornitori.EOF Then
  FattureFornitori.ImpostaFiltroRicerca rsFattureFornitori
  If Not FormVisibile("FattureFornitori") Then
   FattureFornitori.Show
  ElseIf Not FattureFornitori.FormBloccata Then FattureFornitori.CaricaFiltroRicerca
  Else
   MsgBox "Attenzione, la finestra FattureFornitori � impegnata al momento nella modifica o nell'" _
   & "inserimento di un documento." & vbCrLf & "Prima di eseguire la ricerca completare " _
   & "l'operazione in corso !", vbExclamation, "Fattura Pro"
  End If
 Else
  MsgBox "Non � stato trovato nessun documento !", vbExclamation, "Fattura Pro"
 End If
Else
 MsgBox "Attenzione, tutti i campi di ricerca sono vuoti !", vbExclamation, "Fattura Pro"
End If
End Sub
Private Function NumMese() As String
Dim NumeroMese%: NumeroMese = LstMesi.ListIndex + 1
If NumeroMese <> 0 Then
 NumMese = NumeroMese
End If
End Function
Private Sub PopupDitte_Click()
If PopupDitte.ListIndex <> -1 Then
 TxtDitta.Text = PopupDitte.Text
 TxtDitta.Tag = PopupDitte.ItemData(PopupDitte.ListIndex)
 PopupDitte.Visible = False
End If
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
Dim NumCar As Integer
NumCar = Len(TxtDitta.Text) + IIf(KeyAscii <> 8, 1, -1)
Dim PosCursore%
PosCursore = TxtDitta.SelStart + 1
If PosCursore = 1 Or Mid(TxtDitta.Text, PosCursore + 1, 1) = "." Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
ElseIf Mid(TxtDitta.Text, PosCursore - 1, 1) = " " Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
If NumCar >= 3 Then
 Dim rsDitte As New ADODB.Recordset
 If KeyAscii >= 32 Then
  TxtDitta.Text = TxtDitta.Text & Chr(KeyAscii): KeyAscii = 0
 End If
 TxtDitta.SelStart = Len(TxtDitta.Text)
 rsDitte.Open "SELECT * FROM Fornitori WHERE UCASE(Ditta) LIKE '" & UCase(Replace(TxtDitta.Text, "'", "''")) & _
 "%' ORDER BY Ditta ASC", conn, adOpenDynamic
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
  PopupDitte.Height = Me.ScaleHeight - (TxtDitta.Top + TxtDitta.Height)
 Else
  PopupDitte.Visible = False
 End If
 rsDitte.Close
Else
 PopupDitte.Visible = False
End If
End Sub
