VERSION 5.00
Begin VB.Form RicercaFattureClienti 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ricerca Fatture Clienti"
   ClientHeight    =   3465
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "RicercaFattureClienti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   7950
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
      Height          =   510
      ItemData        =   "RicercaFattureClienti.frx":4072
      Left            =   645
      List            =   "RicercaFattureClienti.frx":4074
      TabIndex        =   9
      Top             =   645
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
      Height          =   960
      ItemData        =   "RicercaFattureClienti.frx":4076
      Left            =   5745
      List            =   "RicercaFattureClienti.frx":4078
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   105
      Width           =   1110
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
      Height          =   540
      Left            =   3225
      TabIndex        =   3
      Top             =   2805
      Width           =   1530
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
      Height          =   330
      Left            =   945
      TabIndex        =   0
      Top             =   105
      Width           =   1500
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
      Left            =   885
      TabIndex        =   1
      Top             =   1620
      Width           =   6780
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
      Height          =   1410
      ItemData        =   "RicercaFattureClienti.frx":407A
      Left            =   3180
      List            =   "RicercaFattureClienti.frx":40A5
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   1815
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
      Left            =   225
      TabIndex        =   7
      Top             =   135
      Width           =   660
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
      Left            =   2655
      TabIndex        =   6
      Top             =   135
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      TabIndex        =   5
      Top             =   1650
      Width           =   600
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
      Left            =   5205
      TabIndex        =   4
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "RicercaFattureClienti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PopupDitte_Click()
If PopupDitte.ListIndex <> -1 Then
 TxtDitta.Text = PopupDitte.Text
 TxtDitta.Tag = PopupDitte.ItemData(PopupDitte.ListIndex)
 PopupDitte.Visible = False
End If
End Sub
Private Sub TxtDitta_Change()
TxtDitta.Tag = ""
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
 rsDitte.Open "SELECT * FROM Clienti WHERE UCASE(Ditta) LIKE '" & UCase(Replace(TxtDitta.Text, "'", "''")) & _
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
Private Sub Form_Load()
Me.Move 1000, 1000
Dim rsElencoAnni As New ADODB.Recordset
rsElencoAnni.Open "SELECT Year(Data) As Anno FROM FattureClienti GROUP BY Year(Data) " _
& "ORDER BY Year(Data) ASC", conn, adOpenDynamic
While Not rsElencoAnni.EOF
 LstAnni.AddItem rsElencoAnni("Anno")
 rsElencoAnni.MoveNext
Wend
rsElencoAnni.Close
End Sub
Private Sub TxtDitta_LostFocus()
PopupDitte.Visible = False
End Sub
Private Sub TxtNumFattura_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtNumFattura.Text) = 6 Then
  KeyAscii = 0: Exit Sub
 End If
 If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End If
End Sub
Private Sub LstMesi_DblClick()
LstMesi.ListIndex = -1
End Sub
Private Sub Ricerca_Click()
Dim ValoriCampi As Variant, EsprRicerca As Variant, StrRicerca As String
ValoriCampi = Array(IdDocumento(TxtNumFattura.Text), TxtDitta.Tag, NumMese(), LstAnni.Text)
EsprRicerca = Array("Mid(IdDoc,6)=val", "IdDitta=val", "Month(Data)=val", "Year(Data)=val")
For i = 0 To UBound(ValoriCampi)
 If Trim(ValoriCampi(i)) <> "" Then
  StrRicerca = IIf(StrRicerca <> "", StrRicerca & " AND ", " WHERE ")
  EsprRicerca(i) = Replace(EsprRicerca(i), "val", ValoriCampi(i))
  StrRicerca = StrRicerca & EsprRicerca(i)
 End If
Next i
If StrRicerca <> "" Then
 StrRicerca = StrRicerca & " ORDER BY IdDoc ASC"
 Dim rsFattureClienti As New ADODB.Recordset
 rsFattureClienti.Open "SELECT * FROM FattureClienti" & StrRicerca, conn, adOpenDynamic, adLockOptimistic
 If Not rsFattureClienti.EOF Then
  FattureClienti.ImpostaFiltroRicerca rsFattureClienti
  If Not FormVisibile("FattureClienti") Then
   FattureClienti.Show
  ElseIf Not FattureClienti.FormBloccata Then FattureClienti.CaricaFiltroRicerca
  Else
   MsgBox "Attenzione, la finestra FattureClienti � impegnata al momento nella modifica o nell'" _
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
