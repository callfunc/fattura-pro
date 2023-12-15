VERSION 5.00
Begin VB.Form RicercaNoteCredito 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ricerca Note Credito"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   Icon            =   "RicercaNoteCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7095
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
      ItemData        =   "RicercaNoteCredito.frx":4072
      Left            =   3225
      List            =   "RicercaNoteCredito.frx":409D
      TabIndex        =   5
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
      Height          =   330
      Left            =   915
      TabIndex        =   4
      Top             =   1695
      Width           =   5745
   End
   Begin VB.TextBox TxtNumNota 
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
      Left            =   990
      TabIndex        =   3
      Top             =   105
      Width           =   1500
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
      Left            =   2835
      TabIndex        =   2
      Top             =   2985
      Width           =   1530
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
      ItemData        =   "RicercaNoteCredito.frx":410B
      Left            =   5805
      List            =   "RicercaNoteCredito.frx":410D
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   945
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
      Height          =   735
      ItemData        =   "RicercaNoteCredito.frx":410F
      Left            =   360
      List            =   "RicercaNoteCredito.frx":4111
      TabIndex        =   0
      Top             =   2190
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   5265
      TabIndex        =   9
      Top             =   120
      Width           =   480
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
      Left            =   255
      TabIndex        =   8
      Top             =   1725
      Width           =   600
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
      Left            =   2700
      TabIndex        =   7
      Top             =   135
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° Doc.:"
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
      TabIndex        =   6
      Top             =   135
      Width           =   660
   End
End
Attribute VB_Name = "RicercaNoteCredito"
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
Me.Move 500, 800
Dim rsElencoAnni As New ADODB.Recordset
rsElencoAnni.Open "SELECT Year(Data) As Anno FROM NoteCredito GROUP BY Year(Data) " _
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
Private Sub TxtNumNota_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtNumNota.Text) = 6 Then
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
ValoriCampi = Array(IdDocumento(TxtNumNota.Text), TxtDitta.Tag, NumMese(), LstAnni.Text)
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
 Dim rsNoteCredito As New ADODB.Recordset
 rsNoteCredito.Open "SELECT * FROM NoteCredito" & StrRicerca, conn, adOpenDynamic, adLockOptimistic
 If Not rsNoteCredito.EOF Then
  NoteCredito.ImpostaFiltroRicerca rsNoteCredito
  If Not FormVisibile("NoteCredito") Then
   NoteCredito.Show
  ElseIf Not NoteCredito.FormBloccata Then NoteCredito.CaricaFiltroRicerca
  Else
   MsgBox "Attenzione, la finestra è impegnata al momento nella modifica o nell'" _
   & "inserimento di un documento." & vbCrLf & "Prima di eseguire la ricerca completare " _
   & "l'operazione in corso !", vbExclamation, "Fattura Pro"
  End If
 Else
  MsgBox "Non è stato trovato nessun documento !", vbExclamation, "Fattura Pro"
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
