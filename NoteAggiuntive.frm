VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form NoteAggiuntive 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elenco Note Aggiuntive"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6495
   Icon            =   "NoteAggiuntive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton InsMod 
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
      Height          =   585
      Left            =   2340
      TabIndex        =   2
      Top             =   3465
      Width           =   1695
   End
   Begin VB.TextBox TxtModifica 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Left            =   3810
      TabIndex        =   1
      Top             =   3705
      Visible         =   0   'False
      Width           =   1110
   End
   Begin MSFlexGridLib.MSFlexGrid ElencoNote 
      Height          =   2910
      Left            =   210
      TabIndex        =   0
      Top             =   300
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   5133
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   14145495
      BackColorBkg    =   24576
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "NoteAggiuntive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MaxNoteAgg = 5
Dim rsNoteCliente As ADODB.Recordset
Public ListaNote As Collection, IdDitta As Long
Dim RigaCorr, ElencoMod As Boolean
Public Property Set NoteCliente(rsNote As ADODB.Recordset)
Set rsNoteCliente = rsNote
End Property
Private Sub TxtModifica_Change()
If TxtModifica.Text <> "" Then
 ElencoNote.Text = TxtModifica.Text: ElencoMod = True
End If
End Sub
Private Sub TxtModifica_KeyPress(KeyAscii As Integer)
Dim CarValidi$: CarValidi = "0123456789" & vbBack
If ElencoNote.Col = 0 Then
 CarValidi = CarValidi & "/"
End If
If InStr(CarValidi, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0: Exit Sub
End If
If ElencoNote.Row = ElencoNote.Rows - 1 And KeyAscii <> 0 And ElencoNote.Rows - 1 < MaxNoteAgg Then
 Dim ColSel&: ColSel = ElencoNote.Col: ElencoNote.Col = 2
 ElencoNote.CellPictureAlignment = 4: Set ElencoNote.CellPicture = ImgCancella
 ElencoNote.Col = ColSel: ElencoNote.AddItem ""
End If
End Sub
Private Sub TxtModifica_LostFocus()
TxtModifica.Visible = False
End Sub
Private Sub ElencoNote_Click()
If ElencoNote.Col <> 2 Then
 With ElencoNote
  TxtModifica.Visible = True: Set TxtModifica.Container = .Container
  TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
  TxtModifica.Text = .Text: TxtModifica.SetFocus
 End With
ElseIf ElencoNote.Rows > 2 And ElencoNote.Row < ElencoNote.Rows - 1 Then
 Dim Scelta%
 Scelta = MsgBox("Cancellare questo documento dall'elenco ?" & vbNewLine, vbYesNo + vbQuestion, "Fattura Pro")
 If Scelta = vbYes Then
  If ElencoNote.Row <= ListaNote.Count Then
   EliminaNota ElencoNote.TextMatrix(ElencoNote.Row, 1) & "-" & IdDocumento(ElencoNote.TextMatrix(ElencoNote.Row, 0))
  End If
  ElencoNote.RemoveItem ElencoNote.Row
  If ElencoNote.Row = RigaCorr Then ElencoMod = False
 End If
End If
RigaCorr = ElencoNote.Row
End Sub
Private Sub ElencoNote_Scroll()
TxtModifica.Visible = False
End Sub
Private Sub Form_Load()
IntestazioniGriglia = Array("Numero Nota", "Anno", "")
ElencoNote.ColWidth(0) = 1800: ElencoNote.ColWidth(1) = 1200
ElencoNote.ColWidth(2) = 500

For i = 0 To ElencoNote.Cols - 1
 ElencoNote.TextMatrix(0, i) = IntestazioniGriglia(i): ElencoNote.ColAlignment(i) = 4
Next i

Set ListaNote = CreaFatturaDiff.ElencoNoteAgg
Dim NotaAgg

For i = 1 To ListaNote.Count
 ElencoNote.AddItem ""
 NotaAgg = Split(ListaNote(i), "@")
 ElencoNote.TextMatrix(i, 0) = NotaAgg(0)
 ElencoNote.TextMatrix(i, 1) = NotaAgg(1)
 ElencoNote.Col = 2: ElencoNote.CellPictureAlignment = 4: Set ElencoNote.CellPicture = ImgCancella
Next i
ElencoNote.AddItem ""
End Sub
Private Sub InsMod_Click()
Call InserisciModifica: Unload Me
End Sub
Private Sub EliminaNota(IdNota As String)
ListaNote.Remove IdNota
End Sub
Private Sub Form_Unload(Cancel As Integer)
If ElencoMod Then
 Dim Scelta%
 Scelta = MsgBox("Vuoi salvare le modifiche ?", vbYesNoCancel + vbQuestion, "Chiusura Note Aggiuntive")
 If Scelta = vbYes Then Call InserisciModifica
End If
Set NoteAggiuntive = Nothing
End Sub
Private Sub InserisciModifica()
Dim NumNota$, Anno$
On Error Resume Next
For i = 1 To ElencoNote.Rows - 2
 NumNota = ElencoNote.TextMatrix(i, 0)
 Anno = ElencoNote.TextMatrix(i, 1)
 rsNoteCliente.Find "IdDoc = '" & Anno & "-" & IdDocumento(NumNota) & "'"
 If Not rsNoteCliente.EOF And rsNoteCliente("IdDitta") = IdDitta Then
  ListaNote.Item NumNota & "@" & Anno
  If Err.Number <> 0 Then
   ListaNote.Add Anno & "-" & IdDocumento(NumNota), NumNota & "@" & Anno
  End If
  Err.Clear
 Else
  cd = True
 End If
Next
Dim Msg$
If ListaNote.Count = 0 Then
 Msg = "Attenzione, l'elenco delle note aggiuntive è vuoto !"
End If
If cd Then
 Msg = Msg & vbNewLine & "Una o più note non sono state aggiunte perchè non presenti in archivio o associate ad un cliente " _
 & "diverso da quello selezionato !"
End If
MsgBox Msg, vbExclamation, "Fattura Pro"
ElencoMod = False
End Sub
