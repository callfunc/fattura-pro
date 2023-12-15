VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form LuoghiConsegna 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luoghi di Consegna"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12270
   Icon            =   "LuoghiConsegna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   315
      Left            =   9075
      TabIndex        =   2
      Top             =   2790
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton BtnSalva 
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
      Height          =   570
      Left            =   5190
      TabIndex        =   1
      Top             =   2850
      Width           =   1770
   End
   Begin MSFlexGridLib.MSFlexGrid GrigliaLuoghi 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      RowHeightMin    =   315
      ForeColor       =   0
      BackColorFixed  =   14145495
      ForeColorFixed  =   0
      BackColorSel    =   12937801
      ForeColorSel    =   16777215
      BackColorBkg    =   24576
      GridColorFixed  =   8421504
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
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
Attribute VB_Name = "LuoghiConsegna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormChiamante As Form
Public rsLuoghi As ADODB.Recordset, Ditta As Long
Dim Stato As StatoRecord, ContCampiVoce As Variant
Public ElencoMod As Boolean
Private Sub Form_Load()
Me.Move 2000, 2100
If FormChiamante.Name = "Clienti" Then
 Me.Height = 3990: GrigliaLuoghi.Cols = 6
 GrigliaLuoghi.SelectionMode = flexSelectionFree
 GrigliaLuoghi.HighLight = flexHighlightNever
Else
 Me.Height = 3120: GrigliaLuoghi.Cols = 5
 GrigliaLuoghi.SelectionMode = flexSelectionByRow
 GrigliaLuoghi.HighLight = flexHighlightWithFocus
End If

Dim IntestazioniGriglia As Variant
IntestazioniGriglia = Array("Nome", "Indirizzo", "Località", "Cap", "Provincia")
GrigliaLuoghi.ColWidth(0) = 3315: GrigliaLuoghi.ColWidth(1) = 3800
GrigliaLuoghi.ColWidth(2) = 2000: GrigliaLuoghi.ColWidth(3) = 900
GrigliaLuoghi.ColWidth(4) = 1100

If GrigliaLuoghi.Cols = 6 Then GrigliaLuoghi.ColWidth(5) = 315

For i = 0 To UBound(IntestazioniGriglia)
 GrigliaLuoghi.TextMatrix(0, i) = IntestazioniGriglia(i)
 GrigliaLuoghi.ColAlignment(i) = 4
Next i

i = 0
If Not rsLuoghi.EOF Then
 rsLuoghi.MoveFirst
 While Not rsLuoghi.EOF
  GrigliaLuoghi.AddItem ""
  GrigliaLuoghi.TextMatrix(i + 1, 0) = rsLuoghi("nome")
  GrigliaLuoghi.TextMatrix(i + 1, 1) = rsLuoghi("indirizzo")
  GrigliaLuoghi.TextMatrix(i + 1, 2) = rsLuoghi("loc")
  GrigliaLuoghi.TextMatrix(i + 1, 3) = rsLuoghi("cap")
  GrigliaLuoghi.TextMatrix(i + 1, 4) = rsLuoghi("prov")

  If GrigliaLuoghi.Cols = 6 Then
   GrigliaLuoghi.Row = i + 1: GrigliaLuoghi.Col = 5
   GrigliaLuoghi.CellPictureAlignment = 4: Set GrigliaLuoghi.CellPicture = ImgCancella
  End If
  rsLuoghi.MoveNext
 Wend
End If
If FormChiamante.Name = "Clienti" Then
 GrigliaLuoghi.AddItem ""
Else
 Me.Height = 3120
End If
ContCampiVoce = Array("^.+$", "^[a-z][a-z0-9 ]+[0-9]*$", "^[a-z]+$", "^([0-9]{5})?$", "^[a-z]{2}$")
End Sub
Public Function CaricaLuoghi(IdDitta As Long) As Boolean
Set rsLuoghi = New ADODB.Recordset
rsLuoghi.Open "SELECT * FROM LuoghiConsegna WHERE IdDitta = " & IdDitta, conn, adOpenDynamic, adLockBatchOptimistic
Ditta = IdDitta
CaricaLuoghi = Not rsLuoghi.EOF
End Function
Private Sub GrigliaLuoghi_Click()
If FormChiamante.Name = "Clienti" Then
 If GrigliaLuoghi.Col < 5 Then
  With GrigliaLuoghi
   TxtModifica.Visible = True: Set TxtModifica.Container = .Container
   TxtModifica.Move .CellLeft + .Left, .CellTop + .Top + 45, .CellWidth - 70, 255
   TxtModifica.Text = .Text: TxtModifica.SetFocus
  End With
 ElseIf GrigliaLuoghi.Rows > 2 And GrigliaLuoghi.Row < GrigliaLuoghi.Rows - 1 Then
  Dim Scelta%
  Scelta = MsgBox("Cancellare questo luogo di consegna ?", vbYesNo + vbQuestion, "Fattura Pro")
  If Scelta = vbYes Then
   If GrigliaLuoghi.Row <= rsLuoghi.RecordCount Then
    rsLuoghi.Move GrigliaLuoghi.Row - 1, adBookmarkFirst
    rsLuoghi.Delete
   End If
   GrigliaLuoghi.RemoveItem GrigliaLuoghi.Row: modificato = True
   ElencoMod = True
  End If
 End If
End If
End Sub
Private Sub GrigliaLuoghi_Scroll()
TxtModifica.Visible = False
End Sub
Private Sub GrigliaLuoghi_DblClick()
If GrigliaLuoghi.Rows > 1 And FormChiamante.Name <> "Clienti" Then
 rsLuoghi.Move GrigliaLuoghi.Row - 1, adBookmarkFirst
 FormChiamante.TxtLC.Tag = rsLuoghi("Id")
 FormChiamante.ChkLuogoConsegna.Value = 1
 FormChiamante.TxtLC.Text = GrigliaLuoghi.TextMatrix(GrigliaLuoghi.Row, 0)
 Unload Me
End If
End Sub
Private Sub TxtModifica_Change()
ElencoMod = True
GrigliaLuoghi.Text = TxtModifica.Text
End Sub
Private Sub TxtModifica_KeyPress(KeyAscii As Integer)
Dim PosCursore%
PosCursore = TxtModifica.SelStart + 1
If PosCursore = 1 Or Mid(TxtModifica.Text, PosCursore + 1, 1) = "." Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
ElseIf Mid(TxtModifica.Text, PosCursore - 1, 1) = " " Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
If GrigliaLuoghi.Row = GrigliaLuoghi.Rows - 1 Then
 Dim ColSel&: GrigliaLuoghi.AddItem "": ColSel = GrigliaLuoghi.Col
 GrigliaLuoghi.Col = 5: GrigliaLuoghi.CellPictureAlignment = 4
 Set GrigliaLuoghi.CellPicture = ImgCancella: GrigliaLuoghi.Col = ColSel
End If
End Sub
Private Sub TxtModifica_LostFocus()
TxtModifica.Visible = False
End Sub
Private Sub BtnSalva_Click()
If SalvaElenco Then
 Unload Me
End If
End Sub
Private Function SalvaElenco() As Boolean
If ElencoMod Then
 Dim re As New RegExp, r&, c&, ValCol$, DescErrore As String, e As Boolean
 re.IgnoreCase = True
 For r = 1 To GrigliaLuoghi.Rows - 2
  For c = 0 To GrigliaLuoghi.Cols - 2
   ValCol = GrigliaLuoghi.TextMatrix(r, c)
   re.Pattern = ContCampiVoce(c)
   If Not re.Test(ValCol) Then
    e = True
    Exit For
   End If
  Next c
  If e Then
   With GrigliaLuoghi
   .Row = r: .Col = c
   TxtModifica.Visible = True: Set TxtModifica.Container = GrigliaLuoghi.Container
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
   TxtModifica.Text = .Text: TxtModifica.SetFocus
   End With
   MsgBox "Il campo " & GrigliaLuoghi.TextMatrix(0, c) & " è vuoto o contiene un valore non " _
   & "valido", vbExclamation, "Fattura Pro"
   Exit Function
  End If
  If r > rsLuoghi.RecordCount Then
   rsLuoghi.AddNew
  Else
   rsLuoghi.Move r - 1, adBookmarkFirst
  End If
  With GrigliaLuoghi
   rsLuoghi("IdDitta") = Ditta
   rsLuoghi("Nome") = .TextMatrix(r, 0)
   rsLuoghi("Indirizzo") = .TextMatrix(r, 1)
   rsLuoghi("Loc") = .TextMatrix(r, 2)
   rsLuoghi("Cap") = .TextMatrix(r, 3)
   rsLuoghi("Prov") = .TextMatrix(r, 4)
  End With
 Next r
End If
SalvaElenco = True
End Function
