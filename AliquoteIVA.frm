VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form AliquoteIVA 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aliquote IVA"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10125
   Icon            =   "AliquoteIVA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   10125
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
      Height          =   285
      Left            =   5745
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   765
   End
   Begin MSFlexGridLib.MSFlexGrid ElencoAliquote 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   14145495
      BackColorBkg    =   24576
      GridColorFixed  =   8421504
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
Attribute VB_Name = "AliquoteIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAliquote As ADODB.Recordset, RigaCorr As Long, cm As CoordinateMouse
Dim StatoVoce As StatoRecord, PropColVoci As Variant
Private Sub ElencoAliquote_Click()
If MouseInGriglia(ElencoAliquote, cm) Then
 If ElencoAliquote.Row <> RigaCorr Then
  If Not SalvaVoce() Then Exit Sub
  RigaCorr = ElencoAliquote.Row
 End If
 If ElencoAliquote.Col <> 3 Then
  With ElencoAliquote
   TxtModifica.Visible = True: Set TxtModifica.Container = .Container
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
   TxtModifica.Text = .Text
   TxtModifica.SetFocus
  End With
 Else
  Dim Scelta%
  Scelta = MsgBox("Cancellare l'aliquota ?", vbExclamation + vbYesNo, "Fattura Pro")
  If Scelta = vbYes Then
   If RigaCorr <= rsAliquote.RecordCount Then
    rsAliquote.Move RigaCorr - 1, adBookmarkFirst
    rsAliquote.Delete
   End If
   ElencoAliquote.RemoveItem ElencoAliquote.Row: StatoVoce = NonModificato
  End If
 End If
End If
End Sub
Private Sub ElencoAliquote_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cm.x = x: cm.y = y
End Sub
Private Sub Form_Load()
IntestazioniGriglia = Array("Aliquota", "Classe IVA", "Descrizione")
ElencoAliquote.Cols = 4
ElencoAliquote.ColWidth(0) = 1100: ElencoAliquote.ColWidth(1) = 1300
ElencoAliquote.ColWidth(2) = 5500: ElencoAliquote.ColWidth(3) = 360
ElencoAliquote.HighLight = flexHighlightNever: ElencoAliquote.SelectionMode = flexSelectionFree
For i = 0 To ElencoAliquote.Cols - 2
 ElencoAliquote.TextMatrix(0, i) = IntestazioniGriglia(i): ElencoAliquote.ColAlignment(i) = 4
Next i
Set rsAliquote = New ADODB.Recordset
rsAliquote.Open "SELECT * FROM AliquoteIVA ORDER BY Aliquota ASC", conn, adOpenDynamic, adLockOptimistic
If Not rsAliquote.EOF Then
 rsAliquote.MoveFirst
 While Not rsAliquote.EOF
  With ElencoAliquote
   .AddItem ""
   .TextMatrix(.Rows - 1, 0) = rsAliquote("Aliquota")
   .TextMatrix(.Rows - 1, 1) = rsAliquote("ClasseIVA")
   .TextMatrix(.Rows - 1, 2) = rsAliquote("Descr")
   .Row = .Rows - 1: .Col = 3: .CellPictureAlignment = 4
   Set .CellPicture = ImgCancella
  End With
  rsAliquote.MoveNext
 Wend
End If
ElencoAliquote.AddItem ""
PropColVoci = Array("no", "a", "ao")
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not SalvaVoce() Then
 Cancel = 1
End If
End Sub
Private Sub TxtModifica_Change()
If ElencoAliquote.Text <> TxtModifica.Text Then
 ElencoAliquote.Text = TxtModifica.Text
 If ElencoAliquote.Row > rsAliquote.RecordCount Then
  StatoVoce = inserimento
 Else
  StatoVoce = modifica
 End If
End If
End Sub
Private Sub TxtModifica_LostFocus()
TxtModifica.Visible = False
End Sub
Private Sub TxtModifica_KeyPress(KeyAscii As Integer)
Dim NumCar As Integer
If KeyAscii >= 32 Then
 NumCar = Len(TxtModifica.Text) + 1
 If ElencoAliquote.Col = 0 Then
  If Len(TxtModifica.Text) = 5 Then
   KeyAscii = 0
  End If
 End If
 If RigaCorr = ElencoAliquote.Rows - 1 Then
  Dim ColCorr&: ColCorr = ElencoAliquote.Col: ElencoAliquote.Col = 3: ElencoAliquote.CellPictureAlignment = 4
  Set ElencoAliquote.CellPicture = ImgCancella: ElencoAliquote.Col = ColCorr: ElencoAliquote.AddItem ""
 End If
 If ElencoAliquote.Col >= 1 Then
  Dim PosCursore%
  PosCursore = TxtModifica.SelStart + 1
  If PosCursore = 1 Or Mid(TxtModifica.Text, PosCursore + 1, 1) = "." Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
 End If
ElseIf KeyAscii = vbKeyReturn Then
 If ElencoAliquote.Col < 2 Then
  ElencoAliquote.Col = ElencoAliquote.Col + 1
 Else: ElencoAliquote.Col = 0
 End If
 With ElencoAliquote
  TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
  TxtModifica.Text = .Text: TxtModifica.SetFocus
 End With
End If
End Sub
Private Function SalvaVoce() As Boolean
If StatoVoce <> NonModificato Then
 Dim DescErrore As String, e As Boolean
 For i = 0 To ElencoAliquote.Cols - 2
  If InStr(1, PropColVoci(i), "o") <> 0 And ElencoAliquote.TextMatrix(RigaCorr, i) = "" Then
    e = True
    DescErrore = ElencoAliquote.TextMatrix(0, i) & " è un campo obbligatorio !"
    Exit For
  ElseIf ElencoAliquote.TextMatrix(RigaCorr, i) <> "" Then
   If InStr(1, PropColVoci(i), "a") And IsNumeric(ElencoAliquote.TextMatrix(RigaCorr, i)) Then
    e = True
    DescErrore = ElencoAliquote.TextMatrix(0, i) & " non può essere un numero !"
    Exit For
   ElseIf InStr(1, PropColVoci(i), "n") <> 0 And Not IsNumeric(ElencoAliquote.TextMatrix(RigaCorr, i)) Then
    e = True
    DescErrore = ElencoAliquote.TextMatrix(0, i) & " deve essere un numero !"
    Exit For
   End If
  End If
 Next i
 
 If e Then
  ElencoAliquote.Row = RigaCorr: ElencoAliquote.Col = i
  TxtModifica.Visible = True: Set TxtModifica.Container = ElencoAliquote.Container
  With ElencoAliquote
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
  End With
  TxtModifica.SetFocus: TxtModifica.Text = ElencoAliquote.Text
  MsgBox DescErrore, vbExclamation, "Fattura Pro": Exit Function
 End If
 
 With rsAliquote
  If StatoVoce = inserimento Then
   .AddNew
  Else
   .Move RigaCorr - 1, adBookmarkFirst
  End If
  .Fields("aliquota") = ElencoAliquote.TextMatrix(RigaCorr, 0)
  .Fields("classeiva") = ElencoAliquote.TextMatrix(RigaCorr, 1)
  .Fields("descr") = ElencoAliquote.TextMatrix(RigaCorr, 2)
  .Update
 End With
End If
StatoVoce = NonModificato
SalvaVoce = True
End Function
