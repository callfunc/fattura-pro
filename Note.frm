VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Note 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Note di Consegna"
   ClientHeight    =   5340
   ClientLeft      =   180
   ClientTop       =   525
   ClientWidth     =   15060
   Icon            =   "Note.frx":0000
   LinkTopic       =   "Note"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   15060
   Begin VB.TextBox TotImp 
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
      Left            =   4980
      TabIndex        =   27
      Top             =   4350
      Width           =   1275
   End
   Begin VB.TextBox TotIva 
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
      Left            =   7350
      TabIndex        =   26
      Top             =   4350
      Width           =   1290
   End
   Begin VB.CheckBox ChkLuogoConsegna 
      BackColor       =   &H0033CCFF&
      Caption         =   "Luogo di Consegna"
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
      Left            =   90
      TabIndex        =   25
      Top             =   3870
      Width           =   1875
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
      Left            =   10425
      TabIndex        =   24
      Top             =   3930
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton BtnSelDitta 
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
      Left            =   11760
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Note.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   225
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.TextBox TxtLC 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3840
      TabIndex        =   21
      Top             =   3870
      Width           =   3030
   End
   Begin VB.CommandButton BtnLuoghiConsegna 
      Height          =   315
      Left            =   6915
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Note.frx":41CD
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3870
      UseMaskColor    =   -1  'True
      Width           =   315
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
      Left            =   6435
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   225
      Width           =   5280
   End
   Begin VB.CommandButton BtnVai 
      Caption         =   "Vai"
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
      Left            =   2910
      TabIndex        =   17
      Top             =   225
      Width           =   555
   End
   Begin VB.CommandButton BtnCanc 
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
      Left            =   3780
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Note.frx":4328
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4845
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TotDoc 
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
      Left            =   1755
      TabIndex        =   15
      Top             =   4335
      Width           =   1440
   End
   Begin VB.CommandButton BtnNuovo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3330
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Note.frx":4418
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4845
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnUltimo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2985
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Note.frx":49B2
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4845
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnSucc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Note.frx":4F4C
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4845
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TxtRecCorr 
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
      Left            =   1470
      TabIndex        =   8
      Top             =   4845
      Width           =   1125
   End
   Begin VB.CommandButton BtnPrec 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Note.frx":54E6
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4845
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnPrimo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   735
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Note.frx":5A80
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4845
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox TxtModifica 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   13260
      TabIndex        =   5
      Top             =   4005
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid ElencoVoci 
      Height          =   2985
      Left            =   60
      TabIndex        =   4
      Top             =   765
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   5265
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   315
      BackColor       =   16777215
      BackColorFixed  =   14145495
      BackColorSel    =   16777215
      BackColorBkg    =   24576
      GridColorFixed  =   8421504
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
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
   Begin VB.TextBox Data 
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
      Left            =   4245
      TabIndex        =   2
      Top             =   225
      Width           =   1410
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
      Height          =   315
      Left            =   1770
      TabIndex        =   0
      Top             =   225
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Imponibile:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3420
      TabIndex        =   29
      Top             =   4380
      Width           =   1500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Iva:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6420
      TabIndex        =   28
      Top             =   4380
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Luogo di Consegna:"
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
      Left            =   2190
      TabIndex        =   22
      Top             =   3900
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00646464&
      Height          =   615
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   14895
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
      Left            =   5955
      TabIndex        =   19
      Top             =   255
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Documento:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   14
      Top             =   4365
      Width           =   1605
   End
   Begin VB.Label LblNumRecord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   4200
      TabIndex        =   13
      Top             =   4890
      Width           =   45
   End
   Begin VB.Label Legenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record:"
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
      Left            =   75
      TabIndex        =   12
      Top             =   4875
      Width           =   600
   End
   Begin VB.Label LblData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
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
      Left            =   3765
      TabIndex        =   3
      Top             =   255
      Width           =   435
   End
   Begin VB.Label LblBolla 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Documento:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   225
      TabIndex        =   1
      Top             =   255
      Width           =   1485
   End
End
Attribute VB_Name = "Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNote As ADODB.Recordset, rsVociNota As ADODB.Recordset, rsClienti As ADODB.Recordset, _
rsLuoghiConsegna As ADODB.Recordset
Dim RigaCorr%, VociMod As Boolean
Dim FiltroRicerca As Boolean, re As New RegExp
Attribute FiltroRicerca.VB_VarHelpID = -1
Public FormChiamante As Form, IdDoc As String
Dim StatoDoc As StatoRecord
Dim ElencoControlli As Variant, DescControlli As Variant
Dim PropCampiVoce As Variant, cm As CoordinateMouse
Dim WithEvents ssc As SmartSubClassLib.SmartSubClass
Attribute ssc.VB_VarHelpID = -1
Private Sub BtnLuoghiConsegna_Click()
If ChkLuogoConsegna.Value Then
 If TxtDitta.Tag = "" Then
  ChkLuogoConsegna.Value = 0
  MsgBox "Inserire la ditta destinataria !", vbExclamation, "Fattura Pro"
 Else
  Dim ElencoLuoghiConsegna As New LuoghiConsegna
  If ElencoLuoghiConsegna.CaricaLuoghi(TxtDitta.Tag) Then
   Set ElencoLuoghiConsegna.FormChiamante = Me: ElencoLuoghiConsegna.Show vbModal
  Else
   ChkLuogoConsegna.Value = 0
   MsgBox "La ditta non ha luoghi di consegna associati !", vbExclamation, "Fattura Pro"
  End If
 End If
End If
End Sub
Private Sub BtnCanc_Click()
Dim Scelta%
Scelta = MsgBox("Cancellare il documento corrente ?" & vbNewLine & "Non sarà possibile" _
& " annullare questa modifica !", vbYesNo + vbQuestion, "Fattura Pro")
If Scelta = vbYes Then
 Dim PosRec As Integer
 PosRec = rsNote.AbsolutePosition
 If StatoDoc <> inserimento Then
  rsNote.Delete
  rsNote.Update
  If StatoDoc <> NonModificato Then
   conn.CommitTrans
   StatoDoc = NonModificato
  End If
  If PosRec > rsNote.RecordCount Then
   If rsNote.RecordCount <> 0 Then
    rsNote.MoveLast
   End If
   CreaNuovoRecord
  Else
   rsNote.MoveNext
   VisualizzaRecord
   TxtRecCorr.Text = rsNote.AbsolutePosition
   LblNumRecord.Caption = "di " & rsNote.RecordCount
  End If
 Else
  CreaNuovoRecord
 End If
End If
End Sub
Private Sub BtnNuovo_Click()
PosizioneRecord "nuovo"
End Sub
Private Sub BtnSelDitta_Click()
Set SelezioneDitta.FormChiamante = Me
If TxtDitta.Tag <> "" Then
 SelezioneDitta.TxtDitta.Text = TxtDitta.Text
 SelezioneDitta.TxtDitta.Tag = TxtDitta.Tag
End If
SelezioneDitta.Show vbModal
If Not SelezioneDitta.rsDitte.EOF Then
 TxtDitta.Tag = SelezioneDitta.rsDitte("id")
 TxtDitta.Text = SelezioneDitta.rsDitte("ditta")
 rsClienti.Requery
 rsClienti.Find "Id = " & TxtDitta.Tag, , adSearchForward, adBookmarkFirst
End If
End Sub
Private Sub ChkLuogoConsegna_Click()
If ChkLuogoConsegna.Value = 0 Then
 TxtLC.Tag = "": TxtLC.Text = ""
End If
ModificaDoc
End Sub
Private Sub Data_KeyPress(KeyAscii As Integer)
Dim CarAmmessi$
CarAmmessi = "0123456789-" & vbBack
If InStr(CarAmmessi, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub ElencoVoci_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cm.x = x: cm.y = y
End Sub
Private Sub PopupDitte_Click()
If PopupDitte.ListIndex <> -1 Then
 TxtDitta.Text = PopupDitte.Text
 TxtDitta.Tag = PopupDitte.ItemData(PopupDitte.ListIndex)
 PopupDitte.Visible = False
End If
End Sub
Private Sub TotDoc_Change()
ModificaDoc
End Sub
Private Sub TotDoc_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotDoc_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TotImp_Change()
ModificaDoc
End Sub
Private Sub TotImp_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotImp_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TotIva_Change()
ModificaDoc
End Sub
Private Sub TotIva_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TotIva_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtDitta_LostFocus()
PopupDitte.Visible = False
End Sub
Private Sub TxtRecCorr_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtRecCorr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub ElencoVoci_Scroll()
TxtModifica.Visible = False
End Sub
Private Sub TxtNumNota_Change()
ModificaDoc
End Sub
Private Sub TxtDitta_Change()
TxtLC.Text = "": ChkLuogoConsegna.Value = 0
ModificaDoc
End Sub
Private Sub Data_Change()
ModificaDoc
End Sub
Private Sub TxtNumNota_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtNumNota.Text) = 10 Then
  KeyAscii = 0: Exit Sub
 End If
 If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End If
End Sub
Private Sub ssc_NewMessage(ByVal hWnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
Static m_bLMousePressed As Boolean, m_bLMouseClicked As Boolean
If m_bLMousePressed And uMsg = WM_LBUTTONUP Then
 m_bLMousePressed = False
 m_bLMouseClicked = True
End If
    
If Not (m_bLMousePressed) And uMsg = WM_LBUTTONDOWN Then
 m_bLMousePressed = True
 m_bLMouseClicked = False
End If
    
If m_bLMouseClicked And (uMsg = WM_ERASEBKGND) Then
 If ElencoVoci.ColWidth(2) < 800 Then
  ElencoVoci.ColWidth(2) = 800
 End If
 m_bLMouseClicked = False
End If
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
If UCase$(Chr(KeyAscii)) <> LCase$(Chr(KeyAscii)) Then
 Dim PosCursore%
 PosCursore = TxtDitta.SelStart + 1
 If PosCursore = 1 Or Mid(TxtDitta.Text, PosCursore + 1, 1) = "." Then
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 ElseIf Mid(TxtDitta.Text, PosCursore - 1, 1) = " " Then
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End If
Dim NumCar As Integer
NumCar = Len(TxtDitta.Text) + IIf(KeyAscii <> 8, 1, -1)
If NumCar >= 3 Then
 Dim rsClienti As New ADODB.Recordset
 rsClienti.Open "SELECT * FROM Clienti WHERE Rimosso = False And UCASE(Ditta) LIKE '" & UCase(TxtDitta.Text) _
 & "%' ORDER BY Ditta ASC", conn, adOpenDynamic
 If Not rsClienti.EOF Then
  PopupDitte.Clear
  rsClienti.MoveFirst
  While Not rsClienti.EOF
   With PopupDitte
    .AddItem rsClienti("Ditta")
    .ItemData(.NewIndex) = rsClienti("Id")
   End With
   rsClienti.MoveNext
  Wend
  PopupDitte.Visible = True
  PopupDitte.Left = TxtDitta.Left
  PopupDitte.Top = TxtDitta.Top + TxtDitta.Height + 45
  PopupDitte.Width = TxtDitta.Width
  PopupDitte.Height = 2500
 Else
  PopupDitte.Visible = False
 End If
 rsClienti.Close
Else
 PopupDitte.Visible = False
End If
End Sub
Private Sub TxtModifica_Change()
If TxtModifica.Text <> ElencoVoci.Text Then
 Dim Diff#, DiffIva#, TotCorr#
 VociMod = True: ModificaDoc
 If ElencoVoci.Col = 5 Then
  Dim Sconto#
  If ElencoVoci.TextMatrix(RigaCorr, 7) <> "" Then
   Dim TotNoSconto#: TotCorr = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7))
   If IsNumeric(TxtModifica.Text) Then Sconto = CDbl(TxtModifica.Text)
   If ElencoVoci.Text <> "" Then
    TotNoSconto = TotCorr * 100 / (100 - CDbl(ElencoVoci.TextMatrix(RigaCorr, 5)))
   Else: TotNoSconto = TotCorr
   End If
   NuovoTot = TotNoSconto - (TotNoSconto * Sconto / 100)
   ElencoVoci.TextMatrix(RigaCorr, 7) = FormatNumber(NuovoTot, 2)
   Diff = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) - TotCorr
   ElencoVoci.Text = TxtModifica.Text
   If ElencoVoci.TextMatrix(RigaCorr, 6) <> "" Then
    TotImp.Text = FormatNumber(CDbl(TotImp.Text) + Diff, 2)
    AlIva = CDbl(Replace(ElencoVoci.TextMatrix(RigaCorr, 6), "%", ""))
    IvaCorr = TotCorr * AlIva / 100: NuovaIva = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) * AlIva / 100
    Diff = CDbl(Arrotonda(NuovaIva)) - CDbl(Arrotonda(IvaCorr))
    TotIva.Text = FormatNumber(CDbl(TotIva.Text) + Diff, 2)
    TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text), 2)
   Else
    TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) + Diff, 2)
   End If
  End If
 ElseIf ElencoVoci.Col <> 6 Then
  If ElencoVoci.Col = 3 Then
   If IsNumeric(TxtModifica.Text) Then
    ElencoVoci.Text = FormatNumber(CDbl(TxtModifica.Text), 3)
   Else
    ElencoVoci.Text = ""
   End If
  ElseIf ElencoVoci.Col = 4 Then
   If IsNumeric(TxtModifica.Text) Then
    ElencoVoci.Text = FormatNumber(CDbl(TxtModifica.Text), 2)
   Else
    ElencoVoci.Text = ""
   End If
  Else: ElencoVoci.Text = Trim(TxtModifica.Text)
  End If
  If (ElencoVoci.Col = 3 Or ElencoVoci.Col = 4) And ElencoVoci.TextMatrix(RigaCorr, 3) <> "" _
  And ElencoVoci.TextMatrix(RigaCorr, 4) <> "" Then
   If IsNumeric(ElencoVoci.TextMatrix(RigaCorr, 7)) Then
    TotCorr = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7))
   End If
   Qnt = 1
   If ElencoVoci.TextMatrix(RigaCorr, 3) <> "" Then
    Qnt = CDbl(ElencoVoci.TextMatrix(RigaCorr, 3))
   End If
   ElencoVoci.TextMatrix(RigaCorr, 7) = Arrotonda(Qnt * CDbl(ElencoVoci.TextMatrix(RigaCorr, 4)))
   Diff = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) - TotCorr
   TotImp.Text = FormatNumber(CDbl(TotImp.Text) + Diff, 2)
   If ElencoVoci.TextMatrix(RigaCorr, 6) <> "" Then
    IvaCorr = TotCorr * CDbl(ElencoVoci.TextMatrix(RigaCorr, 6)) / 100
    NuovaIva = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) * _
    CDbl(ElencoVoci.TextMatrix(RigaCorr, 6)) / 100
    DiffIva = CDbl(Arrotonda(NuovaIva)) - CDbl(Arrotonda(IvaCorr))
    TotIva.Text = FormatNumber(CDbl(TotIva.Text) + DiffIva, 2)
   End If
   TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text), 2)
  End If
 Else
  If ElencoVoci.TextMatrix(RigaCorr, 7) <> "" Then
   If ElencoVoci.Text <> "" Then
    IvaCorr = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) * CDbl(ElencoVoci.Text) / 100
   End If
   ElencoVoci.Text = TxtModifica.Text
   If ElencoVoci.Text <> "" Then
    NuovaIva = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)) * CDbl(ElencoVoci.Text) / 100
   End If
   DiffIva = CDbl(Arrotonda(NuovaIva)) - CDbl(Arrotonda(IvaCorr))
   If DiffIva <> 0 Then
    If NuovaIva = 0 Then
     TotImp.Text = FormatNumber(CDbl(TotImp.Text) - CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)), 2)
    ElseIf IvaCorr = 0 Then
     TotImp.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(ElencoVoci.TextMatrix(RigaCorr, 7)), 2)
    End If
    TotIva.Text = FormatNumber(CDbl(TotIva.Text) + DiffIva, 2)
    TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) + DiffIva, 2)
   End If
  Else
   ElencoVoci.Text = TxtModifica.Text
  End If
 End If
End If
End Sub
Private Sub TxtModifica_KeyPress(KeyAscii As Integer)
Dim NumCar As Integer
If KeyAscii >= 32 Then
 NumCar = Len(TxtModifica.Text) + 1
 If ElencoVoci.Col > 2 Then
  If NumCar > 10 Then
   KeyAscii = 0: Exit Sub
  End If
  Dim CarAmmessi$: CarAmmessi = "0123456789,"
  If ElencoVoci.Col = 5 Then
   CarAmmessi = "0123456789"
   If InStr(CarAmmessi, Chr(KeyAscii)) = 0 Or NumCar > 2 Then
    KeyAscii = 0: Exit Sub
   End If
  ElseIf InStr(CarAmmessi, Chr(KeyAscii)) = 0 Then
   KeyAscii = 0: Exit Sub
  End If
  If ElencoVoci.Col = 3 Or ElencoVoci.Col = 4 Then
   If Not ControlloCarIns(TxtModifica, Chr(KeyAscii), 7, 3) Then
    KeyAscii = 0: Exit Sub
   End If
  End If
 End If
 If RigaCorr = ElencoVoci.Rows - 1 Then
  Dim ColSel&: ElencoVoci.AddItem "": ElencoVoci.RowHeight(ElencoVoci.Rows - 1) = 315
  ColSel = ElencoVoci.Col: ElencoVoci.Col = 8: ElencoVoci.CellPictureAlignment = 4
  Set ElencoVoci.CellPicture = ImgCancella: ElencoVoci.Col = ColSel
  If ElencoVoci.Col <> 5 Then ElencoVoci.TextMatrix(RigaCorr, 6) = "4"
  If ElencoVoci.Col <> 2 Then ElencoVoci.TextMatrix(RigaCorr, 2) = "Kg"
 End If
 If ElencoVoci.Col = 0 Then
  If TxtModifica.SelStart = 0 Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
  ElseIf Right(TxtModifica.Text, 1) <> "." And Right(TxtModifica.Text, 1) <> " " Then
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
  End If
  If NumCar = 3 Then
   TxtModifica.Text = TxtModifica.Text & Chr(KeyAscii): KeyAscii = 0: TxtModifica.SelStart = Len(TxtModifica.Text)
   Dim rsArticoli As New ADODB.Recordset
   rsArticoli.Open "SELECT * FROM Articoli WHERE descr LIKE '" & TxtModifica.Text & "%' ORDER BY descr ASC", conn, _
   adOpenDynamic, adLockOptimistic
   If Not rsArticoli.EOF Then
    Set ElencoArticoli.FormChiamante = Me
    Set ElencoArticoli.Articoli = rsArticoli: ElencoArticoli.Show vbModal
   End If
  End If
 End If
ElseIf KeyAscii = vbKeyReturn Then
 If ElencoVoci.Col < 5 Then
  ElencoVoci.Col = ElencoVoci.Col + 1
 Else: ElencoVoci.Col = 0
 End If
 With ElencoVoci
  TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
  TxtModifica.Text = .Text: TxtModifica.SetFocus
 End With
End If
End Sub
Private Sub ElencoVoci_Click()
If MouseInGriglia(ElencoVoci, cm) Then
 If ElencoVoci.Row <> RigaCorr Then
  If Not SalvaVoceCorr Then Exit Sub
 End If
 RigaCorr = ElencoVoci.Row
 If Not rsVociNota.EOF And RigaCorr <> ElencoVoci.Rows - 1 Then
  rsVociNota.Move RigaCorr - 1, adBookmarkFirst
 End If
 With ElencoVoci
  If .Col < 7 Then
   TxtModifica.Visible = True: Set TxtModifica.Container = .Container
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
   If .Col = 3 Or .Col = 4 Then
    TxtModifica.Text = Replace(.Text, ".", "")
   Else
    TxtModifica.Text = .Text
   End If
   TxtModifica.SetFocus
  ElseIf .Col = 8 And .Rows > 2 And .Row < .Rows - 1 Then
   Dim Scelta%
   Scelta = MsgBox("Rimuovere la voce dall' elenco ?", vbExclamation + vbYesNo, "Fattura Pro")
   If Scelta = vbYes Then
    If .TextMatrix(.Row, 7) <> "" Then
     VoceTot = CDbl(.TextMatrix(.Row, 7))
     If .TextMatrix(.Row, 6) <> "" Then
      TotImp.Text = FormatNumber(CDbl(TotImp.Text) - VoceTot)
      VoceIva = CDbl(Arrotonda(VoceTot * (CDbl(.TextMatrix(.Row, 6)) / 100)))
      TotIva.Text = FormatNumber(CDbl(TotIva.Text) - VoceIva, 2)
      TotDoc.Text = FormatNumber(CDbl(TotImp.Text) + CDbl(TotIva.Text), 2)
     Else
      TotDoc.Text = FormatNumber(CDbl(TotDoc.Text) - VoceTot, 2)
     End If
    End If
    If .Row <= rsVociNota.RecordCount Then
     rsVociNota.Move .Row - 1, adBookmarkFirst
     rsVociNota.Delete
    End If
    ElencoVoci.RemoveItem .Row: VociMod = False
   End If
  End If
 End With
End If
End Sub
Private Sub TxtModifica_LostFocus()
TxtModifica.Visible = False
End Sub
Private Function SalvaVoceCorr() As Boolean
SalvaVoceCorr = False: Dim e As Boolean
If VociMod Then
 For i = 0 To ElencoVoci.Cols - 2
  If InStr(1, PropCampiVoce(i), "o") <> 0 And ElencoVoci.TextMatrix(RigaCorr, i) = "" Then
   e = True: Exit For
  ElseIf ElencoVoci.TextMatrix(RigaCorr, i) <> "" Then
   If InStr(1, PropCampiVoce(i), "a") And IsNumeric(ElencoVoci.TextMatrix(RigaCorr, i)) Then
    e = True: Exit For
   ElseIf InStr(1, PropCampiVoce(i), "n") <> 0 And (Not _
   IsNumeric(ElencoVoci.TextMatrix(RigaCorr, i))) Then
    e = True: Exit For
   End If
  End If
 Next i
 If e Then
  ElencoVoci.Row = RigaCorr: ElencoVoci.Col = i
  TxtModifica.Visible = True: Set TxtModifica.Container = ElencoVoci.Container
  With ElencoVoci
   TxtModifica.Move .CellLeft + .Left + 40, .CellTop + .Top + 45, .CellWidth - 70, 255
  End With
  TxtModifica.SetFocus: TxtModifica.Text = ElencoVoci.Text
  MsgBox "Attenzione, ci sono alcuni campi con valori non validi nella voce n. " & ElencoVoci.Row & _
  " del documento corrente !", vbExclamation, "Fattura Pro": Exit Function
 End If
 If RigaCorr > rsVociNota.RecordCount Then
  rsVociNota.AddNew
 Else
  rsVociNota.Move RigaCorr - 1, adBookmarkFirst
 End If
 With rsVociNota
  .Fields("descr") = EliminaSpazi(ElencoVoci.TextMatrix(RigaCorr, 0))
  .Fields("lotto") = ElencoVoci.TextMatrix(RigaCorr, 1)
  .Fields("um") = ElencoVoci.TextMatrix(RigaCorr, 2)
  If ElencoVoci.TextMatrix(RigaCorr, 3) <> "" Then
   .Fields("qnt") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 3))
  Else
   .Fields("qnt") = 0
  End If
  .Fields("prezzo") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 4))
  .Fields("sconto") = ElencoVoci.TextMatrix(RigaCorr, 5)
  .Fields("iva") = ElencoVoci.TextMatrix(RigaCorr, 6)
  .Fields("totale") = CDbl(ElencoVoci.TextMatrix(RigaCorr, 7))
  .Update
 End With
 ModificaDoc
 VociMod = False
End If
SalvaVoceCorr = True
End Function
Private Sub Form_Load()
Me.Move 500, 500
IntestazioniGriglia = Array("Descrizione", "Lotto", "U.M.", "Quantità", "Prezzo", "Sconto %", "IVA %", _
"Totale", "")
ElencoVoci.ColWidth(0) = 6000: ElencoVoci.ColWidth(1) = 1500
ElencoVoci.ColWidth(2) = 800: ElencoVoci.ColWidth(3) = 1200
ElencoVoci.ColWidth(4) = 1200: ElencoVoci.ColWidth(5) = 1200
ElencoVoci.ColWidth(6) = 800: ElencoVoci.ColWidth(7) = 1200
ElencoVoci.ColWidth(8) = 315

For i = 0 To ElencoVoci.Cols - 1
 ElencoVoci.TextMatrix(0, i) = IntestazioniGriglia(i)
 ElencoVoci.ColAlignment(i) = 4
Next i

Set rsNote = New ADODB.Recordset
Set rsVociNota = New ADODB.Recordset
Set rsClienti = New ADODB.Recordset
Set rsLuoghiConsegna = New ADODB.Recordset
rsClienti.Open "Clienti", conn, adOpenDynamic
rsLuoghiConsegna.Open "LuoghiConsegna", conn, adOpenDynamic
If IdDoc = "" Then
 QuerySQL = "SELECT * FROM NoteConsegna ORDER BY IdDoc ASC"
Else
 QuerySQL = "SELECT * FROM NoteConsegna WHERE IdDoc = '" & IdDoc & "'"
End If
rsNote.Open QuerySQL, conn, adOpenDynamic, adLockOptimistic

TxtRecCorr.Text = "1"

BtnPrec.Enabled = False: BtnPrimo.Enabled = False
BtnUltimo.Enabled = rsNote.RecordCount > 1

If Not rsNote.EOF Then
 If rsNote.RecordCount >= 1 Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
  BtnCanc.Enabled = True
 End If
 LblNumRecord.Caption = "di " & rsNote.RecordCount
 rsNote.MoveFirst
 Call VisualizzaRecord
Else
 LblNumRecord.Caption = "di 1"
 CreaNuovoRecord
End If
If IdDoc <> "" Then
 BtnNuovo.Enabled = False: BtnCanc.Enabled = False
 BtnPrec.Enabled = False: BtnPrimo.Enabled = False
 BtnSucc.Enabled = False: BtnUltimo.Enabled = False
End If
Set ssc = New SmartSubClass: ssc.SubClassHwnd ElencoVoci.hWnd, True
PropCampiVoce = Array("ao", "", "ao", "no", "no", "n", "n", "no")
ElencoControlli = Array("TxtNumNota", "Data", "TxtDitta")
DescControlli = Array("Num. Documento", "Data", "Ditta")
End Sub
Private Sub Form_Unload(Cancel As Integer)
If StatoDoc <> NonModificato Then
 Dim Scelta As VbMsgBoxResult
 Scelta = MsgBox("Salvare il documento corrente ?", vbYesNoCancel + vbQuestion, _
 "Chiusura Note di Consegna")
 If Scelta = vbYes Then
  If Not ConvalidaRecord() Then
   Cancel = 1
  ElseIf IdDoc <> "" Then
   FormChiamante.AggiornaTotali IdDoc, (IdDoc = "0")
  End If
 ElseIf Scelta = vbCancel Then
  Cancel = 1
 End If
End If
If Cancel <> 1 And IdDoc = "" Then
 If StatoDoc <> NonModificato Then
  conn.RollbackTrans
 End If
 rsNote.Close
 Set rsNote = Nothing
 ssc.SubClassHwnd ElencoVoci.hWnd, False
 Set Note = Nothing
End If
End Sub
Private Function ConvalidaRecord() As Boolean
Dim IdNota$, d As Variant, e As Boolean

If StatoDoc <> NonModificato Then
 For i = 0 To UBound(ElencoControlli)
  If Trim(Me(ElencoControlli(i))) = "" Then
   MsgBox "Attenzione, " & DescControlli(i) & " è un campo obbligatorio !", _
   vbExclamation, "Fattura Pro"
   Me(ElencoControlli(i)).SetFocus: Exit Function
  ElseIf i = 1 Then
   If IsDate(Me(ElencoControlli(i))) Then
    d = Split(Me(ElencoControlli(i)), "-")
    If UBound(d) = 2 Then
     If Len(d(0)) < 2 Then d(0) = "0" & d(0)
     If Len(d(1)) < 2 Then d(1) = "0" & d(1)
     Data.Text = d(0) & "-" & d(1) & "-" & d(2)
    Else: e = True
    End If
   Else: e = True
   End If
   If e Then
    MsgBox "Attenzione, inserire una data valida !", vbExclamation, "Fattura Pro"
    Data.SetFocus: Data.SelStart = 0: Data.SelLength = Len(Data.Text): Exit Function
   End If
  End If
 Next i
 Dim re As New RegExp, NumDoc1$
 re.Pattern = "^[1-9][0-9]{0,2}/?[0-9]?$"
 NumDoc1 = Split(TxtNumNota, "/")(0)
 IdNota = Mid(Data.Text, 7) & "-" & String(4 - Len(NumDoc1), "0") & TxtNumNota
 If re.Test(TxtNumNota) Then
  Dim CercaDuplicato As Boolean
  If StatoDoc = inserimento Then
   CercaDuplicato = True
  ElseIf IdNota <> rsNote("IdDoc") Then
   CercaDuplicato = True
  End If
  If CercaDuplicato Then
   Dim rsNoteDuplicato As New ADODB.Recordset
   rsNoteDuplicato.Open "SELECT * FROM NoteConsegna WHERE IdDoc = '" & IdNota & "'", conn, adOpenDynamic
   If Not rsNoteDuplicato.EOF Then
    MsgBox "Attenzione, il numero documento corrisponde a quello di un altro documento" _
    & " in archivio !" & vbNewLine, vbExclamation, "Fattura Pro"
    TxtNumNota.SetFocus: TxtNumNota.SelStart = 0
    TxtNumNota.SelLength = Len(TxtNumNota): Exit Function
   End If
  End If
 Else
  MsgBox "Attenzione, il numero documento non è valido !", vbExclamation, "Fattura Pro"
  TxtNumNota.SetFocus: TxtNumNota.SelStart = 0: TxtNumNota.SelLength = Len(TxtNumNota)
  Exit Function
 End If
 If SalvaVoceCorr() Then
  If rsVociNota.RecordCount = 0 Then
   MsgBox "Attenzione, Non è stata inserita nessuna voce !", vbExclamation, "Fattura Pro": Exit Function
  End If
  Dim VerificaNumDoc As Boolean, NumValido As Boolean: NumValido = True
  If StatoDoc = inserimento Then
   VerificaNumDoc = True
  ElseIf IdNota <> rsNote("IdDoc") Then
   VerificaNumDoc = True
  End If
  If VerificaNumDoc Then
   id = Split(TxtNumNota, "/"): NumValido = NumDocValido(id, d(2))
  End If
  If Not NumValido Then
   MsgBox "Attenzione, il numero documento non è valido !", vbExclamation, "Fattura Pro"
   TxtNumNota.SetFocus: TxtNumNota.SelStart = 0
   TxtNumNota.SelLength = Len(TxtNumNota): Exit Function
  End If
  If ChkLuogoConsegna.Value And TxtLC.Text = "" Then
   MsgBox "Attenzione, inserire il luogo di consegna per questo cliente !", vbExclamation, _
   "Fattura Pro"
   Exit Function
  End If
  If TxtDitta.Tag = "" Or (Not VerificaDestDoc) Then
   MsgBox "Attenzione, il campo ditta non contiene un valore valido !", vbExclamation, _
   "Fattura Pro"
   Exit Function
  End If
  If StatoDoc = inserimento Then
   rsNote.AddNew
  End If
  rsNote("IdDoc") = IdNota
  rsNote("Data") = Data.Text
  rsNote("IdDitta") = TxtDitta.Tag
  If TxtLC.Text <> "" Then
   rsNote("IdLC") = TxtLC.Tag
  Else
   rsNote("IdLC") = Null
  End If
  rsNote("TotDoc") = CDbl(TotDoc.Text)
  rsNote("TotImp") = CDbl(TotImp.Text)
  rsNote("TotIva") = CDbl(TotIva.Text)
  rsNote.Update
  rsVociNota.MoveFirst
  While Not rsVociNota.EOF
   If IsNull(rsVociNota("IdNota")) Then
    rsVociNota("IdNota") = IdNota
    rsVociNota.Update
   End If
   rsVociNota.MoveNext
  Wend
  If IdDoc = "" Then
   StatoDoc = NonModificato
   conn.CommitTrans
   EseguiBackup = True
  End If
 Else
  Exit Function
 End If
End If
ConvalidaRecord = True
End Function
Public Sub SalvaNota(Salva As Boolean)
If StatoDoc <> NonModificato Then
 If Salva Then
  conn.CommitTrans
  EseguiBackup = True
 Else
  conn.RollbackTrans
 End If
 StatoDoc = NonModificato
End If
End Sub
Private Sub BtnPrimo_Click()
PosizioneRecord "primo"
End Sub
Private Sub BtnPrec_Click()
PosizioneRecord "precedente"
End Sub
Private Sub BtnSucc_Click()
PosizioneRecord "successivo"
End Sub
Private Sub TxtLC_Change()
ModificaDoc
End Sub
Private Sub TxtLC_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtLC_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub BtnUltimo_Click()
PosizioneRecord "ultimo"
End Sub
Private Sub PosizioneRecord(PosRecord As String)
Dim valido: valido = ConvalidaRecord()
If valido Then
 If rsNote.RecordCount > 1 Then
  Dim NumRec%: NumRec = rsNote.AbsolutePosition
  rsNote.Requery
  rsNote.Move NumRec - 1, adBookmarkFirst
 End If
 If IdDoc = "" Then
  Select Case PosRecord
  Case "primo"
   rsNote.MoveFirst
  Case "ultimo"
   rsNote.MoveLast
  Case "precedente"
   If CInt(TxtRecCorr.Text) <= rsNote.RecordCount Then
    rsNote.MovePrevious
   ElseIf rsNote.AbsolutePosition <> rsNote.RecordCount Then
    rsNote.MoveLast
   End If
  Case "successivo"
   If rsNote.AbsolutePosition = rsNote.RecordCount Then
    CreaNuovoRecord
    Exit Sub
   End If
   rsNote.MoveNext
  Case "nuovo"
   CreaNuovoRecord
   Exit Sub
  End Select
  BtnCanc.Enabled = True
  If rsNote.AbsolutePosition <= rsNote.RecordCount Then
   BtnSucc.Enabled = True: BtnNuovo.Enabled = True
  End If
  If rsNote.AbsolutePosition <> rsNote.RecordCount Then
   BtnUltimo.Enabled = True
  Else
   BtnUltimo.Enabled = False
  End If
  If rsNote.AbsolutePosition <> 1 Then
   BtnPrimo.Enabled = True: BtnPrec.Enabled = True
  Else
   BtnPrimo.Enabled = False: BtnPrec.Enabled = False
  End If
  TxtRecCorr.Text = rsNote.AbsolutePosition
  LblNumRecord.Caption = "di " & rsNote.RecordCount
  rsVociNota.Close
  VisualizzaRecord
 End If
End If
End Sub
Private Sub VisualizzaRecord()
TxtNumNota.Text = NumeroDocumento(rsNote("IdDoc"))
Data.Text = Format$(rsNote("Data"), "dd-mm-yyyy")
rsClienti.Find "Id = " & rsNote("IdDitta"), , adSearchForward, adBookmarkFirst
TxtDitta.Text = rsClienti("Ditta")
TxtDitta.Tag = rsNote("IdDitta")
If Not IsNull(rsNote("IdLC")) Then
 TxtLC.Tag = rsNote("IdLC")
 rsLuoghiConsegna.Find "Id = " & rsNote("IdLc"), , adSearchForward, adBookmarkFirst
 If Not rsLuoghiConsegna.EOF Then
  ChkLuogoConsegna.Value = 1
  TxtLC.Text = rsLuoghiConsegna("Nome")
 Else
  ChkLuogoConsegna.Value = 0
  TxtLC.Tag = "": TxtLC.Text = ""
  rsNote("IdLC") = Null
 End If
End If
TxtModifica.Visible = False: ElencoVoci.Rows = 1
TotDoc.Text = FormatNumber(rsNote("totdoc"), 2)
TotImp.Text = FormatNumber(rsNote("totimp"), 2)
TotIva.Text = FormatNumber(rsNote("totiva"), 2)
If rsVociNota.State <> adStateClosed Then
 rsVociNota.Close
End If
rsVociNota.Open "SELECT * FROM VociNoteConsegna WHERE IdNota = '" & rsNote("IdDoc") & "' ORDER BY Id ASC", conn, adOpenDynamic, adLockOptimistic
rsVociNota.MoveFirst
With ElencoVoci
 Dim CifreDec%, i%
 ElencoVoci.Redraw = False
 While Not rsVociNota.EOF
  .AddItem "": i = 0
  .RowHeight(.Rows - 1) = 315
  For Each Field In rsVociNota.Fields
   If Field.Name <> "Id" And Field.Name <> "IdNota" Then
    If Field.Type = adDouble Then
     If Field.Name = "Qnt" Then
      CifreDec = 3
     Else
      CifreDec = 2
     End If
     If Field.Value <> 0 Then
      .TextMatrix(.Rows - 1, i) = FormatNumber(rsVociNota(Field.Name), CifreDec)
     End If
    Else
     .TextMatrix(.Rows - 1, i) = IIf(IsNull(rsVociNota(Field.Name)), "", rsVociNota(Field.Name))
    End If
    i = i + 1
   End If
  Next
  .Row = .Rows - 1: .Col = 8: .CellPictureAlignment = 4
  Set .CellPicture = ImgCancella
  rsVociNota.MoveNext
 Wend
 ElencoVoci.Redraw = True
 .AddItem "": .RowHeight(.Rows - 1) = 315
 .Col = 0
End With
If StatoDoc <> NonModificato Then
 StatoDoc = NonModificato
 conn.CommitTrans
End If
End Sub
Public Sub ImpostaFiltroRicerca(rsNoteRicerca As ADODB.Recordset)
Set rsNote = rsNoteRicerca
FiltroRicerca = True
End Sub
Public Sub CaricaFiltroRicerca()
PosizioneRecord "primo"
End Sub
Public Function VerificaDestDoc() As Boolean
Dim rsClienti As New ADODB.Recordset
rsClienti.Open "SELECT * FROM Clienti WHERE UCASE(Ditta) = '" & UCase(TxtDitta.Text) & "'", conn, adOpenDynamic
If Not rsClienti.EOF Then
 TxtDitta.Tag = rsClienti("Id")
 VerificaDestDoc = True
End If
End Function
Private Sub CreaNuovoRecord()
TxtNumNota.Text = "": TotDoc.Text = "0,00"
TotImp.Text = "0,00": TotIva.Text = "0,00"
TxtDitta.Text = "": TxtDitta.Tag = ""
Data.Text = ""
LblNumRecord.Caption = "di " & (rsNote.RecordCount + 1): ElencoVoci.Rows = 1
ElencoVoci.AddItem "": ElencoVoci.RowHeight(1) = 315
TxtRecCorr.Text = rsNote.RecordCount + 1: RigaCorr = 0
BtnSucc.Enabled = False: BtnCanc.Enabled = False: BtnNuovo.Enabled = False
If rsNote.RecordCount >= 1 Then
 BtnPrec.Enabled = True: BtnPrimo.Enabled = True: BtnUltimo.Enabled = True
End If
If rsVociNota.State <> adStateClosed Then
 rsVociNota.Close
End If
rsVociNota.Open "SELECT * FROM VociNoteConsegna WHERE IdNota = '0'", conn, adOpenDynamic, adLockOptimistic
StatoDoc = NonModificato
conn.CommitTrans
End Sub
Private Sub ModificaDoc()
Dim AvviaTrans As Boolean
If StatoDoc = NonModificato Then
 AvviaTrans = True
End If
If CInt(TxtRecCorr.Text) > rsNote.RecordCount Then
 If Not FiltroRicerca Then
  StatoDoc = inserimento
  BtnCanc.Enabled = True: BtnSucc.Enabled = True
  BtnNuovo.Enabled = True
 Else
  AvviaTrans = False
 End If
Else
 StatoDoc = modifica
End If
If AvviaTrans Then conn.BeginTrans
End Sub
Public Function FormBloccata() As Boolean
FormBloccata = StatoDoc <> NonModificato
End Function
Private Sub BtnVai_Click()
If Not rsNote.EOF Then
 IdCorr = NumeroDocumento(rsNote("IdDoc")) & "-" & Mid(rsNote("IdDoc"), 1, 4)
End If
If TxtNumNota <> "" And IdCorr <> TxtNumNota Then
 If InStr(TxtNumNota, "-") <> 0 Then
  Dim IdDoc, NumDoc1$, IdNota$, PosRec%
  PosRec = rsNote.AbsolutePosition
  IdDoc = Split(TxtNumNota, "-")
  NumDoc1 = Split(IdDoc(0), "/")(0)
  IdNota = IdDoc(1) & "-" & String(4 - Len(NumDoc1), "0") & IdDoc(0)
  rsNote.Find "IdDoc = '" & IdNota & "'", , adSearchForward, adBookmarkFirst
  If Not rsNote.EOF Then
   VisualizzaRecord
   TxtRecCorr.Text = rsNote.AbsolutePosition
   If rsNote.AbsolutePosition > 1 Then
    BtnPrec.Enabled = True: BtnPrimo.Enabled = True
   End If
   If rsNote.AbsolutePosition < rsNote.RecordCount Then
    BtnUltimo.Enabled = True
   End If
   BtnNuovo.Enabled = True: BtnSucc.Enabled = True
  Else
   If rsNote.RecordCount Then
    rsNote.Move PosRec - 1, adBookmarkFirst
   End If
   MsgBox "Attenzione, documento non presente in archivio !", vbExclamation, "Fattura Pro"
   If StatoDoc <> NonModificato Then
    StatoDoc = NonModificato
    conn.CommitTrans
   End If
  End If
 Else
  If StatoDoc = inserimento Then
   TxtNumNota.Text = ""
  ElseIf StatoDoc = modifica Then
   TxtNumNota.Text = rsNote("Id")
  End If
  MsgBox "Attenzione, devi indicare il documento che vuoi visualizzare nel formato numero-anno (Es: 100-2016) !", _
  vbExclamation, "Fattura Pro"
 End If
End If
End Sub
Private Function NumDocValido(id As Variant, ByVal Anno As String) As Boolean
If UBound(id) <= 1 Then
 If UBound(id) = 1 Then
  Dim IdPrec As String
  If id(1) = "1" Then
   IdPrec = Anno & "-" & String(4 - Len(id(0)), "0") & id(0)
  Else: IdPrec = Anno & "-" & String(4 - Len(id(0)), "0") & id(0) & "/" & (CInt(id(1)) - 1)
  End If
  If IdPrec <> rsNote("IdDoc") Then
   Dim rsRecPrec As New ADODB.Recordset
   rsRecPrec.Open "SELECT * FROM NoteConsegna WHERE IdDoc = " & IdPrec, conn, adOpenDynamic
   NumDocValido = Not rsRecPrec.EOF
   rsRecPrec.Close
   If Not NumDocValido Then Exit Function
  Else
   Exit Function
  End If
 End If
Else
 Exit Function
End If
Dim rsRecInvalido As New ADODB.Recordset, IdDoc$, NumDoc1$, DataDoc$
NumDoc1 = id(0)
IdDoc = Anno & "-" & String(4 - Len(NumDoc1), "0") & TxtNumNota
DataDoc = Format$(Data, "yyyy/mm/dd")
rsRecInvalido.Open "SELECT * FROM NoteConsegna WHERE (IdDoc < '" & IdDoc & "' AND Data > #" & DataDoc & "#) OR " _
& "(IdDoc > '" & IdDoc & "' AND Data < #" & DataDoc & "#)", conn, adOpenDynamic
NumDocValido = rsRecInvalido.EOF
rsRecInvalido.Close
End Function
