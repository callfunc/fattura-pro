VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EsenzioniIva 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Esenzioni IVA"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9735
   Icon            =   "EsenzioneIva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0033CCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   120
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   2655
      Width           =   15
   End
   Begin MSFlexGridLib.MSFlexGrid CasiEsenzione 
      Height          =   2595
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   4577
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   14145495
      BackColorSel    =   12937801
      BackColorBkg    =   24576
      GridColorFixed  =   8421504
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
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
Attribute VB_Name = "EsenzioniIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MouseX As Single, MouseY As Single
Dim rsEsenzioniIVA As ADODB.Recordset
Public FormChiamante As Form
Private Sub CasiEsenzione_DblClick()
If MouseY > (CasiEsenzione.RowPos(0) + CasiEsenzione.RowHeight(0)) And MouseY <= _
(CasiEsenzione.RowPos(CasiEsenzione.Rows - 1) + CasiEsenzione.RowHeight(CasiEsenzione.Rows - 1)) _
And MouseX <= (CasiEsenzione.ColPos(CasiEsenzione.Cols - 1) + _
CasiEsenzione.ColWidth(CasiEsenzione.Cols - 1)) Then
 FormChiamante.ChkEsenteIva.Value = 1
 FormChiamante.NonImpIva.Text = CasiEsenzione.TextMatrix(CasiEsenzione.Row, 1)
 Unload Me
End If
End Sub
Private Sub CasiEsenzione_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MouseX = x: MouseY = y
End Sub
Private Sub Form_Load()
Me.Move 2500, 2000

IntestazioniGriglia = Array("Classe Iva", "Descrizione"): CasiEsenzione.Cols = 2
CasiEsenzione.ColWidth(0) = 1600: CasiEsenzione.ColWidth(1) = 7000
For i = 0 To CasiEsenzione.Cols - 1
 CasiEsenzione.TextMatrix(0, i) = IntestazioniGriglia(i): CasiEsenzione.ColAlignment(i) = 4
Next i

CasiEsenzione.SelectionMode = flexSelectionByRow
CasiEsenzione.HighLight = flexHighlightWithFocus

Set rsEsenzioniIVA = New ADODB.Recordset
rsEsenzioniIVA.open "SELECT * FROM AliquoteIVA WHERE Aliquota = 0 ORDER BY Descr ASC", conn, adOpenDynamic, _
adLockOptimistic
 
CaricaEsenzioniIva
End Sub
Private Sub CaricaEsenzioniIva()
If rsEsenzioniIVA.EOF Then
 CasiEsenzione.Rows = 3
 CasiEsenzione.TextMatrix(1, 0) = "Non Imponibile"
 CasiEsenzione.TextMatrix(2, 0) = "Non Imponibile"
 CasiEsenzione.TextMatrix(1, 1) = "Non imponibile I.V.A. ai sensi dell'art. 41 commma 1, " _
 & "lettera a) L. 427/93"
 CasiEsenzione.TextMatrix(2, 1) = "Non imponibile I.V.A. ai sensi dell''art. 8 comma 1, " _
 & " lettera a) D.P.R. 633/72"
Else
 rsEsenzioniIVA.MoveFirst
 While Not rsEsenzioniIVA.EOF
  CasiEsenzione.AddItem ""
  CasiEsenzione.TextMatrix(CasiEsenzione.Rows - 1, 1) = rsEsenzioniIVA("classeiva")
  CasiEsenzione.TextMatrix(CasiEsenzione.Rows - 1, 2) = rsEsenzioniIVA("descr")
  rsEsenzioniIVA.MoveNext
 Wend
End If
End Sub
