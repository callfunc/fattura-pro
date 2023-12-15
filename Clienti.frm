VERSION 5.00
Begin VB.Form Clienti 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clienti"
   ClientHeight    =   7965
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "Clienti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   8010
   Begin VB.TextBox TxtComune 
      Height          =   315
      Left            =   1500
      TabIndex        =   40
      Top             =   2610
      Width           =   2115
   End
   Begin VB.Frame FrmFattEl 
      BackColor       =   &H0033CCFF&
      Caption         =   "Fatturazione Elettronica"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   210
      TabIndex        =   34
      Top             =   5970
      Width           =   5895
      Begin VB.TextBox TxtPEC 
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
         Left            =   1785
         TabIndex        =   38
         Top             =   780
         Width           =   3735
      End
      Begin VB.TextBox TxtCodDest 
         Height          =   315
         Left            =   1800
         TabIndex        =   35
         Top             =   345
         Width           =   2550
      End
      Begin VB.Label LblPEC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PEC"
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
         Left            =   195
         TabIndex        =   37
         Top             =   840
         Width           =   315
      End
      Begin VB.Label LblCodDest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codice Destinatario"
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
         Left            =   195
         TabIndex        =   36
         Top             =   375
         Width           =   1545
      End
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
      Left            =   1500
      TabIndex        =   32
      Top             =   3420
      Width           =   2910
   End
   Begin VB.CommandButton BtnLuoghiConsegna 
      Caption         =   "Luoghi di Consegna"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4035
      TabIndex        =   31
      Top             =   5055
      Width           =   1950
   End
   Begin VB.TextBox TxtEmail 
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
      Left            =   1500
      TabIndex        =   30
      Top             =   4230
      Width           =   2655
   End
   Begin VB.TextBox TxtFax 
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
      Left            =   1500
      TabIndex        =   28
      Top             =   4635
      Width           =   2430
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
      Left            =   915
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Clienti.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7455
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnPrec 
      Appearance      =   0  'Flat
      DisabledPicture =   "Clienti.frx":460C
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
      Left            =   1260
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Clienti.frx":4BA6
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7455
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
      Left            =   1635
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7455
      Width           =   1125
   End
   Begin VB.CommandButton BtnSucc 
      Height          =   315
      Left            =   2805
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Clienti.frx":5140
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7455
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnUltimo 
      Height          =   315
      Left            =   3150
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Clienti.frx":56DA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7455
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnNuovo 
      Height          =   315
      Left            =   3495
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Clienti.frx":5C74
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7455
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton BtnCanc 
      Height          =   315
      Left            =   3945
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Clienti.frx":620E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7455
      UseMaskColor    =   -1  'True
      Width           =   345
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
      Left            =   1500
      TabIndex        =   8
      Top             =   3015
      Width           =   1260
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
      Left            =   1500
      MaxLength       =   16
      TabIndex        =   2
      Top             =   990
      Width           =   3120
   End
   Begin VB.TextBox TxtPartIVA 
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
      Left            =   1500
      TabIndex        =   1
      Top             =   585
      Width           =   2940
   End
   Begin VB.ListBox LstPag 
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
      Height          =   735
      ItemData        =   "Clienti.frx":62FE
      Left            =   1500
      List            =   "Clienti.frx":6310
      TabIndex        =   6
      Top             =   5040
      Width           =   1710
   End
   Begin VB.TextBox TxtLoc 
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
      Left            =   1500
      TabIndex        =   7
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox TxtTel 
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
      Left            =   1500
      TabIndex        =   4
      Top             =   3825
      Width           =   2400
   End
   Begin VB.TextBox TxtCap 
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
      Left            =   1500
      TabIndex        =   5
      Top             =   2205
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
      Left            =   1500
      TabIndex        =   3
      Top             =   1395
      Width           =   4320
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
      Left            =   1500
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   180
      Width           =   6150
   End
   Begin VB.Label LblComune 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comune"
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
      TabIndex        =   39
      Top             =   2640
      Width           =   690
   End
   Begin VB.Label Label11 
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
      Left            =   240
      TabIndex        =   33
      Top             =   3450
      Width           =   450
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
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
      TabIndex        =   29
      Top             =   4275
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
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
      TabIndex        =   27
      Top             =   4680
      Width           =   300
   End
   Begin VB.Label LblNumRecord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "di 345"
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
      Left            =   4365
      TabIndex        =   26
      Top             =   7515
      Width           =   465
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
      Left            =   210
      TabIndex        =   25
      Top             =   7485
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia:"
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
      Left            =   240
      TabIndex        =   17
      Top             =   3060
      Width           =   780
   End
   Begin VB.Label Label8 
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
      Left            =   255
      TabIndex        =   16
      Top             =   1020
      Width           =   1170
   End
   Begin VB.Label Label9 
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
      Left            =   270
      TabIndex        =   15
      Top             =   630
      Width           =   825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   210
      TabIndex        =   14
      Top             =   5040
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Località:"
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
      TabIndex        =   13
      Top             =   1830
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono:"
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
      TabIndex        =   12
      Top             =   3870
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.A.P.:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2250
      Width           =   525
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   255
      TabIndex        =   10
      Top             =   1440
      Width           =   705
   End
   Begin VB.Label Label1 
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
      Left            =   270
      TabIndex        =   9
      Top             =   210
      Width           =   540
   End
End
Attribute VB_Name = "Clienti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClienti As ADODB.Recordset, rsLuoghiConsegna As ADODB.Recordset, StatoCliente As StatoRecord
Dim ElencoControlli As Variant, DescControlli As Variant, PropControlli As Variant
Dim FiltroRicerca As Boolean
Private Sub BtnLuoghiConsegna_Click()
Dim ElencoLuoghiConsegna As New LuoghiConsegna, IdCliente As Long
If StatoCliente <> inserimento And rsClienti.RecordCount <> 0 Then
 IdCliente = rsClienti.Fields("id")
End If
If rsLuoghiConsegna.RecordCount <> 0 Then
 rsLuoghiConsegna.MoveFirst
End If
Set ElencoLuoghiConsegna.rsLuoghi = rsLuoghiConsegna
ElencoLuoghiConsegna.Ditta = IdCliente
Set ElencoLuoghiConsegna.FormChiamante = Me
ElencoLuoghiConsegna.Show vbModal
If ElencoLuoghiConsegna.ElencoMod Then
 ModificaDitta
End If
End Sub
Private Sub BtnCanc_Click()
Dim Scelta%
Scelta = MsgBox("Cancellare questo cliente ?" & vbNewLine & "Non sarà possibile" _
& " annullare questa modifica !", vbYesNo + vbQuestion, "Fattura Pro")
If Scelta = vbYes Then
 Dim PosRec As Integer
 PosRec = rsClienti.AbsolutePosition
 If StatoCliente <> inserimento Then
  conn.Execute "DELETE * FROM LuoghiConsegna WHERE IdDitta = " & rsClienti("id")
  rsClienti("Rimosso") = True
  rsClienti.Update
  rsClienti.Requery
  If PosRec > rsClienti.RecordCount Then
   If rsClienti.RecordCount <> 0 Then
    rsClienti.MoveLast
   End If
   CreaNuovoRecord
  Else
   rsClienti.Move PosRec - 1, adBookmarkFirst
   VisualizzaRecord
   TxtRecCorr.Text = rsClienti.AbsolutePosition
   LblNumRecord.Caption = "di " & rsClienti.RecordCount
  End If
 Else
  CreaNuovoRecord
 End If
End If
End Sub
Private Sub TxtCap_Change()
ModificaDitta
End Sub
Private Sub TxtCodFisc_Change()
ModificaDitta
End Sub
Private Sub TxtCodFisc_KeyPress(KeyAscii As Integer)
If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVXYWZ" & vbBack, UCase(Chr(KeyAscii))) = 0 Then
 KeyAscii = 0
Else
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtEmail_Change()
ModificaDitta
End Sub
Private Sub TxtFax_Change()
ModificaDitta
End Sub
Private Sub TxtFax_KeyPress(KeyAscii As Integer)
If InStr("0123456789" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
Private Sub TxtRecCorr_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub TxtRecCorr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub TxtDitta_Change()
ModificaDitta
End Sub
Private Sub TxtDitta_KeyPress(KeyAscii As Integer)
If TxtDitta.SelStart = 0 Or Mid(TxtDitta.Text, TxtDitta.SelStart + 1, 1) = "." Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxtLoc_Change()
ModificaDitta
End Sub
Private Sub TxtLoc_KeyPress(KeyAscii As Integer)
If TxtLoc.SelStart = 0 Or Mid(TxtLoc.Text, TxtLoc.SelStart + 1, 1) = "." Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub LstPag_Click()
ModificaDitta
End Sub
Private Sub TxtPartIva_Change()
ModificaDitta
End Sub
Private Sub TxtPartIva_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtPartIva.Text) = 18 Then
  KeyAscii = 0
 End If
End If
End Sub
Private Sub TxtProv_Change()
ModificaDitta
End Sub
Private Sub TxtProv_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If Len(TxtProv.Text) = 2 Then
  KeyAscii = 0
 Else
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End If
End Sub
Private Sub TxtStato_Change()
ModificaDitta
End Sub
Private Sub TxtTel_Change()
ModificaDitta
End Sub
Private Sub TxtIndirizzo_Change()
ModificaDitta
End Sub
Private Sub TxtIndirizzo_KeyPress(KeyAscii As Integer)
If KeyAscii >= 32 Then
 If TxtIndirizzo.SelStart = 0 Or Mid(TxtIndirizzo.Text, TxtIndirizzo.SelStart + 1, 1) = "." Then
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End If
End Sub
Private Sub Form_Load()
Me.Move 600, 450
If rsClienti Is Nothing Then
 Set rsClienti = New ADODB.Recordset
 rsClienti.Open "SELECT * FROM Clienti WHERE Rimosso = False ORDER BY Ditta ASC", conn, adOpenDynamic, _
 adLockOptimistic
End If
Set rsLuoghiConsegna = New ADODB.Recordset
BtnPrec.Enabled = False: BtnPrimo.Enabled = False
BtnUltimo.Enabled = rsClienti.RecordCount > 1
TxtRecCorr.Text = "1"

ElencoControlli = Array("TxtDitta", "TxtPartIva", "TxtCodFisc", "TxtIndirizzo", _
"TxtTel", "TxtFax", "TxtEmail", "LstPag", "TxtLoc", "TxtComune", "TxtProv", "TxtStato", "TxtPEC")
DescControlli = Array("Ditta", "Partita Iva", "Codice Fiscale", "Indirizzo", "Tel", _
"Fax", "Email", "Pagamento", "Località", "Comune", "Provincia", "Stato", "PEC")
PropControlli = Array("ao", "o", "a", "ao", "n", "n", "a", "ao", "a", "a", "a", "a", "ao")

If Not rsClienti.EOF Then
 If rsClienti.RecordCount >= 1 Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
  BtnCanc.Enabled = True
 End If
 LblNumRecord.Caption = "di " & rsClienti.RecordCount
 rsClienti.MoveFirst
 Call VisualizzaRecord
Else
 CreaNuovoRecord
 LblNumRecord.Caption = "di 1"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If StatoCliente <> NonModificato Then
 Dim Scelta As VbMsgBoxResult
 Scelta = MsgBox("Salvare il record corrente ?", _
 vbYesNoCancel + vbQuestion, "Chiusura Archivio Clienti")
 If Scelta = vbYes Then
  If Not ConvalidaRecord() Then
   Cancel = 1
  End If
 ElseIf Scelta = vbCancel Then
  Cancel = 1
 End If
End If
If Cancel <> 1 Then
 Set rsClienti = Nothing
 Set Clienti = Nothing
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
Private Sub BtnUltimo_Click()
PosizioneRecord "ultimo"
End Sub
Private Sub BtnNuovo_Click()
PosizioneRecord "nuovo"
End Sub
Private Function ConvalidaRecord() As Boolean
If StatoCliente <> NonModificato Then
 For i = 0 To UBound(ElencoControlli)
  If Trim(Me(ElencoControlli(i))) = "" Then
   If InStr(1, PropControlli(i), "o") <> 0 Then
    MsgBox "Attenzione, " & DescControlli(i) & " è un campo obbligatorio !", vbExclamation, _
    "Fattura Pro"
    Exit Function
   End If
  ElseIf InStr(1, PropControlli(i), "a") <> 0 And IsNumeric(Trim(Me(ElencoControlli(i)))) Then
   MsgBox "Attenzione, il campo " & DescControlli(i) & " non può contenere valori valori nume" _
   & "rici !", vbExclamation, "Fattura Pro"
   Me(ElencoControlli(i)).Text = "": Exit Function
  ElseIf InStr(1, PropControlli(i), "n") <> 0 And Not IsNumeric(Trim(Me(ElencoControlli(i)))) Then
   MsgBox "Attenzione, il campo " & DescControlli(i) & " deve contenere valori numerici !", _
   vbExclamation, "Fattura Pro"
   Me(ElencoControlli(i)).Text = "": Exit Function
  End If
 Next i
 Dim CercaDuplicato As Boolean
 If StatoCliente = inserimento Then
  CercaDuplicato = True
 ElseIf TxtDitta <> rsClienti("Ditta") Then
  CercaDuplicato = True
 End If
 If TxtProv.Text = "" And TxtStato.Text = "" Then
  MsgBox "Provincia e Stato non possono essere entrambi vuoti !", vbExclamation, "Fattura Pro"
  TxtProv.SetFocus: Exit Function
 End If
 Dim rsDuplicato As New ADODB.Recordset, NumRec As Long
 If CercaDuplicato Then
  rsDuplicato.Open "SELECT * FROM Clienti WHERE Rimosso = true AND UCASE(Ditta) = '" & _
  UCase(Replace(TxtDitta.Text, "'", "''")) & "'", conn, adOpenDynamic
  If Not rsDuplicato.EOF Then
   MsgBox "La ditta che vuoi inserire è già presente in archivio !", vbExclamation, "Fattura Pro"
   TxtDitta.SetFocus: TxtDitta.SelStart = 0: TxtDitta.SelLength = Len(TxtDitta.Text): Exit Function
  End If
 End If
 If StatoCliente = inserimento Then
  rsClienti.AddNew
 End If
 With rsClienti
  .Fields("ditta") = EliminaSpazi(TxtDitta.Text)
  .Fields("partitaiva") = Trim(TxtPartIva.Text)
  .Fields("codfiscale") = Trim(TxtCodFisc.Text)
  .Fields("indirizzo") = Trim(TxtIndirizzo.Text)
  .Fields("cap") = Trim(TxtCap.Text)
  .Fields("tel") = Trim(TxtTel.Text)
  .Fields("email") = Trim(TxtEmail.Text)
  .Fields("fax") = Trim(TxtFax.Text)
  .Fields("loc") = Trim(TxtLoc.Text)
  .Fields("modpag") = LstPag.ListIndex
  .Fields("comune") = TxtComune.Text
  .Fields("prov") = Trim(TxtProv.Text)
  .Fields("stato") = Trim(TxtStato.Text)
  .Fields("coddest") = Trim(TxtCodDest.Text)
  .Fields("pec") = Trim(TxtPEC.Text)
  .Update
  If StatoCliente = inserimento And rsLuoghiConsegna.RecordCount <> 0 Then
   While Not rsLuoghiConsegna.EOF
    rsLuoghiConsegna("idditta") = .Fields("id")
    rsLuoghiConsegna.MoveNext
   Wend
  End If
  rsLuoghiConsegna.UpdateBatch
 End With
 If StatoCliente = inserimento Then
  CancellaUgualiRimossi
 End If
 EseguiBackup = True
 LblNumRecord.Caption = "di " & rsClienti.RecordCount
 StatoCliente = NonModificato
End If
ConvalidaRecord = True
End Function
Private Sub CancellaUgualiRimossi()
Dim rsRimossi As New ADODB.Recordset
rsRimossi.Open "SELECT * FROM Clienti WHERE Rimosso = True AND UCASE(Ditta) = '" & UCase(TxtDitta.Text) & "'" _
& " AND PartitaIva = '" & TxtPartIva.Text & "'", conn, adOpenDynamic, adLockOptimistic
If Not rsRimossi.EOF Then
 rsRimossi.MoveFirst
 conn.Execute "UPDATE FattureClienti, NoteConsegna, NoteCredito SET FattureClienti.IdDitta = " & rsClienti("Id") _
 & ", NoteConsegna.IdDitta = " & rsClienti("Id") & ", NoteCredito.IdDitta = " & rsClienti("Id") & _
 " WHERE FattureClienti.IdDitta = " & rsRimossi("Id") & " Or NoteConsegna.IdDitta=" & rsRimossi("Id") & " Or " _
 & "NoteCredito.IdDitta=" & rsRimossi("Id")
 rsRimossi.Delete
 rsClienti.Requery
 rsClienti.Move rsClienti.RecordCount - 1, adBookmarkFirst
End If
rsRimossi.Close
End Sub
Private Sub PosizioneRecord(PosRecord As String)
Dim valido: valido = ConvalidaRecord()
If valido Then
 Select Case PosRecord
 Case "primo"
  rsClienti.MoveFirst
 Case "ultimo"
  rsClienti.MoveLast
 Case "precedente"
  If CInt(TxtRecCorr.Text) <= rsClienti.RecordCount Then
   rsClienti.MovePrevious
  ElseIf rsClienti.AbsolutePosition <> rsClienti.RecordCount Then
   rsFatture.MoveLast
  End If
 Case "successivo"
  If rsClienti.AbsolutePosition = rsClienti.RecordCount Then
   CreaNuovoRecord
   Exit Sub
  End If
  rsClienti.MoveNext
 Case "nuovo"
  CreaNuovoRecord
  Exit Sub
 End Select
 BtnCanc.Enabled = True
 If rsClienti.AbsolutePosition <= rsClienti.RecordCount Then
  BtnSucc.Enabled = True: BtnNuovo.Enabled = True
 End If
 If rsClienti.AbsolutePosition <> rsClienti.RecordCount Then
  BtnUltimo.Enabled = True
 Else
  BtnUltimo.Enabled = False
 End If
 If rsClienti.AbsolutePosition <> 1 Then
  BtnPrimo.Enabled = True: BtnPrec.Enabled = True
 Else
  BtnPrimo.Enabled = False: BtnPrec.Enabled = False
 End If
 TxtRecCorr.Text = rsClienti.AbsolutePosition
 LblNumRecord.Caption = "di " & rsClienti.RecordCount
 VisualizzaRecord
End If
End Sub
Private Sub CreaNuovoRecord()
TxtRecCorr.Text = rsClienti.RecordCount + 1: LblNumRecord.Caption = "di " & rsClienti.RecordCount + 1
For Each NomeControllo In ElencoControlli
 If TypeOf Me(NomeControllo) Is VB.TextBox Then
  Me(NomeControllo).Text = ""
 End If
Next
TxtCap.Text = ""
LstPag.ListIndex = -1
StatoCliente = NonModificato
BtnSucc.Enabled = False
BtnCanc.Enabled = False: BtnNuovo.Enabled = False
If rsClienti.RecordCount >= 1 Then
 BtnPrimo.Enabled = True: BtnPrec.Enabled = True: BtnUltimo.Enabled = True
End If
If rsLuoghiConsegna.State <> adStateClosed Then
 rsLuoghiConsegna.Close
End If
rsLuoghiConsegna.Open "SELECT * FROM LuoghiConsegna WHERE IdDitta = 0", conn, adOpenDynamic, _
adLockBatchOptimistic
End Sub
Private Sub VisualizzaRecord()
TxtDitta.Text = rsClienti("ditta")
TxtPartIva.Text = rsClienti("partitaiva")
TxtCodFisc.Text = rsClienti("codfiscale")
TxtIndirizzo.Text = rsClienti("indirizzo")
TxtCap.Text = rsClienti("cap")
TxtTel.Text = rsClienti("tel")
TxtFax.Text = rsClienti("fax")
TxtEmail.Text = rsClienti("email")
TxtLoc.Text = rsClienti("loc")
TxtComune.Text = IIf(IsNull(rsClienti("comune")), "", rsClienti("comune"))
TxtProv.Text = rsClienti("prov")
TxtStato.Text = rsClienti("stato")
TxtCodDest.Text = IIf(IsNull(rsClienti("coddest")), "", rsClienti("coddest"))
TxtPEC.Text = IIf(IsNull(rsClienti("pec")), "", rsClienti("pec"))
LstPag.ListIndex = rsClienti("modpag")
If rsLuoghiConsegna.State <> adStateClosed Then
 rsLuoghiConsegna.Close
End If
rsLuoghiConsegna.Open "SELECT * FROM LuoghiConsegna WHERE IdDitta = " & rsClienti("id"), conn, adOpenDynamic, adLockBatchOptimistic
StatoCliente = NonModificato
End Sub
Public Sub ImpostaFiltroRicerca(rsRicerca As ADODB.Recordset)
Set rsClienti = rsRicerca
FiltroRicerca = True
End Sub
Public Sub CaricaFiltroRicerca()
PosizioneRecord "primo"
End Sub
Private Sub ModificaDitta()
If CInt(TxtRecCorr.Text) > rsClienti.RecordCount Then
 If Not FiltroRicerca Then
  StatoCliente = inserimento: BtnSucc.Enabled = True
  BtnCanc.Enabled = True: BtnNuovo.Enabled = True
 End If
Else
 StatoCliente = modifica
End If
End Sub
Public Function FormBloccata() As Boolean
FormBloccata = StatoCliente <> NonModificato
End Function
Private Sub TxtTel_KeyPress(KeyAscii As Integer)
If InStr("0123456789" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub
