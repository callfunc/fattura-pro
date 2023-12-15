VERSION 5.00
Begin VB.Form ElencoArticoli 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elenco Articoli"
   ClientHeight    =   4725
   ClientLeft      =   645
   ClientTop       =   690
   ClientWidth     =   8430
   Icon            =   "ElencoArticoli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8430
   Begin VB.ListBox Prezzi 
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
      ItemData        =   "ElencoArticoli.frx":4072
      Left            =   5775
      List            =   "ElencoArticoli.frx":4074
      TabIndex        =   1
      Top             =   495
      Width           =   2340
   End
   Begin VB.ListBox Elenco 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      ItemData        =   "ElencoArticoli.frx":4076
      Left            =   240
      List            =   "ElencoArticoli.frx":4078
      TabIndex        =   0
      Top             =   495
      Width           =   5190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prezzi"
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
      Left            =   5775
      TabIndex        =   3
      Top             =   225
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Articoli"
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
      TabIndex        =   2
      Top             =   225
      Width           =   585
   End
End
Attribute VB_Name = "ElencoArticoli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormChiamante As Form
Public rsArticoli As ADODB.Recordset
Dim IndCorr As Integer
Dim VecchioTot#, IvaCorr#, VecchiaIva#, Diff#
Private Sub Elenco_Click()
If Elenco.ListIndex <> -1 And IndCorr <> Elenco.ListIndex Then
 Dim Prezzo As Double: Prezzi.Clear
 rsArticoli.Move Elenco.ListIndex, adBookmarkFirst
 For i = 1 To 3
  Prezzo = rsArticoli("prezzo" & i)
  If Prezzo <> 0 Then Prezzi.AddItem FormatNumber(Prezzo, 2)
 Next i
 IndCorr = Elenco.ListIndex
End If
End Sub
Private Sub Elenco_DblClick()
If Elenco.Text <> "" Then Call EseguiScelta
End Sub
Private Sub Elenco_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn And Prezzi.ListIndex <> -1 Then Call EseguiScelta
End Sub
Public Property Set Articoli(ListaArticoli As ADODB.Recordset)
Set rsArticoli = ListaArticoli
rsArticoli.MoveFirst
While Not rsArticoli.EOF
 Elenco.AddItem rsArticoli("descr")
 rsArticoli.MoveNext
Wend
End Property
Private Sub EseguiScelta()
rsArticoli.Move Elenco.ListIndex, adBookmarkFirst
Call InserisciDatiArticolo(FormChiamante)
Unload Me
End Sub
Private Sub Form_Load()
Move 2250, 1800: IndCorr = -1
End Sub
Private Sub InserisciDatiArticolo(ByRef f As Form)
Dim Iva$, EsenteIva As Boolean
If f.Name = "FattureClienti" Then
 EsenteIva = (f.ChkEsenteIva.Value = 1)
End If
Iva = rsArticoli("AliqIva")
Dim ColPrezzo%, ColIva%, ColTot%, ColUM%
If f.Name = "FattureFornitori" Then
 ColQnt = 2
 ColPrezzo = 3
 ColTot = 5
 ColIva = 4
 ColUM = 1
Else
 ColQnt = 3
 ColPrezzo = 4
 ColTot = 7
 ColIva = 6
 ColUM = 2
End If
With f.ElencoVoci
 .Text = Elenco.Text
 .TextMatrix(.Row, ColPrezzo) = Prezzi.Text
 .TextMatrix(.Row, ColUM) = rsArticoli("um")
 If .TextMatrix(.Row, ColTot) <> "" And Prezzi.Text <> "" Then
  VecchioTot = CDbl(.TextMatrix(.Row, ColTot))
  Dim Qnt#: Qnt = 1
  If .TextMatrix(.Row, ColQnt) <> "" Then Qnt = CDbl(.TextMatrix(.Row, ColQnt))
  .TextMatrix(.Row, ColTot) = FormatNumber(Qnt * CDbl(.TextMatrix(.Row, ColPrezzo)), 2)
  Diff = CDbl(.TextMatrix(.Row, ColTot)) - VecchioTot
  If Not EsenteIva Then
   If .TextMatrix(.Row, ColIva) <> "" Then
    VecchiaIva = VecchioTot * CDbl(.TextMatrix(.Row, ColIva)) / 100
   End If
   IvaCorr = CDbl(.TextMatrix(.Row, ColTot)) * CDbl(Iva) / 100
   f.TotImp.Text = FormatNumber(CDbl(f.TotImp.Text) + Diff, 2)
   .TextMatrix(.Row, ColIva) = Iva
   Diff = CDbl(FormatNumber(IvaCorr, 2)) - CDbl(FormatNumber(VecchiaIva, 2))
   f.TotIva.Text = FormatNumber(CDbl(f.TotIva.Text) + Diff, 2)
   f.TotDoc.Text = FormatNumber(CDbl(f.TotImp.Text) + CDbl(f.TotIva.Text), 2)
  Else
   f.TotDoc.Text = FormatNumber(CDbl(f.TotDoc.Text) + Diff, 2)
  End If
  If f.Name = "FattureFornitori" Then
   f.TotNetto.Text = f.TotDoc.Text
  End If
 ElseIf Not EsenteIva Then
  If f.Name = "FattureFornitori" Then
   .TextMatrix(.Row, 5) = Iva
  Else
   .TextMatrix(.Row, 6) = Iva
  End If
 End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
rsArticoli.Close
Set ElencoArticoli = Nothing
End Sub
